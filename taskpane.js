/* Aufgabe in Planner – Outlook-Add-in (ing Burghausen GmbH)
 * Erstellt aus der geöffneten Mail direkt eine Planner-Aufgabe via Microsoft Graph.
 * Auth: Nested App Authentication (MSAL.js), keine Server-Komponente. */

"use strict";

const CONFIG = {
  clientId: "%%CLIENT_ID%%",
  tenantId: "1571141a-75a9-43a3-ad47-8d613cfbb3e6",
  scopes: ["User.Read", "Tasks.ReadWrite", "Mail.ReadWrite"],
  graph: "https://graph.microsoft.com/v1.0",
  plannerWeb: "https://planner.cloud.microsoft/webui/plan/",
  folderParent: "@Aufgabe",
  folderDone: "Verarbeitet",
  planCacheKey: "pk_plans_v2",
  planCacheTtlMs: 6 * 60 * 60 * 1000,
  folderCacheKey: "pk_folder_v2",
};

let pca = null;
let plans = [];          // [{id, title}]
let plansReady = null;   // Promise
let selectedPlan = null;
let activeIndex = -1;

const el = (id) => document.getElementById(id);

/* ---------- Boot ---------- */

Office.onReady(async () => {
  try {
    // NAA im neuen Outlook/OWA/aktuellen klassischen Outlook; ältere klassische
    // Clients fallen automatisch auf den normalen MSAL-Popup-Flow zurück.
    const naa = Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
    const msalConfig = {
      auth: {
        clientId: CONFIG.clientId,
        authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
        redirectUri: window.location.origin + window.location.pathname,
      },
    };
    pca = naa
      ? await msal.createNestablePublicClientApplication(msalConfig)
      : await msal.PublicClientApplication.createPublicClientApplication(msalConfig);

    wireUi();
    plansReady = loadPlans();
    fillFromItem();

    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      if (!Office.context.mailbox.item) { clearForm(); return; }
      fillFromItem();
    });
  } catch (e) {
    showStatus("Startfehler: " + msg(e), "err");
  }
});

/* ---------- Auth + Graph ---------- */

async function getToken() {
  const req = { scopes: CONFIG.scopes };
  try {
    return (await pca.acquireTokenSilent(req)).accessToken;
  } catch (e) {
    const r = await pca.acquireTokenPopup(req);
    return r.accessToken;
  }
}

function myOid() {
  const acct = pca.getActiveAccount() || (pca.getAllAccounts()[0] || null);
  return acct && acct.idTokenClaims ? acct.idTokenClaims.oid : null;
}

async function graph(path, opts = {}, retry = true) {
  const token = await getToken();
  const res = await fetch(CONFIG.graph + path, {
    ...opts,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      ...(opts.headers || {}),
    },
  });
  if (res.status === 429 && retry) {
    const wait = parseInt(res.headers.get("Retry-After") || "2", 10) * 1000;
    await new Promise((r) => setTimeout(r, wait));
    return graph(path, opts, false);
  }
  return res;
}

/* ---------- Pläne laden (Cache + Hintergrund-Refresh) ---------- */

async function loadPlans() {
  try {
    const cached = JSON.parse(localStorage.getItem(CONFIG.planCacheKey) || "null");
    if (cached && Array.isArray(cached.plans) && cached.plans.length) {
      plans = cached.plans;
      el("plan").placeholder = "Projektnummer oder Name tippen …";
      if (Date.now() - cached.ts < CONFIG.planCacheTtlMs) { refreshPlans(); return; }
    }
    await refreshPlans();
    el("plan").placeholder = "Projektnummer oder Name tippen …";
  } catch (e) {
    if (!plans.length) showStatus("Pläne konnten nicht geladen werden: " + msg(e), "err");
  }
}

async function refreshPlans() {
  let url = "/me/planner/plans";
  const acc = [];
  while (url) {
    const res = await graph(url);
    if (!res.ok) throw new Error("Graph " + res.status + " beim Laden der Pläne");
    const j = await res.json();
    (j.value || []).forEach((p) => acc.push({ id: p.id, title: p.title || "" }));
    url = j["@odata.nextLink"] ? j["@odata.nextLink"].replace(CONFIG.graph, "") : null;
  }
  acc.sort((a, b) => b.title.localeCompare(a.title, "de"));
  plans = acc;
  localStorage.setItem(CONFIG.planCacheKey, JSON.stringify({ ts: Date.now(), plans }));
  if (Office.context.mailbox.item) detectProject();
}

/* ---------- Mail lesen + Projekt erkennen ---------- */

function fillFromItem() {
  const item = Office.context.mailbox.item;
  if (!item) return;
  showStatus("", "");
  el("create").disabled = false;
  el("due").value = "";
  el("title").value = cleanSubject(item.subject || "");
  selectedPlan = null;
  el("plan").value = "";
  el("detected").textContent = "";
  item.body.getAsync(Office.CoercionType.Text, (res) => {
    item._bodyText = res.status === Office.AsyncResultStatus.Succeeded ? (res.value || "").slice(0, 4000) : "";
    if (plansReady) plansReady.then(() => detectProject());
  });
}

function cleanSubject(s) {
  return s.replace(/^\s*((AW|RE|WG|FW|FWD|SV|Antwort)\s*:\s*)+/i, "").trim();
}

function detectProject() {
  const item = Office.context.mailbox.item;
  if (!item || !plans.length) return;
  const text = (item.subject || "") + "\n" + (item._bodyText || "");
  const seen = new Set();
  const withSuffix = [], bare = [];
  const re = /\b(\d{5})(-[A-Za-z0-9]{1,6})?\b/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const full = m[1] + (m[2] || "");
    if (seen.has(full)) continue;
    seen.add(full);
    (m[2] ? withSuffix : bare).push(full);
  }
  const candidates = [...withSuffix, ...bare];
  for (const cand of candidates) {
    const hits = plans.filter((p) => p.title.toUpperCase().startsWith(cand.toUpperCase()));
    if (hits.length === 1) {
      choosePlan(hits[0], "erkannt: " + cand);
      return;
    }
    if (hits.length > 1) {
      el("plan").value = cand;
      el("detected").innerHTML = "Nummer <b>" + esc(cand) + "</b> gefunden – bitte Plan wählen:";
      renderPlanList(cand);
      return;
    }
  }
  if (candidates.length) {
    el("detected").textContent = "Nummer " + candidates[0] + " gefunden, aber kein passender Plan – bitte manuell wählen.";
  } else {
    el("detected").textContent = "Keine Projektnummer in der Mail gefunden – bitte Plan wählen.";
  }
}

function choosePlan(p, note) {
  selectedPlan = p;
  el("plan").value = p.title;
  el("planlist").style.display = "none";
  el("detected").innerHTML = note ? "✓ <b>" + esc(p.title) + "</b> (" + esc(note) + ")" : "";
}

/* ---------- Plan-Auswahl (durchsuchbare Liste) ---------- */

function wireUi() {
  const input = el("plan");
  input.addEventListener("input", () => { selectedPlan = null; renderPlanList(input.value); });
  input.addEventListener("focus", () => renderPlanList(input.value));
  input.addEventListener("keydown", (ev) => {
    const list = el("planlist");
    const items = [...list.querySelectorAll("div[data-id]")];
    if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
      ev.preventDefault();
      if (!items.length) return;
      activeIndex = ev.key === "ArrowDown" ? Math.min(activeIndex + 1, items.length - 1) : Math.max(activeIndex - 1, 0);
      items.forEach((d, i) => d.classList.toggle("active", i === activeIndex));
      items[activeIndex].scrollIntoView({ block: "nearest" });
    } else if (ev.key === "Enter") {
      ev.preventDefault();
      const pick = activeIndex >= 0 && items[activeIndex] ? items[activeIndex] : items[0];
      if (pick) choosePlan(plans.find((p) => p.id === pick.dataset.id));
    } else if (ev.key === "Escape") {
      list.style.display = "none";
    }
  });
  document.addEventListener("click", (ev) => {
    if (!ev.target.closest(".planbox")) el("planlist").style.display = "none";
  });
  el("create").addEventListener("click", createTask);
}

function renderPlanList(filter) {
  const list = el("planlist");
  activeIndex = -1;
  const f = (filter || "").trim().toUpperCase();
  const hits = (f ? plans.filter((p) => p.title.toUpperCase().includes(f)) : plans).slice(0, 60);
  list.innerHTML = "";
  if (!hits.length) {
    list.innerHTML = '<div class="none">Kein Plan gefunden</div>';
  } else {
    hits.forEach((p) => {
      const d = document.createElement("div");
      d.textContent = p.title;
      d.dataset.id = p.id;
      d.addEventListener("mousedown", (ev) => { ev.preventDefault(); choosePlan(p); });
      list.appendChild(d);
    });
  }
  list.style.display = "block";
}

/* ---------- Aufgabe erstellen ---------- */

async function createTask() {
  const item = Office.context.mailbox.item;
  if (!item) { showStatus("Keine Mail ausgewählt.", "err"); return; }
  if (!selectedPlan) {
    const f = el("plan").value.trim().toUpperCase();
    const exact = plans.filter((p) => p.title.toUpperCase() === f);
    if (exact.length === 1) selectedPlan = exact[0];
  }
  if (!selectedPlan) { showStatus("Bitte zuerst einen Plan auswählen.", "err"); el("plan").focus(); return; }
  const title = el("title").value.trim();
  if (!title) { showStatus("Bitte einen Titel eingeben.", "err"); el("title").focus(); return; }

  const btn = el("create");
  btn.disabled = true;
  showStatus("Aufgabe wird erstellt …", "");

  const subject = item.subject || "";
  const fromAddr = item.from ? (item.from.displayName + " <" + item.from.emailAddress + ">") : "";
  const received = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" }) : "";
  const planId = selectedPlan.id;

  try {
    // 1) Task anlegen (mir selbst zugewiesen)
    const body = { planId, title };
    const due = el("due").value;
    if (due) body.dueDateTime = due + "T10:00:00Z";
    const oid = myOid();
    if (oid) body.assignments = { [oid]: { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" } };
    const createRes = await graph("/planner/tasks", { method: "POST", body: JSON.stringify(body) });
    if (!createRes.ok) throw new Error("Aufgabe anlegen fehlgeschlagen (Graph " + createRes.status + ")");
    const task = await createRes.json();

    // 2) Mail ggf. verschieben; webLink der (ggf. neuen) Mail holen
    let webLink = "";
    try {
      const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      if (el("move").checked) {
        const folderId = await getDoneFolderId();
        if (folderId) {
          const mv = await graph("/me/messages/" + encodeURIComponent(restId) + "/move", {
            method: "POST", body: JSON.stringify({ destinationId: folderId }),
          });
          if (mv.ok) { const moved = await mv.json(); webLink = moved.webLink || ""; }
        }
      }
      if (!webLink) {
        const g = await graph("/me/messages/" + encodeURIComponent(restId) + "?$select=webLink");
        if (g.ok) webLink = (await g.json()).webLink || "";
      }
    } catch (e) { /* Mail-Teil ist optional – Task existiert bereits */ }

    // 3) Beschreibung + Mail-Link an die Aufgabe hängen (ETag-Pflicht, 1x Retry)
    const description =
      "Von: " + fromAddr + "\nEmpfangen: " + received + "\nBetreff: " + subject +
      (webLink ? "\nOriginal-Mail: " + webLink : "") +
      "\n\n— erstellt mit dem Planner-Knopf aus Outlook";
    await patchDetails(task.id, description, webLink, 2);

    const link = CONFIG.plannerWeb + planId + "/view/board/task/" + task.id;
    showStatus('✓ Aufgabe angelegt in „' + esc(selectedPlan.title) + '".<br><a href="' + link + '" target="_blank" rel="noopener">In Planner öffnen</a>', "ok");
  } catch (e) {
    showStatus(friendlyAuthError(e), "err");
  } finally {
    btn.disabled = false;
  }
}

async function patchDetails(taskId, description, webLink, tries) {
  for (let i = 0; i < tries; i++) {
    const det = await graph("/planner/tasks/" + taskId + "/details");
    if (!det.ok) return;
    const etag = (await det.json())["@odata.etag"];
    const patch = { description };
    if (webLink) {
      patch.references = {
        [encodeRefKey(webLink)]: {
          "@odata.type": "#microsoft.graph.plannerExternalReference",
          alias: "Original-E-Mail",
          type: "Other",
          previewPriority: " !",
        },
      };
    }
    const res = await graph("/planner/tasks/" + taskId + "/details", {
      method: "PATCH",
      headers: { "If-Match": etag },
      body: JSON.stringify(patch),
    });
    if (res.ok || (res.status !== 409 && res.status !== 412)) return;
  }
}

/* Planner-Referenz-Keys: . : % @ # müssen encodiert sein (% zuerst!) */
function encodeRefKey(url) {
  return url.replace(/%/g, "%25").replace(/\./g, "%2E").replace(/:/g, "%3A").replace(/@/g, "%40").replace(/#/g, "%23");
}

async function getDoneFolderId() {
  try {
    const cached = JSON.parse(localStorage.getItem(CONFIG.folderCacheKey) || "null");
    if (cached && cached.id) return cached.id;
    const p = await graph("/me/mailFolders?$filter=displayName eq '" + CONFIG.folderParent.replace("'", "''") + "'");
    if (!p.ok) return null;
    const parent = (await p.json()).value[0];
    if (!parent) return null;
    const c = await graph("/me/mailFolders/" + parent.id + "/childFolders?$filter=displayName eq '" + CONFIG.folderDone + "'");
    if (!c.ok) return null;
    const done = (await c.json()).value[0];
    if (!done) return null;
    localStorage.setItem(CONFIG.folderCacheKey, JSON.stringify({ id: done.id }));
    return done.id;
  } catch (e) { return null; }
}

/* ---------- Helfer ---------- */

function clearForm() {
  el("title").value = ""; el("plan").value = ""; el("due").value = "";
  selectedPlan = null; el("detected").textContent = "Keine Mail ausgewählt.";
  el("create").disabled = true;
}

function showStatus(html, cls) {
  const s = el("status");
  s.className = cls || "";
  s.innerHTML = html;
}

function friendlyAuthError(e) {
  const m = msg(e);
  if (/consent|AADSTS65001|admin approval|AADSTS90094/i.test(m)) {
    return "Zustimmung erforderlich: Bitte im Anmeldefenster zustimmen – falls „Administratorgenehmigung erforderlich" erscheint, muss der M365-Admin der App einmalig zustimmen. (" + esc(m) + ")";
  }
  return "Fehler: " + esc(m);
}

function msg(e) { return (e && (e.message || e.errorMessage)) || String(e); }
function esc(s) { return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }
