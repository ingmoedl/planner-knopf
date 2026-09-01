/* Aufgabe in Planner – Outlook-Add-in (ing Burghausen GmbH)
 * Erstellt aus der geöffneten Mail direkt eine Planner-Aufgabe via Microsoft Graph.
 * Auth: Nested App Authentication (MSAL.js), keine Server-Komponente. */

"use strict";

const CONFIG = {
  clientId: "92b69fe3-9c55-4262-98d2-4d5642aaeebe",
  tenantId: "1571141a-75a9-43a3-ad47-8d613cfbb3e6",
  scopes: ["User.Read", "User.ReadBasic.All", "Tasks.ReadWrite", "Mail.ReadWrite"],
  graph: "https://graph.microsoft.com/v1.0",
  plannerWeb: "https://planner.cloud.microsoft/webui/plan/",
  folderParent: "@Aufgabe",
  folderDone: "Verarbeitet",
  planCacheKey: "pk_plans_v3",
  peopleCacheKey: "pk_people_v3",
  cacheTtlMs: 6 * 60 * 60 * 1000,
  folderCacheKey: "pk_folder_v3",
};

let pca = null;
let plans = [];            // [{id, title}]
let people = [];           // [{id, name, mail}]
let plansReady = null;
let selectedPlan = null;
let selectedPerson = null; // Zuweisung
const el = (id) => document.getElementById(id);

/* ---------- Boot ---------- */

Office.onReady(async () => {
  try {
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
    fillFromItem();
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      if (!Office.context.mailbox.item) { clearForm(); return; }
      fillFromItem();
    });

    // Nur STILLE Anmeldung beim Start (Popups brauchen einen Klick).
    const token = await silentToken(8000);
    if (token) {
      afterAuth();
    } else {
      el("login").style.display = "block";
      el("plan").placeholder = "Bitte zuerst anmelden (Knopf oben)";
    }
  } catch (e) {
    showStatus("Startfehler: " + msg(e), "err");
  }
});

/* ---------- Auth ---------- */

async function silentToken(timeoutMs) {
  const req = { scopes: CONFIG.scopes, account: pca.getActiveAccount() || pca.getAllAccounts()[0] };
  const attempt = pca.acquireTokenSilent(req).then((r) => { pca.setActiveAccount(r.account); return r.accessToken; });
  const timeout = new Promise((res) => setTimeout(() => res(null), timeoutMs || 8000));
  try {
    return await Promise.race([attempt, timeout]);
  } catch (e) {
    return null;
  }
}

async function interactiveToken() {
  const r = await pca.acquireTokenPopup({ scopes: CONFIG.scopes });
  pca.setActiveAccount(r.account);
  return r.accessToken;
}

/* Token für Hintergrund-Aufrufe: still; wenn das scheitert, Login-Knopf zeigen. */
async function getToken() {
  const t = await silentToken(15000);
  if (t) return t;
  el("login").style.display = "block";
  throw new Error("Anmeldung erforderlich – bitte oben auf „Bei Microsoft anmelden" klicken.");
}

function myAccount() {
  return pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
}

function afterAuth() {
  el("login").style.display = "none";
  plansReady = loadPlans();
  loadPeople();
  const acct = myAccount();
  if (acct && !selectedPerson) {
    selectedPerson = { id: acct.idTokenClaims && acct.idTokenClaims.oid, name: acct.name || "Ich" };
    el("assign").value = selectedPerson.name;
  }
}

/* ---------- Graph ---------- */

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

/* ---------- Pläne + Kollegen laden (Cache + Hintergrund-Refresh) ---------- */

async function loadPlans() {
  try {
    const cached = JSON.parse(localStorage.getItem(CONFIG.planCacheKey) || "null");
    if (cached && Array.isArray(cached.plans) && cached.plans.length) {
      plans = cached.plans;
      el("plan").placeholder = "Projektnummer oder Name tippen …";
      if (Office.context.mailbox.item) detectProject();
      if (Date.now() - cached.ts < CONFIG.cacheTtlMs) { refreshPlans().catch(() => {}); return; }
    }
    await refreshPlans();
    el("plan").placeholder = "Projektnummer oder Name tippen …";
  } catch (e) {
    if (!plans.length) {
      el("plan").placeholder = "Laden fehlgeschlagen";
      showStatus("Pläne konnten nicht geladen werden: " + msg(e), "err");
    }
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

async function loadPeople() {
  try {
    const cached = JSON.parse(localStorage.getItem(CONFIG.peopleCacheKey) || "null");
    if (cached && Array.isArray(cached.people) && cached.people.length) {
      people = cached.people;
      if (Date.now() - cached.ts < CONFIG.cacheTtlMs) return;
    }
    const res = await graph("/users?$select=id,displayName,mail&$top=999");
    if (!res.ok) return;
    const j = await res.json();
    people = (j.value || [])
      .filter((u) => u.displayName && u.mail)
      .map((u) => ({ id: u.id, name: u.displayName, mail: u.mail }))
      .sort((a, b) => a.name.localeCompare(b.name, "de"));
    localStorage.setItem(CONFIG.peopleCacheKey, JSON.stringify({ ts: Date.now(), people }));
  } catch (e) { /* Zuweisung bleibt dann auf "mir" */ }
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
    if (plans.length) detectProject();
    else if (plansReady) plansReady.then(() => detectProject());
  });
}

function cleanSubject(s) {
  return s.replace(/^\s*((AW|RE|WG|FW|FWD|SV|Antwort)\s*:\s*)+/i, "").trim();
}

function detectProject() {
  const item = Office.context.mailbox.item;
  if (!item || !plans.length || selectedPlan) return;
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
    if (hits.length === 1) { choosePlan(hits[0], "erkannt: " + cand); return; }
    if (hits.length > 1) {
      el("plan").value = cand;
      el("detected").innerHTML = "Nummer <b>" + esc(cand) + "</b> gefunden – bitte Plan wählen:";
      renderList("plan", cand);
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

function choosePerson(p) {
  selectedPerson = p;
  el("assign").value = p.name;
  el("assignlist").style.display = "none";
}

/* ---------- Durchsuchbare Listen (Plan + Zuweisung) ---------- */

const combos = {
  plan:   { listId: "planlist",   items: () => plans,  label: (p) => p.title, pick: (p) => choosePlan(p), clear: () => { selectedPlan = null; },   idx: -1 },
  assign: { listId: "assignlist", items: () => people, label: (p) => p.name,  pick: (p) => choosePerson(p), clear: () => { selectedPerson = null; }, idx: -1 },
};

function renderList(key, filter) {
  const c = combos[key];
  const list = el(c.listId);
  c.idx = -1;
  const f = (filter || "").trim().toUpperCase();
  const all = c.items();
  const hits = (f ? all.filter((p) => c.label(p).toUpperCase().includes(f)) : all).slice(0, 60);
  list.innerHTML = "";
  if (!hits.length) {
    list.innerHTML = '<div class="none">Nichts gefunden</div>';
  } else {
    hits.forEach((p) => {
      const d = document.createElement("div");
      d.textContent = c.label(p);
      d.dataset.key = key;
      d.addEventListener("mousedown", (ev) => { ev.preventDefault(); c.pick(p); });
      d._item = p;
      list.appendChild(d);
    });
  }
  list.style.display = "block";
}

function wireCombo(key) {
  const c = combos[key];
  const input = el(key);
  input.addEventListener("input", () => { c.clear(); renderList(key, input.value); });
  input.addEventListener("focus", () => renderList(key, ""));
  input.addEventListener("keydown", (ev) => {
    const list = el(c.listId);
    const items = [...list.children].filter((d) => d._item);
    if (ev.key === "ArrowDown" || ev.key === "ArrowUp") {
      ev.preventDefault();
      if (!items.length) return;
      c.idx = ev.key === "ArrowDown" ? Math.min(c.idx + 1, items.length - 1) : Math.max(c.idx - 1, 0);
      items.forEach((d, i) => d.classList.toggle("active", i === c.idx));
      items[c.idx].scrollIntoView({ block: "nearest" });
    } else if (ev.key === "Enter") {
      ev.preventDefault();
      const pick = c.idx >= 0 && items[c.idx] ? items[c.idx] : items[0];
      if (pick) c.pick(pick._item);
    } else if (ev.key === "Escape") {
      list.style.display = "none";
    }
  });
}

function wireUi() {
  wireCombo("plan");
  wireCombo("assign");
  document.addEventListener("click", (ev) => {
    if (!ev.target.closest(".combobox")) {
      el("planlist").style.display = "none";
      el("assignlist").style.display = "none";
    }
  });
  el("create").addEventListener("click", createTask);
  el("loginBtn").addEventListener("click", async () => {
    try {
      showStatus("Anmeldefenster geöffnet …", "");
      await interactiveToken();
      showStatus("", "");
      afterAuth();
    } catch (e) {
      showStatus(friendlyAuthError(e), "err");
    }
  });
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
  if (!selectedPlan) { showStatus("Bitte zuerst einen Plan auswählen (Feld „Projekt / Plan").", "err"); el("plan").focus(); return; }
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
    // Falls noch kein Token da ist: hier ist ein Klick-Kontext, Popup erlaubt.
    try { await getToken(); } catch (_) { await interactiveToken(); afterAuth(); }

    // 1) Task anlegen
    const body = { planId, title };
    const due = el("due").value;
    if (due) body.dueDateTime = due + "T10:00:00Z";
    const acct = myAccount();
    const assigneeId = (selectedPerson && selectedPerson.id) || (acct && acct.idTokenClaims && acct.idTokenClaims.oid);
    if (assigneeId) body.assignments = { [assigneeId]: { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" } };
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

    // 3) Beschreibung + Mail-Link an die Aufgabe hängen
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

/* Planner-Referenz-Keys: % zuerst, dann . : @ # encodieren */
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
    return "Zustimmung erforderlich: Bitte im Anmeldefenster zustimmen – falls „Administratorgenehmigung erforderlich" erscheint, muss der M365-Admin der App „Planner-Knopf" einmalig zustimmen. (" + esc(m) + ")";
  }
  if (/popup_window_error|popup window/i.test(m)) {
    return "Das Anmeldefenster wurde blockiert. Bitte Popups für dieses Add-in erlauben oder Outlook im Web verwenden.";
  }
  return "Fehler: " + esc(m);
}

function msg(e) { return (e && (e.message || e.errorMessage)) || String(e); }
function esc(s) { return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }
