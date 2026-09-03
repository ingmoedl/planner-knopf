/* Aufgabe in Planner – Outlook-Add-in (ing Burghausen GmbH)
 * Erstellt aus der geöffneten Mail direkt eine Planner-Aufgabe via Microsoft Graph.
 * Auth: Nested App Authentication (MSAL.js), keine Server-Komponente.
 * v1.8: alle Pläne über Team-Mitgliedschaften, Bucket-Auswahl, nur interne Personen.
 * v1.9: Diagnosezeile (Anzahl Pläne, gefundene Teams), Versionsanzeige.
 * v2.0: memberOf liefert mit User.Read nur Gruppen-IDs (kein Name, kein groupTypes) → kein Typ-Filter mehr,
 *       alle Gruppen nach Plänen abfragen; Team-Label aus den Projektnummern der Pläne ableiten. */

"use strict";

const CONFIG = {
  version: "2.0",
  clientId: "92b69fe3-9c55-4262-98d2-4d5642aaeebe",
  tenantId: "1571141a-75a9-43a3-ad47-8d613cfbb3e6",
  scopes: ["User.Read", "User.ReadBasic.All", "Tasks.ReadWrite", "Mail.ReadWrite"],
  graph: "https://graph.microsoft.com/v1.0",
  plannerWeb: "https://planner.cloud.microsoft/webui/plan/",
  planCacheKey: "pk_plans_v5",
  peopleCacheKey: "pk_people_v4",
  bucketCachePrefix: "pk_buckets_v1_",
  lastBucketPrefix: "pk_lastbucket_",
  cacheTtlMs: 6 * 60 * 60 * 1000,
  internalDomain: "ing-burghausen.de",  // "Zuweisen an": nur Personen mit dieser Mail-Domain
  personNamePattern: /^[^,]+,\s*\S+/,    // Personen heißen "Nachname, Vorname" – Räume/Funktionspostfächer nicht
  maxListRows: 500,
};

let pca = null;
let plans = [];            // [{id, title}]
let people = [];           // [{id, name, mail}]
let buckets = [];          // [{id, name, orderHint}] des gewählten Plans
let myGroups = null;       // [{id, label, plans}] Gruppen/Teams des Nutzers; null = memberOf fehlgeschlagen
let groupsError = null;    // Fehlertext, falls memberOf fehlgeschlagen ist
let plansReady = null;
let bucketsReady = null;
let selectedPlan = null;
let selectedPerson = null; // Zuweisung
let standalone = false;    // true = außerhalb von Outlook (Browser-Test)
const el = (id) => document.getElementById(id);
const mailItem = () => (Office.context && Office.context.mailbox) ? Office.context.mailbox.item : null;

/* ---------- Boot ---------- */

Office.onReady(async () => {
  try {
    const inOutlook = !!(Office.context && Office.context.mailbox);
    const naa = !!(inOutlook && Office.context.requirements && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1"));
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
    standalone = !inOutlook;

    // Browser-Test: Rückkehr von der Anmeldeseite (Redirect-Flow) auswerten
    if (standalone) {
      try {
        const rr = await pca.handleRedirectPromise();
        if (rr && rr.account) pca.setActiveAccount(rr.account);
      } catch (e) { console.warn("[PK] handleRedirectPromise:", msg(e)); }
    }

    // Alte Cache-Stände (v3) aufräumen
    ["pk_plans_v3", "pk_plans_v4", "pk_people_v3"].forEach((k) => { try { localStorage.removeItem(k); } catch (_) {} });

    wireUi();
    el("version").textContent = "Planner-Knopf v" + CONFIG.version;
    if (inOutlook) {
      fillFromItem();
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
        if (!mailItem()) { clearForm(); return; }
        fillFromItem();
      });
    } else {
      el("title").value = "Standalone-Test";
      el("create").disabled = false;
    }

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
  if (standalone) {
    // Browser-Test: Popups sind dort oft blockiert → ganze Seite zur Anmeldung weiterleiten
    await pca.acquireTokenRedirect({ scopes: CONFIG.scopes });
    return new Promise(() => {}); // Seite wird verlassen
  }
  const r = await pca.acquireTokenPopup({ scopes: CONFIG.scopes });
  pca.setActiveAccount(r.account);
  return r.accessToken;
}

/* Token für Hintergrund-Aufrufe: still; wenn das scheitert, Login-Knopf zeigen. */
async function getToken() {
  const t = await silentToken(15000);
  if (t) return t;
  el("login").style.display = "block";
  throw new Error("Anmeldung erforderlich – bitte oben auf 'Bei Microsoft anmelden' klicken.");
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

async function graph(path, opts = {}, retry = true, progress = null) {
  console.log("[PK] graph:", path);
  if (progress) progress("Token");
  const token = await getToken();
  console.log("[PK] token ok, fetch:", path);
  if (progress) progress("Abruf");
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

/* Alle Seiten einer Graph-Liste einsammeln (folgt @odata.nextLink). */
async function graphAll(path, opts = {}, progress = null) {
  const out = [];
  let url = path;
  while (url) {
    const res = await graph(url, opts, true, progress);
    if (!res.ok) throw new Error("Graph " + res.status + " bei " + path.split("?")[0]);
    const j = await res.json();
    out.push(...(j.value || []));
    url = j["@odata.nextLink"] ? j["@odata.nextLink"].replace(CONFIG.graph, "") : null;
  }
  return out;
}

/* ---------- Pläne + Kollegen laden (Cache + Hintergrund-Refresh) ---------- */

async function loadPlans() {
  try {
    const cached = JSON.parse(localStorage.getItem(CONFIG.planCacheKey) || "null");
    if (cached && Array.isArray(cached.plans) && cached.plans.length) {
      plans = cached.plans;
      myGroups = Array.isArray(cached.groups) ? cached.groups : null;
      groupsError = cached.groupsError || null;
      renderPlanInfo();
      el("plan").placeholder = "Projektnummer oder Name tippen …";
      if (mailItem()) detectProject();
      if (Date.now() - cached.ts < CONFIG.cacheTtlMs) { refreshPlans().catch(() => {}); return; }
    }
    await refreshPlans();
    el("plan").placeholder = "Projektnummer oder Name tippen …";
  } catch (e) {
    console.error("[PK] loadPlans:", e);
    if (!plans.length) {
      el("plan").placeholder = "Laden fehlgeschlagen";
      showStatus("Pläne konnten nicht geladen werden: " + msg(e), "err");
    }
  }
}

/* Vollständige Planliste:
 *  1) /me/planner/plans            – was Planner dem Nutzer direkt zuordnet (oft unvollständig)
 *  2) /me/memberOf (Gruppen)       – alle Gruppen/Teams, in denen der Nutzer Mitglied ist (User.Read reicht,
 *     liefert dann aber NUR die IDs – Name und Gruppentyp bleiben leer, also kein Typ-Filter möglich)
 *     → je Gruppe /groups/{id}/planner/plans, gebündelt per $batch (max. 20 je Batch); 403/404 ignorieren
 *  Zusammenführen nach id, Titel absteigend (neueste Projektnummern oben). */
async function refreshPlans() {
  const setPh = (step) => { el("plan").placeholder = "Pläne werden geladen … (" + step + ")"; };
  setPh("Anmeldung");
  const byId = new Map();
  const add = (p) => { if (p && p.id && !byId.has(p.id)) byId.set(p.id, { id: p.id, title: p.title || "" }); };
  let firstError = null;

  try {
    (await graphAll("/me/planner/plans", {}, setPh)).forEach(add);
  } catch (e) {
    console.warn("[PK] /me/planner/plans:", msg(e));
    firstError = e;
  }

  let groups = [];
  let groupsOk = false;
  try {
    groups = await loadMyGroups(setPh);
    groupsOk = true;
    console.log("[PK] Gruppen-Mitgliedschaften:", groups.length);
  } catch (e) {
    console.warn("[PK] memberOf fehlgeschlagen – es bleibt bei /me/planner/plans:", msg(e));
    myGroups = null;
    groupsError = msg(e);
    firstError = firstError || e;
  }

  // Je Gruppe: Pläne holen und mitzählen, welche Projektnummern-Jahrgänge darin liegen (→ Team-Label)
  const stats = groups.map((g) => ({ id: g.id, name: g.name || "", plans: 0, years: {} }));
  const note = (st, p) => {
    st.plans++;
    const m = /^(\d{2})\d{3}/.exec(p.title || "");
    if (m) st.years[m[1]] = (st.years[m[1]] || 0) + 1;
  };
  let done = 0;
  for (let i = 0; i < groups.length; i += 20) {
    const chunk = groups.slice(i, i + 20);
    setPh("Teams " + done + "/" + groups.length);
    const batch = { requests: chunk.map((g, k) => ({ id: String(k), method: "GET", url: "/groups/" + g.id + "/planner/plans" })) };
    try {
      const res = await graph("/$batch", { method: "POST", body: JSON.stringify(batch) });
      if (!res.ok) throw new Error("Graph " + res.status + " beim Batch-Abruf der Team-Pläne");
      const j = await res.json();
      for (const r of (j.responses || [])) {
        if (r.status !== 200 || !r.body) continue; // 403/404: Gruppe ohne Planner oder ohne Zugriff → ignorieren
        const st = stats[i + parseInt(r.id, 10)];
        let list = r.body.value || [];
        if (r.body["@odata.nextLink"]) {
          try { list = list.concat(await graphAll(r.body["@odata.nextLink"].replace(CONFIG.graph, ""))); } catch (_) {}
        }
        list.forEach((p) => { add(p); if (st) note(st, p); });
      }
    } catch (e) {
      console.warn("[PK] Batch:", msg(e));
      firstError = firstError || e;
    }
    done += chunk.length;
  }
  if (groupsOk) {
    myGroups = stats.map((st) => ({ id: st.id, plans: st.plans, label: groupLabel(st) }));
    groupsError = null;
    console.log("[PK] Teams:", myGroups.filter((g) => g.plans).map((g) => g.label + " (" + g.plans + ")").join(", "));
  }

  const acc = [...byId.values()];
  if (!acc.length && firstError) throw firstError;
  acc.sort((a, b) => b.title.localeCompare(a.title, "de"));
  plans = acc;
  console.log("[PK] Pläne gesamt:", plans.length);
  localStorage.setItem(CONFIG.planCacheKey, JSON.stringify({ ts: Date.now(), plans, groups: myGroups, groupsError }));
  renderPlanInfo();
  if (mailItem()) detectProject();
}

/* Team-Label: echter Name, falls lesbar; sonst Jahrgang aus den Projektnummern der Pläne ("25xxx" → "2025"). */
function groupLabel(st) {
  if (st.name) return st.name;
  const years = Object.keys(st.years).sort((a, b) => st.years[b] - st.years[a]);
  if (!years.length) return st.plans ? "Sonstige" : "";
  const top = years[0];
  return (st.years[top] >= st.plans * 0.5) ? "20" + top : "gemischt";
}

/* Diagnosezeile unter dem Plan-Feld: wie viele Pläne, aus welchen Teams. */
function renderPlanInfo() {
  const info = el("planinfo");
  if (!plans.length) { info.textContent = ""; info.title = ""; return; }
  let s = plans.length + " Pläne";
  if (myGroups && myGroups.length) {
    const withPlans = myGroups.filter((g) => g.plans > 0);
    const labels = [...new Set(withPlans.map((g) => g.label).filter(Boolean))].sort((a, b) => a.localeCompare(b, "de"));
    if (withPlans.length) {
      const shown = labels.slice(0, 8).join(", ") + (labels.length > 8 ? " +" + (labels.length - 8) + " weitere" : "");
      s += " · Teams: " + shown;
      info.title = "Gruppen mit Plänen:\n" + withPlans.map((g) => g.label + " – " + g.plans + " Pläne").join("\n") +
        (myGroups.length > withPlans.length ? "\n+ " + (myGroups.length - withPlans.length) + " Gruppen ohne Pläne" : "");
    } else {
      s += " · " + myGroups.length + " Gruppen ohne Planner-Pläne";
    }
  } else if (myGroups && !myGroups.length) {
    s += " – keine Team-Mitgliedschaft gefunden (nur direkt geteilte Pläne)";
  } else if (groupsError) {
    s += " – Team-Mitgliedschaften nicht lesbar (" + groupsError + ")";
  }
  info.textContent = s;
}

/* Gruppen/Teams des Nutzers. Mit User.Read kommen nur IDs (displayName/groupTypes leer) – das reicht.
 * Liefert Graph 400 ohne Header, erneut mit ConsistencyLevel. */
async function loadMyGroups(setPh) {
  if (setPh) setPh("Teams");
  let url = "/me/memberOf/microsoft.graph.group?$select=id,displayName&$top=999";
  let opts = {};
  const out = [];
  while (url) {
    let res = await graph(url, opts);
    if (res.status === 400 && !opts.headers) {
      opts = { headers: { ConsistencyLevel: "eventual" } };
      url = url + (url.includes("$count=true") ? "" : "&$count=true");
      res = await graph(url, opts);
    }
    if (!res.ok) throw new Error("Graph " + res.status + " beim Lesen der Team-Mitgliedschaften");
    const j = await res.json();
    (j.value || []).forEach((g) => { if (g.id) out.push({ id: g.id, name: g.displayName || "" }); });
    url = j["@odata.nextLink"] ? j["@odata.nextLink"].replace(CONFIG.graph, "") : null;
  }
  return out;
}

async function loadPeople() {
  try {
    const cached = JSON.parse(localStorage.getItem(CONFIG.peopleCacheKey) || "null");
    if (cached && Array.isArray(cached.people) && cached.people.length) {
      people = cached.people;
      if (Date.now() - cached.ts < CONFIG.cacheTtlMs) return;
    }
    const users = await graphAll("/users?$select=id,displayName,mail&$top=999");
    const dom = "@" + CONFIG.internalDomain.toLowerCase();
    people = users
      .filter((u) => u.displayName && u.mail
        && u.mail.toLowerCase().endsWith(dom)               // nur eigene Domain (keine Gäste)
        && CONFIG.personNamePattern.test(u.displayName))    // nur "Nachname, Vorname" (keine Räume/Funktionspostfächer)
      .map((u) => ({ id: u.id, name: u.displayName, mail: u.mail }))
      .sort((a, b) => a.name.localeCompare(b.name, "de"));
    console.log("[PK] interne Personen:", people.length, "von", users.length);
    localStorage.setItem(CONFIG.peopleCacheKey, JSON.stringify({ ts: Date.now(), people }));
  } catch (e) { /* Zuweisung bleibt dann auf "mir" */ }
}

/* ---------- Buckets des gewählten Plans ---------- */

function hideBuckets() {
  buckets = [];
  bucketsReady = null;
  el("bucketRow").style.display = "none";
  el("bucket").innerHTML = "";
}

async function loadBuckets(planId) {
  const sel = el("bucket");
  const row = el("bucketRow");
  const key = CONFIG.bucketCachePrefix + planId;
  let list = null;
  try {
    const cached = JSON.parse(localStorage.getItem(key) || "null");
    if (cached && Array.isArray(cached.buckets) && Date.now() - cached.ts < CONFIG.cacheTtlMs) list = cached.buckets;
  } catch (_) {}

  if (!list) {
    row.style.display = "block";
    sel.innerHTML = '<option value="">Buckets werden geladen …</option>';
    sel.disabled = true;
    try {
      const res = await graph("/planner/plans/" + planId + "/buckets");
      if (!res.ok) throw new Error("Graph " + res.status + " beim Laden der Buckets");
      const j = await res.json();
      list = (j.value || []).map((b) => ({ id: b.id, name: b.name || "", orderHint: b.orderHint || "" }));
      // Planner sortiert Buckets nach orderHint (ordinaler Stringvergleich, aufsteigend)
      list.sort((a, b) => (a.orderHint < b.orderHint ? -1 : a.orderHint > b.orderHint ? 1 : 0));
      localStorage.setItem(key, JSON.stringify({ ts: Date.now(), buckets: list }));
    } catch (e) {
      console.warn("[PK] loadBuckets:", msg(e));
      list = [];
    } finally {
      sel.disabled = false;
    }
  }
  if (!selectedPlan || selectedPlan.id !== planId) return; // inzwischen anderer Plan gewählt
  renderBuckets(planId, list);
}

function renderBuckets(planId, list) {
  buckets = list;
  const sel = el("bucket");
  if (!list.length) { hideBuckets(); return; } // Plan ohne Buckets → Aufgabe ohne Bucket anlegen
  const last = localStorage.getItem(CONFIG.lastBucketPrefix + planId);
  sel.innerHTML = "";
  list.forEach((b) => {
    const o = document.createElement("option");
    o.value = b.id;
    o.textContent = b.name;
    sel.appendChild(o);
  });
  sel.value = (last && list.some((b) => b.id === last)) ? last : list[0].id;
  el("bucketRow").style.display = "block";
}

/* ---------- Mail lesen + Projekt erkennen ---------- */

function fillFromItem() {
  const item = mailItem();
  if (!item) return;
  showStatus("", "");
  el("create").disabled = false;
  el("due").value = "";
  el("title").value = cleanSubject(item.subject || "");
  selectedPlan = null;
  el("plan").value = "";
  el("detected").textContent = "";
  hideBuckets();
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
  const item = mailItem();
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
  bucketsReady = loadBuckets(p.id);
}

function choosePerson(p) {
  selectedPerson = p;
  el("assign").value = p.name;
  el("assignlist").style.display = "none";
}

/* ---------- Durchsuchbare Listen (Plan + Zuweisung) ---------- */

const combos = {
  plan:   { listId: "planlist",   items: () => plans,  label: (p) => p.title, pick: (p) => choosePlan(p), clear: () => { selectedPlan = null; hideBuckets(); }, idx: -1 },
  assign: { listId: "assignlist", items: () => people, label: (p) => p.name,  pick: (p) => choosePerson(p), clear: () => { selectedPerson = null; }, idx: -1 },
};

function renderList(key, filter) {
  const c = combos[key];
  const list = el(c.listId);
  c.idx = -1;
  const f = (filter || "").trim().toUpperCase();
  const all = c.items();
  const hits = (f ? all.filter((p) => c.label(p).toUpperCase().includes(f)) : all).slice(0, CONFIG.maxListRows);
  list.innerHTML = "";
  if (!hits.length) {
    list.innerHTML = '<div class="none">' + (key === "plan"
      ? "Nichts gefunden – sichtbar sind nur Pläne aus Teams, in denen du Mitglied bist."
      : "Nichts gefunden") + "</div>";
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
  const item = mailItem(); // null im Standalone-Test: Aufgabe ohne Mail-Bezug
  if (!selectedPlan) {
    const f = el("plan").value.trim().toUpperCase();
    const exact = plans.filter((p) => p.title.toUpperCase() === f);
    if (exact.length === 1) choosePlan(exact[0]);
  }
  if (!selectedPlan) { showStatus("Bitte zuerst einen Plan auswählen (Feld 'Projekt / Plan').", "err"); el("plan").focus(); return; }
  const title = el("title").value.trim();
  if (!title) { showStatus("Bitte einen Titel eingeben.", "err"); el("title").focus(); return; }

  const btn = el("create");
  btn.disabled = true;
  showStatus("Aufgabe wird erstellt …", "");

  const subject = item ? (item.subject || "") : title;
  const fromAddr = (item && item.from) ? (item.from.displayName + " <" + item.from.emailAddress + ">") : "";
  const received = (item && item.dateTimeCreated) ? new Date(item.dateTimeCreated).toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" }) : "";
  const planId = selectedPlan.id;

  try {
    // Falls noch kein Token da ist: hier ist ein Klick-Kontext, Popup erlaubt.
    try { await getToken(); } catch (_) { await interactiveToken(); afterAuth(); }

    // Buckets ggf. noch fertig laden, damit die Vorauswahl mitgeht
    if (bucketsReady) { try { await bucketsReady; } catch (_) {} }
    const bucketId = el("bucketRow").style.display !== "none" ? (el("bucket").value || "") : "";

    // 1) Task anlegen
    const body = { planId, title };
    if (bucketId) body.bucketId = bucketId;
    const due = el("due").value;
    if (due) body.dueDateTime = due + "T10:00:00Z";
    const acct = myAccount();
    const assigneeId = (selectedPerson && selectedPerson.id) || (acct && acct.idTokenClaims && acct.idTokenClaims.oid);
    if (assigneeId) body.assignments = { [assigneeId]: { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" } };
    const createRes = await graph("/planner/tasks", { method: "POST", body: JSON.stringify(body) });
    if (createRes.status === 403) {
      throw new Error("Keine Berechtigung für diesen Plan: Du bist nicht Mitglied im Team dieses Projekts. Bitte vom Team-Besitzer ins Team aufnehmen lassen – danach klappt es.");
    }
    if (!createRes.ok) throw new Error("Aufgabe anlegen fehlgeschlagen (Graph " + createRes.status + ")");
    const task = await createRes.json();
    if (bucketId) { try { localStorage.setItem(CONFIG.lastBucketPrefix + planId, bucketId); } catch (_) {} }

    // 2) Link zur Original-Mail holen (die Mail bleibt, wo sie ist)
    let webLink = "";
    if (item) try {
      const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      const g = await graph("/me/messages/" + encodeURIComponent(restId) + "?$select=webLink");
      if (g.ok) webLink = (await g.json()).webLink || "";
    } catch (e) { /* Link ist optional – Aufgabe existiert bereits */ }

    // 3) Beschreibung + Mail-Link an die Aufgabe hängen
    const description =
      "Von: " + fromAddr + "\nEmpfangen: " + received + "\nBetreff: " + subject +
      (webLink ? "\nOriginal-Mail: " + webLink : "") +
      "\n\n— erstellt mit dem Planner-Knopf aus Outlook";
    await patchDetails(task.id, description, webLink, 2);

    const bucketName = bucketId ? (buckets.find((b) => b.id === bucketId) || {}).name : "";
    const link = CONFIG.plannerWeb + planId + "/view/board/task/" + task.id;
    showStatus('✓ Aufgabe angelegt in „' + esc(selectedPlan.title) + '"' + (bucketName ? " → Bucket „" + esc(bucketName) + '"' : "") +
      '.<br><a href="' + link + '" target="_blank" rel="noopener">In Planner öffnen</a>', "ok");
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

/* ---------- Helfer ---------- */

function clearForm() {
  el("title").value = ""; el("plan").value = ""; el("due").value = "";
  selectedPlan = null; el("detected").textContent = "Keine Mail ausgewählt.";
  hideBuckets();
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
    return "Zustimmung erforderlich: Bitte im Anmeldefenster zustimmen – falls 'Administratorgenehmigung erforderlich' erscheint, muss der M365-Admin der App 'Planner-Knopf' einmalig zustimmen. (" + esc(m) + ")";
  }
  if (/popup_window_error|popup window/i.test(m)) {
    return "Das Anmeldefenster wurde blockiert. Bitte Popups für dieses Add-in erlauben oder Outlook im Web verwenden.";
  }
  return "Fehler: " + esc(m);
}

function msg(e) { return (e && (e.message || e.errorMessage)) || String(e); }
function esc(s) { return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }
