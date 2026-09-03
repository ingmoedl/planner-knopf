# Übergabe: Planner-Knopf (Outlook-Add-in „Aufgabe in Planner") – Stand 03.09.2026

Einstieg für einen neuen Claude-Code-Chat. Ergänzt das Projektgedächtnis
(`C:\Users\moedl\.claude\projects\C--Users-moedl-Desktop-Code-PPR\memory\aufgaben-kachel-konzept.md`),
das im Projektordner `C:\Users\moedl\Desktop\Code\PPR` automatisch mitgeladen wird.

## 1. Worum es geht

Ingenieurbüro ing Burghausen GmbH (Nutzer: Samuel Mödl, moedl@ing-burghausen.de, **kein Tenant-Admin**).
Ziel: Aus einer geöffneten E-Mail in Outlook mit zwei Klicks eine Planner-Aufgabe im richtigen
Projektplan anlegen – sofort, ohne Umwege (kein Power Automate, keine Teams-Karte).
Planner-Struktur des Büros: Jahres-Teams „2020" … „2026", darin ein Basis-Plan pro Projekt; Planname
beginnt mit der Projektnummer (Muster `25506-09 WAC LP13 VDI`, `23547-G01 …`, teils nur `26506 …`).

## 2. Ist-Zustand (läuft, vom Nutzer bestätigt)

- **Outlook-Add-in v1.7** (Office-Web-Add-in, XML-Manifest, Taskpane): neues Outlook, klassisches Outlook
  und Outlook im Web. Button „→ Planner" (Gruppe „ing Burghausen") in der geöffneten Mail.
- Panel: Projekt/Plan (durchsuchbare Liste, Projektnummer aus Betreff+Text vorausgewählt), Titel
  (bereinigter Betreff), Zuweisen an (Kollegen-Liste, vorbelegt mit sich selbst), Fällig am, Button
  „Aufgabe erstellen" → Aufgabe + Beschreibung (Von/Empfangen/Betreff/Mail-Link) + Referenz
  „Original-E-Mail" + Link „In Planner öffnen". **Die Mail bleibt im Posteingang** (Verschieben wurde
  auf Nutzerwunsch komplett entfernt).
- Auth: Nested App Authentication (MSAL.js v3.30, self-hosted) mit Login-Knopf „Bei Microsoft anmelden"
  (Popup nur nach Klick), Fallback auf Standard-MSAL außerhalb von NAA; Standalone-Testmodus, wenn
  `Office.context.mailbox` fehlt.
- **Skill „planner-knopf-installation"** für Kollegen (Cowork): Installation, erste Anmeldung, Bedienung,
  Fehlerbehebung. Nutzer hat ihn gespeichert; neue Versionen ersetzt er per „Save skill".
- Alter Power-Automate-Flow „Aufgaben-Kachel – Stufe 1" ist deaktiviert (nicht gelöscht). SharePoint-
  Liste „Projektzuordnung" und Outlook-Ordner `@Aufgabe/Verarbeitet/Unklar` sind Altlasten ohne Funktion.

## 3. Alle Orte und IDs

| Was | Wert |
|---|---|
| Quellcode lokal | `C:\Users\moedl\Desktop\Code\PPR\planner-knopf` (git, Branch main) |
| GitHub-Repo | https://github.com/ingmoedl/planner-knopf (Account `ingmoedl`; gh CLI installiert, angemeldet, `gh auth setup-git` aktiv) |
| Hosting | GitHub Pages https://ingmoedl.github.io/planner-knopf/ (taskpane.html, taskpane.js, msal-browser.min.js, assets/icon-16/32/64/80/128.png, manifest.xml) |
| Manifest-Id | `e3b7a1c4-5d2f-4b8e-9a6c-7f1d0e2b3a58` |
| Entra-App | „Planner-Knopf", Client-ID `92b69fe3-9c55-4262-98d2-4d5642aaeebe`, Tenant `1571141a-75a9-43a3-ad47-8d613cfbb3e6` |
| Redirect-URIs (SPA) | `brk-multihub://ingmoedl.github.io`, `https://ingmoedl.github.io/planner-knopf/taskpane.html` |
| Delegated Scopes (Code = Entra) | `User.Read`, `User.ReadBasic.All`, `Tasks.ReadWrite`, `Mail.ReadWrite` – **zentral vom Admin freigegeben; NICHT ändern (siehe 6)** |
| Skill-Quelle | `C:\Users\moedl\Desktop\Code\PPR\skills\planner-knopf-installation\` (SKILL.md, references/hintergrund-und-datenschutz.md, assets/Planner-Knopf-Manifest.xml) |
| Skill-Paket | `C:\Users\moedl\Desktop\Code\PPR\skills\planner-knopf-installation.skill` (Kopie auf dem Desktop) |
| Skill-Creator (Paketierung) | `C:\Users\moedl\AppData\Roaming\Claude\local-agent-mode-sessions\skills-plugin\cec6e79d-df05-458a-ae9a-adcc4bb04263\68334443-4bee-4e5e-9a6a-037dfa55640b\skills\skill-creator` |
| Bekannte IDs | Team „2025": `ae2e1925-6410-4c99-bda5-c8248395eadf`; Plan „25506-09 WAC LP13 VDI": `6c_rLQin_kOYhEWQVBurrZcACsKA` |
| Konzept-Artefakt (historisch) | https://claude.ai/code/artifact/90c38a79-c244-433d-bb41-57b43a1a99bb |

## 4. Deployment-Ablauf

1. Dateien in `planner-knopf\` ändern. In `taskpane.html` den Cache-Buster hochzählen
   (`taskpane.js?v=N` → N+1; aktuell `v=8`), sonst lädt Outlook bis zu 10 Minuten die alte Logik.
2. `git add -A`, `git commit -m "..."`, `git push origin main` (Push läuft über gh-Credentials).
3. Pages baut 30–90 s; prüfen mit `curl -s https://ingmoedl.github.io/planner-knopf/taskpane.js | grep <Marker>`.
4. Nutzer: Outlook komplett neu starten (Webview-Cache), Panel neu öffnen.
5. Nur bei **Manifest**-Änderung (Button, Name, Höhe): `manifest.xml` neu auf den Desktop kopieren; Nutzer
   lädt sie via aka.ms/olksideload → Meine Add-Ins → Benutzerdefiniert → Aus Datei erneut hoch.
6. Skill ändern → neu paketieren (pyyaml ist installiert):
   `cd <Skill-Creator-Pfad>` dann
   `PYTHONUTF8=1 PYTHONIOENCODING=utf-8 python -m scripts.package_skill "<Skill-Ordner>" "<Zielordner>"`,
   Datei an den Nutzer senden (SendUserFile), er ersetzt sie per „Save skill".

## 5. Test-Möglichkeiten

- Echte Funktionsprüfung (Auth, Pläne, Aufgabe) nur in Outlook/OWA des Nutzers – im eingebauten
  Claude-Browser sind Auth-Popups blockiert („Anmeldefenster wurde blockiert" ist dort normal).
- Standalone-Aufruf https://ingmoedl.github.io/planner-knopf/taskpane.html?fresh=N im Claude-Browser zeigt,
  ob das Skript sauber startet (Konsole via read_console_messages; Login-Knopf muss erscheinen).
- Konsolenausgaben des Add-ins sind mit `[PK]` geprägt.

## 6. Harte Regeln und gelernte Fallen

1. **Scope-Set niemals ändern** (auch nicht verkleinern). Der Tenant erlaubt keine Nutzer-Zustimmung; ein
   Admin hat genau die vier Scopes zentral freigegeben. Jede Abweichung → „Administratorgenehmigung
   erforderlich" (ist am 03.09. passiert). Neue Graph-Aufrufe nur mit den vorhandenen Scopes umsetzen.
2. **Keine deutschen Anführungszeichen in JS-Strings mit doppelten Anführungszeichen als Delimiter**: das
   schließende Zeichen ist oft ein ASCII-Anführungszeichen und bricht den String → SyntaxError, das Add-in
   lädt gar nicht (Symptom: leere Felder, ewiges „Pläne werden geladen…"). Gilt genauso für Python-Patch-
   Skripte. Sichere Variante: einfache Anführungszeichen im Text oder Unicode U+201E/U+201C.
3. Manifest: `RequestedHeight` max. 450 (460 → „Installation fehlgeschlagen"). XML-Manifest mit
   VersionOverrides v1.0 **und** v1.1 (SupportsPinning nur in v1.1).
4. MSAL-Popups dürfen nicht beim Laden aufgehen (werden unterdrückt) → immer Login-Knopf, Popup nur auf Klick.
5. GitHub Pages cached ~10 min → Cache-Buster erhöhen; nach Push mit curl auf neuen Inhalt warten.
6. Windows-Umgebung: PowerShell 5.1 kennt kein `&&`; Windows-Python versteht keine `/c/...`-Pfade (Bash-Tool
   ist Git Bash); `PYTHONUTF8=1` für Emoji-Ausgaben; Node ist NICHT installiert (Admin nötig); lange
   Bash-Heredocs mit Sonderzeichen sind fehleranfällig → Dateien lieber mit dem Write-Tool schreiben.
7. Veröffentlichende Aktionen (z. B. Repo anlegen) kann der Sicherheits-Classifier blocken → Befehl dem
   Nutzer in PowerShell-Syntax geben oder mit klarem Nutzer-Go erneut versuchen.
8. Kommunikation: Nutzer ist Bauingenieur, nicht IT. Deutsch, ein Schritt pro Nachricht, Screenshots genau
   lesen, Fachbegriffe erklären. Er will: sofort, simpel, keine Zwischenschritte, keine erneuten
   Zustimmungsdialoge.

## 7. OFFENE AUFGABEN (Nutzerwunsch vom 03.09.2026, noch nicht umgesetzt)

### 7a) „Nicht alle Pläne werden angezeigt – jeder soll in jedem laufenden Projekt Aufgaben anlegen können"
Ursache: `GET /me/planner/plans` liefert nur Pläne, die Planner dem Nutzer als „geteilt" führt – in der
Praxis unvollständig (bekannte Lücken: Aggregations-/Propagierungsverzögerung, Owner-ohne-Member).
Zusätzlich zeigt `renderList` nur die ersten 60 Treffer (`slice(0, 60)`).

Verifizierter Weg **ohne neue Scopes**:
1. `GET /me/memberOf/microsoft.graph.group?$select=id,displayName&$top=999` – laut Microsoft-Doku
   (user-list-memberof, geprüft 03.09.2026) reicht dafür **User.Read** (bereits freigegeben). OData-Cast
   ohne `$filter`/`$search` braucht keinen ConsistencyLevel-Header; liefert Graph trotzdem 400, dann
   `ConsistencyLevel: eventual` und `$count=true` mitsenden.
2. Je Gruppe `GET /groups/{id}/planner/plans` (Tasks.Read ist in Tasks.ReadWrite enthalten), parallel per
   `Promise.all` oder `POST /$batch` (max. 20 Requests pro Batch); 403/404 einzelner Gruppen ignorieren.
3. Ergebnis mit `/me/planner/plans` (Roster-Pläne) nach `id` zusammenführen, nach Titel absteigend sortieren,
   cachen (bestehenden Cache-Key `pk_plans_v3` auf `pk_plans_v4` erhöhen).
4. In `renderList` das 60er-Limit entfernen bzw. auf ~300 erhöhen.

**Grenze, die dem Nutzer erklärt werden muss:** Planner erlaubt das Anlegen von Aufgaben nur Mitgliedern
der jeweiligen Gruppe. „Jeder in jedem Projekt" setzt voraus, dass alle Mitarbeitenden Mitglied aller
Jahres-Teams sind (organisatorisch: alle in die Teams 2020–2026 aufnehmen). Ohne Mitgliedschaft liefert
Graph 403 – der Fall sollte im Add-in mit einer klaren Meldung abgefangen werden („Du bist nicht Mitglied
im Team dieses Projekts – bitte aufnehmen lassen").

### 7b) Bucket-Auswahl
Nach Wahl eines Plans: `GET /planner/plans/{planId}/buckets` (Tasks.Read), nach `orderHint` sortieren, als
`<select>` unter dem Plan-Feld anzeigen, ersten Bucket vorauswählen (optional: zuletzt genutzten Bucket je
Plan in localStorage merken). `bucketId` in `POST /planner/tasks` mitgeben. Buckets je Plan cachen (6 h).
Hat der Plan keine Buckets: Auswahl ausblenden, ohne `bucketId` anlegen. UI-Reihenfolge: Projekt/Plan →
Bucket → Titel → Zuweisen an → Fällig am.

### 7c) „Zuweisen an" nur interne Personen
Aktuell `GET /users?$select=id,displayName,mail&$top=999` → enthält Gäste sowie Raum-/Funktionspostfächer.
Ein `userType`-Filter ist mit User.ReadBasic.All nicht zuverlässig möglich (Property nicht in „basic").
Robuste Heuristik fürs Büro: nur Einträge mit `mail` auf `@ing-burghausen.de` **und** `displayName` im
Muster „Nachname, Vorname" (Regex `^[^,]+,\s*\S+`) – Personen heißen dort so, Räume/Funktionspostfächer
(„Silentroom", „Content Conversion") nicht. Domain in CONFIG konfigurierbar machen.

### 7d) Danach
- Skill aktualisieren (Bucket-Schritt im Probelauf; Hinweis zur Team-Mitgliedschaft schärfen), neu
  paketieren, an Nutzer senden (ersetzt per „Save skill").
- Gedächtnis-Datei fortschreiben.
- Optional/Aufräumen: alte Outlook-Ordner löschen (Nutzer), deaktivierten Flow und SharePoint-Liste
  „Projektzuordnung" löschen (Nutzer-Entscheidung), später zentraler Rollout via M365 Admin Center
  (Exchange-Admin) statt Einzel-Sideload.

## 8. Relevante Code-Stellen (taskpane.js v1.7)

- `CONFIG` (oben): clientId, tenantId, scopes, Cache-Keys.
- `Office.onReady` → MSAL-Init, `fillFromItem`, ItemChanged-Handler, stille Anmeldung → `afterAuth()`.
- `afterAuth()` → `loadPlans()` (→ `refreshPlans()`), `loadPeople()`, Vorbelegung „Zuweisen an".
- `detectProject()` → Regex `\b(\d{5})(-[A-Za-z0-9]{1,6})?\b`, Match auf `plan.title.startsWith`.
- `combos` / `renderList()` / `wireCombo()` → generische durchsuchbare Listen (plan, assign); für den
  Bucket reicht ein einfaches `<select>`.
- `createTask()` → POST /planner/tasks (planId, title, assignments, dueDateTime) → webLink der Mail →
  `patchDetails()` (GET details → ETag → PATCH mit If-Match; references-Key via `encodeRefKey`).
