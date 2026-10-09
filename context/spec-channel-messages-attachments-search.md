# Spec: teams-mcp Erweiterungen — Kanal-Nachrichten, Attachments, Autor-/Such-Fixes

**Auslöser:** Aktenplan-Harvest 01.10.2026 (DL26-00555). Der Teams-Harvest via `tim` deckte Gruppen-/1:1-Chats vollständig ab, scheiterte aber an 4 dokumentierten Lücken. Diese Spec beschreibt die Fixes. Ziel: Der nächste Harvest kann ALLES automatisch ziehen.

## Stand der Umsetzung (aktualisiert 09.10.2026)

| Befund | Status | Commit |
|---|---|---|
| 1 — Kanal-Nachrichten (Baustein A) | ✅ erledigt (01.10.) | `f9d68e1`, `e2da02a` |
| 2 — `teams_search` 0 Treffer (Baustein C) | ✅ erledigt (09.10.) | `fdca916` |
| 3 — Anhänge nicht herunterladbar (Baustein B) | ✅ erledigt (01.10.) | `f9d68e1`, `a59e010` |
| 4 — Autor-Auflösung defekt (Baustein C) | ✅ erledigt (09.10.) | `14c314a` |
| 5 — Chatlisten-Pagination (Baustein C) | ⬜ **offen** — „See more" wird weiterhin nur herausgefiltert, nicht geklickt; `getFilteredChatRows` liefert nur das geladene Fenster | — |

**Nachträge 08.–09.10.2026 (über die Spec hinaus):**

- **Message Reactions** (`81a8e63`, korrigiert `ade65af`): `teams_get_messages`/`teams_get_channel_messages` liefern je Nachricht `reactions: [{ type, emoji, custom, count, byMe, label }]`. Erste Fassung basierte auf Fehlannahmen (aria-label, CounterBadge, Emoji-Text) und wurde nach Live-DOM-Analyse korrigiert — die Pill nutzt `aria-labelledby="message-<type>-<mid>"`, die Zahl steht im Label-Text, das Emoji in `img[alt]`/`img[itemid]`.
- **Reaktions-Emoji-Leak** (`b0cf2e6`): bei 1:1-Chats liegt der Reaktions-Container INNERHALB des Text-Bodys → Emoji landeten als `![✅](…)` im Markdown. Reaktions-Chrome wird vor der Markdown-Wandlung via `isReactionArtifact()` entfernt.
- **Kanal-Zeitstempel ISO** (`4076553`): Subheader liefert nur „6/23 4:44 PM" ohne Jahr → `parseChannelDateTime()` mit Jahres-Ableitung (Anker: vorherige Nachricht bzw. Plausibilität gegen „jetzt", Wochentage, Relativtage, nur-Uhrzeit = heute).
- **Live-Abnahme Freigabe-Workflow**: Birgit Eilers erteilt formale Freigaben ausschließlich per ✅-Reaction (`type=2705_whiteheavycheckmark`, `byMe=false`) auf Nachrichten, die nur die DL-Nummer enthalten. Bestätigt im Chat „Eilers, Birgit" (DL26-02817, DL25-08496, DL23-10443, DL23-05855, DL26-04047, DL25-07719).

**Lessons (für künftige DOM-Arbeit):** Teams-Web (Fluent UI) verlässt sich auf `aria-labelledby`-Ids, `data-tid`-Familien und `itemid` — NIEMALS `aria-label`/Anzeigetext/CounterBadge raten. Wenn ein Element selbst die Rolle trägt, nach der gesucht wird (`role="group"` am Message-Body), hebt `closest()` auf sich selbst zurück und die Suche bleibt im falschen Subtree. Suchergebnisse sind RELEVANZ-sortiert — chronologische Anker (aufsteigende Jahres-Ableitung) dürfen dort nicht angewendet werden.

**Architektur-Kontext:** teams-mcp ist eine Playwright-Bridge gegen das echte Teams-Web-UI (kein Graph-API). Alle Fixes sind DOM-Scraping-Erweiterungen in `src/teamsClient.js` (+ ggf. `browserManager.js`). Es existieren bereits eingeloggte Browser-Profile pro Tenant (`.teams-browser-profile-<tenant>`), inkl. `dvelop`/`d-velop`.

## Befunde vom 01.10.2026 (Evidenz)

| # | Symptom | Detail |
|---|---|---|
| 1 | `teams_get_messages` auf Team-Kanal → immer n=0 | PUBS-Kanal "Fokusteam Aktenplan-Migration > Allgemein" (Team "PUBS: Fokusteam Aktenplan-Migration", d.velop-Tenant) ist im UI aktiv befüllt (24.09. zuletzt), API liefert 0. `chat_name`-Resolver fehlresolvt zu "Communities and storyline"/falschem 1:1-Chat (Probing chat_index 27–34 dokumentiert). |
| 2 | `teams_search` → 0 Treffer in BEIDEN Tenants | Auch für nachweislich existierende Begriffe ("Aktenplanmigration", "Rottenburg") — Tool unbrauchbar. |
| 3 | Anhänge nicht herunterladbar | Message-Export enthält Name+Link (z.B. `260929_Template_Kickoff_Aktenplanmigration.pptx`), aber kein Download. Teams Web embeddet SharePoint-Links; Download-Button im Attachment-Card-DOM ungenutzt. |
| 4 | Autor-Auflösung defekt | Export zeigt nur "Ich"/"Unbekannt"; Personen mussten aus Zitat-Headern rekonstruiert werden. |
| 5 | Chatlisten-Pagination | Chatliste zeigt ~32 Einträge ("See more" existiert); Chats außerhalb des Fensters (z.B. "Behncke, Lukas, +3") nicht erreichbar. |

## Baustein A — Kanal-Nachrichten lesen (PRIORITÄT 1, unser konkreter Blocker)

**Ziel:** `teams_get_messages` (oder neues Tool `teams_get_channel_messages`) liest Nachrichten aus Team-Kanälen.

**Ansatz:**
1. Navigation: App-Bar → Teams-Icon → Team in Teamliste → Kanal in Kanalsidebar. `teams_list_teams` existiert bereits und liefert vermutlich die Team-/Kanalstruktur — prüfen und ggf. Kanal-Navigation darauf aufbauen.
2. Extraktion: Kanal-Message-Pane nutzt weitgehend dasselbe Message-DOM wie Chats (`data-tid="chat-message"`-Familie) — existierenden Extraktor wiederverwenden, Kanal-Kontext (Team/Kanal-Name) in Metadaten aufnehmen.
3. Safety: Nach Navigation verifizieren, dass der geöffnete Kanal dem Ziel entspricht (analog zum bestehenden Send-Sicherheitsnetz) — sonst Fehler statt falscher Daten.

**Acceptance-Kriterien:**
- PUBS-Kanal liefert Nachrichten mit Autor, Zeitstempel, Text (Sync-Flag "Inline-/System-Nachrichten gekennzeichnet").
- Falsch-Param → klarer Fehler, kein Fehlresolvenz-Silent-Fail wie im chat_index.

## Baustein B — Attachment-Download (PRIORITÄT 2)

**Ziel:** Anhänge aus Nachrichten werden als lokale Dateien gesichert.

**Ansatz:**
1. Playwright-native Downloads: `page.on('download')` + `download.saveAs(...)` — Zielordner: `TEAMS_MCP_DOWNLOAD_DIR` (default neben `~/.teams-media/`).
2. `teams_get_messages` Option `downloadAttachments: true` (Limit + Whitelist nach Extension empfehlenswert) — Rückgabe erweitert um lokale Pfade.
3. SharePoint-gatewayte Dateien: Klick auf Attachment-Card löst i.d.R. den Download/Browser-Tab-Flow aus; Falls-Unterscheidung (direkter Download vs. `&download=1`-URL-Aufruf im selben Page-Kontext) implementieren.

**Acceptance-Kriterien:**
- `260929_Template_Kickoff_Aktenplanmigration.pptx` (Gesamt-JF-Chat, 29.09.) landet als Datei auf der Platte.
- Kein Blockieren des Message-Scrapings, wenn ein Download hängt (Timeout + Überspringen mit Warnung).

## Baustein C — Autor-Auflösung + Chatliste/Search-Fixes (PRIORITÄT 3)

1. **Autor:** Message-Header-Aria-Labels / `data-tid`-Attribute auswerten; eigenes Konto pro Tenant ist bekannt (Config) → "Ich" auflösen. Fallback "Unbekannt" nur noch als echten Ausnahmefall.
2. **Chatliste-Pagination:** "See more"/Scroll-Loop bis Listenlänge stabil (2 identische aufeinanderfolgende Scans); `chat_index`-Resolver danach über komplette Liste.
3. **`teams_search`:** Entweder das native Suchfeld des Webclients ansteuern (Ergebnis-Pane scrappen) ODER als lokaler Filter über die vollständig paginierte Chatliste + Nachrichten-Cache neu definieren (Doku anpassen).

## Test-Strategie

- Unit: `node --test tests/unit.test.mjs` (Bestand erweitern: Resolver-Verifikation, Download-Metadaten, Autor-Mapping).
- Integration: `npm run test:integration` gegen echtes Profil — d.velop-Tenant ist bereits eingeloggt; PUBS-Kanal als Live-Case.

## Out of Scope

- Graph-API-Migration (bewusst Browser-Bridge-Architektur, siehe README).
- Write-Ops auf Kanäle (erst lesend).
- UA-Handling: eigenes Playwright-Chromium hat echten Chrome-UA — das UA-Problem des OpenCode-Desktop-Browsers (teams.microsoft.com blockt Electron-UA serverseitig, gbrain #6901) betrifft teams-mcp NICHT.
