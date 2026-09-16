# Teams MCP Server

Microsoft Teams MCP-Server auf Basis einer **Multi-Tenant Playwright Browser-Bridge**.
Steuert den echten Teams-Webclient im Chromium/Chrome über CDP – kein offizielles
API-Integration, dafür mit echten Browser-Sessions und ohne Cloud-Relay.

Ermöglicht LLM-Agenten (opencode, Claude, Cursor, …) den Lese-/Schreibzugriff auf
Microsoft Teams: Chats listen, Nachrichten lesen, durchsuchen, Teams/Channels auflisten
und Nachrichten senden – jeweils in einem isolierten Browser-Profil pro Konto/Tenant.

> **Hinweis**: Dieses Repository ist komplett organisationsneutral. Es enthält keine
> firmenspezifischen Tenants, Accounts oder Zugangsdaten. Alle Zugänge werden von
> dir selbst pro Tenant konfiguriert und per einmaligem Browser-Login eingerichtet.

## Features

- **Multi-Tenant**: Jeder Tenant/Account bekommt ein isoliertes Browser-Profil
  (`.teams-browser-profile-<tenant>`) – keine Session-Kollisionen.
- **Tools**: `teams_status`, `teams_login`, `teams_list_chats`, `teams_get_messages`,
  `teams_search`, `teams_list_teams`, `teams_send_message`, `teams_inspect`, `teams_meeting_status`,
  `teams_start_tracking`, `teams_stop_tracking`, `teams_close`.
- **Aktivitätstab-Analyse**: `teams_get_activity` (Roh-Extraktion), `teams_analyze_activity`
  (Klassifikation), `teams_generate_activity_report` (täglicher Markdown-Report) und
  `teams_schedule_activity_report` (geplanter proaktiver Lauf 2x/Tag) – siehe Abschnitt
  [Aktivitätstab-Analyse & Bericht](#aktivitätstab-analyse--bericht).
- **Dateianhänge unterstützen**: `teams_send_message` akzeptiert optionale lokale Dateipfade
  (`attachments: ["/tmp/file.zip"]` oder `attachment: "/tmp/file.zip"`). Automatischer Upload via
  FileChooser / Input-Bridge inklusive Upload-Fortschritts-Überwachung vor dem Versand.
- **Konsistente Chat- & Adressbuch-Auflösung**: `chat_index` und `chat_name` lösen auf
  bestehende Chats auf. Ist ein Kollege noch nicht in der Chat-Liste, startet `teams_send_message`
  vollautomatisch einen neuen Chat über das globale Unternehmensadressbuch (GAL / People-Picker).
- **Sichere Sendesemantik**: `teams_send_message` sendet nicht blind in den aktiven
  Chat, wenn der Ziel-Chat nicht eindeutig gefunden wird – es wirft stattdessen einen Fehler.
  Vor dem Versand verifiziert ein **Sicherheitsnetz**, dass der tatsächlich geöffnete Chat dem
  Ziel entspricht, sonst wird abgebrochen (kein Fehlversand).
- **Mehrzeilige Nachrichten**: Nachrichten werden per **Clipboard-Paste** eingefügt statt per
  Tastatureingabe – dadurch bleiben Umlaute/Sonderzeichen intakt und Zeilenumbrüche werden
  nicht fälschlich als „Absenden“ interpretiert (kein Zersplittern in mehrere Fragmente).
- **Robust**: Retry gegen Browser-Profil-Lock-Kollisionen (parallele Sessions),
  klare Fehler bei abgelaufener Login-Session.

## Voraussetzungen

- Node.js ≥ 20
- Ein installiertes Chromium/Chrome/Edge. Automatische Erkennung an gängigen Orten
  (Playwright-Cache, `/usr/bin`, macOS-Apps). Alternativ per `TEAMS_MCP_CHROME_PATH` setzen.
- (Optional) WSLg / eine grafische Umgebung für das einmalige Login-Fenster.

## Installation

```bash
git clone <your-repo-url> teams-mcp
cd teams-mcp
npm install
```

## Konfiguration (Umgebungsvariablen)

Alle Optionen sind optional. Im Standardfall funktioniert der Server org-neutral,
indem der `tenant`-Wert, den du an die Tools übergibst, direkt als Konto/Realm
verwendet wird (z.B. `tenant: "deine-org.onmicrosoft.com"`).

| Variable | Beschreibung | Default |
|---|---|---|
| `TEAMS_MCP_CHROME_PATH` | Pfad zur Chrome/Chromium/Edge-Executable | Auto-Detect |
| `TEAMS_MCP_PROFILE_BASE` | Basis-Verzeichnis für die Browser-Profile | `$HOME` |
| `TEAMS_MCP_TENANTS` | Kommagetrennte Whitelist erlaubter Tenant-Keys (für validierte Auswahl in den Tool-Schemas) | – |
| `TEAMS_MCP_TENANT_REALMS` | JSON-Objekt `{ "key": "realm" }` zur Abbildung von Kurznamen auf Realm | – (tenant = realm) |
| `TEAMS_MCP_SELF_NAME` | Eigener Anzeigename für „Ich“-Nachrichten | `"Ich"` |
| `TEAMS_MCP_HEADLESS` | Headless-Default (`true`/`false`) | `true` |

### Beispiele

Kurznamen auf Realms mappen (z.B. für zwei Konten bei zwei Organisationen):

```bash
export TEAMS_MCP_TENANT_REALMS='{"arbeit":"arbeit.onmicrosoft.com","privat":"privat.onmicrosoft.com"}'
export TEAMS_MCP_TENANTS='arbeit,privat'
```

Oder ganz ohne Konfiguration – einfach immer den vollständigen Realm übergeben:

```bash
# tools mit tenant: "meine-org.onmicrosoft.com"
```

## In opencode konfigurieren

Ergänze in `opencode.json` einen MCP-Server-Eintrag (Pfade anpassen):

```json
{
  "mcp": {
    "teams": {
      "type": "local",
      "command": ["node", "/abs/path/zu/teams-mcp/index.js"],
      "enabled": true
    }
  }
}
```

Für andere MCP-Clients (Claude Desktop, Cursor, …) starte den Server entsprechend
über `index.js` bzw. das `teams-mcp`-Binärskript.

## Erste Schritte (Login)

1. `teams_login({ tenant: "deine-org.onmicrosoft.com" })` – öffnet ein **sichtbares**
   Browserfenster.
2. Melde dich dort einmalig an (inkl. MFA). Die Session wird dauerhaft im
   Tenant-Profil (`~/.teams-browser-profile-<tenant>`) gespeichert.
3. Danach sind alle Tools für diesen Tenant bereit.

> **Hinweis**: Das Login ist immer sichtbar. Headless wird nur für die Lese-/
> Schreib-Tools verwendet, nachdem du dich einmalig angemeldet hast.

## Bedienung & Architektur

- **Chat-Adresse**: Nutze `teams_list_chats` für den Chat-`index` ODER den exakten
  Chat-`name`. Beide Wege führen zuverlässig zum selben Chat (exakter Titel-Match,
  danach Präfix-/Token-Match).
- **Tenant-Pflicht**: Jedes Tool erwartet einen `tenant`-Wert. Ohne Konfiguration ist
  das der Realm deiner Organisation (z.B. `deine-org.onmicrosoft.com`).
- **Chrome-Pfad**: Wird kein Browser erkannt, warnt der Server beim Start und erwartet
  `TEAMS_MCP_CHROME_PATH`.
- **Profil-Lock**: Wird ein Profil gerade von einer anderen Instanz genutzt (z.B. eine
  zweite parallele Agenten-Session), wartet der Server mit Backoff und wirft sonst eine
  klare Meldung.

## Aktivitätstab-Analyse & Bericht

Das Toolset kann den Teams-**Aktivitätstab** systematisch auslesen, klassifizieren und dir
als **proaktiven Bericht** präsentieren – damit du den Tab nicht selbst anklicken musst.
Drei-Schichten-Architektur:

| Schicht | Datei | Zweck |
|---|---|---|
| Layer 1 – Extraktion | `src/activityClient.js` | Reads the Activity-Feed (scroll-enabled) und liefert Roh-Items (Text, Autor, Zeitstempel) |
| Layer 2 – Analyse | `src/activityAnalyzer.js` | Klassifiziert Einträge in **Meeting, Task, Entscheidung, Risiko, Sonstiges** mit Relevanz-Score + erzeugt Markdown-Report |
| Layer 3 – Planung | `src/activityReportScheduler.js` | Geplanter proaktiver Lauf 2x/Tag + Steuerung |
| Layer 4 – Zustellung | `src/reportDelivery.js` | Outbox für den `tim`-Agenten (`teams_list_pending_deliveries` / `teams_mark_delivered`) |

**Tools:**

- `teams_get_activity` – Roh-Items aus dem Aktivitätstab (ohne Klassifikation).
- `teams_analyze_activity` – Extrahiert + klassifiziert, liefert gruppierte Analyse (ohne Datei).
- `teams_generate_activity_report` – Extrahiert + klassifiziert + speichert täglichen
  Markdown-Report nach `reports/activity-summary-YYYY-MM-DD.md` (Top-3 je Kategorie, nach Relevanz).
- `teams_schedule_activity_report` – Steuert den geplanten proaktiven Lauf:
  - `action: "status"` – aktueller Scheduler-Zustand (Running? Zeiten? letzter Lauf?)
  - `action: "start"` – aktiviert den 2x/Tag-Scheduler (Standard `["09:00","17:00"]`),
    optional mit `times`, `tenant`, `max_items`
  - `action: "stop"` – deaktiviert den Scheduler
  - `action: "run-once"` – führt sofort eine volle Analyse+Bericht aus (On-Demand)
  - `action: "config"` – zeigt die aktuelle Konfiguration an
- `teams_list_pending_deliveries` – Listet noch nicht zugestellte (pending) Activity-Berichte
  aus der Zustell-Outbox (`reports/outbox/`) auf — für den `tim`-Agenten, damit er weiss,
  welcher Bericht an den Nutzer präsentiert werden soll (chronologisch, inkl. Report-Text).
- `teams_mark_delivered` – Markiert eine Outbox-Zustellung (aus `teams_list_pending_deliveries`)
  als übergeben (rename auf `*.delivered.json`), sodass kein Doppel-Versand erfolgt.

**Kategorien** (breite Content-Range des Feeds): Der Klassifikator sortiert eingehende
Einträge in die fünf Kategorien. `Risiko` (Blocker, Fehler, Fristrisiko) und `Task`
(Aufgaben/Anfragen) werden als `hoch` priorisiert, `Entscheidung` und `Meeting` als
`mittel`, der Rest als `niedrig`. Jeder Lauf speichert einen Markdown-Report nach
`reports/` und legt eine strukturierte **Zustell-Nachricht** in `reports/outbox/` ab.
Der `tim`-Agent holt diese via `teams_list_pending_deliveries` ab, präsentiert den Bericht
proaktiv an den Nutzer (im "Chat mit mir") und markiert ihn via `teams_mark_delivered`
als übergeben — robuste Zustell-Kette ohne Prozesskopplung zwischen MCP-Server und Agent.

**Verifizierte Konfiguration (Stand 16.09.2026):** Ein Smoke-Test über das MCP-Protokoll
(`tools/list` + `tools/call` mit `action:"status"`) bestätigt: `teams_schedule_activity_report`
ist registriert und der Scheduler läuft mit `running: true`, `tenant: "adesso"`,
`times: ["09:00","17:00"]`, `maxItems: 50`, `topN: 3`, Reports nach
`reports/`. Beim Serverstart wird der Scheduler automatisch gestartet
(EnV `TEAMS_MCP_ACTIVITY_SCHEDULER=0` deaktiviert, `TEAMS_MCP_ACTIVITY_TIMES` überschreibt
die Zeiten).

### Zustellung (Layer 4) – Abschlussbericht (Stand 16.09.2026)

Die Zustellkette für den proaktiven Activity-Bericht ist fertiggestellt und verifiziert.
Sie überbrückt robust die Prozessgrenze zwischen `teams-mcp` (MCP-Server) und dem
`tim`-Agenten (opencode-Subagent), ohne dass der Nutzer den Aktivitätstab selbst
anklicken muss. Zentrale Idee: **Outbox als entkoppelter Übergabepunkt**.

**Ablauf (End-to-End):**

1. **Scheduler erzeugt Bericht** – Zur konfigurierten Zeit (`09:00`/`17:00`) führt
   Layer 3 (`activityReportScheduler`) die Analyse aus und speichert den Markdown-Report
   nach `reports/activity-YYYY-MM-DD-HHmm.md`.
2. **Sender legt Outbox-Nachricht an** – Der in `index.js` verdrahtete
   `createOutboxSender()` (aus `src/reportDelivery.js`) serialisiert das Bericht-Ergebnis
   als strukturierte JSON-Nachricht nach `reports/outbox/activity-delivery-*.json`
   (Kind `teams-activity-report`, Recipient `tim`, inkl. `report`, `counts`, `tenant`,
   `date`).
3. **tim-Agent holt ab** – Der `tim`-Agent ruft `teams_list_pending_deliveries` auf,
   erhält die chronologisch sortierten, noch nicht übergebenen Nachrichten und präsentiert
   den Bericht proaktiv an den Nutzer (im "Chat mit mir").
4. **Übergabe bestätigen** – Der `tim`-Agent markiert die Zustellung via
   `teams_mark_delivered` als übergeben. Dadurch wird die Datei auf
   `*.delivered.json` umbenannt und von `listPendingDeliveries` künftig übersprungen –
   **kein Doppel-Versand**.

**Verifikation:** `npm test` 79/79 grün (inkl. Outbox- und Self-Chat-Name-Tests),
`npm run check` syntaktisch sauber. End-to-End-Demo: `node tests/report-delivery.demo.mjs`
(Outbox → tim-Agent → übergeben, ohne echten Browser).

**Design-Merkmale:** Keine Prozesskopplung (nur Datei-Outbox); kanonische
`extractSelfNameFromTitle`-Logik für die Self-Chat-Erkennung in `src/reportDelivery.js`;
kaputte Outbox-Dateien werden über stderr geloggt und übersprungen, ohne die Kette zu
stören.

## Projektstruktur

```
teams-mcp/
├── index.js               # MCP-Server (Tool-Schema + Dispatch)
├── src/
│   ├── config.js          # Zentrale, per Env überschreibbare Konfiguration
│   ├── browserManager.js  # Playwright-Profil-Management (Multi-Tenant, Locks)
│   ├── teamsClient.js     # Teams-Web-Automation (Chats, Messages, Search, Send)
│   ├── activityClient.js  # Layer 1: Activity-Feed-Extraktion
│   ├── activityAnalyzer.js# Layer 2: Klassifikation + Report-Pipeline
│   ├── activityReportScheduler.js # Layer 3: geplanter proaktiver Lauf
│   ├── reportDelivery.js  # Zustellebene: Outbox für den tim-Agenten
│   └── speakerTracker.js  # Active Speaker Tracking (Meetings)
└── package.json
```

## Sicherheit & Compliance

- Dieser Server interagiert mit Microsoft Teams Web über deine eigenen
  Browser-Sessions – es werden keine Zugangsdaten gespeichert oder übertragen.
- Die Browser-Profile mit den Login-Sessions liegen ausschließlich lokal
  (`~/.teams-browser-profile-<tenant>`) und sind in `.gitignore` ausgeschlossen.
- Stelle sicher, dass die Nutzung die Richtlinien deiner Organisation und die
  geltenden Datenschutz-Anforderungen (z.B. DSGVO) erfüllt.

## Tests

Das Projekt verwendet den eingebauten Node-Test-Runner (`node:test`) – keine zusätzlichen
Abhängigkeiten.

**Unit-Tests** (kein Browser, keine Netzwerkzugriffe):

```bash
npm test            # = node --test tests/unit.test.mjs
```

**Integrationstests** (gegen einen echten, bereits angemeldeten Teams-Tenant):

```bash
npm run test:integration          # LANGSAM, sendet NICHT (Standard: Sende-Test gesperrt)
```

Der Integrationstest prüft `status`, `list_chats` (Index-Konsistenz), `get_messages`
(Name↔Index-Auflösung), `search`, `list_teams` und `send_message`.

**Aktivität-Demos** (ohne echten Browser, mit Beispiel-/Simulationsdaten):

```bash
node tests/activity.demo.mjs                # Layer 1: Activity-Feed-Extraktion (Mock)
node tests/activity-analyzer.demo.mjs       # Layer 2: Klassifikation + Report
node tests/scheduler.demo.mjs               # Layer 3: Scheduler-Pipeline (simuliert)
node tests/report-delivery.demo.mjs         # Zustellebene: Outbox -> tim-Agent -> übergeben
```

> ⚠️ **Sicherheit beim Senden**: Standardmäßig wird der Sende-Teil des Integrationstests
> **übersprungen**. Nur wenn du ihn explizit freigibst UND als Empfänger deinen eigenen
> Self-Chat angegeben hast, wird tatsächlich eine Nachricht gesendet:
>
> ```bash
> TMS_TEST_TENANT=deine-org.onmicrosoft.com \
> TMS_TEST_SEND_ALLOWED=true \
> TMS_TEST_RECIPIENT="Dein Name" \      # unbedingt dein eigener Self-Chat!
> node tests/integration.test.mjs
> ```
>
> So wird in Testläufen garantiert **nur an dich selbst** gesendet – niemals in einen
> fremden oder Gruppen-Chat.

## Lizenz

Siehe `LICENSE`. (Standard: zum privaten/internen Gebrauch oder wie in der LICENSE-Datei
angegeben.)
