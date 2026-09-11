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

## Projektstruktur

```
teams-mcp/
├── index.js               # MCP-Server (Tool-Schema + Dispatch)
├── src/
│   ├── config.js          # Zentrale, per Env überschreibbare Konfiguration
│   ├── browserManager.js  # Playwright-Profil-Management (Multi-Tenant, Locks)
│   └── teamsClient.js     # Teams-Web-Automation (Chats, Messages, Search, Send)
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
