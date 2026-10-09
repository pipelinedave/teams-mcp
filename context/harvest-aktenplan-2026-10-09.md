# Aktenplan-Harvest — Auftragsdefinition (DL26-00555, 09.10.2026)

**Zweck:** Wiederholung des Aktenplan-Harvests vom 01.10.2026 — diesmal live gegen die
09.10. implementierten teams-mcp-Features (alle 5 Spec-Befunde erledigt, Commits
`81a8e63`..`cbb127c`). Der Harvest testet die Features END-TO-END über den MCP-Pfad
(direkte Node-Scripte wurden bereits 5x live verifiziert).

**Voraussetzung:** opencode NEU gestartet (der teams-mcp-Server lädt den Code nur beim
Session-Start; laufende Sessions behalten alten Code — nachgewiesen 09.10. 14:45).

**Auszuführen über:** `tim` (teams_* MCP-Tools), Tenant immer `d-velop`. Rein lesend,
KEIN Nachrichtenversand.

## Schritt 0 — Smoke-Test (ABBRUCH-BEDINGUNG)
`teams_list_chats({ tenant: 'd-velop', limit: 10 })`
- Antwort enthält **`totalAvailable`**? → neuer Code aktiv, weiter.
- Fehlt es? → ABBRUCH, melden: "teams-mcp hat alten Code". (Erwartet: ~106 Chats)

## Schritt 1 — Search (neu)
`teams_search({ tenant: 'd-velop', query: 'Aktenplanmigration', limit: 10 })`
Erwartet strukturierte Treffer: kind/chat/channel|members/author/timestamp/snippet.

## Schritt 2 — Kanal (neu: ISO-Zeitstempel, Autoren)
`teams_get_channel_messages({ tenant: 'd-velop', team_name: 'PUBS: Fokusteam Aktenplan-Migration', channel_name: 'Allgemein', limit: 15 })`

## Schritt 3 — Kern-Chats (neu: Pagination, Reactions, Autoren)
- `teams_get_messages({ tenant: 'd-velop', chat_name: 'Gesamt-JF Aktenplanmigration', limit: 25 })`
- `teams_get_messages({ tenant: 'd-velop', chat_name: 'JF Aktenplanmigration: PCs und TCs', limit: 20 })`

## Schritt 4 — Harvest-Auswertung (DL26-00555)
1. **DL-Belegnummern** (Muster `DL\d{2}-\d{5}`) mit Kontext (wer, wann, worum)
2. **Formale Freigaben**: reactions mit `type: '2705_whiteheavycheckmark'` / emoji `✅`
   und `byMe: false` (= Freigabe durch ANDERE, z.B. Birgit Eilers). Liste: Chat,
   DL-Nummer im Nachrichtentext, Freigabedatum
3. **Personen** mit Rollenkontext
4. **Termine/Deadlines** (bekannt: "records management migration training" 22.10.2026,
   Webinar 22.09., Termine 23.10./6.11./13.11./20.11. aus der Suche)
5. **Offene Aufgaben/Blocker**

## Schritt 5 — Feature-Status-Tabelle
Feature | Schritt | OK/FEHLER + Beleg: Search · Channel Messages · Pagination
(Chats hinterm Initialfenster) · Reactions (✅ byMe) · Autoren (kein „Unbekannt") ·
ISO-Zeitstempel

## Nach dem Harvest
- Report ablegen: `reports/aktenplan-harvest-2026-10-09.md` (dieses Repo)
- Shaman-Event für DL26-00555 loggen (shaman_log_event, billable: true)
- Funde mit gbrain abgleichen (Birgit-Eilers-Freigaben: gbrain #7248)

## Bekannte Stolpersteine
- Chat „AC" in der Liste ist KEIN Chat (Communities-Entry) → Safety-Net wirft klaren
  Fehler, kein Bug.
- Gruppen-Chat-Titel im Pane = Mitgliederliste („Borgmann, Katharina, Wobring, Lena"),
  Chatliste zeigt „…, +2" — dasselbe Team, Safety-Net akzeptiert das.
