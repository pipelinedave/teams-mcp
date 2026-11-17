# Handoff: Teams-Reactions-Tool im teams-mcp fehlt

**Erstellt:** 2026-10-19 | **Status:** IMPLEMENTIERT (Live-Verifikation ausstehend) | **Priorität:** niedrig-mittel

## Update (19.10.2026 – Implementierung)
Das Tool **`teams_react`** ist im teams-mcp implementiert und registriert:
- `src/teamsClient.js`: Methode `reactToMessage()` + Hilfsmethoden `_openReactionPickerForMessage()`
  und `_pickEmojiInPicker()`; pure Helfer `resolveReactionTarget`, `pickerElementMatches`,
  `EMOJI_SLUG_MAP` (exportiert, unit-getestet).
- `index.js`: Tool-Schema `{ message_id, emoji, chat_name?, chat_index?, tenant? }` + Dispatch.
- Mismatch-Schutz identisch zu `teams_get_messages` (STRICT-Chat-Auflösung +
  `chatTitlesConsistent`-Sicherheitsnetz).
- Eingebaute Rücklese-Verifikation per `collectReactionsFromRow` (Reaktion `byMe` an der
  korrekten `mid`), sonst klarer Fehler statt falschem Erfolg.
- Tests: 216/216 grün (inkl. neuer Reaktions-Eingabe-/Picker-Match-Tests), `npm run check` sauber.

**Einschränkung (Live-Verifikation):** In der Headless-Umgebung renderte die Teams-Web-App nicht
(leerer Body, auch nach 4 min — neues teams.cloud.microsoft/hydrated App bootet headless hier
nicht). Der exakte Reaktions-Picker-DOM konnte daher NICHT live gegen eine echte Session
reverse-engineert/kalibriert werden. Die Implementierung nutzt defensive Mehrfach-Selektoren auf
Basis des im Repo dokumentierten Teams-DOM. **Bitte einmal in einem renderenden Browser-Fenster
gegenprüfen** (siehe "Definition of Done"): `teams_react` auf eine `mid` aus `teams_get_messages`
setzen und bestätigen, dass die 👍-Reaktion an der richtigen Nachricht sichtbar/nachlesbar ist.


## Problem
Der teams-mcp-Server bietet **kein Reaction-Tool** (kein `teams_react` o.ä. — alle
~20 Tools geprüft). Reaktionen (👍 etc.) auf Teams-Nachrichten können Agenten daher
nicht selbstständig setzen — nur Lesen und Senden von Nachrichten geht.

## Auslöser (realer Fall)
- Gruppenchat „adessi @ d.velop" (adesso-Tenant): Lea-Marie Lenz postete am
  19.10. 09:08 eine @Everyone-FTE-Liste (als Bild) mit der Bitte „mit Daumen hoch
  reagieren, wenn korrekt".
- tüm/Agent konnte die 👍-Reaktion nicht setzen; David musste sie manuell setzen.

## Gewünschtes Ergebnis
Ein Reaktion-Tool in `teams-mcp`, das eine Emoji-Reaktion (z.B. `👍`) an einer
identifizierten Nachricht (Chat + Message-ID) setzt — mit Chat-Verifikation gegen
den bekannten teams_get_messages-Mismatch-Defekt.

## Offene Fragen (vor/nach Implementierung klären)
1. Wie werden Reaktionen gesetzt — über Teams-WebClient-DOM (wie der Rest des
   Servers) oder Graph-API? Konsistent zum bestehenden Ansatz des Repos bleiben.
2. Tool-Name & Schema (z.B. `teams_react { chat?, message_id, emoji }`).
3. Mismatch-Schutz wie bei `teams_get_messages` einbauen (Ziel-Chat verifizieren).
4. Rest-für-Rest: einfache Emoji-Whitelist vs. beliebiger String.

## Kontext / relevante Stellen
- Repo: `/home/dhallmann/projects/teams-mcp`
- `src/teamsClient.js`, `src/browserManager.js` (bestehende Chat-/Message-Logik)
- `index.js` (Tool-Registrierung)
- Bekannter Defekt: `teams_get_messages` stiller Chat-Mismatch (Safety-Net nötig)

## Definition of Done
- Tool im teams-mcp registriert, reagiert verifiziert (Eigen-Lesbarkeit: Reaktion
  hängt an der richtigen Nachricht).
- Kein Adhoc-Hack; sauberer, dem Repo-Stil folgender Ansatz.
- Nach Einbau: im gleichen Chat gegenprüfen, dass Reaktion sichtbar/nachlesbar ist.
