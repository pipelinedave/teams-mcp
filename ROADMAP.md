# Roadmap & Feature Spec: Teams Meeting Active Speaker Tracking

## Kontext & Motivation

Aus dem Ideenaustausch mit Yannick Bülter (10.09.2026):
Standard-Transkriptionen (z.B. MacWhisper oder hardwarebasierte Stereo-Spurtrennung wie im `meeting-recorder`) können bisher nur trennen:
- **Kanal 1:** David (Mikrofon)
- **Kanal 2:** Gegenseite gesammelt (System-Sound / Loopback)

Die Zuordnung konkreter Personen auf der Gegenseite erfordert bisher mühsame manuelle Nachbearbeitung.

## Kernidee

Wenn Teams Meetings im Webclient laufen (über Playwright CDP / `teams-mcp`), signalisiert Microsoft Teams den aktuell sprechenden Teilnehmer visuell:
1. **Visuelles Feedback im DOM:** Sprechende Teilnehmer erhalten einen hervorgehobenen Rahmen (typischerweise grüner Border / Audio-Level-Indikator oder CSS-Klasse wie `speaking`, `active-speaker` bzw. Canvas/Video-Container-Attribute).
2. **Klartext-Namen im DOM:** Neben oder unter jedem Video-/Avatar-Kästchen steht der Anzeigename der Person (`[data-tid="roster-avatar-name"]`, `[aria-label*="..."]`).
3. **Meeting-Metadaten:** Termintitel und Teilnehmerliste können vorab direkt aus der Titelleiste bzw. dem Roster-Panel extrahiert werden.

## Geplante Erweiterungen in `teams-mcp`

### Neue Tools

1. `teams_meeting_status`
   - Erkennt, ob aktuell ein aktiver Anruf / Meeting im Browser-Kontext läuft.
   - Extrahiert: Meeting-Titel, aktive Teilnehmerliste, Startzeit.

2. `teams_track_speakers` (Stream / Polling / Event-Listener)
   - Injiziert einen DOM-MutationObserver oder einen intervalbasierten RequestAnimationFrame-Watcher auf die Roster-/Stage-Container.
   - Loggt Zeitstempel-Events:
     ```json
     {
       "timestamp_ms": 1725956400000,
       "speaker": "Yannick Bülter",
       "event": "start_speaking"
     }
     ```
   - Exportiert die Sprecher-Timeline als `.speakers.json` für das Post-Processing im `meeting-recorder`.

## Zusammenspiel mit `meeting-recorder`

- **Timeline-Alignment:** Der `meeting-recorder` nimmt weiter via WASAPI Loopback mit Stereo-Spuren auf.
- Beim Transkribieren gleicht `transcribe_dual.py` die Segmente der Gegenseite mit der Zeitachse der `teams_track_speakers`-Timeline ab.
- Fällt ein Segment der Gegenseite in das Zeitfenster von `Yannick Bülter`, wird statt `**Gegenseite / Kunde:**` direkt `**Yannick Bülter:**` ins Markdown-Transkript geschrieben.
