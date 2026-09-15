import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { TeamsClient, messageTimeToIso } from '../src/teamsClient.js';
import { browserManager } from '../src/browserManager.js';
import { cleanSpeakerName, aggregateEvents, isBotOrUiName } from '../src/speakerTracker.js';

// Instanz des TeamsClient (nur zur Nutzung der puren pickChatRow-Methode)
const client = new TeamsClient();

describe('pickChatRow - Chat-Auswahl-Logik', () => {
  const filtered = [
    { row: 'r0', title: 'Yannick Bülter', text: 'Yannick Bülter' },
    { row: 'r1', title: 'Theys Schiller', text: 'Theys Schiller' },
    { row: 'r2', title: 'Yannick und Kollegen', text: 'Yannick und Kollegen' },
    { row: 'r3', title: 'Lukas Behncke', text: 'Lukas Behncke' }
  ];

  test('exakter Titel-Match gewinnt gegen Substring', () => {
    const r = client.pickChatRow(filtered, 'Yannick', undefined);
    assert.equal(r.title, 'Yannick Bülter'); // exakter Präfix-Match, nicht "Yannick und Kollegen"
  });

  test('exakter exakter Case-insensitive Match', () => {
    const r = client.pickChatRow(filtered, 'yannick bülter', undefined);
    assert.equal(r.title, 'Yannick Bülter');
  });

  test('chatIndex gewinnt über chatName', () => {
    const r = client.pickChatRow(filtered, 'Lukas Behncke', 1);
    assert.equal(r.title, 'Theys Schiller');
  });

  test('chatIndex out of range -> fällt auf Name zurück', () => {
    const r = client.pickChatRow(filtered, 'Lukas Behncke', 99);
    assert.equal(r.title, 'Lukas Behncke');
  });

  test('negativer chatIndex -> Name', () => {
    const r = client.pickChatRow(filtered, 'Yannick Bülter', -1);
    assert.equal(r.title, 'Yannick Bülter');
  });

  test('kein Treffer -> null', () => {
    const r = client.pickChatRow(filtered, 'GibtEsNicht', undefined);
    assert.equal(r, null);
  });

  test('leerer chatName/kein Index -> null', () => {
    assert.equal(client.pickChatRow(filtered, '', undefined), null);
    assert.equal(client.pickChatRow(filtered, undefined, undefined), null);
    assert.equal(client.pickChatRow(filtered, undefined, null), null);
  });

  test('Präfix-Match: "Yann" -> Yannick', () => {
    const r = client.pickChatRow(filtered, 'Yann', undefined);
    assert.equal(r.title, 'Yannick Bülter');
  });

  test('Token-Match: "Yannick Kollegen" -> "Yannick und Kollegen"', () => {
    const r = client.pickChatRow(filtered, 'Yannick Kollegen', undefined);
    assert.equal(r.title, 'Yannick und Kollegen');
  });
});

describe('normalizeTenant / realm - Org-neutrale Tenant-Logik', () => {
  test('exakter konfigurierter Key', () => {
    assert.equal(browserManager.normalizeTenant('adesso'), 'adesso');
  });

  test('org-neutral: unbekannter Wert zählt direkt', () => {
    assert.equal(browserManager.normalizeTenant('meine-org'), 'meine-org');
  });

  test('case-insensitive', () => {
    assert.equal(browserManager.normalizeTenant('ADESSO'), 'adesso');
  });

  test('realm: Konvention onmicrosoft.com bei purem Namen', () => {
    assert.equal(browserManager.realm('meine-org'), 'meine-org.onmicrosoft.com');
  });

  test('realm: Realm-artiger Wert direkt', () => {
    assert.equal(browserManager.realm('meine-org.onmicrosoft.com'), 'meine-org.onmicrosoft.com');
  });

  test('Smart-Default: leerer Tenant, undefined oder "all" fällt auf Default ("adesso") zurück', () => {
    assert.equal(browserManager.normalizeTenant(''), 'adesso');
    assert.equal(browserManager.normalizeTenant(undefined), 'adesso');
    assert.equal(browserManager.normalizeTenant(null), 'adesso');
    assert.equal(browserManager.normalizeTenant('all'), 'adesso');
  });
});


describe('speakerTracker - Name cleaning & Event aggregation', () => {
  test('cleanSpeakerName entfernt Rollen-Suffixe', () => {
    assert.equal(cleanSpeakerName('Yannick Bülter (Gast)'), 'Yannick Bülter');
    assert.equal(cleanSpeakerName('Mathis Künzel (Organizer)'), 'Mathis Künzel');
    assert.equal(cleanSpeakerName('Pierre (Extern)'), 'Pierre');
    assert.equal(cleanSpeakerName('David Hallmann'), 'David Hallmann');
    assert.equal(cleanSpeakerName(''), '');
  });

  test('isBotOrUiName erkennt Teams-Copilot und UI-Platzhalter', () => {
    assert.equal(isBotOrUiName('Copilot'), true);
    assert.equal(isBotOrUiName('Microsoft Copilot'), true);
    assert.equal(isBotOrUiName('Copilot Notebook'), true);
    assert.equal(isBotOrUiName(''), true);
    assert.equal(isBotOrUiName(null), true);
    assert.equal(isBotOrUiName('Recording in progress'), true);
    assert.equal(isBotOrUiName('Chat'), true);
  });

  test('isBotOrUiName lässt echte Personen durch', () => {
    assert.equal(isBotOrUiName('Theys Schiller'), false);
    assert.equal(isBotOrUiName('Yannick Bülter'), false);
    assert.equal(isBotOrUiName('David Hallmann'), false);
  });

  test('aggregateEvents: leere Samples liefern leeres Array', () => {
    assert.deepEqual(aggregateEvents([]), []);
    assert.deepEqual(aggregateEvents(null), []);
  });

  test('aggregateEvents: kontinuierliche Samples werden zu 1 Intervall gemerged', () => {
    const base = 1726040000000;
    const samples = [
      { speaker: 'Yannick Bülter', timestamp: base },
      { speaker: 'Yannick Bülter', timestamp: base + 100 },
      { speaker: 'Yannick Bülter', timestamp: base + 200 },
      { speaker: 'Yannick Bülter', timestamp: base + 300 },
      { speaker: 'Yannick Bülter', timestamp: base + 400 },
      { speaker: 'Yannick Bülter', timestamp: base + 500 }
    ];

    const intervals = aggregateEvents(samples, { meetingStartEpochMs: base });
    assert.equal(intervals.length, 1);
    assert.equal(intervals[0].speaker, 'Yannick Bülter');
    assert.equal(intervals[0].start_epoch_ms, base);
    assert.equal(intervals[0].end_epoch_ms, base + 500);
    assert.equal(intervals[0].start_offset_sec, 0);
    assert.equal(intervals[0].end_offset_sec, 0.5);
  });

  test('aggregateEvents: Pause > minGapMs erzeugt zwei getrennte Intervalle', () => {
    const base = 1726040000000;
    const samples = [
      // Block 1 (500ms Dauer)
      { speaker: 'Mathis', timestamp: base },
      { speaker: 'Mathis', timestamp: base + 200 },
      { speaker: 'Mathis', timestamp: base + 500 },
      // 1000ms Pause
      // Block 2 (600ms Dauer)
      { speaker: 'Mathis', timestamp: base + 1500 },
      { speaker: 'Mathis', timestamp: base + 1800 },
      { speaker: 'Mathis', timestamp: base + 2100 }
    ];

    const intervals = aggregateEvents(samples, { minGapMs: 400, meetingStartEpochMs: base });
    assert.equal(intervals.length, 2);
    assert.equal(intervals[0].speaker, 'Mathis');
    assert.equal(intervals[0].end_offset_sec, 0.5);
    assert.equal(intervals[1].start_offset_sec, 1.5);
    assert.equal(intervals[1].end_offset_sec, 2.1);
  });

  test('aggregateEvents: Spikes unter minDurationMs werden verworfen', () => {
    const base = 1726040000000;
    const samples = [
      // Nur ein einzelner Spike von 100ms
      { speaker: 'Hintergrundgeräusch', timestamp: base },
      { speaker: 'Hintergrundgeräusch', timestamp: base + 100 }
    ];

    const intervals = aggregateEvents(samples, { minDurationMs: 250, meetingStartEpochMs: base });
    assert.equal(intervals.length, 0);
  });

  test('aggregateEvents: Mehrere Sprecher chronologisch sortiert', () => {
    const base = 1726040000000;
    const samples = [
      { speaker: 'Yannick', timestamp: base },
      { speaker: 'Yannick', timestamp: base + 300 },
      { speaker: 'Mathis', timestamp: base + 500 },
      { speaker: 'Mathis', timestamp: base + 900 }
    ];

    const intervals = aggregateEvents(samples, { meetingStartEpochMs: base });
    assert.equal(intervals.length, 2);
    assert.equal(intervals[0].speaker, 'Yannick');
    assert.equal(intervals[1].speaker, 'Mathis');
  });
});

describe('messageTimeToIso - Teams-Zeitstempel-Normalisierung', () => {
  test('data-mid als Unix-Millis (13-stellig) -> ISO', () => {
    // 1789463782694 ms ~ 2026-09-15
    const iso = messageTimeToIso('1789463782694');
    assert.ok(iso.startsWith('2026-09-15T'));
    assert.ok(iso.endsWith('Z'));
  });

  test('Unix-Sekunden (10-stellig) -> ISO', () => {
    const iso = messageTimeToIso('1789463782');
    assert.ok(iso.startsWith('2026-09-15T'));
  });

  test('<time datetime> ISO-String bleibt erhalten', () => {
    const iso = messageTimeToIso('2026-09-15T11:16:27.330Z');
    assert.equal(iso, new Date('2026-09-15T11:16:27.330Z').toISOString());
  });

  test('leere / null / undefined -> null', () => {
    assert.equal(messageTimeToIso(''), null);
    assert.equal(messageTimeToIso(null), null);
    assert.equal(messageTimeToIso(undefined), null);
  });

  test('Nicht-Zeitstempel (z.B. GUID) -> null', () => {
    assert.equal(messageTimeToIso('1:19c8f2a0-9b3d-4c5e-a1f2-8b7c6d5e4f3a'), null);
    assert.equal(messageTimeToIso('kein timestamp'), null);
  });

  test('Millis außerhalb plausibler Reichweite -> null', () => {
    assert.equal(messageTimeToIso('999999'), null);
  });
});

describe('validateAttachments - Dateianhang-Validierung', () => {
  test('leere oder nicht gesetzte Anhänge liefern leeres Array', () => {
    assert.deepEqual(client.validateAttachments(null), []);
    assert.deepEqual(client.validateAttachments(undefined), []);
    assert.deepEqual(client.validateAttachments(''), []);
    assert.deepEqual(client.validateAttachments([]), []);
  });

  test('existierende Datei wird zu absolutem Pfad normalisiert', () => {
    const res = client.validateAttachments('package.json');
    assert.equal(res.length, 1);
    assert.ok(res[0].endsWith('package.json'));
    assert.ok(res[0].startsWith('/'));
  });

  test('Array aus existierenden Dateien', () => {
    const res = client.validateAttachments(['package.json', 'README.md']);
    assert.equal(res.length, 2);
    assert.ok(res[0].endsWith('package.json'));
    assert.ok(res[1].endsWith('README.md'));
  });

  test('nicht existierende Datei wirft verständlichen Fehler', () => {
    assert.throws(
      () => client.validateAttachments('/tmp/nicht-existierende-datei-12345.xyz'),
      /Anhang-Datei nicht gefunden/
    );
  });

  test('Verzeichnis statt Datei wirft Fehler', () => {
    assert.throws(
      () => client.validateAttachments('src'),
      /Anhang ist keine reguläre Datei/
    );
  });
});
