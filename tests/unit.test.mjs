import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { TeamsClient } from '../src/teamsClient.js';
import { browserManager } from '../src/browserManager.js';

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
});
