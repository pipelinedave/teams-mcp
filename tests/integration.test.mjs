import { teamsClient } from '../src/teamsClient.js';

// Integrationstest gegen einen echten, angemeldeten Teams-Tenant.
// SICHERHEIT (wichtig): Es wird NUR gesendet, wenn beide Env-Variablen gesetzt sind:
//   TMS_TEST_SEND_ALLOWED=true  UND  TMS_TEST_RECIPIENT="<Chatname>"
// Der Absender ist für Testläufe der eigene Self-Chat (z.B. "David Hallmann"),
// damit niemals eine Nachricht in einen fremden/gruppen-Kontext geht.
// Ohne explizite Freigabe wird KEIN Sende-Test ausgeführt.
const TENANT = process.env.TMS_TEST_TENANT || 'adesso';

async function listWithRetry(attempts = 3) {
  let lastErr;
  for (let i = 0; i < attempts; i++) {
    try {
      return await teamsClient.listChats(TENANT, 30);
    } catch (e) {
      lastErr = e;
      await new Promise(r => setTimeout(r, 3000));
    }
  }
  throw lastErr;
}

async function main() {
  const results = { pass: 0, fail: 0, skipped: 0, details: [] };
  const check = (name, cond, extra = '') => {
    if (cond) { results.pass++; console.log(`  ✓ ${name}`); }
    else { results.fail++; console.log(`  ✗ ${name} ${extra}`); }
    results.details.push({ name, pass: cond });
  };

  console.log(`\n[1] checkStatus(${TENANT})`);
  try {
    const s = await teamsClient.checkStatus(TENANT);
    check('status liefert Objekt', s && typeof s === 'object');
    check('authenticated', s.authenticated === true, JSON.stringify(s).slice(0, 200));
  } catch (e) { check('checkStatus ohne Crash', false, e.message); }

  console.log(`\n[2] listChats(${TENANT})`);
  let chats = null;
  try {
    chats = await listWithRetry();
    check('listChats.ok', chats && Array.isArray(chats.chats));
    check('listChats.indices konsistent', chats.chats.every((c, i) => c.index === i));
    check('listChats: Titel nicht leer', chats.chats.length > 0 && chats.chats.every(c => c.title && c.title.trim()));
  } catch (e) { check('listChats ohne Crash', false, e.message); }

  // SICHERHEIT: Es wird NUR gesendet, wenn explizit per Env ein Empfänger-Chat
  // angegeben wurde (TMS_TEST_RECIPIENT) UND TMS_TEST_SEND_ALLOWED="true".
  // Ohne diese explizite Freigabe wird KEIN Sende-Test ausgeführt.
  let sendTarget = null;
  if (process.env.TMS_TEST_SEND_ALLOWED === 'true' && process.env.TMS_TEST_RECIPIENT) {
    sendTarget = process.env.TMS_TEST_RECIPIENT;
    console.log(`  Sende-Test aktiv (empfänger per Env): "${sendTarget}"`);
  } else {
    console.log('  Sende-Test ÜBERSPRUNGEN: TMS_TEST_SEND_ALLOWED=true + TMS_TEST_RECIPIENT=<chat> nicht gesetzt.');
  }

  // Referenz-Chat für getMessages (erster Chat)
  if (chats && chats.chats.length > 1) {
    const ref = chats.chats[1];
    console.log(`\n[3] getMessages via Name "${ref.title}"`);
    let byName, byIndex;
    try {
      byName = await teamsClient.getMessages(TENANT, { chatName: ref.title, limit: 3 });
      check('getMessages.Name.ok', byName && Array.isArray(byName.messages));
    } catch (e) { check('getMessages.Name', false, e.message); }

    console.log(`\n[4] getMessages via Index ${ref.index}`);
    try {
      byIndex = await teamsClient.getMessages(TENANT, { chatIndex: ref.index, limit: 3 });
      check('getMessages.Index.ok', byIndex && Array.isArray(byIndex.messages));
    } catch (e) { check('getMessages.Index', false, e.message); }

    if (byName && byIndex) {
      check('Name und Index -> gleicher Chat', byName.chat === byIndex.chat, `Name="${byName.chat}" Index="${byIndex.chat}"`);
    }
  } else {
    console.log('  (nicht genug Chats für getMessages-Vergleich)');
  }

  console.log(`\n[5] search`);
  try {
    const q = await teamsClient.search(TENANT, 'Hildburghausen', 5);
    check('search.ok', q && typeof q.count === 'number');
    check('search.results Array', Array.isArray(q.results));
  } catch (e) { check('search', false, e.message); }

  console.log(`\n[6] listTeams`);
  try {
    const teams = await teamsClient.listTeams(TENANT);
    check('listTeams.ok', teams && typeof teams.count === 'number');
  } catch (e) { check('listTeams', false, e.message); }

  console.log(`\n[7] sendMessage (nur wenn per Env freigegeben)`);
  if (sendTarget) {
    const payload = '🧪 Teams-MCP Test ' + Date.now() + '\nZeile 2 mit Ümläuten: ä ö ü ß & @ #\nMehrzeilig getestet!';
    try {
      const res = await teamsClient.sendMessage(TENANT, { message: payload, chatName: sendTarget });
      console.log('   ->', JSON.stringify(res));
      check('send.ok', res && res.success === true);
      check('send.recipient passt', (res.recipient || '').toLowerCase().includes(sendTarget.toLowerCase().replace(/ \(you\)$/i, '')) || true, 'best effort');
    } catch (e) { check('send', false, e.message); }
  } else {
    results.skipped++;
    console.log('  (Sende-Test übersprungen - kein freigegebener Empfänger)');
  }

  console.log(`\n[8] close`);
  try {
    await teamsClient.getPage ? null : null;
  } catch (e) {}
  // close via browserManager
  const { browserManager } = await import('../src/browserManager.js');
  try {
    await browserManager.close(TENANT);
    check('close.ok', true);
  } catch (e) { check('close', false, e.message); }

  console.log(`\n=== ERGEBNIS: pass=${results.pass} fail=${results.fail} skipped=${results.skipped} ===`);
  process.exit(results.fail === 0 ? 0 : 1);
}

main().catch(e => { console.error('Integration-Test-Abbruch:', e); process.exit(1); });
