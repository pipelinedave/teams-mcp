import { config } from './config.js';
import { browserManager } from './browserManager.js';

// Timeout-Helfer mit klarer Fehlermeldung (Anti-Stillstand)
const withTimeout = (promise, ms, label) => {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`Timeout nach ${ms}ms: ${label}`)), ms);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
};

/**
 * Activity-Tab-Client für Microsoft Teams (Web).
 *
 * Extrahiert den "Aktivität"-Feed (Activity-Tab) aus dem Teams-Web-DOM über Playwright.
 * Der Feed ist ein virtuell rendender Scroll-Container; die extrahierten Roh-Items werden
 * hier NICHT klassifiziert (das ist Layer 2 / tim-Agent), sondern nur stabil ausgelesen.
 *
 * Fehlerbehandlung: jede Stufe wirft klare, kontextreiche Fehler (mit Tenant + Schritt),
 * damit ein streikendes DOM nicht stillschweigend leere Ergebnisse liefert.
 * Logging: strukturierte Log-Zeilen über console.error (stderr), die den Stdio-JSON-RPC
 * Kanal des MCP-Servers nicht verschmutzen (stdout bleibt für MCP reserviert).
 */
export class ActivityClient {
  constructor({ logger = null } = {}) {
    // Logger als injizierbare Abhängigkeit (testbar). Ohne Logger: no-op + stderr-Fallback.
    this.log = logger || {
      info: (msg, meta) => console.error(`[activity] INFO  ${msg}${meta ? ' ' + JSON.stringify(meta) : ''}`),
      warn: (msg, meta) => console.error(`[activity] WARN  ${msg}${meta ? ' ' + JSON.stringify(meta) : ''}`),
      error: (msg, meta) => console.error(`[activity] ERROR ${msg}${meta ? ' ' + JSON.stringify(meta) : ''}`)
    };
  }

  /**
   * Öffnet den Activity-Tab der Teams-App (via App-Bar-Icon) und liefert die Page.
   * Robuste Kaskade über data-tid / aria-label / title (DE+EN).
   */
  async openActivityTab(page, tenant) {
    // Kaskade für das Activity-Icon in der linken App-Bar
    const iconSelectors = [
      'button[data-tid="activity"]',
      'button[data-tid="app-bar-activity"]',
      'button[data-tid*="activity" i]',
      'button[aria-label*="Aktivität" i]',
      'button[aria-label*="Activity" i]',
      'button[title*="Aktivität" i]',
      'button[title*="Activity" i]',
      'button:has([data-icon-name*="Activity" i])',
      'a[aria-label*="Activity" i]'
    ];

    let clicked = false;
    for (const sel of iconSelectors) {
      try {
        const btn = await page.$(sel);
        if (btn && await btn.isVisible().catch(() => true)) {
          await btn.click();
          clicked = true;
          break;
        }
      } catch (e) { /* nächster Selektor */ }
    }

    if (!clicked) {
      // Fallback: per Tastatur (Teams-Web: Ctrl+Shift+1 bzw. Ctrl+Shift+2 öffnet Chats/Teams;
      // die App-Bar ist per Maus pflicht; wir versuchen es zusätzlich per evaluate-Klick).
      const domClicked = await page.evaluate(() => {
        const all = Array.from(document.querySelectorAll('button, [role="tab"], [role="treeitem"], a'));
        const candidates = all.filter((el) => {
          const s = [
            el.getAttribute('data-tid') || '',
            el.getAttribute('aria-label') || '',
            el.getAttribute('title') || '',
            el.textContent || ''
          ].join(' ').toLowerCase();
          return s.includes('aktivität') || s.includes('activity');
        });
        if (candidates.length > 0) {
          (candidates[0]).click();
          return true;
        }
        return false;
      });
      if (!domClicked) {
        throw new Error(
          `Activity-Tab-Icon nicht gefunden (tenant="${tenant}"). Kein Element mit data-tid/aria-label "Aktivität"/"Activity" in der App-Bar vorhanden.`
        );
      }
    }

    // Warten, bis der Activity-Feed-Container im DOM erscheint
    await page.waitForSelector(
      '[data-tid="activity-feed"], [data-tid*="activity" i], [role="complementary"] [data-tid*="activity" i], [data-tid="left-rail"]',
      { timeout: 12000 }
    ).catch(() => null);

    return page;
  }

  /**
   * Scannt den aktuellen Activity-Feed im Seiten-DOM und extrahiert alle sichtbaren Items
   * als strukturierte Roh-Objekte. Läuft als reine page.evaluate-Funktion im Kontext der
   * Seite (kein Zugriff auf Modul-Scope nötig).
   */
  async extractFeed(page, maxItems) {
    return page.evaluate(({ max }) => extractActivityFromDom(document, { max }), { max: maxItems });
  }

  /**
   * Vollständiger Extraktions-Lauf für den Activity-Tab eines Tenants.
   * Öffnet den Tab, wartet auf den Feed und liefert die Roh-Items + Metadaten.
   */
  async getActivity(tenant = '', { maxItems = 50 } = {}) {
    const t = browserManager.normalizeTenant(tenant);
    if (!t) throw new Error("Kein 'tenant' angegeben. Bitte Tenant-Name/Realm mitgeben.");

    this.log.info(`Extraktion Activity-Tab gestartet`, { tenant: t, maxItems });

    const page = await withTimeout(
      browserManager.ensureContext(t, true).then(({ page }) => page),
      30000,
      `Page-Context für tenant "${t}" öffnen`
    );

    if (!page.url().includes('teams.')) {
      await page.goto(`https://teams.microsoft.com/v2/?realm=${encodeURIComponent(browserManager.realm(t))}`, {
        waitUntil: 'domcontentloaded',
        timeout: 30000
      });
      await page.waitForSelector('[data-tid="app-bar-chat"], [role="treeitem"]', { timeout: 20000 }).catch(() => null);
    }

    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error(`Nicht in Teams (${t}) angemeldet. Bitte teams_login({ tenant: "${t}" }) in sichtbarem Fenster ausführen.`);
    }

    await this.openActivityTab(page, t);
    // Kurz warten, bis virtuelle Liste den sichtbaren Bereich gerendert hat
    await page.waitForTimeout(1500);

    const items = await this.extractFeed(page, maxItems);

    if (items.length === 0) {
      this.log.warn(`Keine Activity-Items extrahiert`, { tenant: t });
    } else {
      this.log.info(`${items.length} Activity-Items extrahiert`, { tenant: t });
    }

    return {
      tenant: t,
      extractedAt: new Date().toISOString(),
      count: items.length,
      items
    };
  }
}

export const activityClient = new ActivityClient();

/**
 * PURE DOM-Extraktionsfunktion für den Activity-Feed.
 *
 * Läuft im Browser-Kontext (via page.evaluate) UND ist direkt in Unit-Tests
 * mit einem minimalen DOM-Mock testbar. Erwartet ein document-artiges Objekt mit
 * querySelector/querySelectorAll (siehe extractActivityFromDom). Alle DOM-Operationen
 * sind lokale Funktionen, keine Modul-Scope-Abhängigkeiten.
 */
export function extractActivityFromDom(doc, { max = 50 } = {}) {
  const toIso = (raw) => {
    if (raw == null || raw === '') return null;
    const s = String(raw).trim();
    if (/^\d{10,13}$/.test(s)) {
      const num = Number(s);
      const ms = s.length >= 13 ? num : num * 1000;
      const d = new Date(ms);
      if (!isNaN(d.getTime()) && d.getFullYear() > 2000 && d.getFullYear() < 2100) return d.toISOString();
    }
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
      const d = new Date(s);
      if (!isNaN(d.getTime())) return d.toISOString();
    }
    return null;
  };

  const qsa = (sel) => (doc.querySelectorAll ? Array.from(doc.querySelectorAll(sel)) : []);

  const items = [];
  const seen = new Set();

  const feedRoot = doc.querySelector
    ? doc.querySelector('[data-tid="activity-feed"], [data-tid*="activity" i][role="list"], [role="list"][aria-label*="Aktivität" i]')
    : null;

  let nodes = [];
  if (feedRoot) {
    const children = feedRoot.querySelectorAll
      ? Array.from(feedRoot.querySelectorAll('[role="listitem"], [data-tid*="activity-item" i], [data-tid*="activity" i]'))
      : [];
    nodes = children.filter((el) => !(el.querySelector && el.querySelector('[role="listitem"]')));
  } else {
    nodes = qsa('[data-tid*="activity" i], [role="listitem"]');
  }

  for (const el of nodes) {
    if (items.length >= max) break;
    const rawText = (el.innerText || el.textContent || '');
    const text = String(rawText).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    if (!text || text.length < 4) continue;

    const key = text.slice(0, 120);
    if (seen.has(key)) continue;
    seen.add(key);

    let timestamp = null;
    let timestampRaw = null;
    let timeEl = null;
    if (el.querySelector) timeEl = el.querySelector('time[datetime]');
    if (timeEl) {
      const dt = timeEl.getAttribute('datetime');
      const iso = toIso(dt);
      if (iso) { timestamp = iso; timestampRaw = dt; }
      else { timestampRaw = (timeEl.innerText || '').trim(); }
    }
    if (!timestamp) {
      const tsText = text.match(/(\d{1,2}:\d{2}\s?(?:AM|PM)?|\b\d{1,2}\.\d{1,2}\.\d{2,4}\b|\b(?:heute|gestern|Yesterday|Today)\b)/i);
      if (tsText) { timestampRaw = tsText[1]; }
    }

    let author;
    let avatar = null;
    if (el.querySelector) avatar = el.querySelector('[data-tid*="avatar" i], [role="img"]');
    if (avatar) {
      const a = (avatar.getAttribute('aria-label') || '').trim();
      if (a) author = a;
    }

    items.push({
      text,
      author: author || undefined,
      timestamp: timestamp || undefined,
      timestampRaw: timestampRaw || undefined,
      raw: (el.outerHTML || '').slice(0, 500)
    });
  }

  return items;
}

