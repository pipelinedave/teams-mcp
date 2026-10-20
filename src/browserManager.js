import { chromium } from 'playwright-core';
import path from 'path';
import os from 'os';
import fs from 'fs';
import { config } from './config.js';

class BrowserManager {
  constructor() {
    this.instances = new Map(); // tenant -> { context, page, isHeadless, lastUsed }
    // Idle-Unload (2026-10-20): Browser je Tenant nach N Minuten ohne Tool-Call
    // schliessen. Profil bleibt erhalten -> naechster Call startet lazy neu (Login bleibt).
    // 0 = deaktiviert (alter Zustand: Browser bleibt bis close/Restart offen).
    this.idleMinutes = Number.parseInt(process.env.TEAMS_MCP_BROWSER_IDLE_MIN ?? '30', 10);
    this.idleTimer = null;
    this.startIdleReaper();
  }

  startIdleReaper() {
    if (!Number.isFinite(this.idleMinutes) || this.idleMinutes <= 0) return;
    const idleMs = this.idleMinutes * 60_000;
    this.idleTimer = setInterval(() => { void this.reapIdle(idleMs); }, 60_000);
    if (typeof this.idleTimer.unref === 'function') this.idleTimer.unref();
  }

  async reapIdle(idleMs) {
    const now = Date.now();
    for (const [t, inst] of [...this.instances.entries()]) {
      if (!inst?.context) continue;
      if (now - (inst.lastUsed || 0) < idleMs) continue;
      if (this.instances.get(t) !== inst) continue;
      // Vor close austragen: eine parallele ensureContext darf den NEUEN Eintrag
      // nicht durch das nachgelagerte delete verlieren.
      this.instances.delete(t);
      try {
        await inst.context.close();
      } catch (e) {}
      console.error(`[teams-mcp] idle-unload: Browser fuer Tenant '${t}' nach ${this.idleMinutes} min Inaktivitaet geschlossen`);
    }
  }

  normalizeTenant(tenant) {
    const raw = (!tenant || tenant === 'all') ? config.defaultTenant : tenant;
    if (!raw) return config.defaultTenant || 'adesso';
    const t = String(raw).toLowerCase().trim();
    const realms = config.tenantRealms;
    if (realms[t]) return t; // exakter konfigurierter Key
    const keys = Object.keys(realms);
    const hit = keys.find(k => t.includes(k) || k.includes(t));
    return hit || t; // org-neutral: übergebener Wert zählt
  }

  realm(tenant) {
    const t = this.normalizeTenant(tenant);
    if (config.tenantRealms[t]) return config.tenantRealms[t];
    // Org-neutral: Falls Tenant wie ein Realm aussieht, direkt nutzen; sonst onmicrosoft.com-Konvention.
    return t.includes('.') || t.includes('onmicrosoft.com') ? t : `${t}.onmicrosoft.com`;
  }

  getProfileDir(tenant) {
    const t = this.normalizeTenant(tenant);
    return path.join(config.profileBase, `.teams-browser-profile-${t}`);
  }

  cleanLockFiles(profileDir) {
    try {
      const files = ['SingletonLock', 'SingletonSocket', 'SingletonCookie'];
      for (const file of files) {
        const p = path.join(profileDir, file);
        if (fs.existsSync(p)) {
          fs.unlinkSync(p);
        }
      }
    } catch (e) {}
  }

  launchOptions() {
    const opts = {
      viewport: { width: 1366, height: 850 },
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36 Edg/133.0.0.0',
      // Baustein B: Playwright-native Downloads erlauben (Anhänge aus Nachrichten)
      acceptDownloads: true,
      env: {
        ...process.env,
        DISPLAY: process.env.DISPLAY || ':0',
        WAYLAND_DISPLAY: process.env.WAYLAND_DISPLAY || 'wayland-0',
        XDG_RUNTIME_DIR: process.env.XDG_RUNTIME_DIR || '/mnt/wslg/runtime-dir',
        PULSE_SERVER: process.env.PULSE_SERVER || '/mnt/wslg/PulseServer'
      },
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-blink-features=AutomationControlled'
      ]
    };
    if (process.env.DISPLAY || process.env.WAYLAND_DISPLAY) {
      opts.args.push('--ozone-platform=wayland');
    }
    return opts;
  }

  async ensureContext(tenant = '', headless = true) {
    const t = this.normalizeTenant(tenant);
    const existing = this.instances.get(t);

    if (existing && existing.context) {
      if (existing.isHeadless !== headless) {
        await this.close(t);
      } else {
        try {
          if (existing.page && !existing.page.isClosed()) {
            existing.lastUsed = Date.now(); // Idle-Timer zuruecksetzen
            return { context: existing.context, page: existing.page };
          }
        } catch (e) {}
      }
    }

    const profileDir = this.getProfileDir(t);
    if (!fs.existsSync(profileDir)) {
      fs.mkdirSync(profileDir, { recursive: true });
    }

    // Retry gegen SingletonLock-Kollision (z.B. wenn eine parallele MCP-Instanz
    // dasselbe Profil gerade kurz nutzt und beendet). Max 5 Versuche / Backoff.
    const launchCfg = {
      headless: headless,
      ...this.launchOptions()
    };
    if (config.chromePath) launchCfg.executablePath = config.chromePath;

    let context = null;
    let lastError = null;
    for (let attempt = 1; attempt <= 5; attempt++) {
      this.cleanLockFiles(profileDir);
      try {
        context = await chromium.launchPersistentContext(profileDir, launchCfg);
        break;
      } catch (e) {
        lastError = e;
        if (!String(e.message).includes('ProcessSingleton') && !String(e.message).includes('SingletonLock')) {
          throw e; // kein Lock-Problem -> sofort werfen
        }
        await new Promise(r => setTimeout(r, 1000 * attempt));
      }
    }

    if (!context) {
      throw new Error(`Profil für tenant '${t}' ist gerade von einer anderen Instanz belegt (${lastError?.message}). Bitte später erneut versuchen oder die andere Session beenden.`);
    }

    const pages = context.pages();
    const page = pages.length > 0 ? pages[0] : await context.newPage();
    this.instances.set(t, { context, page, isHeadless: headless, lastUsed: Date.now() });
    return { context, page };
  }

  async openLoginWindow(tenant = '') {
    const t = this.normalizeTenant(tenant);
    await this.close(t);

    const { page } = await this.ensureContext(t, false);
    const realm = this.realm(t);
    const targetUrl = `https://teams.microsoft.com/v2/?realm=${encodeURIComponent(realm)}`;
    await page.goto(targetUrl, { waitUntil: 'domcontentloaded' });
    return page;
  }

  async close(tenant = 'all') {
    if (tenant === 'all') {
      for (const [t, inst] of this.instances.entries()) {
        try {
          if (inst.context) await inst.context.close();
        } catch (e) {}
      }
      this.instances.clear();
      return;
    }

    const t = this.normalizeTenant(tenant);
    const inst = this.instances.get(t);
    if (inst && inst.context) {
      try {
        await inst.context.close();
      } catch (e) {}
      this.instances.delete(t);
    }
  }
}

export const browserManager = new BrowserManager();
