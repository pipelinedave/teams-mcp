import fs from 'fs';
import path from 'path';
import TurndownService from 'turndown';
import { browserManager } from './browserManager.js';
import { config } from './config.js';
import { speakerTracker } from './speakerTracker.js';
import { activityClient } from './activityClient.js';
import { analyzeItems, buildReportText, writeReport, CATEGORIES } from './activityAnalyzer.js';

const turndown = new TurndownService({
  headingStyle: 'atx',
  hr: '---',
  bulletListMarker: '-'
});
turndown.remove(['script', 'style', 'head']);

export class TeamsClient {
  async getPage(tenant = '', headless = true) {
    const t = browserManager.normalizeTenant(tenant);
    if (!t) {
      throw new Error("Es wurde kein 'tenant' angegeben. Bitte den Tenant-Namen bzw. Realm mitgeben (z.B. teams_list_chats({ tenant: 'meine-org.onmicrosoft.com' })).");
    }
    const { page } = await browserManager.ensureContext(t, headless);
    const url = page.url();

    if (!url.includes('teams.microsoft.com') && !url.includes('login.microsoftonline.com') && !url.includes('teams.cloud.microsoft')) {
      await page.goto(`https://teams.microsoft.com/v2/?realm=${encodeURIComponent(browserManager.realm(t))}`, {
        waitUntil: 'domcontentloaded',
        timeout: 30000
      });
      // Warten, bis Teams geladen ist ODER die Session auf die Login-Seite abgelaufen ist
      await Promise.race([
        page.waitForSelector('[data-tid="app-bar-chat"], [role="treeitem"], div[role="main"]', { timeout: 20000 }).catch(() => null),
        page.waitForURL(/login\.microsoftonline\.com/, { timeout: 20000 }).catch(() => null)
      ]).catch(() => null);
    }

    const currentUrl = page.url();
    if (currentUrl.includes('login.microsoftonline.com')) {
      throw new Error(`Nicht in Teams (${t}) angemeldet. Bitte führe teams_login({ tenant: "${t}" }) in einem sichtbaren Fenster aus.`);
    }
    return page;
  }

  async checkStatus(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    try {
      const page = await this.getPage(t, true);
      const url = page.url();

      if (url.includes('login.microsoftonline.com')) {
        return {
          authenticated: false,
          tenant: t,
          currentUrl: url,
          message: `Nicht in Teams (${t}) angemeldet. Bitte führe 'teams_login({ tenant: "${t}" })' aus.`
        };
      }

      await page.waitForSelector('#main, div[role="main"], [role="treeitem"], [data-tid="app-bar-chat"]', { timeout: 12000 }).catch(() => null);

      const title = await page.title();
      const isLoaded = title.toLowerCase().includes('teams') || page.url().includes('teams.microsoft.com') || page.url().includes('teams.cloud.microsoft');

      return {
        authenticated: isLoaded && !url.includes('login.microsoftonline.com'),
        tenant: t,
        currentUrl: page.url(),
        title: title,
        message: isLoaded && !url.includes('login.microsoftonline.com')
          ? `Microsoft Teams (${t}) ist angemeldet und einsatzbereit.`
          : `Seite lädt noch oder erfordert Anmeldung ('teams_login({ tenant: "${t}" })').`
      };
    } catch (err) {
      return {
        authenticated: false,
        tenant: t,
        error: err.message
      };
    }
  }

  async openChatView(page) {
    const chatBtn = await page.$('button[data-tid="app-bar-chat"], button[aria-label*="Chat"], a[aria-label*="Chat"]');
    if (chatBtn) {
      await chatBtn.click().catch(() => null);
      await page.waitForTimeout(2000);
    }
  }

  // Liefert die gefilterte Liste der Chat-Treeitems (gleiche Logik wie listChats),
  // damit chat_index- und chat_name-Lookups konsistent auf denselben Items arbeiten.
  async getFilteredChatRows(page) {
    await this.openChatView(page);
    await page.waitForSelector('[role="treeitem"]', { timeout: 15000 }).catch(() => null);
    const rows = await page.$$('[role="treeitem"]');

    const filtered = [];
    const ignoredTitles = new Set(['copilot', 'quick views', 'drafts', 'favorites', 'chats', 'teams', 'channels', 'unread', 'more options', 'mentions', 'followed threads', 'discover', 'see more', 'see all channels', 'see all your teams', 'communities', 'join communities']);

    for (const row of rows) {
      const text = (await row.innerText())?.trim() || '';
      if (!text) continue;

      const lines = text.split('\n').map(s => s.trim()).filter(Boolean);
      const firstLine = lines[0] || '';
      if (ignoredTitles.has(firstLine.toLowerCase())) continue;

      if (lines.length === 1 && (firstLine.includes('Favorites') || firstLine.includes('Chats') || firstLine.includes('Quick views') || firstLine.includes('d.velop') || firstLine.startsWith('See'))) {
        continue;
      }

      // Kanal-/Gruppen-Sektionen ohne konkreten Chat (z.B. "General", "See all channels")
      if (firstLine.toLowerCase() === 'general' || firstLine.toLowerCase().startsWith('see ')) continue;

      filtered.push({ row, title: firstLine, text });
    }

    return filtered;
  }

  // Wählt einen Chat aus der gefilterten Liste. Bevorzugt exakten Titel-Match,
  // fällt auf Präfix-/Wort-Start-Match zurück (statt blindem Substring 'includes',
  // der z.B. "Yannick" fälschlich auf "Yannick und Kollegen" mappen kann).
  pickChatRow(filtered, chatName, chatIndex) {
    if (typeof chatIndex === 'number' && chatIndex >= 0 && chatIndex < filtered.length) {
      return { row: filtered[chatIndex].row, title: filtered[chatIndex].title };
    }
    if (!chatName) return null;
    const term = chatName.toLowerCase().trim();
    if (!term) return null;

    // 1) exakter Titel-Match
    let exact = filtered.find(c => c.title.toLowerCase() === term);
    if (exact) return { row: exact.row, title: exact.title };

    // 2) Titel beginnt mit dem Suchbegriff (z.B. "Yannick" -> "Yannick Bülter")
    let prefix = filtered.find(c => c.title.toLowerCase().startsWith(term));
    if (prefix) return { row: prefix.row, title: prefix.title };

    // 3) Fallback: Token-Match (alle Wörter des Begriffs kommen im Titel vor)
    const tokens = term.split(/\s+/).filter(Boolean);
    if (tokens.length > 0) {
      const tokenMatch = filtered.find(c => {
        const t = c.title.toLowerCase();
        return tokens.every(tok => t.includes(tok));
      });
      if (tokenMatch) return { row: tokenMatch.row, title: tokenMatch.title };
    }

    return null;
  }

  async listChats(tenant = '', limit = 15) {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const filtered = await this.getFilteredChatRows(page);

    const chats = filtered.slice(0, limit).map((chat, idx) => ({
      index: idx,
      title: chat.title,
      preview: chat.text.split('\n').map(s => s.trim()).filter(Boolean).slice(1).join(' | ') || undefined
    }));

    return {
      tenant: t,
      count: chats.length,
      chats: chats
    };
  }

  // Liefert den Titel des aktuell geöffneten Chats. Bevorzugt den Chat-Pane-Header
  // aus dem DOM (stabil, z.B. [data-tid="chat-title"]), fällt auf page.title() zurück.
  async getActiveChatTitle(page) {
    try {
      const domTitle = await page.evaluate(() => {
        // Chat-Header im Message-Pane (Teams/Fluent): [data-tid="chat-title"] ist am stabilsten
        const sels = [
          '[data-tid="chat-title"]',
          '[data-tid="chat-header-title"]',
          '[data-tid="entity-header"] [role="heading"]',
          '[data-tid="thread-pane-header"]',
          'header [aria-level]'
        ];
        for (const sel of sels) {
          const el = document.querySelector(sel);
          if (el && el.innerText && el.innerText.trim() && el.innerText.trim().length < 60) {
            const t = el.innerText.trim();
            if (t && !/^Message List$/i.test(t) && !/^Chat$/i.test(t)) {
              return t;
            }
          }
        }
        return null;
      });
      if (domTitle) return domTitle;
    } catch (e) {}

    // Fallback: page.title() -> "Chat | David Hallmann (You) | Microsoft Teams"
    const raw = (await page.title()).replace(' | Microsoft Teams', '').replace(/^\(\d+\)\s*/, '');
    const parts = raw.split(' | ').map(s => s.trim()).filter(Boolean);
    // Zweites Segment ist meist der Chat-Name (erste = "Chat"), sonst das letzte sichtbare
    if (parts.length >= 2 && /^chat$/i.test(parts[0])) return parts[1];
    return parts[0] || raw;
  }

  // Gemeinsamer Message-Scraper für Chats UND Kanäle (Baustein A/B): extrahiert die
  // im Message-Pane sichtbaren Nachrichten (Autor, Zeitstempel, Text, Medien, Anhänge).
  // Die Blob-Medien werden im Browser zu dataURLs aufgelöst und anschließend in Node
  // als Dateien persistiert. contextToken bestimmt den Dateinamen-Prefix der Medien.
  async _scrapeMessages(page, { max = 20, tenant = '', contextToken = 'chat' } = {}) {
    const messagesData = await page.evaluate(async ({ max, selfName, helpers, blobTimeoutMs }) => {
      const toIso = new Function('return (' + helpers.toIso + ')')();
      const collectAttachments = new Function('return (' + helpers.collectAttachments + ')')();
      const collectReactions = new Function('return (' + helpers.collectReactions + ')')();
      const isReactionArtifact = new Function('return (' + helpers.isReactionArtifact + ')')();
      const msgs = [];
      // Chat-DOM: chat-pane-message / fui-ChatMessage__body. Kanal-DOM (neues Teams-Web,
      // Befund 01.10.2026): channel-pane-message-Rows mit message-body + post-message-subheader.
      const msgBodies = Array.from(document.querySelectorAll('[data-tid="chat-pane-message"], [data-tid="message-body"], .fui-ChatMessage__body, .fui-ChatMyMessage__body'))
        // Double-Counting verhindern: message-body innerhalb eines Chat-Message-Containers überspringen
        .filter((el) => el.getAttribute('data-tid') !== 'message-body' ||
          !el.closest('[data-tid="chat-pane-message"], .fui-ChatMessage, .fui-ChatMyMessage'));
      const slice = msgBodies.slice(-max);

      // Message-Container (Row) zu einem Body — Anhänge/Bilder können neben dem Text-Body liegen
      const msgRowFor = (body) => body.closest('[data-tid="channel-pane-message"], .fui-ChatMessage, .fui-ChatMyMessage, [role="listitem"], [role="group"]') || body;

      for (const body of slice) {
        const isMe = String(body.className || '').includes('ChatMyMessage');

        // --- Autorenname (Spec-Befund 4) ---
        // ROOT-CAUSE (Live-Befund 09.10.2026, Chat "Eilers, Birgit"): das Body-Element
        // [data-tid="chat-pane-message"] traegt SELBST role="group". Die bisherige Logik
        //   body.closest('[role="group"], ...')   -> lieferte deshalb das BODY-Element
        //   parent.querySelector('[data-tid="message-author-name"]')  -> blieb im Body.
        // Eigene Nachrichten tragen den Namen inline im Body (funktionierte nur zufaellig),
        // bei FREMDEN Nachrichten liegt der Header aber als Geschwister VOR dem Body im
        // .fui-ChatMessage-Wrapper -> nie gefunden -> "Unbekannt".
        // Fix: eskalierende Suche Body -> Vorfahren, begrenzt an der Nachrichtengrenze
        // [data-tid="chat-pane-item"] (darueber hiinaus wuerde der Autor der NEXT-Meldung
        // gefunden). Reihenfolge: (1) inline im Body, (2) Vorfahren, (3) Geschwister-Notfall.
        const AUTHOR_SEL = '[data-tid="message-author-name"], [data-tid*="author-name" i], [class*="ChatMessage__author" i]';
        const readAuthor = (el) => {
          if (!el) return '';
          const t = String(el.innerText || '').replace(/\s+/g, ' ').trim();
          return t ? t.slice(0, 80) : '';
        };

        let author = readAuthor(body.querySelector(AUTHOR_SEL));

        if (!author) {
          let anc = body.parentElement;
          for (let i = 0; i < 6 && anc && !author; i++) {
            author = readAuthor(anc.querySelector(AUTHOR_SEL));
            if (anc.getAttribute('data-tid') === 'chat-pane-item') break; // Nachrichtengrenze
            anc = anc.parentElement;
          }
        }

        if (!author) {
          let prev = body.previousElementSibling;
          while (prev && !author) {
            const a = prev.querySelector?.(AUTHOR_SEL) || (prev.getAttribute?.('data-tid') === 'message-author-name' ? prev : null);
            if (a) { author = readAuthor(a); if (author) break; }
            prev = prev.previousElementSibling;
          }
        }

        // Eigenes Konto: Anzeigename aus dem DOM hat Vorrang (echter Name statt "Ich"),
        // selfName bleibt der Fallback, wenn das DOM nichts hergibt.
        if (!author && isMe) author = selfName;

        // Reaktionen sind reine UI-Elemente und gehoeren NICHT in den Nachrichten-Text.
        // Live-Befund 08.10.2026: bei 1:1-Chats liegt der Reaktions-Container INNERHALB
        // des [data-tid="chat-pane-message"]-Elements, das zugleich der Text-Body ist
        // (msgRowFor faellt auf `body` zurueck) -> ohne Filter landet das Emoji als
        // ![✅](https://...emoticons/2705_whiteheavycheckmark/...) im Markdown.
        // Deshalb vor der Markdown-Wandlung die Reaktions-Knoten aus einem Klon entfernen.
        const contentHtml = (() => {
          try {
            const clone = body.cloneNode(true);
            const nodes = clone.querySelectorAll('*') || [];
            for (const n of nodes) {
              const cn = typeof n.className === 'string' ? n.className : '';
              if (isReactionArtifact({ className: cn, dataTid: n.getAttribute('data-tid') })) n.remove();
            }
            return clone.innerHTML || '';
          } catch {
            return body.innerHTML || '';
          }
        })();

        const html = contentHtml;
        const text = body.innerText?.trim() || '';
        if (!text && !html) continue;

        // Timestamp-Extraktion mit mehreren Strategien (robust gegen Teams/UI-Variationen).
        const mid = body.getAttribute('data-mid') || '';
        let timestamp = null;
        let timestampRaw = null;

        // Strategie 1: data-mid ist bei Teams-Web oft ein Unix-Millis-Timestamp.
        if (mid) {
          const isoFromMid = toIso(mid);
          if (isoFromMid) {
            timestamp = isoFromMid;
            timestampRaw = mid;
          }
        }

        // Strategie 2: dediziertes Timestamp-Element (aria-labelledby referenziert "timestamp-{mid}").
        if (!timestamp && mid) {
          const tsEl = document.getElementById(`timestamp-${mid}`);
          if (tsEl) {
            const dT = tsEl.getAttribute('datetime') || tsEl.getAttribute('title') || '';
            if (dT) {
              const iso = toIso(dT);
              if (iso) { timestamp = iso; timestampRaw = dT || (tsEl.innerText || tsEl.textContent || '').trim(); }
            } else {
              timestampRaw = (tsEl.innerText || tsEl.textContent || '').trim();
            }
          }
        }

        // Strategie 3: internes <time datetime> im Body.
        if (!timestamp) {
          const timeEl = body.querySelector('time[datetime]');
          if (timeEl) {
            const iso = toIso(timeEl.getAttribute('datetime'));
            if (iso) { timestamp = iso; timestampRaw = timeEl.getAttribute('datetime'); }
          }
        }

        // Strategie 4 (Kanäle): Autor/Zeit aus post-message-subheader ("Messe, Adrian\n6/23 4:44 PM").
        // Subheader liegt als Geschwister neben dem message-body — daher über die Row suchen.
        const row = msgRowFor(body);
        if (!author && row) {
          const sub = row.querySelector('[data-tid="post-message-subheader"]');
          if (sub) {
            const lines = (sub.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
            if (lines.length > 0) author = lines[0];
          }
        }
        if (!timestamp && row) {
          const rowTime = row.querySelector('time[datetime]');
          if (rowTime) {
            const iso = toIso(rowTime.getAttribute('datetime'));
            if (iso) { timestamp = iso; timestampRaw = rowTime.getAttribute('datetime'); }
          }
          if (!timestamp) {
            const sub = row.querySelector('[data-tid="post-message-subheader"]');
            if (sub) {
              const lines = (sub.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
              if (lines.length > 1) timestampRaw = lines[lines.length - 1];
            }
          }
        }

        msgs.push({
          author: author || 'Unbekannt',
          isOwn: isMe || undefined,
          html: html,
          text: text,
          mid: mid || undefined,
          timestamp: timestamp || undefined,
          timestampRaw: timestampRaw || undefined,
          attachments: collectAttachments(row),
          reactions: collectReactions(row) || []
        });
      }

      // Nur Nachrichten-Bilder behandeln: Teams nutzt blob:-URLs auch für flüchtige UI-Icons
      // (Buttons, Avatare, Reaktionen/Sticker). Diese dürfen NICHT als Nachrichten-Medien gelten.
      const isUiControl = (el) => {
        let cur = el.parentElement;
        while (cur && cur !== document.body) {
          const tag = cur.tagName;
          const cls = typeof cur.className === 'string' ? cur.className : '';
          if (tag === 'BUTTON' || tag === 'A' || cur.getAttribute('role') === 'button' ||
              cls.includes('fui-Button') || cls.includes('fui-Avatar') || cls.includes('fui-CounterBadge')) {
            return true;
          }
          if (/react|Reaction|sticker|Picker/i.test(cls)) return true;
          cur = cur.parentElement;
        }
        return false;
      };
      const isInMessageArea = (el) => {
        let cur = el.parentElement;
        while (cur && cur !== document.body) {
          const cls = typeof cur.className === 'string' ? cur.className : '';
          if (cls.includes('fui-ChatMessage') || cls.includes('fui-ChatMyMessage') || cur.getAttribute('role') === 'listitem') return true;
          cur = cur.parentElement;
        }
        return false;
      };

      // Je Nachricht die blob:-URLs aus dem gesamten Message-Container sammeln (Bilder können
      // als Attachment neben dem reinen Text-Body liegen). Reine UI-Icons werden ausgefiltert.
      const rowBlobUrls = msgs.map((m, i) => {
        const row = msgRowFor(slice[i]);
        if (!row) return [];
        return Array.from(row.querySelectorAll('img[src^="blob:"]'))
          .filter((img) => !isUiControl(img))
          .map((img) => img.getAttribute('src') || img.src)
          .filter(Boolean);
      });

      // Inline-Screenshots/Bilder: blob:-URLs sind nur in der laufenden Browser-Session gültig.
      // Sie müssen HIER im Browser per fetch() + FileReader zu dataURLs aufgelöst werden
      // (Node kann Blob-URLs nicht lesen). 3s-Timeout pro Bild, Fehler blockieren nie.
      const blobData = {};
      try {
        const blobImgs = Array.from(document.querySelectorAll('img[src^="blob:"]'))
          .filter((img) => !isUiControl(img) && isInMessageArea(img));
        const urls = Array.from(new Set(blobImgs.map((img) => img.getAttribute('src') || img.src).filter(Boolean)));
        await Promise.all(urls.map(async (url) => {
          let dataUrl = null;
          try {
            const ctrl = new AbortController();
            const timer = setTimeout(() => ctrl.abort(), blobTimeoutMs);
            try {
              const res = await fetch(url, { signal: ctrl.signal, cache: 'no-store' });
              if (res && res.ok) {
                const blob = await res.blob();
                dataUrl = await new Promise((resolve) => {
                  const fr = new FileReader();
                  fr.onload = () => resolve(fr.result || null);
                  fr.onerror = () => resolve(null);
                  fr.readAsDataURL(blob);
                });
              }
            } finally {
              clearTimeout(timer);
            }
          } catch { dataUrl = null; }
          blobData[url] = dataUrl;
        }));
      } catch { /* Blob-Extraktion scheitert nie fatal: Nachrichten bleiben erhalten */ }

      return { msgs, blobData, rowBlobUrls };
    }, { max, selfName: config.selfName, helpers: { toIso: messageTimeToIso.toString(), collectAttachments: collectAttachmentsFromRow.toString(), collectReactions: collectReactionsFromRow.toString(), isReactionArtifact: isReactionArtifact.toString() }, blobTimeoutMs: 3000 });

    const { msgs, blobData, rowBlobUrls } = messagesData;

    // Persistiert erfolgreich aufgelöste Blobs als lokale Dateien und liefert
    // eine Map { blobUrl -> file:///Pfad }, um blob:-URLs im HTML/Markdown zu ersetzen.
    const blobUrlToPath = {};
    const mediaDir = ensureMediaDir(config.mediaDir);
    if (blobData && typeof blobData === 'object') {
      const chatToken = sanitizeFilename(contextToken || 'chat');
      const timeToken = Date.now();
      let index = 0;
      for (const [blobUrl, dataUrl] of Object.entries(blobData)) {
        if (!dataUrl) continue; // fehlgeschlagener Blob -> nicht blockieren
        const filename = `${sanitizeFilename(tenant)}_${chatToken}_${timeToken}_${index}`;
        const res = saveBlobDataUrl(dataUrl, mediaDir, filename);
        if (res.saved) {
          blobUrlToPath[blobUrl] = `file://${res.filePath}`;
          index += 1;
        }
      }
    }

    // Bilder, die in einem Message-Container (Row) liegen, aber NICHT im reinen Text-Body:
    // werden als eigenständige Bild-Referenz ans Ende des Markdown angehängt.
    const appendRowMedia = (mIndex, content) => {
      const rowUrls = (rowBlobUrls && rowBlobUrls[mIndex]) || [];
      const lines = [];
      for (const url of rowUrls) {
        const p = blobUrlToPath[url];
        if (p && !content.includes(p)) lines.push(`![media](${p})`);
      }
      return lines.length ? `${content}\n\n${lines.join('\n')}` : content;
    };

    return msgs.map((m, i) => {
      const htmlWithMedia = replaceBlobUrlsInHtml(m.html, blobUrlToPath);
      let content = turndown.turndown(htmlWithMedia || m.text);
      content = appendRowMedia(i, content);
      return {
        author: m.author,
        isOwn: m.isOwn || false,
        content,
        mid: m.mid,
        timestamp: m.timestamp,
        timestampRaw: m.timestampRaw,
        attachments: m.attachments || [],
        reactions: m.reactions || [],
        mediaFiles: (rowBlobUrls && rowBlobUrls[i] || [])
          .map((url) => blobUrlToPath[url])
          .filter(Boolean)
      };
    });
  }

  async getMessages(tenant = '', { chatIndex, chatName, limit = 20, downloadAttachments = false, downloadLimit, downloadExtensions } = {}) {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const filtered = await this.getFilteredChatRows(page);

    if (filtered.length === 0) {
      throw new Error("Keine Chats in Teams gefunden.");
    }

    if (typeof chatIndex === 'number' || chatName) {
      const picked = this.pickChatRow(filtered, chatName, chatIndex);
      if (picked) {
        await picked.row.click();
        await page.waitForTimeout(2500);
      }
    }

    // Aktiven Chat-Titel aus dem Message-Pane verifizieren (stabil)
    const activeChat = await this.getActiveChatTitle(page);
    await page.waitForSelector('[data-tid="chat-pane-message"], [data-tid="message-body"], .fui-ChatMessage__body, .fui-ChatMyMessage__body', { timeout: 15000 }).catch(() => null);

    const messages = await this._scrapeMessages(page, {
      max: limit,
      tenant: t,
      contextToken: activeChat || chatName || 'chat'
    });

    let downloads = null;
    if (downloadAttachments) {
      downloads = await this.downloadMessageAttachments(page, {
        limit: downloadLimit,
        extensions: downloadExtensions,
        token: t
      });
    }

    return {
      tenant: t,
      chat: activeChat,
      count: messages.length,
      messages,
      ...(downloads ? { downloads } : {})
    };
  }

  async search(tenant = '', query, limit = 10) {
    if (!query) throw new Error("Suchbegriff (query) ist erforderlich.");
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    // Suchfeld finden - Teams Web ändert die DOM-Struktur; mehrere Wege probieren,
    // inkl. Öffnen der Suche per Tastaturkürzel (Ctrl+E / Ctrl+K / Alt+Shift+S).
    let searchInput = null;
    const selectors = [
      '#top-search-input',
      'input[aria-label*="Search"]',
      'input[placeholder*="Search"]',
      'input[aria-label*="Suche"]',
      'input[placeholder*="Suche"]',
      '[data-tid="searchbox"] input',
      'input[type="search"]'
    ];

    for (const sel of selectors) {
      searchInput = await page.$(sel);
      if (searchInput) break;
    }

    if (!searchInput) {
      // Suche per Tastatur öffnen und erneut versuchen
      for (const combo of ['Control+E', 'Control+K', 'Alt+Shift+S']) {
        try {
          await page.keyboard.press(combo);
          await page.waitForTimeout(1200);
        } catch (e) {}
        for (const sel of selectors) {
          searchInput = await page.$(sel);
          if (searchInput) break;
        }
        if (searchInput) break;
      }
    }

    if (!searchInput) {
      throw new Error("Suchfeld in Teams Web nicht gefunden (auch nicht nach Tastatur-Öffnung).");
    }

    await searchInput.click();
    await page.waitForTimeout(400);
    await searchInput.fill('');
    // Query per Clipboard-Paste einfügen (robust für Sonderzeichen/Um&laute)
    await page.evaluate((q) => { navigator.clipboard.writeText(q); }, query);
    await page.keyboard.press('ControlOrMeta+V');
    await page.waitForTimeout(400);
    await page.keyboard.press('Enter');

    await page.waitForTimeout(4000);
    await page.waitForSelector('div[data-tid*="search-result"], [role="listitem"], [data-tid*="searchResult"]', { timeout: 12000 }).catch(() => null);

    const results = await page.evaluate((max) => {
      const items = [];
      const rows = document.querySelectorAll('div[data-tid*="search-result"], [data-tid*="searchResult"], [role="listitem"]');
      for (const r of rows) {
        if (items.length >= max) break;
        const text = r.innerText?.trim() || '';
        if (text && text.length > 10) {
          items.push({
            preview: text.split('\n').filter(Boolean).join(' | ')
          });
        }
      }
      return items;
    }, limit);

    return {
      tenant: t,
      query: query,
      count: results.length,
      results: results
    };
  }

  // Stellt sicher, dass die Team-/Kanal-Struktur sichtbar ist. Zwei UI-Varianten:
  // (a) Neues Teams-Web: "Teams and channels"-Sektion in der Chat-Rail — ggf. aufklappen.
  // (b) Ältere UI: eigenes App-Bar-Icon (data-tid="app-bar-teams") — darauf klicken.
  async openTeamsView(page) {
    const teamsButton = await page.$('[data-tid="app-bar-teams"], button[aria-label="Teams"], a[aria-label="Teams"]');
    if (teamsButton) {
      await teamsButton.click().catch(() => null);
      await page.waitForTimeout(2500);
    }
    await page.waitForSelector('[role="treeitem"]', { timeout: 15000 }).catch(() => null);

    // Sektion "Teams and channels" ggf. aufklappen (sonst sind Teams/Kanäle nicht im Baum)
    try {
      const scan = await this.scanTeamsTree(page);
      const section = scan.items.find(it =>
        /^teams( and channels| und kanäle)?$/i.test(it.title || '') && it.expanded === 'false'
      );
      if (section && scan.handles[section.idx]) {
        await scan.handles[section.idx].click().catch(() => null);
        await page.waitForTimeout(2000);
      }
    } catch (e) { /* Baumscan nicht kritisch — openChannel scannt erneut */ }
  }

  // Scannt den Team-/Kanal-Baum der linken Rails. WICHTIG für die Index-Alignment:
  // items[] enthält ALLE Knoten (auch ohne Titel) in DOM-Reihenfolge — identisch zur
  // ElementHandle-Liste von page.$$(SEL). Ein Klick erfolgt über items[i].idx auf
  // handles[idx]. Level fällt auf die Tiefe verschachtelter Treeitem-Vorfahren zurück.
  async scanTeamsTree(page) {
    const SEL = '[role="treeitem"], [data-tid="team-channel-item"]';
    const handles = await page.$$(SEL);
    const items = await page.evaluate((sel) => {
      const nodes = Array.from(document.querySelectorAll(sel));
      return nodes.map((el) => {
        const text = (el.innerText || '').replace(/\u00a0/g, ' ').trim();
        let level = parseInt(el.getAttribute('aria-level') || '', 10);
        if (!Number.isFinite(level)) {
          // Fallback: Tiefe über verschachtelte Treeitem-Vorfahren schätzen
          let depth = 1;
          let anc = el.parentElement;
          while (anc) {
            if (anc.getAttribute && anc.getAttribute('role') === 'treeitem') depth++;
            anc = anc.parentElement;
          }
          level = depth;
        }
        return {
          title: text ? text.split('\n')[0].trim() : '',
          level,
          expanded: el.getAttribute('aria-expanded'),
          selected: el.getAttribute('aria-selected') === 'true' || el.getAttribute('aria-current') === 'true',
          tid: el.getAttribute('data-tid') || undefined
        };
      });
    }, SEL);
    const withIdx = items.map((it, idx) => ({ ...it, idx }));
    return { handles, items: withIdx };
  }

  async listTeams(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    await this.openTeamsView(page);
    const { items } = await this.scanTeamsTree(page);

    const flat = items
      .filter(it => it.title)
      .map(it => ({ title: it.title, level: String(it.level) }));
    const teams = buildTeamsStructure(items);

    return {
      tenant: t,
      count: flat.length,
      teams,
      channels: flat
    };
  }

  // Navigiert App-Bar -> Team -> Kanal und verifiziert das Ziel (Safety-Net analog
  // zum Send-Flow): Liefert { team, channel, verify } oder wirft einen klaren Fehler
  // mit den tatsächlich sichtbaren Einträgen (kein Silent-Fehlresolvenz wie chat_index).
  async openChannel(page, tenant, teamName, channelName) {
    await this.openTeamsView(page);
    let scan = await this.scanTeamsTree(page);

    const teamItem = findTeamInTree(scan.items, teamName);
    if (!teamItem) {
      const known = scan.items.filter(it => it.title && !isIgnoredTreeTitle(it.title)).map(it => it.title);
      throw new Error(`Team '${teamName}' wurde im Teams-Baum (${tenant}) nicht gefunden. Sichtbare Einträge: ${known.slice(0, 25).join(' | ') || '(keine)'}`);
    }

    // Team anklicken (selektiert + expandiert die Kanalliste)
    const teamHandle = scan.handles[teamItem.idx];
    if (teamHandle) {
      await teamHandle.scrollIntoViewIfNeeded?.().catch(() => {});
      await teamHandle.click().catch(async () => {
        await teamHandle.click({ force: true }).catch(() => null);
      });
      await page.waitForTimeout(2000);
    }

    // Falls das Team zugeklappt war, erneut scannen und ggf. expandieren
    scan = await this.scanTeamsTree(page);
    const teamItem2 = findTeamInTree(scan.items, teamName) || { ...teamItem, idx: scan.items.findIndex(it => it.title === teamItem.title) };
    if (teamItem2 && teamItem2.expanded === 'false' && scan.handles[teamItem2.idx]) {
      await scan.handles[teamItem2.idx].click().catch(() => null);
      await page.waitForTimeout(2000);
      scan = await this.scanTeamsTree(page);
    }
    const teamFinal = findTeamInTree(scan.items, teamName) || teamItem2;

    const channelItem = findChannelInTree(scan.items, teamFinal, channelName);
    if (!channelItem) {
      const chans = collectChannelTitles(scan.items, teamFinal);
      throw new Error(`Kanal '${channelName}' wurde im Team '${teamFinal.title}' (${tenant}) nicht gefunden. Verfügbare Kanäle: ${chans.join(' | ') || '(keine sichtbaren)'}`);
    }

    const channelHandle = scan.handles[channelItem.idx];
    if (!channelHandle) {
      throw new Error(`Kanal '${channelItem.title}' gefunden, aber kein klickbares Element im Baum (${tenant}).`);
    }
    await channelHandle.scrollIntoViewIfNeeded?.().catch(() => {});
    await channelHandle.click().catch(async () => {
      await channelHandle.click({ force: true }).catch(() => null);
    });
    await page.waitForTimeout(2500);

    // Safety-Net: Verifizieren, dass der geöffnete Kanal dem Ziel entspricht.
    // Befund 01.10.2026 (neues Teams-Web): Kanal-Header = [data-tid="channelTitle-text"],
    // document.title = "(n) Chat | <Team> | <Kanal> | Microsoft Teams".
    const verify = await page.evaluate(() => {
      const selected = Array.from(document.querySelectorAll('[role="treeitem"][aria-selected="true"], [role="treeitem"][aria-current="true"]'))
        .map(el => (((el.innerText || '').split('\n')[0]) || '').trim())
        .filter(Boolean);
      const headerSels = [
        '[data-tid="channelTitle-text"]',
        '[data-tid="entity-header"]',
        '[data-tid="channel-title"]',
        '[data-tid="thread-pane-header"]',
        'header [aria-level]',
        '[data-tid="chat-title"]'
      ];
      let header = null;
      for (const s of headerSels) {
        const el = document.querySelector(s);
        const t = el && el.innerText ? el.innerText.trim() : '';
        if (t && t.length < 120) { header = t.split('\n')[0].trim() || t; break; }
      }
      return { selected, header, title: document.title || '' };
    });

    const chanNorm = channelItem.title.toLowerCase().trim();
    const teamNorm = teamFinal.title.toLowerCase().trim();
    const titleNorm = (verify.title || '').toLowerCase();
    const headerMatch = !!verify.header && verify.header.toLowerCase().includes(chanNorm);
    const selectedMatch = (verify.selected || []).some(s => s.toLowerCase().trim() === chanNorm);
    const titleMatch = titleNorm.includes(chanNorm) && titleNorm.includes(teamNorm);
    if (!headerMatch && !selectedMatch && !titleMatch) {
      throw new Error(`Sicherheitsnetz: Der geöffnete Kanal stimmt nicht mit dem Ziel überein. Ziel: '${channelItem.title}' (Team '${teamFinal.title}'), gefunden: Header='${verify.header || '(keiner)'}', Titel='${verify.title || '(keiner)'}', selektiert=[${(verify.selected || []).join(', ') || '(keins)'}]. Es wurden KEINE Kanal-Nachrichten gelesen.`);
    }

    return { team: teamFinal.title, channel: channelItem.title, verify: { header: verify.header, title: verify.title, selected: verify.selected } };
  }

  async getChannelMessages(tenant = '', { team, channel = 'Allgemein', limit = 20, downloadAttachments = false, downloadLimit, downloadExtensions } = {}) {
    if (!team || !String(team).trim()) {
      throw new Error("Team-Name (team_name) ist erforderlich, z.B. teams_get_channel_messages({ tenant: 'd-velop', team_name: 'Mein Team', channel_name: 'Allgemein' }).");
    }
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const nav = await this.openChannel(page, t, String(team).trim(), String(channel || 'Allgemein').trim());

    await page.waitForSelector('[data-tid="channel-pane-message"], [data-tid="message-body"], [data-tid="chat-pane-message"], .fui-ChatMessage__body, .fui-ChatMyMessage__body', { timeout: 15000 }).catch(() => null);

    const messages = await this._scrapeMessages(page, {
      max: limit,
      tenant: t,
      contextToken: `${nav.team}_${nav.channel}`
    });

    let downloads = null;
    if (downloadAttachments) {
      downloads = await this.downloadMessageAttachments(page, {
        limit: downloadLimit,
        extensions: downloadExtensions,
        token: t
      });
    }

    return {
      tenant: t,
      team: nav.team,
      channel: nav.channel,
      verified: nav.verify,
      count: messages.length,
      messages,
      ...(downloads ? { downloads } : {})
    };
  }

  // Baustein B: Lädt Anhänge aus den sichtbaren Nachrichten als lokale Dateien herunter.
  // Playwright-native Downloads (page.on('download') + saveAs) in config.downloadDir.
  // Zwei Kaskaden: (1) Download-Buttons an Attachment-Karten, (2) dateiartige Links via
  // Temp-Page mit ?download=1. Ein hängender Download wird nach timeoutMs übersprungen
  // (Warnung in skipped[]) — das Scraping blockiert nie.
  async downloadMessageAttachments(page, { limit = 5, extensions, timeoutMs = 25000, dir, token = 'teams' } = {}) {
    const exts = normalizeExtensionWhitelist(extensions);
    const downloadDir = ensureMediaDir(dir || config.downloadDir);
    const result = { dir: downloadDir, attempted: 0, saved: [], skipped: [] };

    const raceTimeout = (promise, label) => {
      let timer;
      const timeoutPromise = new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error(`Timeout nach ${timeoutMs}ms: ${label}`)), timeoutMs);
      });
      // No-op-Handler: Verhindert unhandled rejections, wenn die Gegenseite gewinnt
      // (der MCP-Server-Prozess dürfte bei einer späten Rejection nicht crashen).
      timeoutPromise.catch(() => {});
      if (promise && typeof promise.catch === 'function') promise.catch(() => {});
      return Promise.race([promise, timeoutPromise]).finally(() => clearTimeout(timer));
    };

    // --- Kaskade 0: Attachment-Karten über "More actions" -> Menü "Download" ---
    // Befund 01.10.2026: Karten (data-tid="file-attachment-grid") haben Hover-Buttons
    // ("Copy link", "More actions"); das Menü enthält den Download-Eintrag. Direkte
    // SharePoint-Links (:p:/r/) öffnen dagegen den Office-Viewer — kein Download-Event.
    let cardHandles = [];
    try {
      cardHandles = await page.$$('[data-tid="file-attachment-grid"], [data-tid*="attachment-card" i], [data-tid*="file-card" i]');
    } catch { cardHandles = []; }

    for (const card of cardHandles) {
      if (result.saved.length >= limit) break;
      const cardName = await card.evaluate((el) => ((el.innerText || '').split('\n')[0] || '').trim()).catch(() => '');
      if (exts !== null && !extensionAllowed(cardName, exts)) {
        if (cardName) result.skipped.push({ name: cardName, reason: 'Endung nicht in Whitelist' });
        continue;
      }
      if (!cardName) continue;

      result.attempted++;
      try {
        await card.scrollIntoViewIfNeeded().catch(() => {});
        await card.hover().catch(() => {});
        await page.waitForTimeout(400);

        // More-Actions-/Download-Button innerhalb der Karte (erscheint teils erst beim Hover)
        const actionBtn = await card.$([
          'button[aria-label*="more actions" i]',
          'button[aria-label*="weitere aktionen" i]',
          'button[aria-label*="mehr aktionen" i]',
          'button[aria-label*="herunterladen" i]',
          'button[aria-label*="download" i]',
          'button[title*="ownload" i]'
        ].join(', '));
        if (!actionBtn) {
          result.skipped.push({ name: cardName, reason: 'Kein More-Actions-/Download-Button in der Karte sichtbar' });
          continue;
        }

        const directPromise = page.waitForEvent('download', { timeout: timeoutMs });
        directPromise.catch(() => {}); // kein unhandled rejection bei Timeout
        await actionBtn.click({ timeout: 5000 });

        // Menü-Eintrag "Herunterladen"/"Download" suchen (Playwright :has-text ist case-insensitive)
        const menuItem = await page.waitForSelector(
          '[role="menuitem"]:has-text("Herunterladen"), [role="menuitem"]:has-text("Download")',
          { timeout: 4000 }
        ).catch(() => null);

        let download = null;
        if (menuItem) {
          const menuPromise = page.waitForEvent('download', { timeout: timeoutMs });
          menuPromise.catch(() => {});
          await menuItem.click({ timeout: 5000 });
          download = await menuPromise.then(d => d, () => null);
        } else {
          // Kein Menü — evtl. direkter Download durch den Button-Klick (kurz abwarten)
          download = await Promise.race([directPromise, new Promise(r => setTimeout(() => r(null), 2000))]);
        }
        await page.keyboard.press('Escape').catch(() => {});

        if (!download) {
          result.skipped.push({ name: cardName, reason: menuItem ? 'Download-Eintrag geklickt, aber kein Download ausgelöst' : 'Kein Download-Menüeintrag gefunden' });
          continue;
        }
        const saved = await raceTimeout(this._saveDownload(download, downloadDir, token), `saveAs(${cardName})`);
        result.saved.push(saved);
      } catch (e) {
        await page.keyboard.press('Escape').catch(() => {});
        result.skipped.push({ name: cardName || '(unbenannt)', reason: `Karten-Download fehlgeschlagen: ${String(e.message).slice(0, 120)}` });
      }
    }

    // --- Kaskade 1: direkte Download-Buttons in den sichtbaren Nachrichten ---
    const buttonSelectors = [
      '[data-tid*="attachment" i] button[aria-label*="erunterladen" i]',
      '[data-tid*="attachment" i] button[aria-label*="ownload" i]',
      '[data-tid*="attachment" i] button[title*="erunterladen" i]',
      '[data-tid*="attachment" i] button[title*="ownload" i]',
      '[role="listitem"] button[aria-label*="erunterladen" i]',
      '[role="listitem"] button[aria-label*="ownload" i]',
      '[role="group"] button[aria-label*="erunterladen" i]',
      '[role="group"] button[aria-label*="ownload" i]',
      'button[data-tid*="download" i]',
      'button[data-tid*="attachment-download" i]'
    ].join(', ');

    let buttons = [];
    try {
      buttons = await page.$$(buttonSelectors);
    } catch { buttons = []; }

    for (const btn of buttons) {
      if (result.saved.length >= limit) break;
      // Dateiname aus dem Karten-Kontext (Button-Label selbst enthält meist nur "Herunterladen")
      const cardName = await btn.evaluate((el) => {
        const card = el.closest('[data-tid*="attachment" i], [class*="attachment" i], [role="listitem"], [role="group"]');
        const label = (card && (card.getAttribute('aria-label') || card.innerText)) || '';
        return (label + '').trim().split('\n')[0].trim();
      }).catch(() => '');
      const name = extractFilenameFromText(cardName) || extractFilenameFromText(cardName && cardName.replace(/[\\/]/g, ' '));
      if (exts !== null && !name) {
        result.skipped.push({ name: cardName || '(unbenannt)', reason: 'Kein Dateiname erkennbar (Extension-Whitelist aktiv)' });
        continue;
      }
      if (exts !== null && !extensionAllowed(name, exts)) {
        result.skipped.push({ name, reason: 'Endung nicht in Whitelist' });
        continue;
      }

      result.attempted++;
      try {
        await btn.scrollIntoViewIfNeeded().catch(() => {});
        await btn.hover().catch(() => {});
        const downloadPromise = page.waitForEvent('download', { timeout: timeoutMs });
        downloadPromise.catch(() => {}); // kein unhandled rejection bei Timeout/Klick-Fehler
        await btn.click({ timeout: 5000 });
        const download = await downloadPromise;
        const saved = await raceTimeout(this._saveDownload(download, downloadDir, token), `saveAs(${name || 'anhang'})`);
        result.saved.push(saved);
      } catch (e) {
        result.skipped.push({ name: name || cardName || '(unbenannt)', reason: `Download-Button fehlgeschlagen: ${String(e.message).slice(0, 120)}` });
      }
    }

    // --- Kaskade 2: dateiartige Links (SharePoint) via Temp-Page mit &download=1 ---
    // Letzter Ausweg: :p:/r-Links öffnen meist den Office-Viewer (kein Download-Event);
    // für direkte /download=1-URLs funktioniert dieser Weg.
    if (result.saved.length < limit) {
      let links = [];
      try {
        links = await page.evaluate(() => {
          const NAME_RE = /([A-Za-z0-9\u00c0-\u024f][\w\u00c0-\u024f. ()[\]#&'`-]*\.[A-Za-z0-9]{2,5})(?=\s|$|\n)/;
          const out = [];
          const seen = new Set();
          for (const a of document.querySelectorAll('[role="listitem"] a[href], [role="group"] a[href], [data-tid*="attachment" i] a[href]')) {
            const href = a.getAttribute('href') || '';
            if (!/^https?:\/\//i.test(href)) continue;
            if (!/sharepoint|1drv|onedrive|personalcontent|download=1|\.pdf|\.docx?|\.xlsx?|\.pptx?|\.zip|\.txt|\.csv|\.png|\.jpe?g/i.test(href)) continue;
            let bare = href;
            try { bare = decodeURIComponent(href.split('?')[0]); } catch {}
            const m = NAME_RE.exec((a.innerText || '').trim()) || NAME_RE.exec(bare);
            const name = m ? m[1].trim() : null;
            const key = name + '|' + href;
            if (seen.has(key)) continue;
            seen.add(key);
            out.push({ name, href });
            if (out.length >= 20) break;
          }
          return out;
        });
      } catch { links = []; }

      for (const link of links) {
        if (result.saved.length >= limit) break;
        if (exts !== null && !extensionAllowed(link.name || link.href, exts)) {
          result.skipped.push({ name: link.name || link.href, reason: 'Endung nicht in Whitelist' });
          continue;
        }
        result.attempted++;
        let tmpPage = null;
        try {
          tmpPage = await page.context().newPage();
          const downloadPromise = tmpPage.waitForEvent('download', { timeout: timeoutMs });
          downloadPromise.catch(() => {}); // kein unhandled rejection bei Timeout
          await tmpPage.goto(urlForcedDownload(link.href), { waitUntil: 'commit', timeout: timeoutMs }).catch(() => {});
          const download = await downloadPromise.then(d => d, () => null);
          if (!download) {
            result.skipped.push({ name: link.name || link.href, reason: 'Kein Download ausgelöst (öffnete vermutlich den Office-Viewer)' });
            continue;
          }
          const saved = await raceTimeout(this._saveDownload(download, downloadDir, token), `saveAs(${link.name || 'anhang'})`);
          result.saved.push(saved);
        } catch (e) {
          result.skipped.push({ name: link.name || link.href, reason: `Link-Download fehlgeschlagen: ${String(e.message).slice(0, 120)}` });
        } finally {
          if (tmpPage) await tmpPage.close().catch(() => {});
        }
      }
    }

    if (result.saved.length === 0 && result.attempted === 0) {
      result.note = 'Keine herunterladbaren Anhänge in den sichtbaren Nachrichten gefunden (Karten ohne Download-Button/Datei-Link).';
    }
    return result;
  }

  // Speichert einen Playwright-Download kollisionsfrei im Zielverzeichnis.
  async _saveDownload(download, downloadDir, token = 'teams') {
    let suggested = 'anhang';
    try { suggested = download.suggestedFilename() || suggested; } catch {}
    const target = nextAvailablePath(downloadDir, suggested);
    await download.saveAs(target);
    const size = fs.existsSync(target) ? fs.statSync(target).size : 0;
    return {
      name: path.basename(target),
      suggestedName: suggested,
      path: target,
      file: `file://${target}`,
      size,
      token
    };
  }

  // Fügt Text per Clipboard-Paste statt keyboard.type ein.
  // Verhindert, dass Zeilenumbrüche (\n) in Teams als Enter/Ab-Senden interpretiert
  // werden und die Nachricht in mehrere Fragmente zersplittert.
  async insertTextViaPaste(page, text) {
    await page.evaluate((t) => {
      navigator.clipboard.writeText(t);
    }, text);
    await page.waitForTimeout(150);
    // Fokus im Compose-Feld + einfügen
    await page.keyboard.press('ControlOrMeta+A');
    await page.keyboard.press('ControlOrMeta+V');
    await page.waitForTimeout(400);
  }

  // Validiert und normalisiert Dateipfade für Anhänge
  validateAttachments(attachments) {
    if (!attachments) return [];
    const list = Array.isArray(attachments) ? attachments : [attachments];
    const resolved = [];
    for (const item of list) {
      if (!item || typeof item !== 'string') continue;
      const trimmed = item.trim();
      if (!trimmed) continue;
      const abs = path.resolve(trimmed);
      if (!fs.existsSync(abs)) {
        throw new Error(`Anhang-Datei nicht gefunden: "${trimmed}" (aufgelöst: "${abs}")`);
      }
      const stat = fs.statSync(abs);
      if (!stat.isFile()) {
        throw new Error(`Anhang ist keine reguläre Datei: "${trimmed}"`);
      }
      resolved.push(abs);
    }
    return resolved;
  }

  // Hängt eine oder mehrere lokale Dateien im Teams Chat-Compose-Bereich an
  async attachFiles(page, filePaths) {
    if (!filePaths || filePaths.length === 0) return;

    // 1. Kaskade: Existiert bereits ein input[type="file"] auf der Seite?
    const fileInputs = await page.$$('input[type="file"]');
    if (fileInputs.length > 0) {
      for (const input of fileInputs) {
        try {
          await input.setInputFiles(filePaths);
          await this.waitForAttachmentUpload(page, filePaths);
          return;
        } catch (e) {
          // Fallback zum nächsten Weg
        }
      }
    }

    // 2. Kaskade: Büroklammer / Attach-Button / Aktionen-Button im Compose-Footer
    const attachButtonSelectors = [
      // Spezifische Teams Test-IDs
      'button[data-tid="newMessageCommands-FilePicker"]',
      'button[data-tid*="FilePicker" i]',
      'button[data-tid*="file-picker" i]',
      'button[data-tid*="file-upload" i]',
      'button[data-tid*="attachment" i]',
      'button[data-tid="compose-attach-button"]',
      'button[data-tid="attach-button"]',
      'button[data-tid*="attach" i]',
      '[data-tid="chat-pane-compose-message-footer"] button[data-tid="attach-button"]',
      '[data-tid="chat-pane-compose-message-footer"] button[data-tid="compose-attach-button"]',
      '[data-tid="chat-pane-compose"] button[data-tid*="attach" i]',

      // Deutsche Labels (Teams Web DE)
      'button[aria-label*="anfügen" i]',
      'button[aria-label*="anhängen" i]',
      'button[aria-label*="anheften" i]',
      'button[aria-label*="Dateien" i]',
      'button[aria-label*="Datei" i]',
      'button[title*="anfügen" i]',
      'button[title*="anhängen" i]',
      'button[title*="anheften" i]',
      'button[title*="Datei" i]',

      // Englische Labels (Teams Web EN)
      'button[aria-label*="Attach" i]',
      'button[aria-label*="File" i]',
      'button[aria-label*="Upload" i]',
      'button[title*="Attach" i]',
      'button[title*="File" i]',

      // Icons (Paperclip / Büroklammer / Attach)
      'button:has(svg[data-icon-name*="Attach" i])',
      'button:has(svg[data-icon-name*="Paperclip" i])',
      'button:has(i[data-icon-name*="Attach" i])',
      'button:has(i[data-icon-name*="Paperclip" i])',

      // Modern Teams v2: "Aktionen und Apps" / "+" Button im Compose-Footer
      'button[aria-label*="Aktionen und Apps" i]',
      'button[aria-label*="Aktionen" i]',
      'button[aria-label*="Actions and apps" i]',
      'button[aria-label*="Add an action" i]',
      'button[aria-label*="Weitere Aktionen" i]',
      'button[data-tid*="action-overflow" i]',
      'button[data-tid*="actions-and-apps" i]',
      'button[data-tid="plus-button"]',
      'button[data-tid="expand-compose-actions-button"]'
    ];

    let attachBtn = null;
    for (const sel of attachButtonSelectors) {
      try {
        const btn = await page.$(sel);
        if (btn && await btn.isVisible().catch(() => true)) {
          attachBtn = btn;
          break;
        }
      } catch (e) {}
    }

    // Falls Selektoren keinen Treffer brachten: DOM-Scan über Toolbar-/Footer-Buttons
    if (!attachBtn) {
      const handle = await page.evaluateHandle(() => {
        const containers = Array.from(document.querySelectorAll('[role="toolbar"], [data-tid*="compose"], footer, [data-tid*="chat-pane"]'));
        for (const c of containers) {
          const btns = Array.from(c.querySelectorAll('button'));
          for (const b of btns) {
            const txt = (b.innerText || '').toLowerCase();
            const label = (b.getAttribute('aria-label') || '').toLowerCase();
            const title = (b.getAttribute('title') || '').toLowerCase();
            const tid = (b.getAttribute('data-tid') || '').toLowerCase();
            const props = `${txt} ${label} ${title} ${tid}`;
            if (
              props.includes('anfüg') ||
              props.includes('anhäng') ||
              props.includes('anheft') ||
              props.includes('attach') ||
              props.includes('datei') ||
              props.includes('file') ||
              props.includes('aktion') ||
              props.includes('action') ||
              props.includes('paperclip')
            ) {
              return b;
            }
          }
        }
        return null;
      });
      attachBtn = handle.asElement();
    }

    let fileChooser = null;
    if (attachBtn) {
      // Prüfen, ob Klick direkt den FileChooser öffnet
      const chooserPromise = page.waitForEvent('filechooser', { timeout: 2500 }).catch(() => null);
      await attachBtn.click().catch(() => null);
      fileChooser = await chooserPromise;

      if (!fileChooser) {
        // Möglicherweise hat sich ein Dropdown/Flyout-Menü geöffnet ("Von diesem Gerät hochladen")
        await page.waitForTimeout(600);
        const menuSelectors = [
          '[data-tid="upload-from-computer"]',
          '[data-tid*="upload-from-computer" i]',
          '[data-tid*="attach-upload-from-computer" i]',
          '[data-tid*="upload" i]',
          '[data-tid*="device" i]',
          '[role="menuitem"]:has-text("diesem Gerät")',
          '[role="menuitem"]:has-text("diesem Computer")',
          '[role="menuitem"]:has-text("Computer")',
          '[role="menuitem"]:has-text("this device")',
          '[role="menuitem"]:has-text("computer")',
          '[role="menuitem"]:has-text("Gerät")',
          '[role="menuitem"]:has-text("Upload")',
          '[role="menuitem"]:has-text("Dateien anfügen")',
          '[role="menuitem"]:has-text("Datei anfügen")',
          '[role="menuitem"]:has-text("Dateien hochladen")',
          '[role="menuitem"]:has-text("Datei hochladen")',
          'button:has-text("diesem Gerät")',
          'button:has-text("this device")',
          'button:has-text("Computer")',
          'button:has-text("Upload")'
        ];

        let menuOption = null;
        for (const sel of menuSelectors) {
          try {
            menuOption = await page.$(sel);
            if (menuOption && await menuOption.isVisible().catch(() => true)) break;
          } catch (e) {}
        }

        // Falls menuOption noch nicht gefunden: alle Menüeinträge per Text durchsuchen
        if (!menuOption) {
          const handle = await page.evaluateHandle(() => {
            const items = Array.from(document.querySelectorAll(
              '[role="menu"] [role="menuitem"], [role="menu"] button, [role="listbox"] [role="option"], div[role="menuitem"], div[class*="menu"] [role="menuitem"], div[class*="popover"] button, div[class*="flyout"] button'
            ));
            return items.find(el => {
              const txt = (el.innerText || el.textContent || el.getAttribute('aria-label') || '').toLowerCase();
              return txt.includes('gerät') || txt.includes('device') || txt.includes('computer') || txt.includes('upload') || txt.includes('hochladen') || txt.includes('datei');
            }) || null;
          });
          menuOption = handle.asElement();
        }

        if (menuOption) {
          const [uploadChooser] = await Promise.all([
            page.waitForEvent('filechooser', { timeout: 8000 }).catch(() => null),
            menuOption.click().catch(() => null)
          ]);
          fileChooser = uploadChooser;
        }
      }
    }

    if (fileChooser) {
      await fileChooser.setFiles(filePaths);
      await this.waitForAttachmentUpload(page, filePaths);
      return;
    }

    // 3. Kaskade: Eventuell wurde das input[type="file"] erst durch Interaktion ins DOM gehängt
    const lateInputs = await page.$$('input[type="file"]');
    if (lateInputs.length > 0) {
      for (const lateInput of lateInputs) {
        try {
          await lateInput.setInputFiles(filePaths);
          await this.waitForAttachmentUpload(page, filePaths);
          return;
        } catch (e) {}
      }
    }

    // 4. Kaskade: Drag & Drop Fallback via HTML5 DataTransfer auf das Compose-Feld
    try {
      const filesData = filePaths.map(fp => ({
        name: path.basename(fp),
        type: 'application/octet-stream',
        content: fs.readFileSync(fp).toString('base64')
      }));

      const dropSuccess = await page.evaluate(({ selector, files }) => {
        const target = document.querySelector(selector) || document.querySelector('div[role="textbox"]');
        if (!target) return false;

        const dt = new DataTransfer();
        for (const file of files) {
          const byteCharacters = atob(file.content);
          const byteNumbers = new Array(byteCharacters.length);
          for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
          }
          const byteArray = new Uint8Array(byteNumbers);
          const blob = new Blob([byteArray], { type: file.type });
          const domFile = new File([blob], file.name, { type: file.type });
          dt.items.add(domFile);
        }

        target.dispatchEvent(new DragEvent('dragenter', { dataTransfer: dt, bubbles: true, cancelable: true }));
        target.dispatchEvent(new DragEvent('dragover', { dataTransfer: dt, bubbles: true, cancelable: true }));
        target.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, bubbles: true, cancelable: true }));
        return true;
      }, { selector: '[data-tid="ckeditor"], div[role="textbox"], [contenteditable]', files: filesData });

      if (dropSuccess) {
        await this.waitForAttachmentUpload(page, filePaths);
        return;
      }
    } catch (dropErr) {
      // Ignorieren und unten verständlichen Fehler werfen
    }

    throw new Error("Dateianhang fehlgeschlagen: Kein Datei-Upload-Button, Datei-Input oder Drop-Ziel im Compose-Bereich ansprechbar.");
  }

  // Wartet auf das Fertigstellen des Uploads in Teams
  async waitForAttachmentUpload(page, filePaths = []) {
    await page.waitForTimeout(1500);

    // Auf Verschwinden von Progressbars warten
    try {
      const progressbars = page.locator('[data-tid*="compose"] [role="progressbar"], div[role="progressbar"], [data-tid*="progress"]');
      const count = await progressbars.count();
      if (count > 0) {
        await progressbars.first().waitFor({ state: 'detached', timeout: 60000 });
      }
    } catch (e) {}

    // Warten, bis Dateikarte(n) oder Anhang-Vorschau im Compose-Bereich sichtbar sind
    try {
      await page.waitForFunction(() => {
        const compose = document.querySelector('[data-tid*="compose"], footer');
        if (!compose) return true;
        const cards = compose.querySelectorAll('[data-tid*="attachment"], [data-tid*="file-card"], [class*="attachment"], [class*="fileCard"]');
        return cards.length > 0;
      }, { timeout: 8000 }).catch(() => null);
    } catch (e) {}

    // Sicherstellen, dass der Senden-Button aktiv / nicht disabled ist
    await page.waitForFunction(() => {
      const btn = document.querySelector('button[data-tid="newMessageCommands-send"], button[data-tid="send-message-button"], button[aria-label*="Send"], button[aria-label*="Senden"]');
      return btn && !btn.disabled && btn.getAttribute('aria-disabled') !== 'true';
    }, { timeout: 15000 }).catch(() => null);

    await page.waitForTimeout(800);
  }

  async inspectCompose(tenant = '') {
    const raw = (!tenant || tenant === 'all') ? config.defaultTenant : tenant;
    const t = browserManager.normalizeTenant(raw);
    const page = await this.getPage(t, true);

    const activeChat = await this.getActiveChatTitle(page);

    const inspectData = await page.evaluate(() => {
      const fileInputs = Array.from(document.querySelectorAll('input[type="file"]')).map(i => ({
        id: i.id || null,
        name: i.name || null,
        dataTid: i.getAttribute('data-tid') || null,
        multiple: i.multiple,
        accept: i.accept || null,
        visible: i.offsetWidth > 0 && i.offsetHeight > 0
      }));

      const composeFooter = document.querySelector('[data-tid="chat-pane-compose-message-footer"], [data-tid="chat-pane-compose"], footer, [role="region"][aria-label*="compose" i]');
      const toolbars = Array.from(document.querySelectorAll('[role="toolbar"], footer, [data-tid*="compose"]'));
      const allButtons = [];
      for (const tb of toolbars) {
        for (const b of tb.querySelectorAll('button')) {
          allButtons.push({
            dataTid: b.getAttribute('data-tid') || null,
            ariaLabel: b.getAttribute('aria-label') || null,
            title: b.getAttribute('title') || null,
            innerText: b.innerText?.trim() || '',
            disabled: b.disabled || b.getAttribute('aria-disabled') === 'true'
          });
        }
      }

      const fileCards = Array.from(document.querySelectorAll('[data-tid*="attachment"], [class*="Attachment"], [class*="file-card"], a[href], [role="button"]')).map(el => ({
        tag: el.tagName,
        text: el.innerText ? el.innerText.trim().slice(0, 200) : "",
        href: el.getAttribute("href"),
        ariaLabel: el.getAttribute("aria-label"),
        dataTid: el.getAttribute("data-tid"),
        className: el.className ? String(el.className) : ""
      })).filter(c => (c.text && c.text.includes("xlsx")) || (c.href && c.href.includes("xlsx")) || (c.ariaLabel && c.ariaLabel.includes("xlsx")));

      const lastMsg = document.querySelectorAll('[data-tid="chat-pane-message"], .fui-ChatMessage__body, .fui-ChatMyMessage__body');
      const lastMsgEl = lastMsg.length > 0 ? lastMsg[lastMsg.length - 1] : null;
      const lastMsgHtml = lastMsgEl ? lastMsgEl.innerHTML : null;

      return {
        hasComposeFooter: !!composeFooter,
        fileInputs,
        composeButtons: allButtons.slice(0, 30),
        fileCards,
        lastMsgHtml
      };
    });

    return {
      tenant: t,
      activeChat,
      ...inspectData
    };
  }

  // Startet einen neuen Chat über das To:-Feld (Adressbuch/GAL), falls kein bestehender Chat vorhanden ist
  async startNewChat(page, personName) {
    // 1. Zuerst sicherstellen, dass wir in der Chat-App sind (Ctrl+Shift+4)
    await page.keyboard.press('Control+Shift+4');
    await page.waitForTimeout(1000);

    // 2. Neuer Chat per Shortcut Alt+Shift+N oder Button
    const newMsgBtn = await page.$('button[aria-label*="New message" i], button[aria-label*="Neuer Chat" i]');
    if (newMsgBtn) {
      await newMsgBtn.click();
    } else {
      await page.keyboard.press('Alt+Shift+N');
    }
    await page.waitForTimeout(1500);

    // 3. To-Input finden
    const toInput = await page.$(
      'input[aria-label*="name" i], input[placeholder*="name" i], input[aria-label*="Namen" i], input[placeholder*="Namen" i], input[data-tid*="people-picker"]'
    );
    if (!toInput) return null;

    await toInput.click();
    await toInput.fill('');
    await toInput.type(personName, { delay: 60 });
    await page.waitForTimeout(2000);

    // 4. Suggestions abwarten
    const optionLocator = page.locator('[role="listbox"] [role="option"], [role="option"], [data-tid*="suggestion"]');
    await optionLocator.first().waitFor({ state: 'visible', timeout: 8000 }).catch(() => null);

    const count = await optionLocator.count();
    if (count === 0) return null;

    // Finde passende Option
    let targetOption = null;
    let pickedTitle = personName;
    const norm = personName.toLowerCase();
    for (let i = 0; i < count; i++) {
      const opt = optionLocator.nth(i);
      const text = (await opt.innerText().catch(() => '')) || '';
      if (text.toLowerCase().includes(norm)) {
        targetOption = opt;
        pickedTitle = text.split('\n')[0].trim();
        break;
      }
    }
    if (!targetOption) {
      targetOption = optionLocator.first();
      const text = (await targetOption.innerText().catch(() => '')) || '';
      if (text) pickedTitle = text.split('\n')[0].trim();
    }

    await targetOption.click();
    await page.waitForTimeout(2500);
    return { title: pickedTitle };
  }

  async sendMessage(tenant = '', { message, chatName, attachments } = {}) {
    if (!message) throw new Error("Nachrichtentext (message) ist erforderlich.");
    if (!chatName) throw new Error("Empfänger (chat_name) ist erforderlich. Es wird bewusst nicht in einen unbestimmten 'aktiven' Chat gesendet.");
    const validAttachments = this.validateAttachments(attachments);
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);

    const filtered = await this.getFilteredChatRows(page);

    let picked = this.pickChatRow(filtered, chatName, undefined);
    if (!picked) {
      // Chat nicht in den geladenen Zeilen -> versuche neuen Chat über das Adressbuch
      const newChatResult = await this.startNewChat(page, chatName);
      if (!newChatResult) {
        throw new Error(`Chat oder Kollege '${chatName}' wurde in Teams (${t}) weder in bestehenden Chats noch im Adressbuch gefunden.`);
      }
      picked = newChatResult;
    } else {
      await picked.row.click();
      await page.waitForTimeout(2500);
    }

    // Compose-Feld GEZIELT im geöffneten Chat-Pane finden. WICHTIG: Darf NICHT das
    // globale div[role="textbox"] verwenden, sonst greift der Klick auf das
    // "New Message"-Compose-Feld und navigiert aus dem Chat heraus.
    const composeCtl = page
      .locator('[data-tid="chat-pane-compose-message-footer"], [data-tid="chat-pane-compose"]')
      .locator('[data-tid="ckeditor"], div[role="textbox"], [contenteditable], textarea, div[id^="new-message-"]')
      .first();

    try {
      await composeCtl.waitFor({ state: 'visible', timeout: 10000 });
    } catch (e) {
      throw new Error("Chat-Komposefeld im geöffneten Chat nicht gefunden. Keine Nachricht gesendet.");
    }
    // ElementHandle holen ist optional (Locator.click funktioniert direkt)
    await composeCtl.click();
    await page.waitForTimeout(500);
    await this.insertTextViaPaste(page, message);
    await page.waitForTimeout(600);

    // Anhänge hinzufügen, falls übergeben
    if (validAttachments.length > 0) {
      await this.attachFiles(page, validAttachments);
    }

    const sendButton = await page.$('button[data-tid="newMessageCommands-send"], button[data-tid="send-message-button"], button[aria-label*="Send"], button[aria-label*="Senden"]');
    if (sendButton) {
      await sendButton.click();
    } else {
      await page.keyboard.press('Enter');
    }
    await page.waitForTimeout(2000);

    const activeChat = await this.getActiveChatTitle(page);

    // Verifizieren, dass der aktive Chat dem Ziel entspricht (Sicherheitsnetz)
    const targetNorm = picked.title.toLowerCase().replace(/ \(you\)$/i, '');
    const activeNorm = activeChat.toLowerCase().replace(/ \(you\)$/i, '');
    const nameParts = targetNorm.split(/\s+/).filter(p => p.length > 2);
    const matchAny = nameParts.some(p => activeNorm.includes(p));
    if (!activeNorm.includes(targetNorm) && !targetNorm.includes(activeNorm) && !matchAny) {
      throw new Error(`Sicherheitsnetz: Der geöffnete Chat-Titel '${activeChat}' stimmt nicht mit Ziel '${picked.title}' überein. Nachricht wurde NICHT gesendet.`);
    }

    return {
      success: true,
      tenant: t,
      recipient: picked.title,
      confirmedActiveChat: activeChat,
      messageSent: message,
      attachments: validAttachments.map(p => path.basename(p)),
      status: validAttachments.length > 0
        ? `Nachricht mit ${validAttachments.length} Anhang/Anhängen erfolgreich gesendet.`
        : "Nachricht erfolgreich gesendet."
    };
  }
  async getMeetingStatus(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);
    const info = await speakerTracker.inspectMeeting(page);
    return {
      tenant: t,
      ...info,
      isTracking: speakerTracker.isTracking(t)
    };
  }

  async startSpeakerTracking(tenant = '', outputPath = null, options = {}) {
    const t = browserManager.normalizeTenant(tenant);
    const page = await this.getPage(t, true);
    return await speakerTracker.startTracking(page, t, outputPath, options);
  }

  async stopSpeakerTracking(tenant = '') {
    const t = browserManager.normalizeTenant(tenant);
    return await speakerTracker.stopTracking(t);
  }

  // Layer 1: Extrahiert den "Aktivität"-Tab (Activity-Feed) aus dem Teams-Web-DOM.
  // Die Roh-Items werden hier nicht klassifiziert (das macht der tim-Agent / die
  // Nachbearbeitung in Layer 2/3) — dieses Tool liefert nur die stabilen Rohdaten.
  async getActivity(tenant = '', options = {}) {
    return await activityClient.getActivity(tenant, options);
  }

  // Layer 2: Extrahiert den Activity-Feed UND klassifiziert die Einträge in
  // Kategorien (Meeting/Task/Entscheidung/Risiko/Sonstiges). Liefert die Gruppen
  // mit Relevanz-Scores, ohne eine Datei zu schreiben (reine Analyse).
  async getAnalyzedActivity(tenant = '', { maxItems = 50 } = {}) {
    const activity = await activityClient.getActivity(tenant, { maxItems });
    const groups = analyzeItems(activity.items);
    const counts = {};
    for (const cat of CATEGORIES) counts[cat] = groups[cat]?.length || 0;
    return {
      tenant: activity.tenant,
      extractedAt: activity.extractedAt,
      total: activity.count,
      counts,
      groups
    };
  }
}

// PURE Hilfsfunktion: konvertiert einen Teams-Zeitwert (data-mid als Unix-Millis/-Sekunden,
// <time datetime> ISO) zu einem ISO-8601-String. Muss reine Funktion bleiben (ohne Zugriff auf
// Modul-Scope), damit sie via .toString() für page.evaluate serialisierbar ist.
export function messageTimeToIso(raw) {
  if (raw == null || raw === '') return null;
  const s = String(raw).trim();
  if (!s) return null;
  // Unix-Epoche: 13-stellig = Millis, 10-stellig = Sekunden
  if (/^\d{10,13}$/.test(s)) {
    const num = Number(s);
    const ms = s.length >= 13 ? num : num * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime()) && d.getFullYear() > 2000 && d.getFullYear() < 2100) {
      return d.toISOString();
    }
  }
  // <time datetime="..."> ISO
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.toISOString();
  }
  return null;
}

// Pseudo-Einträge im Team-/Kanal-Baum, die keine echten Teams/Kanäle sind
// (Navigations- und Sektions-Header der linken Rails). Lowercase.
// Befund 01.10.2026 (d.velop-Tenant, neues Teams-Web): Die Team-/Kanal-Struktur liegt
// als Sektion "Teams and channels" INNERHALB der Chat-Rail — Level 1 = Sektions-Header
// (Copilot, Quick views, Favorites, Chats, Teams and channels), Level 2 = Teams,
// Level 3 = Kanäle. Ein data-tid="app-bar-teams" existiert dort NICHT mehr.
export const TREE_IGNORED_TITLES = new Set([
  'see all channels', 'see all your teams', 'see more', 'manage teams', 'create a team',
  'your teams', 'hidden teams', 'communities', 'communities and storyline', 'favorites',
  'favoriten', 'alle kanäle anzeigen', 'alle teams anzeigen', 'mehr anzeigen',
  'team erstellen', 'teams verwalten', 'meine teams', 'ausgeblendete teams',
  'teams and channels', 'teams und kanäle', 'copilot', 'quick views', 'drafts',
  'mentions', 'chats', 'activity', 'aktivität', 'calendar', 'kalender'
]);

// Prüft, ob ein Baum-Titel ein Navigations-/Sammel-Entry statt echtem Team/Kanal ist.
export function isIgnoredTreeTitle(title) {
  const t = String(title || '').toLowerCase().trim();
  if (!t) return true;
  if (TREE_IGNORED_TITLES.has(t)) return true;
  return t.startsWith('see ') || t.startsWith('alle kanäle') || t.startsWith('alle teams');
}

// Wählt ein Item aus einer Baum-Liste ({ title, idx?, level? }) — gleiche Match-Kaskade
// wie pickChatRow: exakter Titel -> Präfix -> Token-Match. Ignorierte/leere Titel werden
// übersprungen. Optionaler numerischer index adressiert die nutzbare Liste direkt.
export function pickTreeItem(items, name, index = null) {
  const usable = (items || []).filter(it => it && it.title && !isIgnoredTreeTitle(it.title));
  if (typeof index === 'number' && index >= 0 && index < usable.length) return usable[index];
  if (!name) return null;
  const term = String(name).toLowerCase().trim();
  if (!term) return null;

  const exact = usable.find(it => it.title.toLowerCase() === term);
  if (exact) return exact;

  const prefix = usable.find(it => it.title.toLowerCase().startsWith(term));
  if (prefix) return prefix;

  const tokens = term.split(/\s+/).filter(Boolean);
  if (tokens.length > 0) {
    const tokenMatch = usable.find(it => {
      const t = it.title.toLowerCase();
      return tokens.every(tok => t.includes(tok));
    });
    if (tokenMatch) return tokenMatch;
  }
  return null;
}

// Findet ein Team im gescannten Baum. Teams sind aufklappbare Einträge (aria-expanded
// gesetzt) — im neuen Teams-Web auf Level 2 unter der Sektion "Teams and channels",
// in älteren UIs auf Level 1. Sektions-Header werden via isIgnoredTreeTitle ausgefiltert.
export function findTeamInTree(items, teamName) {
  const list = items || [];
  const expandable = list.filter(it => it.expanded != null && !isIgnoredTreeTitle(it.title));
  if (expandable.length > 0) return pickTreeItem(expandable, teamName);
  const level1 = list.filter(it => Number.isFinite(it.level) && it.level === 1);
  if (level1.length > 0) return pickTreeItem(level1, teamName);
  return pickTreeItem(list, teamName);
}

// Liefert die Kandidaten-Kanal-Einträge unterhalb eines Team-Eintrags (DOM-Reihenfolge).
// Stoppt am nächsten aufklappbaren Eintrag (aria-expanded = nächstes Team) bzw. — wenn
// keine Expand-Infos vorliegen — am nächsten Eintrag mit level <= team.level. Ohne beide
// Struktur-Infos: alle Einträge nach dem Team (DOM-Reihenfolge); das Safety-Net beim
// Öffnen fängt Fehlklicks ab.
export function channelPoolForTeam(items, teamItem) {
  if (!teamItem) return [];
  const list = items || [];
  const hasExpandedInfo = list.some(it => it.expanded != null);
  const pool = [];
  for (const it of list) {
    if (it.idx <= teamItem.idx) continue;
    if (hasExpandedInfo) {
      // Nächstes aufklappbares Element = nächstes Team -> Teilbaum zu Ende
      if (it.expanded != null) break;
      pool.push(it);
    } else if (Number.isFinite(it.level) && Number.isFinite(teamItem.level)) {
      if (it.level <= teamItem.level) break;
      pool.push(it);
    } else {
      pool.push(it);
    }
  }
  return pool;
}

// Findet einen Kanal innerhalb des Team-Teilbaums.
export function findChannelInTree(items, teamItem, channelName) {
  return pickTreeItem(channelPoolForTeam(items, teamItem), channelName);
}

// Kanaltitel eines Teams (für hilfreiche Fehlermeldungen).
export function collectChannelTitles(items, teamItem) {
  return channelPoolForTeam(items, teamItem)
    .filter(it => it.title && !isIgnoredTreeTitle(it.title))
    .map(it => it.title);
}

// Gruppiert den flachen Baum in eine strukturierte Team->Kanäle-Ansicht.
// Strategie 1 (neues Teams-Web): Teams = aufklappbare Einträge (aria-expanded gesetzt),
// Kanäle = der Teilbaum darunter (bis zum nächsten aufklappbaren Eintrag).
// Strategie 2 (Fallback, ältere UI): Level 1 sind Teams, Rest Kanäle (DOM-Reihenfolge).
export function buildTeamsStructure(items) {
  const list = (items || []).filter(it => it && it.title && !isIgnoredTreeTitle(it.title));
  if (list.length === 0) return [];

  const expandableTeams = list.filter(it => it.expanded != null);
  if (expandableTeams.length > 0) {
    return expandableTeams.map(team => ({
      name: team.title,
      channels: collectChannelTitles(items, team)
    }));
  }

  const hasLevel1 = list.some(it => Number.isFinite(it.level) && it.level === 1);
  const teams = [];
  let current = null;

  for (const it of list) {
    const isTeam = hasLevel1
      ? (Number.isFinite(it.level) && it.level === 1)
      : it.expanded != null;
    if (isTeam) {
      current = { name: it.title, channels: [] };
      teams.push(current);
    } else if (current) {
      current.channels.push(it.title);
    }
  }
  return teams;
}

// Extrahiert Dateianhang-Metadaten ({ name, href }) aus einer Message-Row.
// Läuft serialisiert im Browser-Kontext (page.evaluate) und ist mit einem
// minimalen DOM-Mock unit-testbar. Bewusst self-contained (keine Modul-Scope-Referenzen).
export function collectAttachmentsFromRow(row) {
  const out = [];
  if (!row || typeof row.querySelectorAll !== 'function') return out;
  const seen = new Set();
  // Dateiname "name.ext" am Ende eines Labels / URL-Pfads
  const NAME_RE = /([A-Za-z0-9\u00c0-\u024f][\w\u00c0-\u024f. ()[\]#&'`-]*\.[A-Za-z0-9]{2,5})(?=\s|$|\n)/;
  const safeDecode = (s) => { try { return decodeURIComponent(s); } catch { return s; } };
  const push = (name, href) => {
    if (!name) return;
    const key = `${name}|${href || ''}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ name, href: href || null });
  };

  // 1) Links mit dateiartigem Ziel (SharePoint/OneDrive) oder dateiartigem Label
  const links = row.querySelectorAll('a[href]') || [];
  for (const a of links) {
    const href = (a.getAttribute && a.getAttribute('href')) || '';
    const label = (((a.getAttribute && a.getAttribute('aria-label')) || a.innerText || '') + '').trim();
    const bare = safeDecode(String(href).split('?')[0]);
    const m = NAME_RE.exec(label) || NAME_RE.exec(bare);
    const fileish = /sharepoint|1drv|onedrive|personalcontent|download=1/i.test(href) || m;
    if (!fileish) continue;
    const name = m ? m[1].trim() : (label.split('\n')[0].trim() || null);
    push(name, href);
  }

  // 2) Attachment-Cards (auch ohne <a>-Link: Download läuft über den Karten-Button)
  const cardSel = '[data-tid*="attachment" i], [data-tid*="file-card" i], [class*="attachmentCard" i]';
  const cards = row.querySelectorAll(cardSel) || [];
  for (const card of cards) {
    const label = (((card.getAttribute && card.getAttribute('aria-label')) || card.innerText || '') + '').trim();
    const m = NAME_RE.exec(label);
    if (m) push(m[1].trim(), null);
  }
  return out;
}

// Message-Reactions (Checkmark, Daumen, Herz, ...) einer Nachricht extrahieren.
// Laeuft serialisiert im Browser-Kontext (page.evaluate) und ist mit einem minimalen
// DOM-Mock unit-testbar — gleiches Muster wie collectAttachmentsFromRow.
//
// KERN BEFUND (Live-DOM-Analyse 08.10.2026, Tenant adesso + d-velop, Teams-Web):
// Reaktionen stehen als eigenstaendige Controls NEBEN dem Text-Body in der Message-ROW
// (Container .fui-ChatMessage__reactions) und tauchen daher NICHT in body.innerText auf —
// nur ein Scan ueber die ROW findet sie. Reale Struktur (verifiziert):
//
//   <div class="fui-ChatMessage__reactions">
//     <div role="toolbar" data-tid="diverse-reaction-summary" aria-label="Reaction summary">
//       <button data-tid="diverse-reaction-pill-button"
//               aria-labelledby="message-heart-1790155757631"   <-- KEIN aria-label!
//               aria-pressed="true"                             <-- eigene Reaktion
//               class="... diverse-reaction-pill-checked ...">  <-- eigene Reaktion
//         <span data-tid="emoticon-renderer">
//           <img alt="&#10084;&#65039;" itemid="heart" itemtype=".../Emoji" src=".../emoticons/heart/...">
//         </span>
//         <div id="message-heart-1790155757631">2 Heart reactions.</div>
//         <span class="fui-StyledText">2</span>                <-- nur ab count>1
//       </button>
//       <button data-tid="add-reaction-picker-entry-point-button">…  <-- PLUS-Button, KEINE Reaktion
//
// Daraus folgt die Extraktionslogik:
//  - type   = Slug aus aria-labelledby ("message-<type>-<mid>") bzw. img[itemid].
//            SPRACHUNABHAENGIG (Slug bleibt auch bei deutschem UI englisch: heart/yes/laugh/cry).
//  - emoji  = img[alt], NUR wenn es ein echtes Emoji-Glyph ist (❤️ / 👍 / 😢).
//            Bei Custom-Emoji liefert Teams stattdessen eine ID ("adesso-cody_hooray",
//            ASCII) -> dann emoji=null und custom=true; der Slug steht dann in type.
//  - count  = fuehrende Zahl im Label-Text ("2 Heart reactions." -> 2; ohne Zahl -> 1).
//  - byMe   = aria-pressed="true" oder Klasse diverse-reaction-pill-checked.
//  - label  = Original-Label fuer die Anzeige (lokalisiert, daher nur informativ).
// Es wird bewusst KEINE Positiv-/Negativ-Whitelist angewendet: das freigaberelevante
// Checkmark ist je nach Teams-Version/Region ein eigener Slug — die Entscheidung, ob eine
// Reaktion eine Freigabe ist, trifft der aufrufende Agent anhand von type/emoji/count/byMe.
// Kein CounterBadge existiert in dieser Struktur (Befund: count kommt NUR aus dem Label-Text).
// Klassifiziert ein DOM-Element als "Reaktions-Artefakt" (Reaktions-Container/-Pills/
// Plus-Button) — also UI-Chrome, der NICHT zum Nachrichten-Text gehoert.
// Reine Funktion ohne Modul-Scope, damit sie via .toString() in page.evaluate
// serialisierbar und hiermit unit-testbar ist (gleiches Muster wie messageTimeToIso).
// Live-Befund 08.10.2026: .fui-ChatMessage__reactions liegt innerhalb von
// [data-tid="chat-pane-message"] und landet sonst als ![✅](...) im Markdown.
export function isReactionArtifact({ className = '', dataTid = '' } = {}) {
  const cls = String(className || '').toLowerCase();
  const tid = String(dataTid || '').toLowerCase();
  if (cls.includes('reactions')) return true;              // fui-ChatMessage__reactions
  if (/diverse-reaction|add-reaction-picker|reaction-picker|reactions-summary/.test(tid)) return true;
  return false;
}

export function collectReactionsFromRow(row) {
  const out = [];
  if (!row || typeof row.querySelectorAll !== 'function') return out;
  const seen = new Set();

  const clean = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

  const push = (rec) => {
    if (!rec || (!rec.type && !rec.emoji)) return;
    const key = `${rec.type}|${rec.emoji}`;
    if (seen.has(key)) return;
    seen.add(key);
    out.push(rec);
  };

  // Sprachunabhaengiger Reaktions-Slug aus "message-heart-1790155757631".
  // Bei Custom-Emoji ist die Id laenger und enthaelt ein ';':
  //   message-adesso-cody_hooray;0-weu-d7-<hash>-<mid>  ->  "adesso-cody_hooray"
  const typeFromLabelledby = (id) => {
    const m = /^message-(.+)-\d{4,}$/.exec(clean(id));
    if (!m) return null;
    const t = m[1].split(';')[0].trim().toLowerCase();
    return t || null;
  };

  // Ist der alt-Wert ein echtes Emoji-Glyph oder nur eine Emoji-ID?
  // Custom-Emoji liefern "adesso-cody_hooray" (ASCII) statt eines Zeichens.
  const isGlyph = (s) => /[^\x20-\x7E]/.test(String(s || ''));

  // Fuehrende Zahl im Label ("2 Heart reactions." -> 2). Ohne Zahl => 1.
  const countFromLabel = (txt) => {
    const m = /(\d+)/.exec(String(txt || ''));
    const n = m ? Number(m[1]) : 1;
    return Number.isFinite(n) && n > 0 ? n : 1;
  };

  // Emoji + itemid aus dem emoticon-renderer des Pill-Buttons.
  // Custom-Emoji: alt ist die Emoji-ID (kein Glyph) -> emoji=null, custom=true.
  const emojiFrom = (pill) => {
    const img = pill.querySelector && pill.querySelector('img[alt]');
    if (!img) return { emoji: null, itemType: null, custom: false };
    const alt = clean(img.getAttribute('alt'));
    const itemType = clean(img.getAttribute('itemid')) || null;
    if (!alt) return { emoji: null, itemType, custom: false };
    if (isGlyph(alt)) return { emoji: alt, itemType, custom: false };
    return { emoji: null, itemType: alt, custom: true };
  };

  // --- Strategie 1: reale Teams-Struktur (data-tid=diverse-reaction-pill-button) ---
  const pills = row.querySelectorAll('[data-tid="diverse-reaction-pill-button"]') || [];
  for (const pill of pills) {
    const labelledby = pill.getAttribute('aria-labelledby') || '';
    const { emoji, itemType, custom } = emojiFrom(pill);
    const labelText = clean(pill.innerText);
    const cls = String(pill.getAttribute('class') || '');
    push({
      type: typeFromLabelledby(labelledby) || (itemType || '').split(';')[0].toLowerCase() || null,
      emoji: emoji || null,
      custom: custom || false,
      count: countFromLabel(labelText),
      byMe: pill.getAttribute('aria-pressed') === 'true' || /reaction-pill-checked/.test(cls),
      label: labelText || null
    });
  }
  if (out.length > 0) return out;

  // --- Strategie 2 (Fallback, aeltere/alternative Markup-Varianten): ---
  // Emoji-Bilder (img[itemid][itemtype*=Emoji]) in Reaktions-Containern. Der
  // Plus-Button (add-reaction-picker-entry-point-button) wird explizit ausgeschlossen.
  const containers = row.querySelectorAll(
    '[class*="reactionSummary" i], [class*="ChatMessage__reactions" i], [data-tid*="reaction-summary" i], [data-tid*="reactions-summary" i]'
  ) || [];
  const handled = new Set();
  for (const c of containers) {
    if (handled.has(c)) continue;
    handled.add(c);
    const contText = clean(c.innerText);
    const imgs = Array.from(c.querySelectorAll('img[alt], img[itemid]')) || [];
    for (const img of imgs) {
      // Plus-/Picker-Button ist keine Reaktion
      let cur = img.parentElement;
      let inPicker = false;
      while (cur && cur !== c) {
        const tid = String((cur.getAttribute && cur.getAttribute('data-tid')) || '');
        if (tid.includes('add-reaction-picker') || tid.includes('reaction-picker')) { inPicker = true; break; }
        cur = cur.parentElement;
      }
      if (inPicker) continue;
      const alt = clean(img.getAttribute('alt'));
      if (!alt) continue;
      const glyph = isGlyph(alt);
      push({
        type: (clean(img.getAttribute('itemid')) || alt).split(';')[0].toLowerCase() || null,
        emoji: glyph ? alt : null,
        custom: !glyph,
        count: countFromLabel(contText),
        byMe: false,
        label: contText || null
      });
    }
  }

  return out;
}

// Extrahiert einen "name.ext"-Dateinamen aus beliebigem Text (Karten-Label, Button-Kontext).
export function extractFilenameFromText(text) {
  const NAME_RE = /([A-Za-z0-9\u00c0-\u024f][\w\u00c0-\u024f. ()[\]#&'`-]*\.[A-Za-z0-9]{2,5})(?=\s|$|\n)/;
  const m = NAME_RE.exec(String(text || ''));
  return m ? m[1].trim() : null;
}

// Normalisiert eine Extension-Whitelist: Array kleinbuchstaben ohne führenden Punkt.
// undefined/null -> Default-Whitelist; leeres Array -> null (= alle Endungen erlaubt).
export function normalizeExtensionWhitelist(extensions) {
  if (extensions === undefined || extensions === null) {
    return [...DEFAULT_DOWNLOAD_EXTENSIONS];
  }
  let raw = extensions;
  if (typeof raw === 'string') raw = raw.split(',');
  if (!Array.isArray(raw)) return [...DEFAULT_DOWNLOAD_EXTENSIONS];
  const cleaned = [...new Set(raw
    .map(e => String(e || '').trim().toLowerCase().replace(/^\./, ''))
    .filter(Boolean))];
  return cleaned.length === 0 ? null : cleaned;
}

// Prüft anhand des Dateinamens gegen die Whitelist (null = alles erlaubt).
export function extensionAllowed(filename, exts) {
  if (exts === null || exts === undefined) return true;
  const m = /\.([A-Za-z0-9]{1,10})$/.exec(String(filename || '').trim());
  if (!m) return false;
  return exts.includes(m[1].toLowerCase());
}

// Hängt ?download=1 bzw. &download=1 an eine Datei-URL an (SharePoint-Forced-Download),
// sofern nicht bereits vorhanden.
export function urlForcedDownload(href) {
  const raw = String(href || '').trim();
  if (!raw || !/^https?:\/\//i.test(raw)) return raw;
  if (/([?&])download=1/i.test(raw)) return raw;
  return raw + (raw.includes('?') ? '&' : '?') + 'download=1';
}

// Liefert einen kollisionsfreien absoluten Zielpfad im Download-Verzeichnis
// (datei.ext, datei_1.ext, ... datei_999.ext, sonst Zeitstempel-Suffix).
export function nextAvailablePath(dir, filename) {
  const clean = sanitizeFilename(path.basename(String(filename || 'datei')), 80);
  let candidate = path.join(dir, clean);
  if (!fs.existsSync(candidate)) return candidate;
  const ext = path.extname(clean);
  const base = ext ? clean.slice(0, -ext.length) : clean;
  for (let i = 1; i < 1000; i++) {
    candidate = path.join(dir, `${base}_${i}${ext}`);
    if (!fs.existsSync(candidate)) return candidate;
  }
  return path.join(dir, `${base}_${Date.now()}${ext}`);
}

// Standard-Whitelist herunterladbarer Anhänge (Baustein B).
export const DEFAULT_DOWNLOAD_EXTENSIONS = [
  'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv',
  'zip', 'png', 'jpg', 'jpeg', 'gif', 'msg', 'vsdx', 'one', 'heic'
];

// Mime-Type -> Dateiendung für persistierte Teams-Medien (Fallback .bin).
export function mimeToExt(mime) {
  if (!mime) return 'bin';
  const m = String(mime).split(';')[0].trim().toLowerCase();
  const map = {
    'image/png': 'png',
    'image/jpeg': 'jpg',
    'image/jpg': 'jpg',
    'image/gif': 'gif',
    'image/webp': 'webp',
    'image/bmp': 'bmp',
    'image/svg+xml': 'svg',
    'image/avif': 'avif'
  };
  return map[m] || 'bin';
}

// Ersetzt Zeichen, die in Dateinamen problematisch sind (auch für tenant/chatName-Anteile).
export function sanitizeFilename(name, maxLen = 50) {
  const cleaned = String(name ?? '')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .trim();
  return cleaned.slice(0, maxLen) || 'unknown';
}

// Konvertiert eine dataURL ("data:<mime>;base64,<...>") und schreibt sie als Datei.
// Liefert { filePath, ext, saved } oder { saved: false } bei ungültiger dataURL.
export function saveBlobDataUrl(dataUrl, mediaDir, filename) {
  if (!dataUrl || typeof dataUrl !== 'string' || !dataUrl.startsWith('data:')) {
    return { saved: false };
  }
  const match = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(dataUrl);
  if (!match) return { saved: false };

  const mime = match[1] || '';
  const isBase64 = Boolean(match[2]);
  const payload = match[3] || '';
  let buf;
  try {
    buf = isBase64 ? Buffer.from(payload, 'base64') : Buffer.from(decodeURIComponent(payload), 'utf8');
  } catch {
    return { saved: false };
  }
  if (!buf || buf.length === 0) return { saved: false };

  const ext = mimeToExt(mime);
  const filePath = path.join(mediaDir, `${filename}.${ext}`);
  try {
    fs.writeFileSync(filePath, buf);
    return { saved: true, filePath, ext };
  } catch {
    return { saved: false };
  }
}

// Ersetzt in HTML alle img-src="blob:..." (und srcset) durch die lokalen Dateipfade aus dem Mapping.
// blobUrlToPath: Map/Objekt { blobUrl -> file:// Pfad }. Nicht gemappte Blobs bleiben unverändert.
export function replaceBlobUrlsInHtml(html, blobUrlToPath = {}) {
  if (!html) return html;
  const mapEntries = Object.entries(blobUrlToPath);
  if (mapEntries.length === 0) return html;

  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  let out = html;
  for (const [blobUrl, filePath] of mapEntries) {
    if (!filePath) continue;
    const qEsc = esc(blobUrl);
    // src="blob:..." und src='blob:...'
    out = out.replace(new RegExp(`src=["']${qEsc}["']`, 'g'), `src="${filePath}"`);
    // srcset (Comma-getrennt "url 1x, url 2x") — nur den exakten URL-Teil ersetzen
    out = out.replace(new RegExp(qEsc, 'g'), filePath);
  }
  return out;
}

// Erstellt (falls nötig) das Medienverzeichnis und liefert den absoluten Pfad zurück.
export function ensureMediaDir(mediaDir) {
  fs.mkdirSync(mediaDir, { recursive: true });
  return mediaDir;
}

export const teamsClient = new TeamsClient();
