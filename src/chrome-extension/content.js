// FieldCap BHA Equipment Exporter — Content Script v2.5.0
// Watches the DOM as the user navigates FieldCap and captures:
//   • BHA summary list   (from the BHAs tab)
//   • Tools master list  (from the Tools tab)
//   • BHA components     (when the user has drilled into a BHA detail page)
//
// Also intercepts FieldCap's own OData fetch calls (page-context injection)
// so that hour/footage statistics the app loads natively are forwarded to
// the background and cached — these fields do NOT exist on the plain
// ToolAssemblies response that the extension fetches directly.
//
// Hour capture workflow:
//   1. User opens FieldCap, navigates into the job's BHAs tab
//   2. User clicks each BHA to open its dialog — the app fetches hour data
//   3. This script intercepts those fetches and caches the hour fields
//   4. User opens extension popup → Fetch & Build CSVs
//   5. buildBhaCsv merges cached hours into the CSV output

(function () {
  "use strict";

  // ── Inject page-context fetch interceptor ─────────────────────────────────
  // Chrome content scripts run in an isolated context; to intercept window.fetch
  // we must inject a real <script> into the page DOM, then communicate back via
  // window.postMessage.
  (function injectFetchSpy() {
    if (document.getElementById("__fc_spy")) return;
    const s = document.createElement("script");
    s.id = "__fc_spy";
    s.src = chrome.runtime.getURL("injected-spy.js");
    (document.head || document.documentElement).appendChild(s);
    s.onload = () => s.remove();
    s.onerror = () => s.remove();
  })();

  // ── Utilities ─────────────────────────────────────────────────────────────
  const normalize = (v) => String(v ?? "").replace(/\s+/g, " ").trim();
  const compactKey = (v) => normalize(v).toLowerCase().replace(/[^a-z0-9]/g, "");
  const getCellText = (cell) => normalize(cell.innerText ?? cell.textContent);

  // ── Click-memory: remember which BHA row was clicked last ─────────────────
  // FieldCap's "Update Tool Assembly" detail page doesn't show the BHA # in
  // the URL or heading, so we capture it at click time on the BHA list page.
  let lastClickedBha = null;
  let lastClickedAt  = 0;
  const CLICK_MEMORY_MS = 5 * 60 * 1000; // 5 min

  // Manual override from popup
  let manualBhaContext = null;

  document.addEventListener("click", (ev) => {
    // Walk up to the nearest <tr>
    let row = ev.target;
    while (row && row.tagName !== "TR" && row !== document.body) row = row.parentElement;
    if (!row || row.tagName !== "TR") return;

    // Only count clicks on rows inside a table that looks like a BHA list.
    const table = row.closest("table");
    if (!table) return;
    const headerCells = [...table.querySelectorAll("tr")][0]?.querySelectorAll("th, td") ?? [];
    const headerKeys = [...headerCells].map((c) => compactKey(c.innerText));
    const hasHeaderLike = (re) => headerKeys.some((h) => re.test(h));
    const looksLikeBhaList =
      hasHeaderLike(/^bha\d*$/) ||
      hasHeaderLike(/^toolassembly\d*$/) ||
      hasHeaderLike(/^toolassemblynumber$/);
    if (!looksLikeBhaList) return;

    // Find which column has the BHA #
    const headerArr = [...headerCells].map((c) => compactKey(c.innerText));
    const bhaColIdx = headerArr.findIndex((h) =>
      /^bha\d*$/.test(h) || /^toolassembly\d*$/.test(h) || h === "toolassemblynumber"
    );
    if (bhaColIdx < 0) return;

    const cells = row.querySelectorAll("td");
    const cell = cells[bhaColIdx];
    if (!cell) return;
    const num = normalize(cell.innerText).match(/\d+/)?.[0];
    if (!num) return;

    lastClickedBha = num;
    lastClickedAt  = Date.now();
  }, true); // capture phase so we see clicks even if SPA stops propagation

  // ── Detect "which BHA am I currently viewing?" ────────────────────────────
  const detectBhaContext = () => {
    // 1. Manual override wins
    if (manualBhaContext) return { bha: manualBhaContext, source: "manual" };

    // 2. URL patterns
    const url = location.href;
    const urlPatterns = [
      /[/#]BHAs?[/_-]?#?(\d+)\b/i,
      /[?&](?:bha|bhaNum|bhaNumber|bhaId)=([^&]+)/i,
      /toolAssemblyNumber[/=:](\d+)/i,
    ];
    for (const re of urlPatterns) {
      const m = url.match(re);
      if (m && m[1]) return { bha: decodeURIComponent(m[1]), source: "url" };
    }

    // 3. Page heading
    const headingSelectors =
      "h1, h2, h3, h4, .page-title, .header-title, .panel-title, [class*='breadcrumb']";
    for (const h of document.querySelectorAll(headingSelectors)) {
      const text = normalize(h.innerText ?? h.textContent);
      const m = text.match(/BHA\s*#?\s*(\d+)\b/i);
      if (m) return { bha: m[1], source: "heading" };
    }

    // 4. Click memory (within 5 min)
    if (lastClickedBha && (Date.now() - lastClickedAt) < CLICK_MEMORY_MS) {
      // Sanity check: are we likely on a detail page (not the list itself)?
      const onListPage = !!document.querySelector("table th, table td")
        ? [...document.querySelectorAll("table tr")].length > 5
          && [...document.querySelectorAll("th, td")].some((c) =>
            /tool assembly #|bha\s*#?$/i.test(normalize(c.innerText))
          )
        : false;
      if (!onListPage) {
        return { bha: lastClickedBha, source: "click" };
      }
    }

    return null;
  };

  // ── Table scraping ────────────────────────────────────────────────────────
  const scrapeAllTables = () => {
    const tables = [...document.querySelectorAll("table")];
    const result = [];

    for (const table of tables) {
      const allRows = [...table.querySelectorAll("tr")];
      if (allRows.length < 2) continue;

      let headerIdx = -1;
      let headers = [];
      for (let i = 0; i < allRows.length; i++) {
        const cells = [...allRows[i].querySelectorAll("th, td")];
        const texts = cells.map(getCellText).filter(Boolean);
        if (texts.length >= 2) { headerIdx = i; headers = cells.map(getCellText); break; }
      }
      if (headerIdx < 0) continue;

      const bodyRows = allRows.slice(headerIdx + 1);
      const tableRows = [];
      for (const tr of bodyRows) {
        const cells = [...tr.querySelectorAll("th, td")];
        if (cells.every((c) => !getCellText(c))) continue;
        const rowObj = {};
        cells.forEach((cell, ci) => {
          rowObj[headers[ci] ?? `Col${ci}`] = getCellText(cell);
        });
        tableRows.push(rowObj);
      }
      if (tableRows.length === 0) continue;

      const hKeys = headers.map((h) => compactKey(h));
      const hasHdr = (re) => hKeys.some((k) => re.test(k));
      let tableType = "unknown";

      if (hasHdr(/^bha\d*$/) || hasHdr(/^toolassembly\d*$/) || hasHdr(/^toolassemblynumber$/)) {
        tableType = "bha";
      } else if (hasHdr(/^jobhours$/) || hasHdr(/^hsls$/)) {
        tableType = "tools";
      } else if (
        (hasHdr(/^activitycode$/) || hasHdr(/^activity$/)) &&
        (hasHdr(/^duration$/) || hasHdr(/^course$/))
      ) {
        tableType = "activities";
      } else if (hasHdr(/^serial$/) || hasHdr(/^serialnumber$/)) {
        tableType = "components";
      }

      result.push({ tableType, headers, rows: tableRows });
    }
    return result;
  };

  // ── Forward intercepted OData responses to background ────────────────────
  // The page-context fetch spy posts messages here; we relay them to the
  // background service worker which caches the assembly hour data.
  window.addEventListener("message", (ev) => {
    if (!ev.data?.__FC_ODATA__ || !isAlive()) return;
    const url  = ev.data.url ?? "";
    const data = ev.data.data;
    if (!data) return;
    // Forward all OData responses; background now filters/extracts relevant fields.
    // FieldCap endpoints vary by tenant/version, and strict URL filters can miss
    // the actual source powering the BHA grid.
    safeSend({ type: "INTERCEPTED_ODATA", url, data });
  });

  // ── Extension-context guard ───────────────────────────────────────────────
  // After an extension reload the old injected script's runtime context becomes
  // invalid. chrome.runtime.id is undefined when that happens; any further call
  // to chrome.runtime.sendMessage / onMessage throws synchronously. We check
  // here before every call and tear down the observer + interval on first fault.
  const isAlive = () => {
    try { return !!chrome.runtime?.id; } catch (_) { return false; }
  };

  let observerActive = true;
  const killScript = () => {
    if (!observerActive) return;
    observerActive = false;
    observer.disconnect();
    clearInterval(urlWatcher);
    clearTimeout(debounceTimer);
  };

  const safeSend = (msg) => {
    if (!isAlive()) { killScript(); return; }
    try {
      chrome.runtime.sendMessage(msg).catch(() => {});
    } catch (_) {
      killScript();
    }
  };

  // ── Auto-scrape via MutationObserver ──────────────────────────────────────
  let debounceTimer = null;
  const DEBOUNCE_MS = 800;

  const tryAutoScrape = () => {
    if (!isAlive()) { killScript(); return; }

    const tables    = scrapeAllTables();
    const bhaRows   = tables.filter((t) => t.tableType === "bha").flatMap((t) => t.rows);
    const toolRows  = tables.filter((t) => t.tableType === "tools").flatMap((t) => t.rows);
    const activityRows = tables.filter((t) => t.tableType === "activities").flatMap((t) => t.rows);
    const ctx       = detectBhaContext();

    const componentTables = tables.filter((t) => t.tableType === "components");
    const componentRows   = ctx
      ? componentTables.flatMap((t) => t.rows.map((r) => ({ ...r, __bha: ctx.bha })))
      : [];

    if (bhaRows.length === 0 && toolRows.length === 0 && activityRows.length === 0 && componentRows.length === 0
        && !ctx) {
      return;
    }

    safeSend({
      type: "AUTO_SCRAPE",
      bhaRows,
      toolRows,
      activityRows,
      componentRows,
      bhaContext:    ctx?.bha   ?? null,
      contextSource: ctx?.source ?? null,
      url: location.href,
    });
  };

  const observer = new MutationObserver(() => {
    if (!isAlive()) { killScript(); return; }
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(tryAutoScrape, DEBOUNCE_MS);
  });

  observer.observe(document.body, { childList: true, subtree: true });

  // Initial scrape on page load
  setTimeout(tryAutoScrape, 1500);

  // Also re-scan when URL changes (SPA route changes)
  let lastUrl = location.href;
  const urlWatcher = setInterval(() => {
    if (!isAlive()) { killScript(); return; }
    if (location.href !== lastUrl) {
      lastUrl = location.href;
      setTimeout(tryAutoScrape, 600);
    }
  }, 500);

  // ── Message handler ────────────────────────────────────────────────────────
  if (isAlive()) {
    chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
      if (message.type === "PING") {
        sendResponse({ ok: true });
        return false;
      }

      if (message.type === "SET_MANUAL_BHA") {
        manualBhaContext = message.bha ? String(message.bha).trim() : null;
        setTimeout(tryAutoScrape, 50);
        sendResponse({ ok: true, bha: manualBhaContext });
        return false;
      }

      if (message.type === "SCRAPE_NOW") {
        const tables   = scrapeAllTables();
        const bhaRows  = tables.filter((t) => t.tableType === "bha").flatMap((t) => t.rows);
        const toolRows = tables.filter((t) => t.tableType === "tools").flatMap((t) => t.rows);
        const activityRows = tables.filter((t) => t.tableType === "activities").flatMap((t) => t.rows);
        const ctx      = detectBhaContext();
        const componentRows = ctx
          ? tables.filter((t) => t.tableType === "components")
                  .flatMap((t) => t.rows.map((r) => ({ ...r, __bha: ctx.bha })))
          : [];

        if (bhaRows.length > 0 || toolRows.length > 0 || activityRows.length > 0 || componentRows.length > 0) {
          safeSend({
            type: "AUTO_SCRAPE",
            bhaRows, toolRows, activityRows, componentRows,
            bhaContext:    ctx?.bha   ?? null,
            contextSource: ctx?.source ?? null,
            url: location.href,
          });
        }
        sendResponse({
          ok: true, bhaRows, toolRows, activityRows, componentRows,
          bhaContext: ctx?.bha ?? null, contextSource: ctx?.source ?? null,
        });
        return false;
      }

      return false;
    });
  }
})();
