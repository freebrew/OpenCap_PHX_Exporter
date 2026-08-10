// FieldCap BHA Equipment Exporter — Content Script v2.5.10
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
  const KEY_SUPPRESS_BOTTOM_LINE_MODAL = "fieldcap_suppress_bottom_line_modal";
  const BOTTOM_LINE_DIALOG_SELECTORS = [
    "[role='dialog']",
    "[aria-modal='true']",
    ".modal-dialog",
    ".modal-content",
    ".ui-dialog",
    ".k-window",
    ".k-dialog",
    ".cdk-overlay-pane",
  ].join(",");

  // Daily Activities footer (dark strip): distance totals use R:/S:/T: with decimal
  // metres; duration columns use H:MM — match only decimal patterns so we do not
  // confuse R: 3:35 (time) with metres.
  const parseActivityMetreFooter = (tableEl) => {
    const text = tableEl.innerText ?? "";
    const rm = text.match(/\bR:\s*([\d,]+\.\d{2})\b/);
    const sm = text.match(/\bS:\s*([\d,]+\.\d{2})\b/);
    const rot = rm ? Number(rm[1].replace(/,/g, "")) : null;
    const slide = sm ? Number(sm[1].replace(/,/g, "")) : null;
    return {
      rot: Number.isFinite(rot) ? rot : null,
      slide: Number.isFinite(slide) ? slide : null,
    };
  };

  const findActivityMetreFooterOnPage = () => {
    for (const table of document.querySelectorAll("table")) {
      const ft = parseActivityMetreFooter(table);
      if (ft.slide != null || ft.rot != null) return ft;
    }
    return null;
  };

  // FieldCap occasionally blocks workflow with a bottom-hole verification modal.
  // Suppression is intentionally text-gated so unrelated dialogs remain visible.
  let suppressBottomLineModal = true;
  const suppressedBottomLineEls = new WeakSet();

  const isBottomLineVerificationText = (text) =>
    /bottom line verification/i.test(text) &&
    /bottom hole measurements/i.test(text) &&
    /\bSD\b/i.test(text) &&
    /\bTVD\b/i.test(text);

  const elementArea = (el) => {
    const rect = el.getBoundingClientRect();
    return Math.max(rect.width, 0) * Math.max(rect.height, 0);
  };

  const isSafeDialogElement = (el) => {
    if (!el || el === document.body || el === document.documentElement) return false;

    const text = normalize(el.innerText ?? el.textContent);
    if (!isBottomLineVerificationText(text)) return false;

    const rect = el.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return false;

    const viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
    const viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);
    const pageSized = rect.width >= viewportWidth * 0.96 && rect.height >= viewportHeight * 0.85;
    return !pageSized;
  };

  const pickSmallestSafeDialog = (candidates) =>
    [...candidates]
      .filter(isSafeDialogElement)
      .sort((a, b) => elementArea(a) - elementArea(b))[0] ?? null;

  const isLikelyBackdrop = (el) => {
    const rect = el.getBoundingClientRect();
    const viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
    const viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);
    const text = normalize(el.innerText ?? el.textContent);
    return text.length < 40 && rect.width >= viewportWidth * 0.6 && rect.height >= viewportHeight * 0.6;
  };

  const findBottomLineVerificationModal = () => {
    const directMatch = pickSmallestSafeDialog(document.querySelectorAll(BOTTOM_LINE_DIALOG_SELECTORS));
    if (directMatch) return directMatch;

    for (const heading of document.querySelectorAll("h1, h2, h3, h4, h5, strong, b")) {
      const headingText = normalize(heading.innerText ?? heading.textContent);
      if (!/^Bottom Line Verification$/i.test(headingText)) continue;

      const candidates = [];
      let el = heading.parentElement;
      for (let depth = 0; el && depth < 6; depth++) {
        candidates.push(el);
        el = el.parentElement;
      }
      const ancestorMatch = pickSmallestSafeDialog(candidates);
      if (ancestorMatch) return ancestorMatch;
    }

    return null;
  };

  const hideBottomLineModalElement = (modal) => {
    modal.setAttribute("data-open-cap-suppressed", "bottom-line-verification");
    modal.setAttribute("aria-hidden", "true");
    modal.style.setProperty("display", "none", "important");
    modal.style.setProperty("visibility", "hidden", "important");

    for (const backdrop of document.querySelectorAll(".modal-backdrop, .cdk-overlay-backdrop, .k-overlay, .MuiBackdrop-root")) {
      if (!isLikelyBackdrop(backdrop)) continue;
      backdrop.style.setProperty("display", "none", "important");
      backdrop.style.setProperty("visibility", "hidden", "important");
    }
    document.body.classList.remove("modal-open");
    document.body.style.removeProperty("overflow");
  };

  const dismissBottomLineVerificationModal = () => {
    if (!suppressBottomLineModal) return;

    const modal = findBottomLineVerificationModal();
    if (!modal || suppressedBottomLineEls.has(modal)) return;
    suppressedBottomLineEls.add(modal);

    const controls = [...modal.querySelectorAll("button, a, input[type='button'], input[type='submit']")];
    const remindControl = controls.find((el) =>
      /remind me in 10 minutes/i.test(normalize(el.innerText ?? el.value ?? el.textContent))
    );

    if (remindControl) {
      remindControl.click();
      setTimeout(() => {
        if (modal.isConnected) hideBottomLineModalElement(modal);
      }, 250);
      return;
    }

    hideBottomLineModalElement(modal);
  };

  try {
    chrome.storage.local.get(KEY_SUPPRESS_BOTTOM_LINE_MODAL, (data) => {
      suppressBottomLineModal = data[KEY_SUPPRESS_BOTTOM_LINE_MODAL] !== false;
      dismissBottomLineVerificationModal();
    });

    chrome.storage.onChanged.addListener((changes, areaName) => {
      if (areaName !== "local" || !changes[KEY_SUPPRESS_BOTTOM_LINE_MODAL]) return;
      suppressBottomLineModal = changes[KEY_SUPPRESS_BOTTOM_LINE_MODAL].newValue !== false;
      dismissBottomLineVerificationModal();
    });
  } catch (_) {}

  const looksLikeRigNameValue = (v) => {
    const s = normalize(v);
    if (!s || s.length > 80) return false;
    if (/^rig\s*name$/i.test(s)) return false;
    // Reject following-field labels accidentally scooped as the value.
    if (/^(rig\s*type|data\s*rec|pipe\s*arm|top\s*drive|loader|ground\s*msl|rkb)\b/i.test(s)) {
      return false;
    }
    return /[a-z0-9]/i.test(s);
  };

  const detectRigNameFromPage = () => {
    // 1) Inline "Rig Name: Akita-520" on one element.
    for (const el of document.querySelectorAll("div, span, p, td, th, li, h1, h2, h3, h4, label, dt, dd")) {
      const text = normalize(el.innerText ?? el.textContent);
      if (!text || text.length > 120) continue;
      const inline = text.match(/^Rig\s*Name\s*:?\s+(.+)$/i);
      if (inline && looksLikeRigNameValue(inline[1])) return normalize(inline[1]);
    }

    // 2) Label/value split across siblings (FieldCap Well Parameters / Rig panel).
    for (const el of document.querySelectorAll("div, span, p, td, th, label, dt")) {
      const label = normalize(el.innerText ?? el.textContent);
      if (!/^Rig\s*Name\s*:?$/i.test(label)) continue;

      let sib = el.nextElementSibling;
      while (sib) {
        const value = normalize(sib.innerText ?? sib.textContent);
        if (looksLikeRigNameValue(value)) return value;
        sib = sib.nextElementSibling;
      }

      const parent = el.parentElement;
      if (parent) {
        const kids = Array.from(parent.children);
        const idx = kids.indexOf(el);
        for (let i = idx + 1; i < kids.length; i++) {
          const value = normalize(kids[i].innerText ?? kids[i].textContent);
          if (looksLikeRigNameValue(value)) return value;
        }
      }
    }

    // 3) Whole-panel text fallback: "Rig Name" … value … next known Rig label.
    const body = normalize(document.body?.innerText ?? "");
    const m = body.match(
      /\bRig\s*Name\s*:?\s+(.+?)(?=\s{2,}|\s+Rig\s*Type\b|\s+Data\s*Rec\.?\b|\s+Iron\s*Rgh|\s+Pipe\s*Arm\b|$)/i
    );
    if (m && looksLikeRigNameValue(m[1])) return normalize(m[1]);

    return "";
  };

  // Per-row BHA from the Activities grid (e.g. "4", "5" on the same day). Page
  // context (__bha) must NOT override these — otherwise all rows inherit the
  // last-clicked / route BHA and slide/rotate metres pile onto one assembly.
  const rowExplicitBha = (row) => {
    if (!row || typeof row !== "object") return "";
    for (const [k, v] of Object.entries(row)) {
      if (String(k).startsWith("__")) continue;
      const nk = compactKey(k);
      const looksBhaCol =
        /^bha\d*$/.test(nk) ||
        nk === "toolassemblynumber" ||
        /^toolassembly\d*$/.test(nk) ||
        nk === "bhano" ||
        nk === "assemblyno" ||
        nk === "assemblynumber";
      if (!looksBhaCol) continue;
      const m = normalize(String(v ?? "")).match(/^(\d+)/);
      if (m) return m[1];
    }
    return "";
  };

  const distinctExplicitBhas = (rows) => {
    const s = new Set();
    for (const r of rows ?? []) {
      const b = rowExplicitBha(r);
      if (b) s.add(b);
    }
    return s;
  };

  const attachActivityBhaContext = (rawRows, ctx) =>
    rawRows.map((r) => {
      const explicit = rowExplicitBha(r);
      if (explicit) {
        const { __bha, ...rest } = r;
        return rest;
      }
      if (ctx?.bha) return { ...r, __bha: ctx.bha };
      return r;
    });

  // Footer R:/S: totals are for the whole visible daily table; only merge them
  // when a single BHA is represented in the grid — otherwise they double-count
  // across BHAs or attach the full-day total to ctx.bha only.
  const buildActivityRowsWithOptionalFooter = (activityRowsRaw, ctx, footerFt) => {
    const distinct = distinctExplicitBhas(activityRowsRaw);
    const multiBhaDay = distinct.size >= 2;
    const footerOk =
      ctx &&
      !multiBhaDay &&
      footerFt &&
      (footerFt.slide != null || footerFt.rot != null);
    const footerRow = footerOk
      ? [{
          __bha: ctx.bha,
          __activityFooter: true,
          __footerSlideMetres: footerFt.slide != null ? String(footerFt.slide) : "",
          __footerRotateMetres: footerFt.rot != null ? String(footerFt.rot) : "",
        }]
      : [];
    const body = attachActivityBhaContext(activityRowsRaw, ctx);
    return [...footerRow, ...body];
  };

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
      } else if (
        (hasHdr(/^ticketdate$/) || hasHdr(/^ticketday$/)) &&
        (hasHdr(/^tickettotal$/) || hasHdr(/^total$/) || hasHdr(/^amount$/))
      ) {
        tableType = "tickets";
      } else if (hasHdr(/^jobhours$/) || hasHdr(/^hsls$/)) {
        tableType = "tools";
      } else if (
        (hasHdr(/^activitycode$/) || hasHdr(/^activity$/)) &&
        (hasHdr(/^duration$/) || hasHdr(/^course$/))
      ) {
        tableType = "activities";
      } else if (
        hasHdr(/course|metre|meter/i) &&
        hasHdr(/activity|code/i) &&
        hasHdr(/bha|toolassembly/i)
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
    const activityRowsRaw = tables.filter((t) => t.tableType === "activities").flatMap((t) => t.rows);
    const ticketRows = tables.filter((t) => t.tableType === "tickets").flatMap((t) => t.rows);
    const ctx       = detectBhaContext();
    const footerFt  = findActivityMetreFooterOnPage();
    const activityRows = buildActivityRowsWithOptionalFooter(activityRowsRaw, ctx, footerFt);
    const rigName = detectRigNameFromPage();

    const componentTables = tables.filter((t) => t.tableType === "components");
    const componentRows   = ctx
      ? componentTables.flatMap((t) => t.rows.map((r) => ({ ...r, __bha: ctx.bha })))
      : [];

    // Well Parameters / Rig / Tickets pages may have no BHA tables — still cache.
    if (bhaRows.length === 0 && toolRows.length === 0 && activityRows.length === 0
        && componentRows.length === 0 && ticketRows.length === 0 && !ctx && !rigName) {
      return;
    }

    safeSend({
      type: "AUTO_SCRAPE",
      bhaRows,
      toolRows,
      activityRows,
      ticketRows,
      rigName,
      componentRows,
      bhaContext:    ctx?.bha   ?? null,
      contextSource: ctx?.source ?? null,
      url: location.href,
    });
  };

  const observer = new MutationObserver(() => {
    if (!isAlive()) { killScript(); return; }
    dismissBottomLineVerificationModal();
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(tryAutoScrape, DEBOUNCE_MS);
  });

  observer.observe(document.body, { childList: true, subtree: true });

  // Initial scrape on page load
  setTimeout(dismissBottomLineVerificationModal, 500);
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

      if (message.type === "SET_BOTTOM_LINE_SUPPRESSION") {
        suppressBottomLineModal = message.enabled !== false;
        dismissBottomLineVerificationModal();
        sendResponse({ ok: true, enabled: suppressBottomLineModal });
        return false;
      }

      if (message.type === "SCRAPE_NOW") {
        const tables   = scrapeAllTables();
        const bhaRows  = tables.filter((t) => t.tableType === "bha").flatMap((t) => t.rows);
        const toolRows = tables.filter((t) => t.tableType === "tools").flatMap((t) => t.rows);
        const activityRowsRaw = tables.filter((t) => t.tableType === "activities").flatMap((t) => t.rows);
        const ticketRows = tables.filter((t) => t.tableType === "tickets").flatMap((t) => t.rows);
        const ctx      = detectBhaContext();
        const footerFt = findActivityMetreFooterOnPage();
        const activityRows = buildActivityRowsWithOptionalFooter(activityRowsRaw, ctx, footerFt);
        const rigName = detectRigNameFromPage();
        const componentRows = ctx
          ? tables.filter((t) => t.tableType === "components")
                  .flatMap((t) => t.rows.map((r) => ({ ...r, __bha: ctx.bha })))
          : [];

        if (bhaRows.length > 0 || toolRows.length > 0 || activityRows.length > 0
            || componentRows.length > 0 || ticketRows.length > 0 || rigName) {
          safeSend({
            type: "AUTO_SCRAPE",
            bhaRows, toolRows, activityRows, ticketRows, rigName, componentRows,
            bhaContext:    ctx?.bha   ?? null,
            contextSource: ctx?.source ?? null,
            url: location.href,
          });
        }
        sendResponse({
          ok: true, bhaRows, toolRows, activityRows, ticketRows, rigName, componentRows,
          bhaContext: ctx?.bha ?? null, contextSource: ctx?.source ?? null,
        });
        return false;
      }

      return false;
    });
  }
})();
