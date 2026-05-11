// OpenCap Data Exporter — Popup v3.1.1

(async () => {
  "use strict";

  // ── Version ──────────────────────────────────────────────────────────────
  const vEl = document.getElementById("extVersion");
  if (vEl) vEl.textContent = chrome.runtime.getManifest().version;

  // ── DOM refs ─────────────────────────────────────────────────────────────
  const jobIdInput      = document.getElementById("jobId");
  const jobIdHint       = document.getElementById("jobIdHint");
  const folderNameEl    = document.getElementById("folderName");
  const btnChooseFolder = document.getElementById("btnChooseFolder");
  const btnClearFolder  = document.getElementById("btnClearFolder");

  const chkJob       = document.getElementById("chkJob");
  const chkCrew      = document.getElementById("chkCrew");
  const chkBha       = document.getElementById("chkBha");
  const chkSlide     = document.getElementById("chkSlide");
  const chkInventory = document.getElementById("chkInventory");
  const btnFetch = document.getElementById("btnFetch");
  const btnClear = document.getElementById("btnClear");

  const badgeJob       = document.getElementById("badgeJob");
  const badgeCrew      = document.getElementById("badgeCrew");
  const badgeBha       = document.getElementById("badgeBha");
  const badgeSlide     = document.getElementById("badgeSlide");
  const badgeInventory = document.getElementById("badgeInventory");
  const metaJob        = document.getElementById("metaJob");
  const metaCrew       = document.getElementById("metaCrew");
  const metaBha        = document.getElementById("metaBha");
  const metaSlide      = document.getElementById("metaSlide");
  const metaInventory  = document.getElementById("metaInventory");

  const progressWrap  = document.getElementById("progressWrap");
  const progressLabel = document.getElementById("progressLabel");
  const progressPct   = document.getElementById("progressPct");
  const progressFill  = document.getElementById("progressFill");

  const resultArea = document.getElementById("resultArea");
  const resultText = document.getElementById("resultText");
  const resultMeta = document.getElementById("resultMeta");

  // ── IndexedDB helpers (for FileSystemDirectoryHandle persistence) ─────────
  const IDB_NAME    = "fieldcap-exporter";
  const IDB_VERSION = 1;
  const IDB_STORE   = "handles";

  const openIdb = () => new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = (e) => e.target.result.createObjectStore(IDB_STORE);
    req.onsuccess       = (e) => resolve(e.target.result);
    req.onerror         = (e) => reject(e.target.error);
  });

  const idbGet = async (key) => {
    const db = await openIdb();
    return new Promise((resolve, reject) => {
      const req = db.transaction(IDB_STORE, "readonly").objectStore(IDB_STORE).get(key);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror   = () => reject(req.error);
    });
  };

  const idbSet = async (key, value) => {
    const db = await openIdb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror    = () => reject(tx.error);
    });
  };

  const idbDel = async (key) => {
    const db = await openIdb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror    = () => reject(tx.error);
    });
  };

  // ── Auto-detect Job ID from active FieldCap tab URL ───────────────────────
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tab?.url) {
      const m = tab.url.match(/[#/][Jj]obs?[/=](\d+)/);
      if (m) {
        jobIdInput.value = m[1];
        jobIdHint.textContent = "↑ from active tab";
      }
    }
  } catch (_) {}

  // ── Folder picker state ───────────────────────────────────────────────────
  let dirHandle = null;

  const updateFolderUI = () => {
    if (dirHandle) {
      folderNameEl.textContent = dirHandle.name;
      folderNameEl.classList.remove("default");
      btnClearFolder.hidden = false;
    } else {
      folderNameEl.textContent = "No folder selected";
      folderNameEl.classList.add("default");
      btnClearFolder.hidden = true;
    }
  };

  try {
    const stored = await idbGet("dirHandle");
    if (stored) dirHandle = stored;
  } catch (_) {
    dirHandle = null;
  }
  updateFolderUI();

  btnChooseFolder.addEventListener("click", async () => {
    try {
      const handle = await window.showDirectoryPicker({ mode: "readwrite" });
      dirHandle = handle;
      await idbSet("dirHandle", handle);
      updateFolderUI();
    } catch (e) {
      if (e.name !== "AbortError") showErr(`Could not open folder: ${e.message}`);
    }
  });

  btnClearFolder.addEventListener("click", async () => {
    dirHandle = null;
    await idbDel("dirHandle");
    updateFolderUI();
  });

  // ── Badge helpers ─────────────────────────────────────────────────────────
  const setBadge = (el, text, cls) => {
    el.textContent = text;
    el.className = `card-badge${cls ? ` ${cls}` : ""}`;
  };

  // ── Progress bar ──────────────────────────────────────────────────────────
  const showProgress = (pct, label) => {
    progressWrap.classList.add("visible");
    progressLabel.textContent = label ?? "";
    progressPct.textContent   = `${Math.round(pct)}%`;
    progressFill.style.width  = `${pct}%`;
    progressFill.classList.remove("complete", "error");
  };
  const completeProgress = (label = "Done") => {
    progressLabel.textContent = label;
    progressPct.textContent   = "100%";
    progressFill.style.width  = "100%";
    progressFill.classList.add("complete");
  };
  const errorProgress = () => {
    progressFill.classList.add("error");
  };
  const hideProgress = () => {
    progressWrap.classList.remove("visible");
    progressFill.style.width = "0%";
    progressFill.classList.remove("complete", "error");
  };

  // ── Status bar ────────────────────────────────────────────────────────────
  const showOk = (msg, meta = "") => {
    resultArea.classList.add("visible");
    resultText.className = "result-ok";
    resultText.textContent = `✓ ${msg}`;
    resultMeta.textContent = meta;
  };
  const showErr = (msg) => {
    resultArea.classList.add("visible");
    resultText.className = "result-err";
    resultText.textContent = `✗ ${msg}`;
    resultMeta.textContent = "";
  };
  const showInfo = (msg) => {
    resultArea.classList.add("visible");
    resultText.className = "meta";
    resultText.textContent = msg;
    resultMeta.textContent = "";
  };
  const hideResult = () => resultArea.classList.remove("visible");

  // ── Storage keys ─────────────────────────────────────────────────────────
  const KEY_META          = "fieldcap_meta";
  const KEY_CSV_JOB       = "fieldcap_csv_job";
  const KEY_CSV_CREW      = "fieldcap_csv_crew";
  const KEY_CSV_BHA       = "fieldcap_csv_bha";
  const KEY_CSV_SLIDE_DAY = "fieldcap_csv_slide_by_day";
  const KEY_CSV_INVENTORY = "fieldcap_csv_inventory";

  // ── Restore cached CSV state ──────────────────────────────────────────────
  const storedAll = await chrome.storage.local.get([
    KEY_META, KEY_CSV_JOB, KEY_CSV_CREW, KEY_CSV_BHA, KEY_CSV_SLIDE_DAY, KEY_CSV_INVENTORY,
  ]);

  const meta = storedAll[KEY_META];
  if (meta) {
    if (meta.jobId && !jobIdInput.value) jobIdInput.value = meta.jobId;
    const ts = meta.builtAt ? new Date(meta.builtAt).toLocaleString() : "";
    showInfo(`Last fetch: ${ts}`);

    if (storedAll[KEY_CSV_JOB]) {
      setBadge(badgeJob, "ready", "ok");
      metaJob.textContent = `fieldcap-job-${meta.jobId}-job-details.csv`;
    }
    if (storedAll[KEY_CSV_CREW]) {
      setBadge(badgeCrew, `${meta.crewRows ?? "?"} rows`, "ok");
      metaCrew.textContent = `fieldcap-job-${meta.jobId}-crew.csv`;
    }
    if (storedAll[KEY_CSV_BHA]) {
      setBadge(badgeBha, `${meta.bhaRows ?? "?"} rows`, "ok");
      metaBha.textContent = `fieldcap-job-${meta.jobId}-bha-equipment.csv · ${meta.bhaCount ?? "?"} BHAs`;
    }
    if (storedAll[KEY_CSV_SLIDE_DAY]) {
      setBadge(badgeSlide, `${meta.slideByDayRows ?? "?"} rows`, "ok");
      metaSlide.textContent = `fieldcap-job-${meta.jobId}-slide-rotate-metres-by-day.csv`;
    }
    if (storedAll[KEY_CSV_INVENTORY]) {
      setBadge(badgeInventory, `${meta.inventoryRows ?? "?"} rows`, "ok");
      metaInventory.textContent = `fieldcap-job-${meta.jobId}-inventory.csv`;
    }
  }

  // ── File write helpers ────────────────────────────────────────────────────
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const downloadViaChromeDownloads = (csvText, filename) =>
    new Promise((resolve, reject) => {
      const blob = new Blob([csvText], { type: "text/csv;charset=utf-8" });
      const url  = URL.createObjectURL(blob);
      chrome.downloads.download({ url, filename, saveAs: false, conflictAction: "overwrite" }, () => {
        URL.revokeObjectURL(url);
        const err = chrome.runtime.lastError;
        if (err) reject(new Error(err.message));
        else resolve(true);
      });
    });

  const writeFile = async (csv, filename) => {
    if (!dirHandle) throw new Error("No folder selected.");

    let perm = await dirHandle.queryPermission({ mode: "readwrite" });
    if (perm !== "granted") perm = await dirHandle.requestPermission({ mode: "readwrite" });
    if (perm !== "granted") throw new Error("Write permission was denied.");

    const writeOnce = async () => {
      const fh       = await dirHandle.getFileHandle(filename, { create: true });
      const writable = await fh.createWritable();
      await writable.write(csv);
      await writable.close();
      return `${dirHandle.name}/${filename}`;
    };

    let lastMsg = "";
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        return await writeOnce();
      } catch (e) {
        const msg = String(e?.message ?? e ?? "");
        lastMsg = msg;
        if (/denied|used by another process|being used|lock|busy/i.test(msg))
          throw new Error(`${filename} is open/locked. Close it in Excel and retry.`);
        if (/aborted due to security policy/i.test(msg) && attempt < 3) {
          if (attempt === 2) { try { await dirHandle.removeEntry(filename); } catch (_) {} }
          await sleep(350 * attempt);
          continue;
        }
        if (attempt < 3) { await sleep(250 * attempt); continue; }
      }
    }

    await downloadViaChromeDownloads(csv, filename);
    return `${filename} → Chrome Downloads (folder write blocked by OS policy)`;
  };

  const saveAllCachedCsvs = async (jobId) => {
    const keyMap = {
      job:       KEY_CSV_JOB,
      crew:      KEY_CSV_CREW,
      bha:       KEY_CSV_BHA,
      slideDay:  KEY_CSV_SLIDE_DAY,
      inventory: KEY_CSV_INVENTORY,
    };
    const nameMap = {
      job:       `fieldcap-job-${jobId}-job-details.csv`,
      crew:      `fieldcap-job-${jobId}-crew.csv`,
      bha:       `fieldcap-job-${jobId}-bha-equipment.csv`,
      slideDay:  `fieldcap-job-${jobId}-slide-rotate-metres-by-day.csv`,
      inventory: `fieldcap-job-${jobId}-inventory.csv`,
    };

    const stored = await chrome.storage.local.get(Object.values(keyMap));
    const saved  = [];
    const failed = [];

    for (const which of ["job", "crew", "bha", "slideDay", "inventory"]) {
      const csv = stored[keyMap[which]];
      if (!csv) continue;
      try {
        const name = await writeFile(csv, nameMap[which]);
        saved.push(name);
      } catch (e) {
        failed.push(String(e?.message ?? e));
      }
      await sleep(100);
    }
    return { saved, failed };
  };

  // ── Ensure folder is available before starting ────────────────────────────
  const ensureFolder = async () => {
    if (!dirHandle) {
      const handle = await window.showDirectoryPicker({ mode: "readwrite" });
      dirHandle = handle;
      await idbSet("dirHandle", handle);
      updateFolderUI();
    }
    let perm = await dirHandle.queryPermission({ mode: "readwrite" });
    if (perm !== "granted") perm = await dirHandle.requestPermission({ mode: "readwrite" });
    if (perm !== "granted") throw new Error("Write permission was denied for the folder.");
  };

  // ── Fetch & Save (single action) ──────────────────────────────────────────
  btnFetch.addEventListener("click", async () => {
    const jobId = parseInt(jobIdInput.value, 10);
    if (!jobId) { showErr("Enter a valid Job ID first."); return; }

    const flags = {
      jobDetails: chkJob.checked,
      crew:       chkCrew.checked,
      bha:        chkBha.checked,
      slideDay:   chkSlide.checked,
      inventory:  chkInventory.checked,
    };

    if (!flags.jobDetails && !flags.crew && !flags.bha && !flags.slideDay && !flags.inventory) {
      showErr("Select at least one data type to fetch.");
      return;
    }

    // Ensure folder is ready before starting the long fetch
    try {
      await ensureFolder();
    } catch (e) {
      if (e.name === "AbortError") return;
      showErr(e.message);
      return;
    }

    btnFetch.disabled = true;
    btnClear.disabled = true;
    hideResult();

    if (flags.jobDetails) { setBadge(badgeJob,       "Fetching…", "busy"); metaJob.textContent       = ""; }
    if (flags.crew)       { setBadge(badgeCrew,      "Fetching…", "busy"); metaCrew.textContent      = ""; }
    if (flags.bha)        { setBadge(badgeBha,       "Fetching…", "busy"); metaBha.textContent       = ""; }
    if (flags.slideDay)   { setBadge(badgeSlide,     "Fetching…", "busy"); metaSlide.textContent     = ""; }
    if (flags.inventory)  { setBadge(badgeInventory, "Fetching…", "busy"); metaInventory.textContent = ""; }

    showProgress(2, "Connecting to FieldCap OData API…");

    // Open long-lived port for progress streaming
    let port;
    try {
      port = chrome.runtime.connect({ name: "fieldcap-fetch" });
    } catch (e) {
      showErr(`Could not connect to background: ${e.message}`);
      btnFetch.disabled = false;
      btnClear.disabled = false;
      hideProgress();
      return;
    }

    port.onDisconnect.addListener(() => {
      const err = chrome.runtime.lastError;
      if (err) {
        showErr(`Background disconnected: ${err.message}`);
        errorProgress();
        btnFetch.disabled = false;
        btnClear.disabled = false;
      }
    });

    port.onMessage.addListener(async (msg) => {
      if (msg.type === "PROGRESS") {
        showProgress(msg.pct, msg.label);
        return;
      }

      if (msg.type !== "DONE") return;

      port.disconnect();

      if (!msg.ok) {
        errorProgress();
        setTimeout(hideProgress, 1500);
        showErr(msg.error ?? "Fetch failed. Make sure you are logged into FieldCap.");
        if (flags.jobDetails) setBadge(badgeJob,       "Error", "err");
        if (flags.crew)       setBadge(badgeCrew,      "Error", "err");
        if (flags.bha)        setBadge(badgeBha,       "Error", "err");
        if (flags.slideDay)   setBadge(badgeSlide,     "Error", "err");
        if (flags.inventory)  setBadge(badgeInventory, "Error", "err");
        btnFetch.disabled = false;
        btnClear.disabled = false;
        return;
      }

      // Update status cards
      if (flags.jobDetails) {
        setBadge(badgeJob, "ready", "ok");
        metaJob.textContent = `fieldcap-job-${jobId}-job-details.csv`;
      }
      if (flags.crew) {
        setBadge(badgeCrew, `${msg.crewRows} rows`, "ok");
        metaCrew.textContent = `fieldcap-job-${jobId}-crew.csv`;
      }
      if (flags.bha) {
        setBadge(badgeBha, `${msg.bhaRows} rows`, "ok");
        metaBha.textContent = `fieldcap-job-${jobId}-bha-equipment.csv · ${msg.bhaCount} BHAs`;
      }
      const slideDayRows = Number(msg.slideByDayRows ?? 0);
      if (flags.slideDay) {
        if (slideDayRows > 0) {
          setBadge(badgeSlide, `${slideDayRows} rows`, "ok");
          metaSlide.textContent = `fieldcap-job-${jobId}-slide-rotate-metres-by-day.csv`;
        } else {
          setBadge(badgeSlide, "0 rows", "warn");
          metaSlide.textContent = "No ActivityLogs found for this job.";
        }
      }
      const inventoryRows = Number(msg.inventoryRows ?? 0);
      if (flags.inventory) {
        if (inventoryRows > 0) {
          setBadge(badgeInventory, `${inventoryRows} rows`, "ok");
          metaInventory.textContent = `fieldcap-job-${jobId}-inventory.csv`;
        } else {
          setBadge(badgeInventory, "0 rows", "warn");
          metaInventory.textContent = "No JobTools found for this job.";
        }
      }

      // Auto-save all CSVs to folder
      showProgress(96, "Saving files to disk…");
      try {
        const { saved, failed } = await saveAllCachedCsvs(jobId);

        if (saved.length === 0 && failed.length === 0) {
          showErr("No CSV data was cached after fetch.");
          errorProgress();
          setTimeout(hideProgress, 1500);
        } else if (failed.length > 0) {
          showErr(`Saved ${saved.length} file(s). Failed: ${failed.join(" | ")}`);
          errorProgress();
          setTimeout(hideProgress, 2000);
        } else {
          completeProgress(`Saved ${saved.length} file(s) → ${dirHandle.name}`);
          setTimeout(hideProgress, 2500);
          const liveBha = Number(msg.liveBhaRows ?? 0);
          const liveAct = Number(msg.liveActivityRows ?? 0);
          showOk(
            `${saved.length} files → ${dirHandle.name}`,
            `${new Date().toLocaleString()} · live bhaRows=${liveBha} · activityRows=${liveAct}`
          );
        }
      } catch (saveErr) {
        showErr(saveErr.message);
        errorProgress();
        setTimeout(hideProgress, 1500);
      }

      btnFetch.disabled = false;
      btnClear.disabled = false;
    });

    port.postMessage({ type: "FETCH_ALL", jobId, flags });
  });

  // ── Clear ─────────────────────────────────────────────────────────────────
  btnClear.addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "CLEAR_CACHE" }, () => {
      setBadge(badgeJob,       "\u2014", "");
      setBadge(badgeCrew,      "\u2014", "");
      setBadge(badgeBha,       "\u2014", "");
      setBadge(badgeSlide,     "\u2014", "");
      setBadge(badgeInventory, "\u2014", "");
      metaJob.textContent       = "";
      metaCrew.textContent      = "";
      metaBha.textContent       = "";
      metaSlide.textContent     = "";
      metaInventory.textContent = "";
      hideResult();
      hideProgress();
    });
  });

  // ── Schema Probe ──────────────────────────────────────────────────────────
  const btnProbe = document.getElementById("btnProbe");
  if (btnProbe) {
    btnProbe.addEventListener("click", () => {
      const jobId = parseInt(jobIdInput.value, 10);
      if (!jobId) { showErr("Enter a Job ID first."); return; }
      btnProbe.disabled = true;
      showInfo("Probing ToolAssembly schema…");
      chrome.runtime.sendMessage({ type: "PROBE_SCHEMA", jobId }, (res) => {
        btnProbe.disabled = false;
        if (chrome.runtime.lastError) { showErr(chrome.runtime.lastError.message); return; }
        if (!res?.ok) { showErr(res?.error ?? "Schema probe failed"); return; }
        const lines = [`ToolAssembly fields (${res.keys.length}):`];
        lines.push(res.keys.join("  |  "));
        if (res.statsKeys?.length > 0) {
          lines.push(``, `Statistics sub-fields (${res.statsKeys.length}):`);
          lines.push(res.statsKeys.join("  |  "));
        } else {
          lines.push(``, `(no Statistics nav property)`);
        }
        resultArea.classList.add("visible");
        resultText.className = "meta";
        resultText.style.whiteSpace = "pre-wrap";
        resultText.style.fontSize   = "9px";
        resultText.textContent = lines.join("\n");
        resultMeta.textContent = "Copy these field names to identify correct hour columns";
      });
    });
  }

  // ── Inventory Probe ───────────────────────────────────────────────────────
  const btnInventoryProbe = document.getElementById("btnInventoryProbe");
  if (btnInventoryProbe) {
    btnInventoryProbe.addEventListener("click", () => {
      btnInventoryProbe.disabled = true;
      showInfo("Probing inventory endpoints (ItemSerials → Items → JobTools)…");

      chrome.runtime.sendMessage({ type: "PROBE_INVENTORY" }, (res) => {
        btnInventoryProbe.disabled = false;
        if (chrome.runtime.lastError) { showErr(chrome.runtime.lastError.message); return; }
        if (!res?.ok) { showErr(res?.error ?? "Inventory probe failed"); return; }

        const lines = [];
        for (const r of res.results) {
          if (!r.ok) {
            lines.push(`[${r.label}] ✗ ${r.reason}`);
            continue;
          }
          lines.push(`[${r.label}] ✓  ${r.keys.length} fields${r.count != null ? `  ·  ~${r.count} records` : ""}`);
          if (r.serialNumbers?.length) {
            lines.push(`  Sample serials: ${r.serialNumbers.join("  |  ")}`);
          }
          lines.push(`  All fields: ${r.keys.join("  |  ")}`);
          if (r.hourKeys.length > 0) {
            const hourDisplay = r.hourKeys.filter((k) => !/clientsynctime|timestamp/i.test(k));
            lines.push(`  Hour / metric fields (${hourDisplay.length}): ${hourDisplay.join("  |  ")}`);
            const nonZero = Object.entries(r.sampleValues).filter(([, v]) => typeof v === "number" && v > 0);
            const allZero = Object.entries(r.sampleValues).filter(([, v]) => v === 0);
            if (nonZero.length > 0) {
              lines.push(`  ✓ Non-zero values found (max across sample rows):`);
              for (const [k, v] of nonZero) lines.push(`    ${k}: ${v}`);
            } else if (allZero.length > 0) {
              lines.push(`  ⚠ All hour fields are 0 on every sampled row`);
              lines.push(`    → hours may accumulate only after job completion, or`);
              lines.push(`    → TotalHours1 may map to a specific "hours type" configured per company`);
              lines.push(`    → try filtering by a known serial number to confirm`);
            } else {
              lines.push(`  (hour fields present but all null on sample rows)`);
            }
          } else {
            lines.push(`  ⚠ No hour/metric fields detected — hours may be in a sub-entity`);
          }
          lines.push("");
        }

        resultArea.classList.add("visible");
        resultText.className = "meta";
        resultText.style.whiteSpace = "pre-wrap";
        resultText.style.fontSize   = "9px";
        resultText.textContent = lines.join("\n");
        resultMeta.textContent = "Review fields above — share output to plan the full inventory export";
      });
    });
  }

  // ── Network Sniffer ───────────────────────────────────────────────────────
  const btnSniff      = document.getElementById("btnSniff");
  const btnSniffLog   = document.getElementById("btnSniffLog");
  const btnSniffClear = document.getElementById("btnSniffClear");

  const setSniffUI = (active) => {
    if (!btnSniff) return;
    btnSniff.textContent = active ? "⏹ Sniffing…" : "📶 Sniff";
    btnSniff.classList.toggle("btn-sniff-active", active);
    if (btnSniffLog)   btnSniffLog.style.display   = active ? "inline-block" : "inline-block";
    if (btnSniffClear) btnSniffClear.style.display = active ? "inline-block" : "inline-block";
  };

  // Restore sniff state on popup open
  chrome.runtime.sendMessage({ type: "SNIFF_STATUS" }, (res) => {
    if (res?.ok) setSniffUI(res.active);
  });

  if (btnSniff) {
    btnSniff.addEventListener("click", () => {
      chrome.runtime.sendMessage({ type: "SNIFF_STATUS" }, (res) => {
        const wasActive = res?.active ?? false;
        const nextMsg   = wasActive ? "SNIFF_STOP" : "SNIFF_START";
        chrome.runtime.sendMessage({ type: nextMsg }, (r2) => {
          if (chrome.runtime.lastError) { showErr(chrome.runtime.lastError.message); return; }
          setSniffUI(r2?.active ?? false);
          if (r2?.active) {
            showInfo("Sniffing active — browse FieldCap now. Open the Inventory tab, BHAs, etc. Then click Log.");
          } else {
            showInfo("Sniffing stopped. Click Log to review captured endpoints.");
          }
        });
      });
    });
  }

  if (btnSniffLog) {
    btnSniffLog.addEventListener("click", () => {
      chrome.runtime.sendMessage({ type: "SNIFF_GET_LOG" }, (res) => {
        if (chrome.runtime.lastError) { showErr(chrome.runtime.lastError.message); return; }
        const log = res?.log ?? [];
        if (!log.length) { showInfo("Sniff log is empty — start sniffing then browse FieldCap."); return; }

        // Group by entity (first path segment) so inventory-related calls stand out
        const byEntity = {};
        for (const entry of log) {
          const entity = entry.path.split("?")[0].split("(")[0].trim() || "other";
          (byEntity[entity] = byEntity[entity] ?? []).push(entry);
        }

        const lines = [`Captured ${log.length} unique OData calls:\n`];
        for (const [entity, entries] of Object.entries(byEntity)) {
          lines.push(`── ${entity} (${entries.length}) ──`);
          for (const e of entries) {
            const t = e.ts.substring(11, 19);
            lines.push(`  [${t}] ${e.method}  ${e.path}`);
          }
          lines.push("");
        }

        resultArea.classList.add("visible");
        resultText.className = "meta";
        resultText.style.whiteSpace = "pre-wrap";
        resultText.style.fontSize   = "8.5px";
        resultText.textContent = lines.join("\n");
        resultMeta.textContent = "Copy the entity name of interest → share here to build the export";
      });
    });
  }

  if (btnSniffClear) {
    btnSniffClear.addEventListener("click", () => {
      chrome.runtime.sendMessage({ type: "SNIFF_CLEAR" }, () => {
        showInfo("Sniff log cleared.");
      });
    });
  }

  // ── Debug State ───────────────────────────────────────────────────────────
  const btnDebugState = document.getElementById("btnDebugState");
  if (btnDebugState) {
    btnDebugState.addEventListener("click", async () => {
      const jobId = parseInt(jobIdInput.value, 10);
      if (!jobId) { showErr("Enter a Job ID first."); return; }
      btnDebugState.disabled = true;
      try {
        const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
        let liveBhaRows = [];
        if (tab?.id) {
          const live = await new Promise((resolve) => {
            chrome.tabs.sendMessage(tab.id, { type: "SCRAPE_NOW" }, (res) => {
              if (chrome.runtime.lastError) return resolve({ ok: false, error: chrome.runtime.lastError.message });
              resolve(res ?? { ok: false });
            });
          });
          liveBhaRows = Array.isArray(live?.bhaRows) ? live.bhaRows : [];
        }

        const keys   = ["fieldcap_bha_grid_rows", "fieldcap_intercepted_assemblies", "fieldcap_schema_probe", "fieldcap_meta"];
        const stored = await chrome.storage.local.get(keys);
        const gridByJob     = stored.fieldcap_bha_grid_rows ?? {};
        const gridForJob    = gridByJob[String(jobId)] ?? {};
        const intercept     = stored.fieldcap_intercepted_assemblies ?? {};
        const sampleGrid    = Object.values(gridForJob)[0] ?? null;
        const sampleLive    = liveBhaRows[0] ?? null;
        const sampleIntercept = Object.values(intercept)[0] ?? null;

        const metricKeysOf = (obj) => {
          if (!obj || typeof obj !== "object") return [];
          return Object.keys(obj).filter((k) => /met|meter|metre|foot|depth|drill|distance|\bmd\b/i.test(k));
        };

        const lines = [
          `Debug State for job ${jobId}`,
          ``,
          `Live SCRAPE_NOW bhaRows: ${liveBhaRows.length}`,
          `Cached grid rows: ${Object.keys(gridForJob).length}`,
          `Cached intercepted assemblies: ${Object.keys(intercept).length}`,
          `Last fetch meta liveBhaRows: ${stored.fieldcap_meta?.liveBhaRows ?? "(none)"}`,
          ``,
          `Active tab: ${tab?.url ?? "(none)"}`,
        ];

        if (sampleLive) {
          lines.push(``, `Sample LIVE row keys:`, Object.keys(sampleLive).join(" | "));
          const mk = metricKeysOf(sampleLive);
          if (mk.length) {
            lines.push(`Sample LIVE metric values:`);
            for (const k of mk.slice(0, 8)) lines.push(`  ${k}: ${sampleLive[k]}`);
          }
        }
        if (sampleGrid) {
          lines.push(``, `Sample CACHED row keys:`, Object.keys(sampleGrid).join(" | "));
          const mk = metricKeysOf(sampleGrid);
          if (mk.length) lines.push(`Sample CACHED metric keys:`, mk.join(" | "));
        }
        if (sampleIntercept) {
          lines.push(``, `Sample INTERCEPT keys:`, Object.keys(sampleIntercept).join(" | "));
          const mk = metricKeysOf(sampleIntercept);
          if (mk.length) {
            lines.push(`Sample INTERCEPT metric values:`);
            for (const k of mk.slice(0, 8)) lines.push(`  ${k}: ${sampleIntercept[k]}`);
          }
        }

        resultArea.classList.add("visible");
        resultText.className = "meta";
        resultText.style.whiteSpace = "pre-wrap";
        resultText.style.fontSize   = "9px";
        resultText.textContent = lines.join("\n");
        resultMeta.textContent = "If live rows are 0, parser is not matching table headers in runtime.";
      } catch (e) {
        showErr(`Debug failed: ${e.message}`);
      } finally {
        btnDebugState.disabled = false;
      }
    });
  }
})();
