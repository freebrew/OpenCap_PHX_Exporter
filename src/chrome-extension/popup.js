// FieldCap Data Exporter — Popup v3.1.0

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

  const chkJob   = document.getElementById("chkJob");
  const chkCrew  = document.getElementById("chkCrew");
  const chkBha   = document.getElementById("chkBha");
  const chkSlide = document.getElementById("chkSlide");
  const btnFetch = document.getElementById("btnFetch");
  const btnClear = document.getElementById("btnClear");

  const badgeJob   = document.getElementById("badgeJob");
  const badgeCrew  = document.getElementById("badgeCrew");
  const badgeBha   = document.getElementById("badgeBha");
  const badgeSlide = document.getElementById("badgeSlide");
  const metaJob    = document.getElementById("metaJob");
  const metaCrew   = document.getElementById("metaCrew");
  const metaBha    = document.getElementById("metaBha");
  const metaSlide  = document.getElementById("metaSlide");

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

  // ── Restore cached CSV state ──────────────────────────────────────────────
  const storedAll = await chrome.storage.local.get([
    KEY_META, KEY_CSV_JOB, KEY_CSV_CREW, KEY_CSV_BHA, KEY_CSV_SLIDE_DAY,
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
      job:      KEY_CSV_JOB,
      crew:     KEY_CSV_CREW,
      bha:      KEY_CSV_BHA,
      slideDay: KEY_CSV_SLIDE_DAY,
    };
    const nameMap = {
      job:      `fieldcap-job-${jobId}-job-details.csv`,
      crew:     `fieldcap-job-${jobId}-crew.csv`,
      bha:      `fieldcap-job-${jobId}-bha-equipment.csv`,
      slideDay: `fieldcap-job-${jobId}-slide-rotate-metres-by-day.csv`,
    };

    const stored = await chrome.storage.local.get(Object.values(keyMap));
    const saved  = [];
    const failed = [];

    for (const which of ["job", "crew", "bha", "slideDay"]) {
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
    };

    if (!flags.jobDetails && !flags.crew && !flags.bha && !flags.slideDay) {
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

    if (flags.jobDetails) { setBadge(badgeJob,   "Fetching…", "busy"); metaJob.textContent   = ""; }
    if (flags.crew)       { setBadge(badgeCrew,  "Fetching…", "busy"); metaCrew.textContent  = ""; }
    if (flags.bha)        { setBadge(badgeBha,   "Fetching…", "busy"); metaBha.textContent   = ""; }
    if (flags.slideDay)   { setBadge(badgeSlide, "Fetching…", "busy"); metaSlide.textContent = ""; }

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
        if (flags.jobDetails) setBadge(badgeJob,   "Error", "err");
        if (flags.crew)       setBadge(badgeCrew,  "Error", "err");
        if (flags.bha)        setBadge(badgeBha,   "Error", "err");
        if (flags.slideDay)   setBadge(badgeSlide, "Error", "err");
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
      setBadge(badgeJob,   "\u2014", "");
      setBadge(badgeCrew,  "\u2014", "");
      setBadge(badgeBha,   "\u2014", "");
      setBadge(badgeSlide, "\u2014", "");
      metaJob.textContent   = "";
      metaCrew.textContent  = "";
      metaBha.textContent   = "";
      metaSlide.textContent = "";
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
