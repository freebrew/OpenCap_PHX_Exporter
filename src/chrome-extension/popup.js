// FieldCap Data Exporter — Popup v2.5.0

(async () => {
  "use strict";

  // ── Version ──────────────────────────────────────────────────────────────
  const vEl = document.getElementById("extVersion");
  if (vEl) vEl.textContent = chrome.runtime.getManifest().version;

  // ── DOM refs ─────────────────────────────────────────────────────────────
  const jobIdInput      = document.getElementById("jobId");
  const folderNameEl    = document.getElementById("folderName");
  const btnChooseFolder = document.getElementById("btnChooseFolder");
  const btnClearFolder  = document.getElementById("btnClearFolder");

  const chkJob  = document.getElementById("chkJob");
  const chkCrew = document.getElementById("chkCrew");
  const chkBha  = document.getElementById("chkBha");
  const btnFetch       = document.getElementById("btnFetch");
  const btnClear       = document.getElementById("btnClear");
  const btnDownloadAll = document.getElementById("btnDownloadAll");

  const badgeJob  = document.getElementById("badgeJob");
  const badgeCrew = document.getElementById("badgeCrew");
  const badgeBha  = document.getElementById("badgeBha");
  const metaJob   = document.getElementById("metaJob");
  const metaCrew  = document.getElementById("metaCrew");
  const metaBha   = document.getElementById("metaBha");

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
    const db  = await openIdb();
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

  // ── Folder picker state ───────────────────────────────────────────────────
  let dirHandle = null;

  const updateFolderUI = () => {
    if (dirHandle) {
      folderNameEl.textContent = dirHandle.name;
      folderNameEl.classList.remove("default");
      btnClearFolder.hidden = false;
    } else {
      folderNameEl.textContent = "Chrome Downloads (default)";
      folderNameEl.classList.add("default");
      btnClearFolder.hidden = true;
    }
  };

  // Restore persisted handle from IndexedDB
  try {
    const stored = await idbGet("dirHandle");
    if (stored) dirHandle = stored;
  } catch (_) {
    dirHandle = null;
  }
  updateFolderUI();

  // Choose folder
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

  // Clear chosen folder
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
  const KEY_META     = "fieldcap_meta";
  const KEY_CSV_JOB  = "fieldcap_csv_job";
  const KEY_CSV_CREW = "fieldcap_csv_crew";
  const KEY_CSV_BHA  = "fieldcap_csv_bha";

  // ── Restore cached CSV state ──────────────────────────────────────────────
  const storedAll = await chrome.storage.local.get([
    KEY_META, KEY_CSV_JOB, KEY_CSV_CREW, KEY_CSV_BHA,
  ]);

  const meta = storedAll[KEY_META];
  if (meta) {
    if (meta.jobId) jobIdInput.value = meta.jobId;
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
    if (storedAll[KEY_CSV_JOB] || storedAll[KEY_CSV_CREW] || storedAll[KEY_CSV_BHA]) {
      btnDownloadAll.disabled = false;
    }
  }

  // ── Fetch & Build ─────────────────────────────────────────────────────────
  btnFetch.addEventListener("click", () => {
    const jobId = parseInt(jobIdInput.value, 10);
    if (!jobId) { showErr("Enter a valid Job ID first."); return; }

    const flags = {
      jobDetails: chkJob.checked,
      crew:       chkCrew.checked,
      bha:        chkBha.checked,
    };

    if (!flags.jobDetails && !flags.crew && !flags.bha) {
      showErr("Select at least one data type to fetch.");
      return;
    }

    btnFetch.disabled = true;
    hideResult();

    if (flags.jobDetails) { setBadge(badgeJob,  "Fetching…", "busy"); metaJob.textContent  = ""; }
    if (flags.crew)       { setBadge(badgeCrew, "Fetching…", "busy"); metaCrew.textContent = ""; }
    if (flags.bha)        { setBadge(badgeBha,  "Fetching…", "busy"); metaBha.textContent  = ""; }

    showInfo("Fetching from FieldCap OData API…");

    chrome.runtime.sendMessage({ type: "FETCH_ALL", jobId, flags }, (res) => {
      btnFetch.disabled = false;

      if (chrome.runtime.lastError) {
        showErr(chrome.runtime.lastError.message);
        if (flags.jobDetails) setBadge(badgeJob,  "Error", "err");
        if (flags.crew)       setBadge(badgeCrew, "Error", "err");
        if (flags.bha)        setBadge(badgeBha,  "Error", "err");
        return;
      }

      if (!res?.ok) {
        showErr(res?.error ?? "Fetch failed. Make sure you are logged into FieldCap.");
        if (flags.jobDetails) setBadge(badgeJob,  "Error", "err");
        if (flags.crew)       setBadge(badgeCrew, "Error", "err");
        if (flags.bha)        setBadge(badgeBha,  "Error", "err");
        return;
      }

      if (flags.jobDetails) {
        setBadge(badgeJob, "ready", "ok");
        metaJob.textContent = `fieldcap-job-${jobId}-job-details.csv`;
      }
      if (flags.crew) {
        setBadge(badgeCrew, `${res.crewRows} rows`, "ok");
        metaCrew.textContent = `fieldcap-job-${jobId}-crew.csv`;
      }
      if (flags.bha) {
        setBadge(badgeBha, `${res.bhaRows} rows`, "ok");
        metaBha.textContent = `fieldcap-job-${jobId}-bha-equipment.csv · ${res.bhaCount} BHAs`;
      }
      btnDownloadAll.disabled = false;

      const folderNote = dirHandle ? ` → ${dirHandle.name}/` : "";
      const parts = [];
      if (flags.jobDetails) parts.push("job details");
      if (flags.crew)       parts.push(`${res.crewRows} crew`);
      if (flags.bha)        parts.push(`${res.bhaRows} BHA rows (${res.bhaCount} BHAs)`);
      const liveBha = Number(res.liveBhaRows ?? 0);
      const liveAct = Number(res.liveActivityRows ?? 0);

      let fetchMeta = `${new Date().toLocaleString()} · live bhaRows=${liveBha} · live activityRows=${liveAct}`;
      if (flags.bha && liveAct === 0) {
        fetchMeta += " · Open the FieldCap Activities tab for this job before Fetch to populate Mtrs Slid/Rot.";
      }

      showOk(`Fetched: ${parts.join(" · ")}${folderNote}`, fetchMeta);
    });
  });

  // ── Download All ──────────────────────────────────────────────────────────
  const ensureWritableFolder = async () => {
    if (!dirHandle) {
      // Bypass chrome.downloads policy restrictions by writing directly via
      // File System Access API.
      const handle = await window.showDirectoryPicker({ mode: "readwrite" });
      dirHandle = handle;
      await idbSet("dirHandle", handle);
      updateFolderUI();
    }

    let perm = await dirHandle.queryPermission({ mode: "readwrite" });
    if (perm !== "granted") perm = await dirHandle.requestPermission({ mode: "readwrite" });
    if (perm !== "granted") throw new Error("Write permission was denied.");
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const download = async (which) => {
    const jobId = parseInt(jobIdInput.value, 10);

    const keyMap  = { job: KEY_CSV_JOB, crew: KEY_CSV_CREW, bha: KEY_CSV_BHA };
    const stored  = await chrome.storage.local.get([keyMap[which], KEY_META]);
    const csv     = stored[keyMap[which]];
    const cached  = stored[KEY_META] ?? {};

    if (!csv) return null; // skip silently if not cached

    const id = jobId || cached.jobId || "export";
    const names = {
      job:  `fieldcap-job-${id}-job-details.csv`,
      crew: `fieldcap-job-${id}-crew.csv`,
      bha:  `fieldcap-job-${id}-bha-equipment.csv`,
    };
    const filename = names[which];

    if (!dirHandle) {
      throw new Error("Choose a download folder first.");
    }

    let perm = await dirHandle.queryPermission({ mode: "readwrite" });
    if (perm !== "granted") perm = await dirHandle.requestPermission({ mode: "readwrite" });
    if (perm !== "granted") throw new Error("Write permission was denied.");

    const writeOnce = async () => {
      const fh = await dirHandle.getFileHandle(filename, { create: true });
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

        if (/denied|used by another process|being used|lock|busy/i.test(msg)) {
          throw new Error(`${filename} is open/locked. Close it in Excel and retry.`);
        }

        // Some environments intermittently throw policy aborts on overwrite.
        // Retry in-place, then delete+recreate once, then final retry.
        if (/aborted due to security policy/i.test(msg) && attempt < 3) {
          if (attempt === 2) {
            try { await dirHandle.removeEntry(filename); } catch (_) {}
          }
          await sleep(350 * attempt);
          continue;
        }

        if (attempt < 3) {
          await sleep(250 * attempt);
          continue;
        }
      }
    }

    throw new Error(`${filename} failed to save: ${lastMsg}`);
  };

  btnDownloadAll.addEventListener("click", async () => {
    btnDownloadAll.disabled = true;
    try {
      await ensureWritableFolder();
      const saved = [];
      const failed = [];
      for (const which of ["job", "crew", "bha"]) {
        try {
          const name = await download(which);
          if (name) saved.push(name);
        } catch (e) {
          failed.push(String(e?.message ?? e));
        }
      }
      if (saved.length === 0 && failed.length === 0) {
        showErr("No cached data. Click Fetch first.");
        return;
      }
      if (failed.length > 0) {
        showErr(`Saved ${saved.length}/3 files. ${failed.join(" | ")}`);
        return;
      }
      const dest = dirHandle ? dirHandle.name : "Downloads";
      showOk(`Saved ${saved.length} file${saved.length > 1 ? "s" : ""} → ${dest}`);
    } catch (e) {
      showErr(e.message);
    } finally {
      btnDownloadAll.disabled = false;
    }
  });

  // ── Clear ─────────────────────────────────────────────────────────────────
  btnClear.addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "CLEAR_CACHE" }, () => {
      setBadge(badgeJob,  "\u2014", "");
      setBadge(badgeCrew, "\u2014", "");
      setBadge(badgeBha,  "\u2014", "");
      metaJob.textContent  = "";
      metaCrew.textContent = "";
      metaBha.textContent  = "";
      btnDownloadAll.disabled = true;
      hideResult();
    });
  });

  // ── Schema Probe — shows raw ToolAssembly field names from the API ────────
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
        if (res.statsKeys && res.statsKeys.length > 0) {
          lines.push(``, `Statistics sub-fields (${res.statsKeys.length}):`);
          lines.push(res.statsKeys.join("  |  "));
        } else {
          lines.push(``, `(no Statistics nav property — hours may be direct fields above)`);
        }

        resultArea.classList.add("visible");
        resultText.className  = "meta";
        resultText.style.whiteSpace = "pre-wrap";
        resultText.style.fontSize   = "9px";
        resultText.textContent = lines.join("\n");
        resultMeta.textContent = "Copy these field names to identify correct hour columns";
      });
    });
  }

  // ── Debug State — show what capture pipeline currently has ─────────────────
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

        const keys = [
          "fieldcap_bha_grid_rows",
          "fieldcap_intercepted_assemblies",
          "fieldcap_schema_probe",
          "fieldcap_meta",
        ];
        const stored = await chrome.storage.local.get(keys);
        const gridByJob = stored.fieldcap_bha_grid_rows ?? {};
        const gridForJob = gridByJob[String(jobId)] ?? {};
        const intercept = stored.fieldcap_intercepted_assemblies ?? {};
        const sampleGrid = Object.values(gridForJob)[0] ?? null;
        const sampleLive = liveBhaRows[0] ?? null;
        const sampleIntercept = Object.values(intercept)[0] ?? null;

        const metricKeysOf = (obj) => {
          if (!obj || typeof obj !== "object") return [];
          return Object.keys(obj).filter((k) => /met|meter|metre|foot|depth|drill|distance|\bmd\b/i.test(k));
        };

        const lines = [
          `Debug State for job ${jobId}`,
          ``,
          `Live SCRAPE_NOW bhaRows: ${liveBhaRows.length}`,
          `Cached grid rows (fieldcap_bha_grid_rows[jobId]): ${Object.keys(gridForJob).length}`,
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
          lines.push(``, `Sample CACHED canonical row keys:`, Object.keys(sampleGrid).join(" | "));
          const mk = metricKeysOf(sampleGrid);
          if (mk.length) lines.push(`Sample CACHED metric keys:`, mk.join(" | "));
        }
        if (sampleIntercept) {
          lines.push(``, `Sample INTERCEPT patch keys:`, Object.keys(sampleIntercept).join(" | "));
          const mk = metricKeysOf(sampleIntercept);
          if (mk.length) {
            lines.push(`Sample INTERCEPT metric values:`);
            for (const k of mk.slice(0, 8)) lines.push(`  ${k}: ${sampleIntercept[k]}`);
          }
        }

        resultArea.classList.add("visible");
        resultText.className = "meta";
        resultText.style.whiteSpace = "pre-wrap";
        resultText.style.fontSize = "9px";
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
