// FieldCap Data Exporter — Background Service Worker v2.5.0
// Fetches data directly from FieldCap's OData API using the active session.
// Produces 3 clean, properly-typed CSVs:
//   • fieldcap-job-{id}-job-details.csv   (core + custom fields, all labels from FieldCap UI)
//   • fieldcap-job-{id}-crew.csv          (one row per crew member)
//   • fieldcap-job-{id}-bha-equipment.csv (one row per BHA component)

"use strict";

// ── Constants ─────────────────────────────────────────────────────────────────
const BASE = "https://fieldcap-cdn.phxtech.com/odata";

const KEY_META     = "fieldcap_meta";
const KEY_CSV_JOB  = "fieldcap_csv_job";
const KEY_CSV_CREW = "fieldcap_csv_crew";
const KEY_CSV_BHA  = "fieldcap_csv_bha";
const KEY_INTERCEPT = "fieldcap_intercepted_assemblies"; // keyed by ToolAssemblyId
const KEY_BHA_GRID = "fieldcap_bha_grid_rows"; // { [jobId]: { [bhaNumber]: canonicalRow } }

// ── CSV helpers ───────────────────────────────────────────────────────────────
const csvEscape = (v) => {
  if (v === null || v === undefined) return "";
  const t = String(v);
  return /[",\r\n]/.test(t) ? `"${t.replaceAll('"', '""')}"` : t;
};

const csvRow = (cols, obj) => cols.map((c) => csvEscape(obj[c] ?? "")).join(",");

const buildCsvString = (columns, rows) => {
  const header = columns.join(",");
  const body = rows.map((r) => csvRow(columns, r)).join("\r\n");
  return body.length ? `${header}\r\n${body}` : header;
};

// ── Date / duration helpers ───────────────────────────────────────────────────
// OData dates arrive as:
//  • ISO strings:   "2026-04-19T00:00:00Z"
//  • Int64 (ms):    1777910647481  (FieldCap StartDate/EndDate fields)
//  • Legacy format: "/Date(ms)/"
const toDateStr = (v) => {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") {
    const n = Number(v);

    // Excel serial date (days since 1899-12-30)
    if (n > 20000 && n < 90000) {
      const dExcel = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      const y = dExcel.getUTCFullYear();
      if (y >= 2000 && y <= 2100) return dExcel.toISOString().slice(0, 10);
      return "";
    }

    // Unix seconds
    if (n >= 1000000000 && n < 10000000000) {
      const dSec = new Date(n * 1000);
      const y = dSec.getUTCFullYear();
      if (y >= 2000 && y <= 2100) return dSec.toISOString().slice(0, 10);
      return "";
    }

    // Unix milliseconds
    const dMs = new Date(n);
    const yMs = dMs.getUTCFullYear();
    if (yMs >= 2000 && yMs <= 2100) return dMs.toISOString().slice(0, 10);

    // 8-digit yyyymmdd integer
    const s = String(Math.trunc(n));
    if (/^\d{8}$/.test(s)) {
      const yyyy = Number(s.slice(0, 4));
      const mm   = Number(s.slice(4, 6));
      const dd   = Number(s.slice(6, 8));
      if (yyyy >= 2000 && yyyy <= 2100 && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
        return `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}`;
      }
    }

    // Unknown numeric date encoding for this tenant/version -> blank (not fake 1976)
    return "";
  }
  const legacyMatch = String(v).match(/\/Date\((-?\d+)\)\//);
  if (legacyMatch) {
    const dLegacy = new Date(parseInt(legacyMatch[1], 10));
    const y = dLegacy.getUTCFullYear();
    if (y >= 2000 && y <= 2100) return dLegacy.toISOString().slice(0, 10);
    return "";
  }
  const d = new Date(v);
  if (isNaN(d.getTime())) return "";
  const y = d.getUTCFullYear();
  if (y < 2000 || y > 2100) return "";
  return d.toISOString().slice(0, 10);
};

// Convert decimal hours or HH:MM string to 2-decimal hours
const toHours = (v) => {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "number") return v.toFixed(2);
  const m = String(v).match(/^(\d+):(\d{2})$/);
  if (m) return (parseInt(m[1], 10) + parseInt(m[2], 10) / 60).toFixed(2);
  return String(v);
};

const toNum = (v) => {
  if (v === null || v === undefined || v === "") return "";
  if (typeof v === "string") {
    const cleaned = v.replace(/,/g, "").trim();
    const nStr = cleaned.match(/-?\d+(\.\d+)?/)?.[0];
    if (nStr) {
      const n2 = Number(nStr);
      if (!isNaN(n2)) return n2;
    }
  }
  const n = Number(v);
  return isNaN(n) ? String(v) : n;
};

const parseHoursNum = (v) => {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return v;
  const m = String(v).match(/^(\d+):(\d{2})$/);
  if (m) return parseInt(m[1], 10) + (parseInt(m[2], 10) / 60);
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
};

const minutesToHours = (v) => {
  const n = parseHoursNum(v);
  if (n === null) return "";
  return (n / 60).toFixed(2);
};

const rangeDelta = (v) => {
  const s = String(v ?? "");
  const nums = s
    .replace(/,/g, "")
    .match(/-?\d+(\.\d+)?/g)
    ?.map(Number)
    .filter((n) => Number.isFinite(n));
  if (!nums || nums.length < 2) return "";
  const delta = Math.abs(nums[nums.length - 1] - nums[0]);
  return Number.isFinite(delta) ? Number(delta.toFixed(2)) : "";
};

const pickMetresHeuristic = (obj) => {
  const buckets = [obj, obj?.Statistics].filter(Boolean);
  const directNums = [];
  const startNums = [];
  const endNums = [];

  for (const b of buckets) {
    for (const [k, raw] of Object.entries(b)) {
      const key = String(k ?? "").toLowerCase();
      if (!key) continue;
      const n = parseHoursNum(raw);
      if (n === null || n <= 0) continue;

      const looksHour = /(hour|hours|hrs|toolhours|slide|rot|circ|below)/i.test(key);
      if (looksHour) continue;

      if (/(metre|meter|footage|drilled|distance)/i.test(key)) {
        directNums.push(n);
        continue;
      }
      if (/(start|from).*(depth|md)|\bfrom(md|depth)\b/i.test(key)) {
        startNums.push(n);
        continue;
      }
      if (/(end|to).*(depth|md)|\bto(md|depth)\b/i.test(key)) {
        endNums.push(n);
      }
    }
  }

  if (directNums.length) {
    const candidate = Math.max(...directNums);
    if (Number.isFinite(candidate)) return Number(candidate.toFixed(2));
  }

  let best = "";
  for (const s of startNums) {
    for (const e of endNums) {
      const d = Math.abs(e - s);
      if (!Number.isFinite(d) || d <= 0) continue;
      if (best === "" || d > best) best = d;
    }
  }
  return best === "" ? "" : Number(best.toFixed(2));
};

const pickMetres = (obj) => {
  const direct = pickNum(
    obj,
    "MetresDrilled",
    "MetersDrilled",
    "Metres Drilled",
    "Meters Drilled",
    "Meters Driled",
    "Meter Drilled",
    "Footage",
    "FootageDrilled",
    "DepthDrilled",
    "Metres",
    "Meters",
    "DrilledMeters",
    "DrilledMetres",
    "DistanceDrilled",
    "TotalMeters",
    "TotalMetres"
  );
  if (direct !== "") return direct;

  return (
    rangeDelta(
      pickRaw(
        obj,
        "StartEndDepth",
        "Start/End Depth",
        "Start-End Depth",
        "Depth Range"
      )
    ) ||
    rangeDelta(
      `${pickRaw(obj, "StartDepth", "DepthStart", "FromDepth")} - ${pickRaw(
        obj,
        "EndDepth",
        "DepthEnd",
        "ToDepth"
      )}`
    ) ||
    pickMetresHeuristic(obj) ||
    ""
  );
};

const pickRaw = (obj, ...keys) => {
  for (const k of keys) {
    const v = obj?.[k] ?? obj?.Statistics?.[k];
    if (v !== null && v !== undefined && v !== "") return v;
  }
  return "";
};

const normalizeKey = (v) => String(v ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const isFilled = (v) => !(v === null || v === undefined || v === "");
const mergeNonEmpty = (base, patch) => {
  const out = { ...(base ?? {}) };
  for (const [k, v] of Object.entries(patch ?? {})) {
    if (isFilled(v)) out[k] = v;
  }
  return out;
};

const getByAliases = (row, aliases) => {
  const index = new Map();
  for (const [k, v] of Object.entries(row ?? {})) index.set(normalizeKey(k), v);
  for (const a of aliases) {
    const hit = index.get(normalizeKey(a));
    if (hit !== null && hit !== undefined && hit !== "") return hit;
  }
  return "";
};

const toCanonicalBhaGridRow = (row) => {
  const bha = getByAliases(row, ["BHA", "BHA #", "Tool Assembly #", "ToolAssemblyNumber"]);
  if (bha === "") return null;
  const bhaNum = String(toNum(bha));

  return {
    bhaNum,
    row: {
      ToolAssemblyNumber: bhaNum,
      Section:          getByAliases(row, ["Section"]),
      Status:           getByAliases(row, ["Status"]),
      Motor:            getByAliases(row, ["Motor"]),
      GuidanceType:     getByAliases(row, ["Guidance"]),
      MetresDrilled:    getByAliases(row, ["Metres Drilled", "Meters Drilled", "Meters Driled", "Meter Drilled", "Metres", "Meters"]),
      StartEndDepth:    getByAliases(row, ["Start/End Depth", "Start-End Depth", "Depth Range"]),
      StartDepth:       getByAliases(row, ["Start Depth", "Depth Start", "From Depth"]),
      EndDepth:         getByAliases(row, ["End Depth", "Depth End", "To Depth"]),
      SlideHours:       getByAliases(row, ["Hrs Sld", "Hrs Slid", "Slide Hrs", "Slide Hours"]),
      RotateHours:      getByAliases(row, ["Hrs Rot", "Rotate Hrs", "Rotate Hours"]),
      CirculateHours:   getByAliases(row, ["Hrs Circ", "Circ Hrs", "Circulate Hours"]),
      TotalHours:       getByAliases(row, ["Total Hrs", "Total Hours"]),
      BelowRotateHours: getByAliases(row, ["Below Rot", "Below Rotate"]),
      ActivatedOn:      getByAliases(row, ["Date In", "Activated", "Activated On"]),
      CompletedOn:      getByAliases(row, ["Date Out", "Completed", "Completed On"]),
    },
  };
};

const jobIdFromUrl = (url) => {
  const m = String(url ?? "").match(/\/Jobs\/(\d+)\b/i);
  return m?.[1] ?? "";
};

const toBhaGridMap = (rows) => {
  const map = {};
  for (const raw of rows ?? []) {
    const c = toCanonicalBhaGridRow(raw);
    if (!c?.bhaNum) continue;
    map[c.bhaNum] = mergeNonEmpty(map[c.bhaNum] ?? {}, c.row);
  }
  return map;
};

const toActivityMetresMap = (rows) => {
  const valueFromKeyIncludes = (row, needles) => {
    for (const [k, v] of Object.entries(row ?? {})) {
      const nk = normalizeKey(k);
      if (needles.some((n) => nk.includes(n))) {
        if (v !== null && v !== undefined && v !== "") return v;
      }
    }
    return "";
  };

  const map = {};
  for (const row of rows ?? []) {
    const bha = getByAliases(row, ["BHA", "BHA #", "Tool Assembly #", "ToolAssemblyNumber"])
      || valueFromKeyIncludes(row, ["bha", "toolassembly"]);
    if (!bha) continue;
    const bhaNum = String(toNum(bha));
    if (!bhaNum) continue;

    const activity = String(
      getByAliases(row, ["Activity/Code", "Activity Code", "Activity", "Code"])
      || valueFromKeyIncludes(row, ["activitycode", "activity", "code"])
    ).toLowerCase();

    const courseRaw = String(
      getByAliases(row, ["Course", "Metres", "Meters", "Metres Drilled", "Meters Drilled"])
      || valueFromKeyIncludes(row, ["course", "metre", "meter", "distance", "drill"])
      || ""
    );
    let course = Number((courseRaw.match(/-?\d+(\.\d+)?/) ?? [])[0]);

    // If "Course" is blank, derive meters from depth range.
    if (!Number.isFinite(course) || course <= 0) {
      const depthRange = String(
        getByAliases(row, ["Start - End Depth", "Start-End Depth", "Start/End Depth", "Depth Range"])
        || valueFromKeyIncludes(row, ["startenddepth", "depthrange", "startend"])
        || ""
      );
      const delta = rangeDelta(depthRange);
      course = Number(delta);
    }

    if (!Number.isFinite(course) || course <= 0) continue;

    const prev = map[bhaNum] ?? { slide: 0, rot: 0 };
    if (/\bslid(e|ing)?\b|code\s*2a\b/.test(activity)) {
      prev.slide += course;
    } else if (/\brotat(e|ing)?\b|code\s*2\b/.test(activity)) {
      prev.rot += course;
    }
    map[bhaNum] = prev;
  }

  const out = {};
  for (const [bhaNum, v] of Object.entries(map)) {
    out[bhaNum] = {
      SlideMetres: Number(v.slide.toFixed(2)),
      RotateMetres: Number(v.rot.toFixed(2)),
    };
  }
  return out;
};

const collectObjectsDeep = (root, out = [], depth = 0) => {
  if (depth > 8 || root === null || root === undefined) return out;
  if (Array.isArray(root)) {
    for (const v of root) collectObjectsDeep(v, out, depth + 1);
    return out;
  }
  if (typeof root === "object") {
    out.push(root);
    for (const v of Object.values(root)) collectObjectsDeep(v, out, depth + 1);
  }
  return out;
};

const toCanonicalAssemblyPatch = (obj) => {
  const id = getByAliases(obj, ["ToolAssemblyId", "Tool Assembly Id", "ToolAssemblyKey"]);
  const bhaNum = getByAliases(obj, ["ToolAssemblyNumber", "Tool Assembly #", "BHA", "BHA #"]);

  const patch = {
    ToolAssemblyNumber: bhaNum,
    Section:          getByAliases(obj, ["Section", "BHA Section"]),
    Status:           getByAliases(obj, ["Status", "Assembly Status"]),
    Motor:            getByAliases(obj, ["Motor", "Motor Serial", "MotorDescription"]),
    GuidanceType:     getByAliases(obj, ["Guidance", "Guidance Type", "GuidanceType"]),
    MetresDrilled:    getByAliases(obj, ["Metres Drilled", "Meters Drilled", "Meters Driled", "Meter Drilled", "MetresDrilled", "MetersDrilled", "Footage"]),
    SlideHours:       getByAliases(obj, ["Hrs Sld", "Hrs Slid", "Slide Hrs", "Slide Hours", "SlideHours"]),
    RotateHours:      getByAliases(obj, ["Hrs Rot", "Rotate Hrs", "Rotate Hours", "RotateHours"]),
    CirculateHours:   getByAliases(obj, ["Hrs Circ", "Circ Hrs", "Circulate Hours", "CirculateHours"]),
    TotalHours:       getByAliases(obj, ["Total Hrs", "Total Hours", "TotalHours"]),
    BelowRotateHours: getByAliases(obj, ["Below Rot", "Below Rotate", "BelowRotateHours"]),
    ActivatedOn:      getByAliases(obj, ["Activated", "Activated On", "Date In", "ActivatedOn"]),
    CompletedOn:      getByAliases(obj, ["Completed", "Completed On", "Date Out", "CompletedOn"]),
    StartEndDepth:    getByAliases(obj, ["Start/End Depth", "Start-End Depth", "Depth Range"]),
    StartDepth:       getByAliases(obj, ["StartDepth", "Start Depth", "DepthStart", "FromDepth", "From Depth"]),
    EndDepth:         getByAliases(obj, ["EndDepth", "End Depth", "DepthEnd", "ToDepth", "To Depth"]),
  };

  // Preserve tenant-specific depth/meter fields so pickMetresHeuristic can evaluate
  // them later, even when names differ from our known aliases.
  const addMetricCandidates = (source) => {
    if (!source || typeof source !== "object") return;
    for (const [k, v] of Object.entries(source)) {
      if (v === null || v === undefined || v === "") continue;
      const key = String(k);
      const lk = key.toLowerCase();
      if (/(hour|hours|hrs|toolhours|slide|rot|circ|below)/i.test(lk)) continue;
      if (!/(met|meter|metre|foot|depth|drill|distance|\bmd\b|from|to)/i.test(lk)) continue;
      if (!(key in patch)) patch[key] = v;
    }
  };
  addMetricCandidates(obj);
  addMetricCandidates(obj?.Statistics);

  const hasUseful = Object.values(patch).some((v) => v !== null && v !== undefined && v !== "");
  if (!hasUseful) return null;

  return {
    id: id !== "" ? String(id) : "",
    bhaNum: bhaNum !== "" ? String(toNum(bhaNum)) : "",
    patch,
  };
};

const scrapeFromTab = async (tabId) => new Promise((resolve) => {
  chrome.tabs.sendMessage(tabId, { type: "SCRAPE_NOW" }, (res) => {
    const firstErr = chrome.runtime.lastError?.message ?? "";
    if (!firstErr) {
      return resolve({
        bhaRows: Array.isArray(res?.bhaRows) ? res.bhaRows : [],
        activityRows: Array.isArray(res?.activityRows) ? res.activityRows : [],
      });
    }

    // If the content script is not attached yet, inject and retry once.
    if (!/receiving end does not exist|could not establish connection/i.test(firstErr)) {
      return resolve({ bhaRows: [], activityRows: [] });
    }
    if (!chrome.scripting?.executeScript) return resolve({ bhaRows: [], activityRows: [] });

    chrome.scripting.executeScript(
      { target: { tabId }, files: ["content.js"] },
      () => {
        chrome.tabs.sendMessage(tabId, { type: "SCRAPE_NOW" }, (res2) => {
          if (chrome.runtime.lastError) return resolve({ bhaRows: [], activityRows: [] });
          resolve({
            bhaRows: Array.isArray(res2?.bhaRows) ? res2.bhaRows : [],
            activityRows: Array.isArray(res2?.activityRows) ? res2.activityRows : [],
          });
        });
      }
    );
  });
});

const scrapeNowFromFieldCapTabs = async (jobId) => {
  const tabs = await chrome.tabs.query({ url: ["*://*.phxtech.com/*"] });
  if (!tabs?.length) return { bhaRows: [], activityRows: [] };
  const fieldcapTabs = tabs.filter((t) => /fieldcap/i.test(String(t.url ?? "")));
  if (!fieldcapTabs.length) return { bhaRows: [], activityRows: [] };

  const results = await Promise.all(
    fieldcapTabs
      .filter((t) => !!t?.id)
      .map((t) => new Promise((resolve) => {
        scrapeFromTab(t.id).then((data) => resolve({ tab: t, ...data }));
      }))
  );

  // Prefer tabs whose URL clearly includes this job id; otherwise choose max rows.
  const forJob = results
    .filter((r) => new RegExp(`/Jobs/${jobId}\\b`, "i").test(String(r.tab?.url ?? "")))
    .sort((a, b) => b.bhaRows.length - a.bhaRows.length);

  const source = forJob.length ? forJob : results;
  const best = source.sort((a, b) => b.bhaRows.length - a.bhaRows.length)[0];
  const activityRows = source.flatMap((r) => r.activityRows ?? []);
  return {
    bhaRows: best?.bhaRows ?? [],
    activityRows,
  };
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// ── OData fetch helper ────────────────────────────────────────────────────────
const odataGet = async (path) => {
  const url = `${BASE}/${path}`;
  const res = await fetch(url, {
    credentials: "include",
    headers: { Accept: "application/json;odata.metadata=minimal" },
  });
  if (!res.ok) throw new Error(`OData ${res.status}: ${path}`);
  const json = await res.json();
  return json.value ?? json;
};

// ── Job Details — core entity ─────────────────────────────────────────────────
const fetchJobDetails = async (jobId) => {
  const filter = encodeURIComponent(
    `(ClientJobId eq ${jobId}) and (null eq DeletedBy)`
  );
  const rows = await odataGet(
    `ClientJobs?$expand=Client,JobStatus&$filter=${filter}&$top=1`
  );
  return rows[0] ?? null;
};

// ── Job Details — custom key-value pairs ──────────────────────────────────────
// Each record: { ClientJobCustomValueKey, ClientJobId, KeyName, Value, ... }
const fetchJobCustomValues = async (jobId) => {
  const filter = encodeURIComponent(
    `(ClientJobId eq ${jobId}) and (null eq DeletedBy)`
  );
  return odataGet(`ClientJobCustomValues?$filter=${filter}`);
};

// ── Crew / Personnel ──────────────────────────────────────────────────────────
const fetchCrew = async (jobId) => {
  const filter = encodeURIComponent(
    `(null eq DeletedBy) and (ClientJobId eq ${jobId})`
  );
  return odataGet(
    `JobSchedules?$expand=User,WorkType&$filter=${filter}&$orderby=LineNumber`
  );
};

// ── BHA Assemblies ────────────────────────────────────────────────────────────
// Try to expand Statistics navigation property (holds aggregated hour/footage data).
// Falls back to bare fetch if the server doesn't expose that nav property.
const fetchToolAssemblies = async (jobId) => {
  const filter = encodeURIComponent(
    `(ClientJobId eq ${jobId}) and (DeletedOn eq null)`
  );
  try {
    return await odataGet(
      `ToolAssemblies?$filter=${filter}&$expand=${encodeURIComponent("Statistics")}`
    );
  } catch (_) {
    return odataGet(`ToolAssemblies?$filter=${filter}`);
  }
};

// ── BHA Component Items ───────────────────────────────────────────────────────
const fetchToolAssemblyItems = async (jobId) => {
  const expand = [
    "JobTool",
    "JobTool($expand=Item)",
    "JobTool($expand=ItemSerial)",
    "JobTool($expand=ItemSerial($expand=OwnerBusinessPartner))",
    "ToolAssembly",
  ].join(",");
  const filter = encodeURIComponent(
    `(null eq DeletedBy) and (ToolAssembly/ClientJobId eq ${jobId})`
  );
  return odataGet(`ToolAssemblyItems?$expand=${expand}&$filter=${filter}`);
};

// ── Tools master list ─────────────────────────────────────────────────────────
const fetchJobTools = async (jobId) => {
  const expand = [
    "Item",
    "ItemSerial",
    "ItemSerial($expand=Item)",
  ].join(",");
  const filter = encodeURIComponent(
    `(null eq DeletedBy) and (ClientJobId eq ${jobId})`
  );
  return odataGet(`JobTools?$expand=${expand}&$filter=${filter}`);
};

// ── Normalize: Job Details ────────────────────────────────────────────────────
// Returns a flat object merging:
//   • Core named fields from ClientJobs (always present, labeled clearly)
//   • Custom key-value pairs from ClientJobCustomValues (tenant-configured labels)
// Column order: core fields first, then custom fields in the order they arrive.
const normalizeJobDetails = (j, customValues) => {
  if (!j) return {};

  const row = {
    "Job ID":             j.ClientJobId ?? "",
    "Job Name":           j.ClientJobName ?? "",
    "Job Code":           j.JobCode ?? "",
    "Client":             j.Client?.ClientName ?? j.ClientName ?? "",
    "Job Type":           j.JobType ?? "",
    "Ops Status":         j.JobStatus?.JobStatusName ?? "",
    "AFE (core)":         j.Afe ?? "",
    "Well (core)":        j.Well ?? "",
    "Description":        j.Description ?? "",
    "PO":                 j.Po ?? "",
    "Work Order":         j.WorkOrder ?? "",
    "Cost Center":        j.CostCenter ?? "",
    "County":             j.County ?? "",
    "City":               j.City ?? "",
    "Start Date":         j.StartDate    ? toDateStr(j.StartDate)         : "",
    "End Date":           j.EndDate      ? toDateStr(j.EndDate)           : "",
    "Planned Start Date": j.PlannedStartDate ? toDateStr(j.PlannedStartDate) : "",
    "Planned End Date":   j.PlannedEndDate   ? toDateStr(j.PlannedEndDate)   : "",
  };

  // Merge custom fields — KeyName becomes the column header, Value is the cell
  for (const cv of customValues ?? []) {
    const label = (cv.KeyName ?? "").trim();
    if (!label) continue;
    // Skip blanks but keep explicit empty strings
    row[label] = cv.Value ?? "";
  }

  return row;
};

// ── Build Job Details CSV ─────────────────────────────────────────────────────
const buildJobDetailsCsv = (jobDetails, customValues) => {
  const row = normalizeJobDetails(jobDetails, customValues);
  const columns = Object.keys(row);
  return buildCsvString(columns, [row]);
};

// ── Normalize: Crew row ───────────────────────────────────────────────────────
const CREW_COLUMNS = [
  "Job ID", "Name", "Employee Code", "Email", "Phone", "Role", "Work Type",
  "Start Date", "End Date", "Line Number",
];

const normalizeCrewRow = (s, jobId) => ({
  "Job ID":        jobId,
  "Name":          [s.User?.FirstName, s.User?.LastName].filter(Boolean).join(" "),
  "Employee Code": s.User?.EmployeeCode ?? s.User?.UserCode ?? "",
  "Email":         s.User?.Email ?? "",
  "Phone":         s.User?.Phone ?? s.User?.MobilePhone ?? "",
  "Role":          s.WorkType?.WorkTypeName ?? s.Role ?? "",
  "Work Type":     s.WorkType?.WorkTypeName ?? "",
  "Start Date":    toDateStr(s.StartDate),
  "End Date":      toDateStr(s.EndDate),
  "Line Number":   toNum(s.LineNumber),
});

// ── Multi-variant field resolvers ─────────────────────────────────────────────
// FieldCap's OData schema uses inconsistent naming across API versions.
// Each helper walks through candidate key names (including inside a nested
// Statistics sub-object if the $expand returned one) and returns the first hit.
const pickHours = (obj, ...keys) => {
  for (const k of keys) {
    const v = obj?.[k] ?? obj?.Statistics?.[k];
    if (v !== null && v !== undefined && v !== "") return toHours(v);
  }
  return "";
};

const pickNum = (obj, ...keys) => {
  for (const k of keys) {
    const v = obj?.[k] ?? obj?.Statistics?.[k];
    if (v !== null && v !== undefined && v !== "") return toNum(v);
  }
  return "";
};

// ── Normalize: BHA Equipment ──────────────────────────────────────────────────
const BHA_COLUMNS = [
  "Job ID", "BHA #", "Section", "Status", "Motor", "Guidance",
  "Metres Drilled", "BHA Mtrs Slid", "BHA Mtrs Rot", "BHA Total Hrs", "BHA Hrs Slid", "BHA Hrs Rot", "BHA Hrs Circ", "BHA Below Rot",
  "Activated On", "Completed On",
  "Serial #", "Item Code", "Description", "Sub Description",
  "Length", "Accum Length", "Top", "Bottom", "Max OD", "Min ID",
  "Job Hours", "HSLS", "Strapped", "Shipping Status", "Dispatched On", "Returned On",
];

const normalizeBhaRow = (assembly, item, jobTool) => {
  const itemSerial = item?.JobTool?.ItemSerial;
  const itemRecord = item?.JobTool?.Item ?? itemSerial?.Item;

  // Tenant-specific fallback observed in this FieldCap instance:
  // ToolHours1=Rotate (minutes), ToolHours2=Slide (minutes),
  // ToolHours3=Circ (minutes), ToolHours4=Below Rot (minutes)
  const tH1 = pickRaw(assembly, "ToolHours1");
  const tH2 = pickRaw(assembly, "ToolHours2");
  const tH3 = pickRaw(assembly, "ToolHours3");
  const tH4 = pickRaw(assembly, "ToolHours4");

  let hrsSlide = pickHours(assembly,
    "SlideHours", "SlideHrs", "SlidingHours", "HoursSliding", "HrsSlide"
  );
  let hrsRot = pickHours(assembly,
    "RotateHours", "RotateHrs", "RotatingHours", "HoursRotating", "HrsRotate"
  );
  let hrsCirc = pickHours(assembly,
    "CirculateHours", "CircHours", "CirculatingHours", "HoursCirculating", "HrsCirc", "CircHrs"
  );
  if (!hrsSlide) hrsSlide = minutesToHours(tH2);
  if (!hrsRot)   hrsRot   = minutesToHours(tH1);
  if (!hrsCirc)  hrsCirc  = minutesToHours(tH3);

  let hrsTotal = pickHours(assembly,
    "TotalHours", "TotalHrs", "TotalDrillingHours", "HoursTotal", "DrillHours"
  );
  let hrsBelowRot = pickHours(assembly,
    "BelowRotateHours", "BelowRotHours", "BelowRotatingHours", "HoursBelowRotate", "BelowRotHrs"
  );
  if (!hrsBelowRot) hrsBelowRot = minutesToHours(tH4);

  // If tenant stores only slide/rot/circ in ToolHours slots, derive total hours.
  if (!hrsTotal) {
    const s = parseHoursNum(hrsSlide);
    const r = parseHoursNum(hrsRot);
    const c = parseHoursNum(hrsCirc);
    if (s !== null || r !== null || c !== null) {
      hrsTotal = ((s ?? 0) + (r ?? 0) + (c ?? 0)).toFixed(2);
    }
  }

  const totalMetresVal = parseHoursNum(pickMetres(assembly));
  let mtrsSlide = pickNum(assembly,
    "SlideMetres", "SlideMeters", "SlidingMetres", "SlidingMeters", "SlideAmount", "Slide Meter", "Slide Meters"
  );
  let mtrsRot = pickNum(assembly,
    "RotateMetres", "RotateMeters", "RotatingMetres", "RotatingMeters", "RotateAmount", "Rotate Meter", "Rotate Meters"
  );
  const sMeters = parseHoursNum(mtrsSlide);
  const rMeters = parseHoursNum(mtrsRot);
  if ((rMeters === null || rMeters === 0) && totalMetresVal !== null && sMeters !== null && totalMetresVal >= sMeters) {
    mtrsRot = Number((totalMetresVal - sMeters).toFixed(2));
  }
  if ((sMeters === null || sMeters === 0) && totalMetresVal !== null && rMeters !== null && totalMetresVal >= rMeters) {
    mtrsSlide = Number((totalMetresVal - rMeters).toFixed(2));
  }

  // Last-resort fallback: if split metres are still missing, infer from
  // slide/rotate hour proportions for this BHA.
  const sMeters2 = parseHoursNum(mtrsSlide);
  const rMeters2 = parseHoursNum(mtrsRot);
  const sHrsNum = parseHoursNum(hrsSlide);
  const rHrsNum = parseHoursNum(hrsRot);
  if (
    totalMetresVal !== null && totalMetresVal > 0 &&
    (sMeters2 === null || sMeters2 === 0) &&
    (rMeters2 === null || rMeters2 === 0) &&
    sHrsNum !== null && rHrsNum !== null && (sHrsNum + rHrsNum) > 0
  ) {
    const inferredSlide = Number((totalMetresVal * (sHrsNum / (sHrsNum + rHrsNum))).toFixed(2));
    const inferredRot = Number((totalMetresVal - inferredSlide).toFixed(2));
    mtrsSlide = inferredSlide;
    mtrsRot = inferredRot;
  }

  return {
    "Job ID":          assembly.ClientJobId ?? "",
    "BHA #":           toNum(assembly.ToolAssemblyNumber),
    "Section":         assembly.Section ?? assembly.SectionName ?? assembly.BHASection ?? "",
    "Status":          assembly.Status ?? assembly.AssemblyStatus ?? assembly.AssemblyStatusName ?? "",
    "Motor":           assembly.Motor ?? assembly.MotorSerial ?? assembly.MotorDescription ?? "",
    "Guidance":        assembly.GuidanceType ?? assembly.Guidance ?? assembly.GuidanceName ?? "",
    "Metres Drilled":  pickMetres(assembly),
    "BHA Mtrs Slid":   mtrsSlide,
    "BHA Mtrs Rot":    mtrsRot,
    "BHA Total Hrs":   hrsTotal,
    "BHA Hrs Slid":    hrsSlide,
    "BHA Hrs Rot":     hrsRot,
    "BHA Hrs Circ":    hrsCirc,
    "BHA Below Rot":   hrsBelowRot,
    "Activated On":    toDateStr(assembly.ActivatedOn),
    "Completed On":    toDateStr(assembly.CompletedOn),
    "Serial #":        itemSerial?.SerialNumber ?? item?.SerialNumber ?? "",
    "Item Code":       itemRecord?.ItemCode ?? item?.ItemCode ?? "",
    "Description":     itemRecord?.ItemName ?? item?.Description ?? "",
    "Sub Description": itemRecord?.Description ?? item?.SubDescription ?? item?.JobTool?.Notes ?? "",
    "Length":          toNum(item?.Length),
    "Accum Length":    toNum(item?.AccumulatedLength),
    "Top":             item?.TopConnection ?? "",
    "Bottom":          item?.BottomConnection ?? "",
    "Max OD":          toNum(item?.MaxOD ?? itemRecord?.OD),
    "Min ID":          toNum(item?.MinID ?? itemRecord?.ID),
    "Job Hours":       toHours(jobTool?.JobHours),
    "HSLS":            toHours(jobTool?.HslsHours),
    "Strapped":        jobTool?.Strapped ?? "",
    "Shipping Status": jobTool?.ShippingStatus ?? "",
    "Dispatched On":   toDateStr(jobTool?.DispatchedOn),
    "Returned On":     toDateStr(jobTool?.ReturnedOn),
  };
};

// ── Build BHA CSV ─────────────────────────────────────────────────────────────
const buildBhaCsv = (assemblies, assemblyItems, jobTools, interceptedMap = {}, bhaGridMap = {}) => {
  const toolBySerial = new Map();
  for (const t of jobTools) {
    const serial = t.ItemSerial?.SerialNumber ?? t.SerialNumber;
    if (serial) toolBySerial.set(serial, t);
  }

  const itemsByAssembly = new Map();
  for (const item of assemblyItems) {
    const aid = item.ToolAssemblyId ?? item.ToolAssembly?.ToolAssemblyId;
    if (!aid) continue;
    if (!itemsByAssembly.has(aid)) itemsByAssembly.set(aid, []);
    itemsByAssembly.get(aid).push(item);
  }

  const sorted = [...assemblies].sort((a, b) => {
    const an = Number(a.ToolAssemblyNumber) || 0;
    const bn = Number(b.ToolAssemblyNumber) || 0;
    return bn - an;
  });

  const rows = [];
  for (const asm of sorted) {
    const asmId = asm.ToolAssemblyId;
    // Merge in hours from both sources:
    // 1) intercepted OData payloads keyed by ToolAssemblyId
    // 2) BHA grid rows keyed by ToolAssemblyNumber
    const intercepted = interceptedMap[String(asmId)] ?? {};
    const grid = bhaGridMap[String(toNum(asm.ToolAssemblyNumber))] ?? {};
    // IMPORTANT: asm goes first, then enrichments override null/empty values.
    const merged = mergeNonEmpty(mergeNonEmpty(asm, intercepted), grid);
    const items = itemsByAssembly.get(asmId) ?? [];

    if (items.length === 0) {
      rows.push(normalizeBhaRow(merged, {}, null));
      continue;
    }

    for (const item of items) {
      const serial = item.JobTool?.ItemSerial?.SerialNumber ?? item.SerialNumber;
      const jobTool = serial ? (toolBySerial.get(serial) ?? null) : null;
      rows.push(normalizeBhaRow(merged, item, jobTool));
    }
  }

  return buildCsvString(BHA_COLUMNS, rows);
};

// ── Main fetch orchestrator ───────────────────────────────────────────────────
const fetchAll = async (jobId, flags, liveBhaRows = [], liveActivityRows = []) => {
  const results = {};

  if (flags.jobDetails) {
    const [raw, customValues] = await Promise.all([
      fetchJobDetails(jobId),
      fetchJobCustomValues(jobId),
    ]);
    results.jobDetailsCsv = buildJobDetailsCsv(raw, customValues);
  }

  if (flags.crew) {
    const raw = await fetchCrew(jobId);
    results.crew = raw.map((s) => normalizeCrewRow(s, jobId));
    results.crewCsv = buildCsvString(CREW_COLUMNS, results.crew);
  }

  if (flags.bha) {
    const [assemblies, items, tools, stored] = await Promise.all([
      fetchToolAssemblies(jobId),
      fetchToolAssemblyItems(jobId),
      fetchJobTools(jobId),
      new Promise((res) => chrome.storage.local.get([KEY_INTERCEPT, KEY_BHA_GRID], res)),
    ]);
    const interceptedMap = stored[KEY_INTERCEPT] ?? {};
    const bhaGridRowsByJob = stored[KEY_BHA_GRID] ?? {};
    const cachedBhaGridMap = bhaGridRowsByJob[String(jobId)] ?? {};
    const liveBhaGridMap = toBhaGridMap(liveBhaRows);
    const liveActivityMap = toActivityMetresMap(liveActivityRows);
    const bhaGridMap = {};
    for (const [k, v] of Object.entries(cachedBhaGridMap)) bhaGridMap[k] = mergeNonEmpty(bhaGridMap[k], v);
    for (const [k, v] of Object.entries(liveBhaGridMap)) bhaGridMap[k] = mergeNonEmpty(bhaGridMap[k], v);
    for (const [k, v] of Object.entries(liveActivityMap)) bhaGridMap[k] = mergeNonEmpty(bhaGridMap[k], v);
    results.bhaCsv      = buildBhaCsv(assemblies, items, tools, interceptedMap, bhaGridMap);
    results.bhaRowCount = results.bhaCsv.split("\r\n").length - 1;
    results.bhaCount    = assemblies.length;
  }

  return results;
};

// ── Message handler ───────────────────────────────────────────────────────────
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {

  if (message.type === "PING") {
    sendResponse({ ok: true });
    return false;
  }

  // ── Fetch all data from OData API ─────────────────────────────────────────
  if (message.type === "FETCH_ALL") {
    const { jobId, flags = { jobDetails: true, crew: true, bha: true } } = message;
    if (!jobId) {
      sendResponse({ ok: false, error: "Job ID is required." });
      return false;
    }

    (async () => {
      let liveData = flags.bha ? await scrapeNowFromFieldCapTabs(jobId) : { bhaRows: [], activityRows: [] };
      if (flags.bha && liveData.bhaRows.length === 0) {
        // FieldCap is an SPA; tables can render shortly after initial request.
        await sleep(1200);
        liveData = await scrapeNowFromFieldCapTabs(jobId);
      }

      if (flags.bha && (liveData.bhaRows.length > 0 || liveData.activityRows.length > 0)) {
        const liveMap = toBhaGridMap(liveData.bhaRows);
        const liveActivityMap = toActivityMetresMap(liveData.activityRows);
        chrome.storage.local.get([KEY_BHA_GRID], (storedGrid) => {
          const byJob = storedGrid[KEY_BHA_GRID] ?? {};
          const mergedJob = { ...(byJob[String(jobId)] ?? {}) };
          for (const [k, v] of Object.entries(liveMap)) mergedJob[k] = mergeNonEmpty(mergedJob[k], v);
          for (const [k, v] of Object.entries(liveActivityMap)) mergedJob[k] = mergeNonEmpty(mergedJob[k], v);
          byJob[String(jobId)] = mergedJob;
          chrome.storage.local.set({ [KEY_BHA_GRID]: byJob });
        });
      }

      const results = await fetchAll(jobId, flags, liveData.bhaRows, liveData.activityRows);
        const meta = {
          jobId,
          builtAt:      new Date().toISOString(),
          crewRows:     results.crew?.length ?? 0,
          bhaRows:      results.bhaRowCount  ?? 0,
          bhaCount:     results.bhaCount     ?? 0,
          liveBhaRows:  liveData.bhaRows.length,
          liveActivityRows: liveData.activityRows.length,
          flags,
        };

        const stored = {};
        if (results.jobDetailsCsv) stored[KEY_CSV_JOB]  = results.jobDetailsCsv;
        if (results.crewCsv)       stored[KEY_CSV_CREW] = results.crewCsv;
        if (results.bhaCsv)        stored[KEY_CSV_BHA]  = results.bhaCsv;
        stored[KEY_META] = meta;

        chrome.storage.local.set(stored, () => {
          sendResponse({ ok: true, ...meta });
        });
      })().catch((err) => {
      sendResponse({ ok: false, error: err.message });
    });

    return true;
  }

// ── Download a cached CSV ────────────────────────────────────────────────
  if (message.type === "DOWNLOAD_CSV") {
    const { which, jobId } = message;
    const keyMap = { job: KEY_CSV_JOB, crew: KEY_CSV_CREW, bha: KEY_CSV_BHA };
    const key = keyMap[which];
    if (!key) {
      sendResponse({ ok: false, error: `Unknown CSV type: ${which}` });
      return false;
    }

    chrome.storage.local.get([key, KEY_META], (stored) => {
      const csv  = stored[key];
      const meta = stored[KEY_META] ?? {};

      if (!csv) {
        sendResponse({ ok: false, error: `No ${which} CSV cached. Click Fetch first.` });
        return;
      }

      const jid = jobId ?? meta.jobId ?? "export";
      const filename = {
        job:  `fieldcap-job-${jid}-job-details.csv`,
        crew: `fieldcap-job-${jid}-crew.csv`,
        bha:  `fieldcap-job-${jid}-bha-equipment.csv`,
      }[which];

      const blob   = new Blob([csv], { type: "text/csv;charset=utf-8" });
      const reader = new FileReader();
      reader.onload = () => {
        chrome.downloads.download({ url: reader.result, filename, saveAs: true });
        sendResponse({ ok: true, filename });
      };
      reader.readAsDataURL(blob);
    });
    return true;
  }

  // ── Cache OData responses intercepted from the FieldCap app itself ────────
  // When the user browses BHA pages, content.js forwards the app's own API
  // responses here. We extract assembly-level hour/footage fields and store
  // them keyed by ToolAssemblyId so buildBhaCsv can merge them in.
  if (message.type === "INTERCEPTED_ODATA") {
    const { data, url = "" } = message;
    if (!data) { sendResponse({ ok: true }); return false; }

    const objects = collectObjectsDeep(data);
    if (objects.length === 0) { sendResponse({ ok: true }); return false; }
    const jobId = jobIdFromUrl(url);

    chrome.storage.local.get([KEY_INTERCEPT, KEY_BHA_GRID], (stored) => {
      const cacheById = stored[KEY_INTERCEPT] ?? {};
      const gridByJob = stored[KEY_BHA_GRID] ?? {};
      const jobMap = jobId ? (gridByJob[jobId] ?? {}) : {};
      let changedId = false;
      let changedGrid = false;
      let captured = 0;

      for (const obj of objects) {
        const c = toCanonicalAssemblyPatch(obj);
        if (!c) continue;
        captured += 1;
        if (c.id) {
          cacheById[c.id] = mergeNonEmpty(cacheById[c.id] ?? {}, c.patch);
          changedId = true;
        }
        if (jobId && c.bhaNum) {
          jobMap[c.bhaNum] = mergeNonEmpty(jobMap[c.bhaNum] ?? {}, c.patch);
          changedGrid = true;
        }
      }

      if (changedGrid && jobId) gridByJob[jobId] = jobMap;

      const setObj = {};
      if (changedId) setObj[KEY_INTERCEPT] = cacheById;
      if (changedGrid) setObj[KEY_BHA_GRID] = gridByJob;

      if (Object.keys(setObj).length > 0) chrome.storage.local.set(setObj);
      sendResponse({ ok: true, captured, byId: changedId, byGrid: changedGrid });
    });
    return true;
  }

  // ── Cache BHA grid rows from content AUTO_SCRAPE ──────────────────────────
  if (message.type === "AUTO_SCRAPE") {
    const { bhaRows = [], activityRows = [], url = "" } = message;
    const jobId = jobIdFromUrl(url);
    if (!jobId || ((!Array.isArray(bhaRows) || bhaRows.length === 0) && (!Array.isArray(activityRows) || activityRows.length === 0))) {
      sendResponse({ ok: true });
      return false;
    }

    chrome.storage.local.get([KEY_BHA_GRID], (stored) => {
      const byJob = stored[KEY_BHA_GRID] ?? {};
      const jobMap = byJob[jobId] ?? {};
      let changed = false;

      for (const rawRow of bhaRows) {
        const c = toCanonicalBhaGridRow(rawRow);
        if (!c?.bhaNum) continue;
        jobMap[c.bhaNum] = mergeNonEmpty(jobMap[c.bhaNum] ?? {}, c.row);
        changed = true;
      }

      const actMap = toActivityMetresMap(activityRows);
      for (const [bhaNum, m] of Object.entries(actMap)) {
        jobMap[bhaNum] = mergeNonEmpty(jobMap[bhaNum] ?? {}, m);
        changed = true;
      }

      if (changed) {
        byJob[jobId] = jobMap;
        chrome.storage.local.set({ [KEY_BHA_GRID]: byJob }, () =>
          sendResponse({ ok: true, cachedBhaRows: Object.keys(jobMap).length })
        );
      } else {
        sendResponse({ ok: true, cachedBhaRows: Object.keys(jobMap).length });
      }
    });
    return true;
  }

  // ── Clear cached data ─────────────────────────────────────────────────────
  if (message.type === "CLEAR_CACHE") {
    chrome.storage.local.remove(
      [KEY_META, KEY_CSV_JOB, KEY_CSV_CREW, KEY_CSV_BHA, KEY_INTERCEPT, KEY_BHA_GRID],
      () => sendResponse({ ok: true })
    );
    return true;
  }

  // ── Schema probe — fetch 1 raw assembly and return its actual field names ──
  if (message.type === "PROBE_SCHEMA") {
    const { jobId } = message;
    if (!jobId) {
      sendResponse({ ok: false, error: "Job ID required" });
      return false;
    }
    const filter = encodeURIComponent(
      `(ClientJobId eq ${jobId}) and (DeletedOn eq null)`
    );
    odataGet(`ToolAssemblies?$filter=${filter}&$top=1`)
      .then((records) => {
        const sample    = Array.isArray(records) ? (records[0] ?? {}) : (records ?? {});
        const keys      = Object.keys(sample).sort();
        const statsKeys = sample.Statistics ? Object.keys(sample.Statistics).sort() : null;
        chrome.storage.local.set(
          { fieldcap_schema_probe: { keys, statsKeys, jobId } },
          () => sendResponse({ ok: true, keys, statsKeys })
        );
      })
      .catch((err) => sendResponse({ ok: false, error: err.message }));
    return true;
  }

  return false;
});
