// FieldCap Data Exporter — Background Service Worker v3.1.0
// Fetches data directly from FieldCap's OData API using the active session.
// Produces typed CSVs including:
//   • job-details, crew, bha-equipment
//   • slide/rotate metres by calendar day (SurveySheetEntries × ActivityLogs — all days in one pull)

"use strict";

// ── Constants ─────────────────────────────────────────────────────────────────
const BASE = "https://fieldcap-cdn.phxtech.com/odata";

const KEY_META     = "fieldcap_meta";
const KEY_CSV_JOB  = "fieldcap_csv_job";
const KEY_CSV_CREW = "fieldcap_csv_crew";
const KEY_CSV_BHA  = "fieldcap_csv_bha";
const KEY_CSV_SLIDE_DAY = "fieldcap_csv_slide_by_day";
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

// Coerce a FieldCap date/time value to milliseconds since epoch.
// Parse FieldCap's compact YYYYMMDDHHMI integer (e.g. 202605062255 → 2026-05-06 22:55).
// Returns ms-from-epoch, or NaN if not a valid date in 2000-2099.
const parseYYYYMMDDHHMI = (n) => {
  const s = String(Math.round(n));
  if (s.length !== 12) return NaN;
  const y = parseInt(s.slice(0, 4), 10);
  const mo = parseInt(s.slice(4, 6), 10);
  const d = parseInt(s.slice(6, 8), 10);
  const h = parseInt(s.slice(8, 10), 10);
  const mi = parseInt(s.slice(10, 12), 10);
  if (y < 2000 || y > 2099 || mo < 1 || mo > 12 || d < 1 || d > 31 || h > 23 || mi > 59) return NaN;
  return new Date(y, mo - 1, d, h, mi).getTime();
};

// Accepts ISO strings, compact YYYYMMDDHHMI integers, Int64 ms, Unix seconds,
// Excel serial days, and legacy "/Date(ms)/" strings.
const toMs = (v) => {
  if (v === null || v === undefined || v === "") return NaN;
  if (typeof v === "number" && Number.isFinite(v)) {
    // YYYYMMDDHHMI compact format: 2000-01-01 00:00 = 200001010000 ≈ 2.0e11
    // Real ms timestamps for year 2000+ start at ≈ 9.46e11 — no overlap.
    if (v >= 2e11 && v < 2.1e11) {
      const t = parseYYYYMMDDHHMI(v);
      if (Number.isFinite(t)) return t;
    }
    if (v >= 9e11 && v < 1e14) return v;           // ms (year 2000–2286)
    if (v >= 1e9 && v < 9e11) return v * 1000;     // Unix seconds
    if (v > 20000 && v < 90000) {                  // Excel serial day
      return new Date(Date.UTC(1899, 11, 30) + v * 86400000).getTime();
    }
    return NaN;
  }
  const s = String(v);
  const legacy = s.match(/\/Date\((-?\d+)\)\//);
  if (legacy) return parseInt(legacy[1], 10);
  // Numeric string — check YYYYMMDDHHMI first, then other numeric types
  if (/^\d{12}$/.test(s)) {
    const t = parseYYYYMMDDHHMI(Number(s));
    if (Number.isFinite(t)) return t;
  }
  const n = Number(s);
  if (Number.isFinite(n)) {
    if (n >= 9e11 && n < 1e14) return n;
    if (n >= 1e9 && n < 9e11) return n * 1000;
    if (n > 20000 && n < 90000) {
      return new Date(Date.UTC(1899, 11, 30) + n * 86400000).getTime();
    }
  }
  const d = new Date(s);
  const t = d.getTime();
  return Number.isFinite(t) ? t : NaN;
};

// Local calendar date (browser timezone) — use for bucketing survey rows into “daily” totals.
const toLocalDateStr = (v) => {
  const ms = toMs(v);
  if (!Number.isFinite(ms)) return "";
  const d = new Date(ms);
  const y = d.getFullYear();
  if (y < 2000 || y > 2100) return "";
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};

// Calendar bucket for reporting — reject garbage epochs (e.g. Unix seconds mis-read as ms → 1976).
const surveyDayBucketKey = (surveyDateTime) => {
  const ms = toMs(surveyDateTime);
  if (!Number.isFinite(ms)) return "";
  const y = new Date(ms).getFullYear();
  if (y < 2000 || y > 2100) return "";
  const local = toLocalDateStr(surveyDateTime);
  if (local) return local;
  return new Date(ms).toISOString().slice(0, 10);
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

// Pick slide vs rotate distance fields when OData uses tenant-specific names.
const keyLooksTotalMetresOnly = (k) => {
  const nk = normalizeKey(k);
  if (/total|drilled|footage|distance/i.test(k) && !/sld|slide|rot/i.test(k)) return true;
  return /^(metresdrilled|metersdrilled|metres|meters)$/i.test(nk);
};

// Merge Statistics onto the root so pickNum sees SlideMetres nested under Statistics
// or inside one-level nested metric objects.
const augmentAssemblyForPicking = (asm) => {
  if (!asm || typeof asm !== "object") return asm;
  const out = { ...asm };
  const st = asm.Statistics;
  if (!st) return out;
  if (Array.isArray(st)) {
    for (const row of st) {
      if (row && typeof row === "object") Object.assign(out, row);
    }
  } else if (typeof st === "object") {
    Object.assign(out, st);
    for (const v of Object.values(st)) {
      if (v && typeof v === "object" && !Array.isArray(v)) Object.assign(out, v);
    }
  }
  return out;
};

const pickSlideOrRotateMetres = (obj, mode) => {
  const buckets = [obj, obj?.Statistics].filter(Boolean);
  for (const b of buckets) {
    for (const [k, raw] of Object.entries(b ?? {})) {
      if (keyLooksTotalMetresOnly(k)) continue;
      const lk = String(k ?? "").toLowerCase();
      if (!/(metre|meter|\bmtrs\b|footage|distance|\bft\b|feet)/i.test(lk)) continue;
      if (/(hour|hrs|toolhours|minute)/i.test(lk)) continue;
      const hasSlide = /(sld|slide|slid)/i.test(k);
      const hasRot = /(rot|rotate|rotat)/i.test(k) && !/motor/i.test(lk);
      if (mode === "slide") {
        if (!hasSlide || (hasRot && !hasSlide)) continue;
      } else {
        if (!hasRot || hasSlide) continue;
      }
      const n = parseHoursNum(raw);
      if (n !== null && n > 0) return raw;
    }
  }
  return "";
};

const normalizeKey = (v) => String(v ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const isFilled = (v) => !(v === null || v === undefined || v === "");
const mergeNonEmpty = (base, patch) => {
  const out = { ...(base ?? {}) };
  for (const [k, v] of Object.entries(patch ?? {})) {
    if (!isFilled(v)) continue;
    // ActivityLogs⨯Survey join can legitimately compute SlideMetres=0 while slide
    // metres exist on the daily Activities UI; 0 must not wipe scraped/API values.
    if (
      (k === "SlideMetres" || k === "RotateMetres") &&
      Number(v) === 0
    ) {
      const prev = parseHoursNum(out[k]);
      if (prev !== null && prev > 0) continue;
    }
    out[k] = v;
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

// Prefer explicit grid columns over injected __bha (multi-BHA daily reports).
const ACTIVITY_BHA_ALIASES = [
  "BHA", "BHA #", "BHA#", "Bha", "BHANo", "BHA Number", "BHA No",
  "Tool Assembly #", "ToolAssemblyNumber", "__bha",
];

const pickActivityRowBha = (row, valueFromKeyIncludes) => {
  const direct = getByAliases(row, ACTIVITY_BHA_ALIASES);
  if (direct !== "") return direct;
  return valueFromKeyIncludes(row, ["toolassemblynumber", "toolassembly"]);
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
      SlideMetres:      getByAliases(row, [
        "SLD METERS", "SLD Metres", "SLD Mtrs", "Slide Metres", "Slide Meters",
        "Mtrs Slid", "Mtrs Slide", "BHA Mtrs Slid", "Slide Mtrs", "SLD",
      ]),
      RotateMetres:     getByAliases(row, [
        "ROTATE METERS", "ROT Metres", "ROT METERS", "Rotate Metres", "Rotate Meters",
        "Mtrs Rot", "Mtrs Rotate", "BHA Mtrs Rot", "Rotate Mtrs", "ROT",
      ]),
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

const activityClassifyBlob = (row) => {
  const chunks = [];
  for (const [k, v] of Object.entries(row ?? {})) {
    const nk = normalizeKey(k);
    if (!/(activity|code|desc|type|name|operation|phase)/i.test(k) && !/(activity|code)/i.test(nk)) continue;
    if (v !== null && v !== undefined && v !== "") chunks.push(String(v));
  }
  chunks.push(
    getByAliases(row, ["Activity/Code", "Activity Code", "Activity", "Code", "Operation"]),
    row?.Description ?? "",
    row?.ActivityName ?? "",
    row?.ActivityCode ?? ""
  );
  return chunks.filter(Boolean).join(" ").toLowerCase();
};

const isSlideActivityText = (activityBlob, codeOnly) => {
  const a = `${activityBlob} ${codeOnly}`.toLowerCase();
  if (/motor|rotor|circulation|circ\b|rig\b/i.test(a) && !/slide|sld|slid/i.test(a)) return false;
  return (
    /\bsld\b|slide|sliding|\bslid\b|code\s*2\s*a|\b2\s*a\b|\b2a\b/i.test(a) ||
    /^2\s*a$/i.test(String(codeOnly).trim())
  );
};

const isRotateActivityText = (activityBlob, codeOnly) => {
  const a = `${activityBlob} ${codeOnly}`.toLowerCase();
  if (/motor|rotor|electro|prop\b/i.test(a)) return false;
  const co = String(codeOnly).trim();
  return (
    /\brotate\b|\brotating\b|\brotation\b|\brot\b|rotating|code\s*2(?!\d)/i.test(a) ||
    /^2$/i.test(co)
  );
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

  const footerByBha = {};
  for (const row of rows ?? []) {
    if (!row?.__activityFooter) continue;
    const bha = pickActivityRowBha(row, valueFromKeyIncludes);
    if (!bha) continue;
    const bhaNum = String(toNum(bha));
    if (!bhaNum) continue;
    const sRaw = row.__footerSlideMetres;
    const rRaw = row.__footerRotateMetres;
    const s = sRaw !== "" && sRaw != null ? Number(toNum(sRaw)) : NaN;
    const r = rRaw !== "" && rRaw != null ? Number(toNum(rRaw)) : NaN;
    footerByBha[bhaNum] = {
      slide: Number.isFinite(s) ? s : 0,
      rot: Number.isFinite(r) ? r : 0,
    };
  }

  const map = {};
  for (const row of rows ?? []) {
    if (row?.__activityFooter) continue;
    const bha = pickActivityRowBha(row, valueFromKeyIncludes);
    if (!bha) continue;
    const bhaNum = String(toNum(bha));
    if (!bhaNum) continue;
    if (footerByBha[bhaNum]) continue;

    const codeOnly = String(
      getByAliases(row, ["Activity/Code", "Activity Code", "Code"]) || row?.Code || ""
    ).trim();

    const activityBlob = activityClassifyBlob(row);

    const courseRaw = String(
      getByAliases(row, ["Course", "Metres", "Meters", "Metres Drilled", "Meters Drilled", "Length"])
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
    const slideHit = isSlideActivityText(activityBlob, codeOnly);
    const rotHit = isRotateActivityText(activityBlob, codeOnly);
    if (slideHit && !rotHit) prev.slide += course;
    else if (rotHit && !slideHit) prev.rot += course;
    else if (slideHit && rotHit) {
      if (/sld|slide|slid|2\s*a|2a/i.test(activityBlob + codeOnly)) prev.slide += course;
      else prev.rot += course;
    }
    map[bhaNum] = prev;
  }

  for (const [bhaNum, v] of Object.entries(footerByBha)) {
    map[bhaNum] = v;
  }

  const out = {};
  for (const [bhaNum, v] of Object.entries(map)) {
    const s = v.slide ?? 0;
    const r = v.rot ?? 0;
    if (s <= 0 && r <= 0) continue;
    const patch = {};
    if (s > 0) patch.SlideMetres = Number(s.toFixed(2));
    if (r > 0) patch.RotateMetres = Number(r.toFixed(2));
    if (Object.keys(patch).length) out[bhaNum] = patch;
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
      // Keep slide/rotate metre keys; only exclude duration-like fields.
      if (/(hour|hours|hrs|toolhours|minute|minut|circ|below)/i.test(lk)) continue;
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

// Follow @odata.nextLink — FieldCap often paginates; missing pages meant missing custom fields / assemblies.
const odataGetAll = async (relativePath) => {
  const out = [];
  let url = `${BASE}/${relativePath}`;
  let guard = 0;
  while (url && guard++ < 100) {
    const res = await fetch(url, {
      credentials: "include",
      headers: { Accept: "application/json;odata.metadata=minimal" },
    });
    if (!res.ok) throw new Error(`OData ${res.status}`);
    const json = await res.json();
    if (Array.isArray(json.value)) out.push(...json.value);
    url = json["@odata.nextLink"] || null;
  }
  return out;
};

const ACTIVITY_ENTITY_DAY_KEYS = [
  "ReportDate",
  "ActivityDate",
  "JobActivityDate",
  "DailyReportDate",
  "StartDateTime",
  "EndDateTime",
  "Date",
  "SurveyDateTime",
  "ActivityStartDateTime",
  "ActivityEndDateTime",
];

const pickActivityEntityCalendarDay = (r) => {
  for (const k of ACTIVITY_ENTITY_DAY_KEYS) {
    const v = r[k];
    if (v == null || v === "") continue;
    const day = surveyDayBucketKey(v);
    if (day) return day;
  }
  return "";
};

const mapODataActivitiesToSyntheticRows = (rows, idToBha) => {
  const fake = [];
  for (const r of rows ?? []) {
    const tid =
      r.ToolAssemblyId ??
      r.ToolAssembly?.ToolAssemblyId ??
      r.ToolAssembly?.ToolAssembly?.ToolAssemblyId;
    const taNum =
      r.ToolAssembly?.ToolAssemblyNumber ??
      r.ToolAssemblyNumber ??
      r.BHA ??
      r.BhaNumber;
    const bha =
      (tid ? idToBha[String(tid)] : null) ??
      (taNum !== undefined && taNum !== "" ? String(toNum(taNum)) : "");
    if (!bha) continue;

    const courseRaw =
      r.Course ??
      r.Metres ??
      r.Meters ??
      r.MetresDrilled ??
      r.MetersDrilled ??
      r.Distance ??
      r.Footage ??
      r.Length ??
      "";
    const parts = [
      r.ActivityCode,
      r.ActivityName,
      r.ActivityTypeName,
      r.Activity,
      r.Description,
      r.ActivityDescription,
      r.Code,
      r.Operation,
    ].filter((x) => x !== null && x !== undefined && String(x).trim() !== "");

    const calendarDay = pickActivityEntityCalendarDay(r);
    const row = {
      BHA: bha,
      "Activity Code": parts.join(" "),
      Activity: parts.join(" "),
      Course: courseRaw,
      Code: r.Code ?? r.ActivityCode ?? "",
    };
    if (calendarDay) {
      row.Date = calendarDay;
      row.__calendarDay = calendarDay;
    }
    fake.push(row);
  }
  return fake;
};

const fetchJobActivitiesOData = async (jobId, assemblies) => {
  const idToBha = {};
  for (const a of assemblies ?? []) {
    if (a?.ToolAssemblyId) idToBha[String(a.ToolAssemblyId)] = String(toNum(a.ToolAssemblyNumber));
  }
  const filters = [
    `(ClientJobId eq ${jobId}) and (null eq DeletedBy)`,
    `(ClientJobId eq ${jobId})`,
  ];
  const entities = ["JobActivities", "Activities", "JobActivityRecords"];
  for (const entity of entities) {
    for (const f of filters) {
      for (const expand of ["", `&$expand=${encodeURIComponent("ToolAssembly")}`]) {
        try {
          const filter = encodeURIComponent(f);
          const rows = await odataGetAll(`${entity}?$filter=${filter}${expand}`);
          if (rows.length > 0) return mapODataActivitiesToSyntheticRows(rows, idToBha);
        } catch (_) {
          /* try next shape */
        }
      }
    }
  }
  return [];
};

// ── Slide / Rotate metres from ActivityLogs ⨯ SurveySheetEntries ───────────────
// FieldCap's daily report bottom-of-page slide/rotate totals are not stored as a
// single field anywhere. They are computed by joining each ActivityLog row's
// time window (per BHA) against the SurveySheetEntry CourseLength advances that
// fall within that window, then summing per slide-vs-rotate classification.
const activityIsSlide = (text) => {
  const t = String(text ?? "").toLowerCase();
  if (!t) return false;
  if (/(motor|rotor|circulation|circ\b)/i.test(t) && !/slid|sld|slide/i.test(t)) return false;
  return /\bsld\b|\bslid\b|slide|sliding/i.test(t);
};

const activityIsRotate = (text) => {
  const t = String(text ?? "").toLowerCase();
  if (!t) return false;
  if (/slid|sld|slide/i.test(t)) return false;
  return /\brot\b|rotate|rotating|rotation|\bdrg\b|\bdrl\b|drilling/i.test(t);
};

const fetchAllActivityLogsForJob = async (jobId) => {
  const filters = [
    `(ClientJobId eq ${jobId}) and (DeletedOn eq null)`,
    `(ClientJobId eq ${jobId}) and (null eq DeletedBy)`,
    `(ClientJobId eq ${jobId})`,
  ];
  const select = encodeURIComponent(
    "ActivityLogId,ActivityType,Comments,StartDateTime,EndDateTime,ToolAssemblyId," +
      "Custom1,Custom2,Custom3,Custom4,Custom5,Custom6,Custom7,Custom8,Custom9,Custom10"
  );
  for (const f of filters) {
    try {
      const filter = encodeURIComponent(f);
      const rows = await odataGetAll(`ActivityLogs?$filter=${filter}&$select=${select}`);
      if (Array.isArray(rows)) return rows;
    } catch (_) { /* try next filter */ }
  }
  return [];
};

const fetchSurveyEntriesForJob = async (jobId) => {
  const filters = [
    `(SurveySheet/ClientJobId eq ${jobId}) and (null eq DeletedBy) and (null eq OriginalSurveyId)`,
    `(SurveySheet/ClientJobId eq ${jobId}) and (null eq DeletedBy)`,
    `(SurveySheet/ClientJobId eq ${jobId})`,
  ];
  // SurveyDateTime is a compact YYYYMMDDHHMI integer (e.g. 202605062255 = 2026-05-06 22:55).
  // SlidingDistance = slide metres per interval; CourseLength = total interval metres.
  // ToolAssemblyId is a flat scalar field — NO $expand needed (idToBha already maps it).
  // $expand=ToolAssembly causes row multiplication in FieldCap's OData — do NOT use it here.
  const variants = [
    { sel: "SurveyDateTime,CourseLength,SlidingDistance,ToolAssemblyId", exp: "" },
    { sel: "SurveyDateTime,CourseLength,SlidingDistance", exp: "" },
  ];
  for (const f of filters) {
    const encFilter = encodeURIComponent(f);
    for (const v of variants) {
      try {
        const path = `SurveySheetEntries?$filter=${encFilter}&$select=${encodeURIComponent(v.sel)}${v.exp}`;
        const rows = await odataGetAll(path);
        if (Array.isArray(rows) && rows.length > 0) return rows;
      } catch (_) { /* try next */ }
    }
  }
  return [];
};

// SurveySheetEntries × ActivityLogs: same join as the daily UI, but OData returns
// every survey stamp for the job — no need to click each day on the calendar.
const buildLogsByAsmIndex = (logs) => {
  const logsByAsm = {};
  for (const log of logs) {
    const tid = log?.ToolAssemblyId ? String(log.ToolAssemblyId) : "";
    if (!tid) continue;
    const blob = [
      log.ActivityType, log.Comments,
      log.Custom1, log.Custom2, log.Custom3, log.Custom4, log.Custom5,
      log.Custom6, log.Custom7, log.Custom8, log.Custom9, log.Custom10,
    ].filter((x) => x !== null && x !== undefined && String(x).trim() !== "").join(" ");
    const start = toMs(log.StartDateTime);
    const end = toMs(log.EndDateTime);
    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;
    const slide = activityIsSlide(blob);
    const rotate = !slide && activityIsRotate(blob);
    if (!slide && !rotate) continue;
    (logsByAsm[tid] ??= []).push({ start, end, kind: slide ? "slide" : "rot" });
  }
  return logsByAsm;
};

const surveyCellShape = () => ({ slide: 0, rot: 0, total: 0, classified: 0 });

const bumpSurveyCell = (cell, course, tid, tMs, logsByAsm) => {
  cell.total += course;
  if (!Number.isFinite(tMs)) return;
  const candidates = logsByAsm[tid] ?? [];
  let kind = null;
  for (const c of candidates) {
    if (c.start <= tMs && tMs <= c.end) { kind = c.kind; break; }
  }
  if (kind === "slide") {
    cell.slide += course;
    cell.classified += course;
  } else if (kind === "rot") {
    cell.rot += course;
    cell.classified += course;
  }
};

const surveyRowToolAssemblyId = (s) => {
  const raw = s?.ToolAssemblyId ?? s?.ToolAssembly?.ToolAssemblyId;
  if (raw !== null && raw !== undefined && raw !== "") return String(raw);
  return "";
};

const surveyRowCourseLength = (s) => {
  const raw = s?.CourseLength ?? s?.Length ?? s?.Metres ?? s?.Meters ?? s?.Distance ?? "";
  const n = Number(raw);
  return Number.isFinite(n) && n > 0 ? n : NaN;
};

// SlidingDistance = slide metres for this survey interval.
// FieldCap's own $apply aggregation queries use "SlidingDistance with sum as TotalSlide".
// Rotate metres = CourseLength - SlidingDistance.
const surveyRowSlideMetres = (s) => {
  const raw = s?.SlidingDistance ?? "";
  const n = Number(raw);
  return Number.isFinite(n) && n >= 0 ? n : NaN;
};

// FieldCap sometimes omits SurveyDateTime on SurveySheetEntries; try sheet-level and audit dates last.
const surveyRowDateTimeValue = (s) => {
  if (!s || typeof s !== "object") return "";
  const sheet = s.SurveySheet;
  const candidates = [
    s.SurveyDateTime,
    s.EntryDateTime,
    s.StationDateTime,
    s.SurveyTime,
    s.Date,
    s.SurveyDate,
    s.ReportDate,
    s.DailyDate,
    s.RigDate,
    sheet?.ReportDate,
    sheet?.SurveyDate,
    sheet?.Date,
    sheet?.DailyReportDate,
    s.CreatedOn,
    s.ModifiedOn,
  ];
  for (const c of candidates) {
    if (c !== null && c !== undefined && c !== "") return c;
  }
  return "";
};

// Accumulate slide + rotate totals directly from ActivityLogs.
// The Slide Sheet "CL where Mode=S" = ActivityType='Sliding', CL = Custom2 - Custom1.
// The Slide Sheet "CL where Mode=R" = ActivityType='Drilling', CL = Custom2 - Custom1.
// Custom1 = Start MD (string), Custom2 = End MD (string).
const SLIDE_ACTIVITY_TYPE = "Sliding";
const ROTATE_ACTIVITY_TYPE = "Drilling";

const accumulateActivityLogMetres = (logs, idToBha) => {
  const lifetime = {};
  const byDay = {};

  for (const log of logs) {
    const tid = log.ToolAssemblyId ? String(log.ToolAssemblyId) : "";
    const bha = tid ? idToBha[tid] : "";
    if (!bha) continue;

    const actType = String(log.ActivityType ?? "").trim();
    const isSlide = actType === SLIDE_ACTIVITY_TYPE;
    const isRot = actType === ROTATE_ACTIVITY_TYPE;
    if (!isSlide && !isRot) continue;

    const startMd = parseFloat(log.Custom1);
    const endMd = parseFloat(log.Custom2);
    if (!Number.isFinite(startMd) || !Number.isFinite(endMd)) continue;
    const cl = endMd - startMd;
    if (cl <= 0) continue;

    const slide = isSlide ? cl : 0;
    const rot = isRot ? cl : 0;

    const dayStr = surveyDayBucketKey(log.StartDateTime);

    const lf = (lifetime[bha] ??= surveyCellShape());
    lf.slide += slide;
    lf.rot += rot;
    lf.total += cl;

    if (dayStr) {
      const dc = ((byDay[bha] ??= {})[dayStr] ??= surveyCellShape());
      dc.slide += slide;
      dc.rot += rot;
      dc.total += cl;
    }
  }
  return { lifetime, byDay };
};

const finalizeSurveyPatchMap = (totals) => {
  const out = {};
  for (const [bha, v] of Object.entries(totals)) {
    if (v.slide <= 0 && v.rot <= 0 && v.total <= 0) continue;
    const patch = {};
    if (v.slide > 0) patch.SlideMetres = Number(v.slide.toFixed(2));
    if (v.rot > 0) patch.RotateMetres = Number(v.rot.toFixed(2));
    if (v.total > 0) patch.MetresDrilled = Number(v.total.toFixed(2));
    if (Object.keys(patch).length) out[bha] = patch;
  }
  return out;
};

const SLIDE_DAY_COLUMNS = [
  "Job ID", "Date", "BHA #", "Slide Metres", "Rotate Metres", "Survey Course Sum",
];

const buildSlideRotateMetresByDayCsv = (jobId, byDay, lifetimeTotals = {}) => {
  const rows = [];
  for (const [bha, days] of Object.entries(byDay)) {
    for (const [date, v] of Object.entries(days)) {
      if (v.slide <= 0 && v.rot <= 0 && v.total <= 0) continue;
      rows.push({
        "Job ID": jobId,
        "Date": date,
        "BHA #": bha,
        "Slide Metres": v.slide > 0 ? Number(v.slide.toFixed(2)) : "",
        "Rotate Metres": v.rot > 0 ? Number(v.rot.toFixed(2)) : "",
        "Survey Course Sum": v.total > 0 ? Number(v.total.toFixed(2)) : "",
      });
    }
  }
  if (rows.length === 0 && lifetimeTotals && typeof lifetimeTotals === "object") {
    for (const [bha, v] of Object.entries(lifetimeTotals)) {
      if (v.slide <= 0 && v.rot <= 0 && v.total <= 0) continue;
      rows.push({
        "Job ID": jobId,
        "Date": "(lifetime — no calendar date on survey rows in OData)",
        "BHA #": bha,
        "Slide Metres": v.slide > 0 ? Number(v.slide.toFixed(2)) : "",
        "Rotate Metres": v.rot > 0 ? Number(v.rot.toFixed(2)) : "",
        "Survey Course Sum": v.total > 0 ? Number(v.total.toFixed(2)) : "",
      });
    }
  }
  rows.sort((a, b) =>
    (a.Date === b.Date ? 0 : a.Date.localeCompare(b.Date)) ||
    String(a["BHA #"]).localeCompare(String(b["BHA #"]), undefined, { numeric: true })
  );
  return buildCsvString(SLIDE_DAY_COLUMNS, rows);
};

const pickActivityRowCalendarDay = (row) => {
  const pre = String(row?.__calendarDay ?? "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(pre)) return pre;
  const fromAliases = getByAliases(row, ["Date", "Report Date", "Activity Date", "Daily Date", "Day"]);
  if (fromAliases !== "") {
    const k = surveyDayBucketKey(fromAliases);
    if (k) return k;
  }
  return "";
};

// Buckets slide/rotate metres by calendar day from scraped + OData activity rows (matches daily Activities when dated).
const buildSlideRotateMetresByDayCsvFromActivities = (rows, jobId) => {
  const valueFromKeyIncludes = (row, needles) => {
    for (const [k, v] of Object.entries(row ?? {})) {
      const nk = normalizeKey(k);
      if (needles.some((n) => nk.includes(n))) {
        if (v !== null && v !== undefined && v !== "") return v;
      }
    }
    return "";
  };

  const footerByBha = {};
  for (const row of rows ?? []) {
    if (!row?.__activityFooter) continue;
    const bha = pickActivityRowBha(row, valueFromKeyIncludes);
    if (!bha) continue;
    const bhaNum = String(toNum(bha));
    if (!bhaNum) continue;
    footerByBha[bhaNum] = true;
  }

  const nested = {};
  for (const row of rows ?? []) {
    if (row?.__activityFooter) continue;
    const bha = pickActivityRowBha(row, valueFromKeyIncludes);
    if (!bha) continue;
    const bhaNum = String(toNum(bha));
    if (!bhaNum) continue;
    if (footerByBha[bhaNum]) continue;

    const day = pickActivityRowCalendarDay(row);
    if (!day) continue;

    const codeOnly = String(
      getByAliases(row, ["Activity/Code", "Activity Code", "Code"]) || row?.Code || ""
    ).trim();
    const activityBlob = activityClassifyBlob(row);
    const courseRaw = String(
      getByAliases(row, ["Course", "Metres", "Meters", "Metres Drilled", "Meters Drilled", "Length"])
      || valueFromKeyIncludes(row, ["course", "metre", "meter", "distance", "drill"])
      || ""
    );
    let course = Number((courseRaw.match(/-?\d+(\.\d+)?/) ?? [])[0]);
    if (!Number.isFinite(course) || course <= 0) {
      const depthRange = String(
        getByAliases(row, ["Start - End Depth", "Start-End Depth", "Start/End Depth", "Depth Range"])
        || valueFromKeyIncludes(row, ["startenddepth", "depthrange", "startend"])
        || ""
      );
      course = Number(rangeDelta(depthRange));
    }
    if (!Number.isFinite(course) || course <= 0) continue;

    const slideHit = isSlideActivityText(activityBlob, codeOnly);
    const rotHit = isRotateActivityText(activityBlob, codeOnly);
    let slide = 0;
    let rot = 0;
    if (slideHit && !rotHit) slide = course;
    else if (rotHit && !slideHit) rot = course;
    else if (slideHit && rotHit) {
      if (/sld|slide|slid|2\s*a|2a/i.test(activityBlob + codeOnly)) slide = course;
      else rot = course;
    }
    if (slide <= 0 && rot <= 0) continue;

    const slot = ((nested[bhaNum] ??= {})[day] ??= { slide: 0, rot: 0 });
    slot.slide += slide;
    slot.rot += rot;
  }

  const outRows = [];
  for (const [bha, days] of Object.entries(nested)) {
    for (const [date, v] of Object.entries(days)) {
      const sum = v.slide + v.rot;
      if (sum <= 0) continue;
      outRows.push({
        "Job ID": jobId,
        "Date": date,
        "BHA #": bha,
        "Slide Metres": v.slide > 0 ? Number(v.slide.toFixed(2)) : "",
        "Rotate Metres": v.rot > 0 ? Number(v.rot.toFixed(2)) : "",
        "Survey Course Sum": Number(sum.toFixed(2)),
      });
    }
  }
  outRows.sort((a, b) =>
    (a.Date === b.Date ? 0 : a.Date.localeCompare(b.Date)) ||
    String(a["BHA #"]).localeCompare(String(b["BHA #"]), undefined, { numeric: true })
  );
  return buildCsvString(SLIDE_DAY_COLUMNS, outRows);
};

const fetchSurveyDerivedSlideRotate = async (jobId, assemblies) => {
  const idToBha = {};
  for (const a of assemblies ?? []) {
    if (a?.ToolAssemblyId) idToBha[String(a.ToolAssemblyId)] = String(toNum(a.ToolAssemblyNumber));
  }
  if (Object.keys(idToBha).length === 0) {
    return { activityLogMap: {}, slideByDayCsv: "", slideByDayRowCount: 0 };
  }

  let logs = [];
  try {
    logs = await fetchAllActivityLogsForJob(jobId);
  } catch (_) {
    return { activityLogMap: {}, slideByDayCsv: "", slideByDayRowCount: 0 };
  }
  if (!logs.length) {
    return { activityLogMap: {}, slideByDayCsv: "", slideByDayRowCount: 0 };
  }

  const { lifetime, byDay } = accumulateActivityLogMetres(logs, idToBha);
  const activityLogMap = finalizeSurveyPatchMap(lifetime);
  const slideByDayCsv = buildSlideRotateMetresByDayCsv(jobId, byDay, lifetime);
  const lines = slideByDayCsv.split("\r\n").filter((ln) => ln.length > 0);
  const slideByDayRowCount = Math.max(0, lines.length - 1);
  return { activityLogMap, slideByDayCsv, slideByDayRowCount };
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
  const filters = [
    encodeURIComponent(`(ClientJobId eq ${jobId}) and (null eq DeletedBy)`),
    encodeURIComponent(`(ClientJobId eq ${jobId}) and (DeletedOn eq null)`),
  ];
  const expand = encodeURIComponent("Statistics");
  for (const enc of filters) {
    try {
      const rows = await odataGetAll(`ToolAssemblies?$filter=${enc}&$expand=${expand}`);
      if (rows.length > 0) return rows;
    } catch (_) {}
    try {
      const rows = await odataGetAll(`ToolAssemblies?$filter=${enc}`);
      if (rows.length > 0) return rows;
    } catch (_) {}
  }
  try {
    const enc = encodeURIComponent(`(ClientJobId eq ${jobId}) and (DeletedOn eq null)`);
    return await odataGetAll(`ToolAssemblies?$filter=${enc}&$expand=${expand}`);
  } catch (_) {
    return [];
  }
};

// Single-assembly fetch — FieldCap often omits slide/rotate metres on the list query only.
const enrichToolAssembliesFromDetail = async (assemblies) => {
  const list = Array.isArray(assemblies) ? assemblies : [];
  return Promise.all(
    list.map(async (asm) => {
      const id = asm.ToolAssemblyId;
      if (!id) return asm;
      const expand = encodeURIComponent("Statistics");
      const filters = [
        encodeURIComponent(`(ToolAssemblyId eq ${id}) and (null eq DeletedBy)`),
        encodeURIComponent(`(ToolAssemblyId eq ${id}) and (DeletedOn eq null)`),
      ];
      for (const filter of filters) {
        try {
          const rows = await odataGet(`ToolAssemblies?$filter=${filter}&$expand=${expand}&$top=1`);
          const detail = Array.isArray(rows) ? rows[0] : rows;
          if (!detail || typeof detail !== "object") continue;
          return mergeNonEmpty(asm, detail);
        } catch (_) {}
      }
      return asm;
    })
  );
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
  return odataGetAll(`ToolAssemblyItems?$expand=${expand}&$filter=${filter}`);
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

// ── BHA custom fields (tenant-defined labels such as "SLD METERS") ────────────
const labelFromToolAssemblyCustomRow = (row) => {
  const cf = row.CustomField ?? row.CustomFieldDefinition ?? row.Field ?? {};
  return String(
    row.KeyName ??
    row.FieldName ??
    row.Name ??
    row.CustomFieldName ??
    row.Label ??
    cf.Name ??
    cf.FieldName ??
    cf.DisplayName ??
    cf.Title ??
    row.ToolAssemblyCustomValueName ??
    ""
  ).trim();
};

const valueFromToolAssemblyCustomRow = (row) =>
  row.Value ??
  row.TextValue ??
  row.StringValue ??
  row.NumberValue ??
  row.DecimalValue ??
  row.DoubleValue ??
  row.FloatValue ??
  row.CustomValue ??
  "";

const assemblyPatchFromCustomRows = (rows) => {
  const patch = {};
  for (const row of rows ?? []) {
    const label = labelFromToolAssemblyCustomRow(row);
    const val = valueFromToolAssemblyCustomRow(row);
    if (!label) continue;
    patch[label] = val;
    const slug = label.replace(/\s+/g, "");
    if (slug && slug !== label) patch[slug] = val;
  }
  return patch;
};

const customValueFilterVariants = (toolAssemblyId) => [
  `(null eq DeletedBy) and (ToolAssemblyId eq ${toolAssemblyId})`,
  `(DeletedOn eq null) and (ToolAssemblyId eq ${toolAssemblyId})`,
  `(ToolAssemblyId eq ${toolAssemblyId})`,
];

const fetchToolAssemblyCustomValuesForJob = async (assemblies) => {
  const map = {};
  const list = Array.isArray(assemblies) ? assemblies : [];
  await Promise.all(
    list.map(async (asm) => {
      const id = asm.ToolAssemblyId;
      if (!id) return;
      let rows = [];
      for (const f of customValueFilterVariants(id)) {
        try {
          const filter = encodeURIComponent(f);
          const got = await odataGetAll(`ToolAssemblyCustomValues?$filter=${filter}`);
          if (Array.isArray(got) && got.length > 0) {
            rows = got;
            break;
          }
        } catch (_) {
          /* try next filter */
        }
      }
      map[String(id)] = assemblyPatchFromCustomRows(rows);
    })
  );
  return map;
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
  const assemblyPick = augmentAssemblyForPicking(assembly);
  const itemSerial = item?.JobTool?.ItemSerial;
  const itemRecord = item?.JobTool?.Item ?? itemSerial?.Item;

  // Tenant-specific fallback observed in this FieldCap instance:
  // ToolHours1=Rotate (minutes), ToolHours2=Slide (minutes),
  // ToolHours3=Circ (minutes), ToolHours4=Below Rot (minutes)
  const tH1 = pickRaw(assemblyPick, "ToolHours1");
  const tH2 = pickRaw(assemblyPick, "ToolHours2");
  const tH3 = pickRaw(assemblyPick, "ToolHours3");
  const tH4 = pickRaw(assemblyPick, "ToolHours4");

  let hrsSlide = pickHours(assemblyPick,
    "SlideHours", "SlideHrs", "SlidingHours", "HoursSliding", "HrsSlide"
  );
  let hrsRot = pickHours(assemblyPick,
    "RotateHours", "RotateHrs", "RotatingHours", "HoursRotating", "HrsRotate"
  );
  let hrsCirc = pickHours(assemblyPick,
    "CirculateHours", "CircHours", "CirculatingHours", "HoursCirculating", "HrsCirc", "CircHrs"
  );
  if (!hrsSlide) hrsSlide = minutesToHours(tH2);
  if (!hrsRot)   hrsRot   = minutesToHours(tH1);
  if (!hrsCirc)  hrsCirc  = minutesToHours(tH3);

  let hrsTotal = pickHours(assemblyPick,
    "TotalHours", "TotalHrs", "TotalDrillingHours", "HoursTotal", "DrillHours"
  );
  let hrsBelowRot = pickHours(assemblyPick,
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

  const totalMetresVal = parseHoursNum(pickMetres(assemblyPick));
  let mtrsSlide = pickNum(assemblyPick,
    "SlideMetres", "SlideMeters", "SlidingMetres", "SlidingMeters", "SlideAmount",
    "Slide Meter", "Slide Meters", "SldMetres", "SldMeters", "SLDMetres", "SLDMeters",
    "SLD METERS", "SLD Metres", "SLD Mtrs", "SLD Metres Drilled", "SLD Meters Drilled",
    "SlidingDistance", "SlideDistance", "SlidingFootage", "SlideFootage",
    "ToolMetres2", "ToolMeters2",
    "MetresSlid", "MetersSlid", "MetresSlide", "MetersSlide", "BhaSlideMetres", "BHASlideMetres"
  );
  let mtrsRot = pickNum(assemblyPick,
    "RotateMetres", "RotateMeters", "RotatingMetres", "RotatingMeters", "RotateAmount",
    "Rotate Meter", "Rotate Meters", "RotMetres", "RotMeters",
    "ROTATE METERS", "ROT Metres", "ROT METERS", "Rotate Mtrs",
    "RotatingDistance", "RotateDistance", "RotatingFootage", "RotateFootage",
    "ToolMetres1", "ToolMeters1",
    "MetresRot", "MetersRot", "MetresRotate", "MetersRotate", "BhaRotateMetres", "BHARotateMetres"
  );
  if (mtrsSlide === "") mtrsSlide = pickSlideOrRotateMetres(assemblyPick, "slide");
  if (mtrsRot === "") mtrsRot = pickSlideOrRotateMetres(assemblyPick, "rot");
  const sMeters = parseHoursNum(mtrsSlide);
  const rMeters = parseHoursNum(mtrsRot);
  if ((rMeters === null || rMeters === 0) && totalMetresVal !== null && sMeters !== null && totalMetresVal >= sMeters) {
    mtrsRot = Number((totalMetresVal - sMeters).toFixed(2));
  }
  if ((sMeters === null || sMeters === 0) && totalMetresVal !== null && rMeters !== null && totalMetresVal >= rMeters) {
    mtrsSlide = Number((totalMetresVal - rMeters).toFixed(2));
  }

  return {
    "Job ID":          assembly.ClientJobId ?? "",
    "BHA #":           toNum(assembly.ToolAssemblyNumber),
    "Section":         assembly.Section ?? assembly.SectionName ?? assembly.BHASection ?? "",
    "Status":          assembly.Status ?? assembly.AssemblyStatus ?? assembly.AssemblyStatusName ?? "",
    "Motor":           assembly.Motor ?? assembly.MotorSerial ?? assembly.MotorDescription ?? "",
    "Guidance":        assembly.GuidanceType ?? assembly.Guidance ?? assembly.GuidanceName ?? "",
    "Metres Drilled":  pickMetres(assemblyPick),
    "BHA Mtrs Slid":   mtrsSlide,
    "BHA Mtrs Rot":    mtrsRot,
    "BHA Total Hrs":   hrsTotal,
    "BHA Hrs Slid":    hrsSlide,
    "BHA Hrs Rot":     hrsRot,
    "BHA Hrs Circ":    hrsCirc,
    "BHA Below Rot":   hrsBelowRot,
    "Activated On":    toDateStr(assemblyPick.ActivatedOn),
    "Completed On":    toDateStr(assemblyPick.CompletedOn),
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
const mapAssemblyPatchesFromItems = (assemblyItems) => {
  const byAssemblyId = {};
  for (const item of assemblyItems ?? []) {
    const ta = item?.ToolAssembly;
    if (!ta || typeof ta !== "object") continue;
    const c = toCanonicalAssemblyPatch(ta);
    if (!c?.id) continue;
    byAssemblyId[String(c.id)] = mergeNonEmpty(byAssemblyId[String(c.id)] ?? {}, c.patch);
  }
  return byAssemblyId;
};

const buildBhaCsv = (
  assemblies,
  assemblyItems,
  jobTools,
  interceptedMap = {},
  bhaGridMap = {},
  customAsmMap = {},
  itemAsmMap = {}
) => {
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
    const custom = customAsmMap[String(asmId)] ?? {};
    const fromItems = itemAsmMap[String(asmId)] ?? {};
    // IMPORTANT: asm first; intercepted/UI grid/custom fields override with non-empty values.
    const merged = mergeNonEmpty(
      mergeNonEmpty(mergeNonEmpty(mergeNonEmpty(asm, intercepted), grid), custom),
      fromItems
    );
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
const fetchAll = async (jobId, flags, liveBhaRows = [], liveActivityRows = [], onProgress = null) => {
  const prog = (pct, label) => { try { onProgress?.(pct, label); } catch (_) {} };
  const results = {};

  prog(10, "Job details…");
  if (flags.jobDetails) {
    const [raw, customValues] = await Promise.all([
      fetchJobDetails(jobId),
      fetchJobCustomValues(jobId),
    ]);
    results.jobDetailsCsv = buildJobDetailsCsv(raw, customValues);
  }

  prog(28, "Crew / Personnel…");
  if (flags.crew) {
    const raw = await fetchCrew(jobId);
    results.crew = raw.map((s) => normalizeCrewRow(s, jobId));
    results.crewCsv = buildCsvString(CREW_COLUMNS, results.crew);
  }

  prog(45, "BHA assemblies…");
  if (flags.bha) {
    const [assembliesList, items, tools, stored] = await Promise.all([
      fetchToolAssemblies(jobId),
      fetchToolAssemblyItems(jobId),
      fetchJobTools(jobId),
      new Promise((res) => chrome.storage.local.get([KEY_INTERCEPT, KEY_BHA_GRID], res)),
    ]);
    prog(60, "BHA components…");
    const assemblies = await enrichToolAssembliesFromDetail(assembliesList);
    const customAsmMap = await fetchToolAssemblyCustomValuesForJob(assemblies);
    prog(72, "Activity data…");
    let odataActivityRows = [];
    try {
      odataActivityRows = await fetchJobActivitiesOData(jobId, assemblies);
    } catch (_) {
      odataActivityRows = [];
    }
    const mergedActivityRows = [...liveActivityRows, ...odataActivityRows];
    prog(83, "Slide / rotate metres…");
    let activityLogMap = {};
    let slideByDayCsv = "";
    let slideByDayRowCount = 0;
    try {
      const derived = await fetchSurveyDerivedSlideRotate(jobId, assemblies);
      activityLogMap = derived.activityLogMap ?? {};
      slideByDayCsv = derived.slideByDayCsv ?? "";
      slideByDayRowCount = derived.slideByDayRowCount ?? 0;
    } catch (_) {
      activityLogMap = {};
    }
    const activityDayCsv = buildSlideRotateMetresByDayCsvFromActivities(mergedActivityRows, jobId);
    const activityDayLines = activityDayCsv.split("\r\n").filter((ln) => ln.length > 0);
    if (activityDayLines.length > 1) {
      slideByDayCsv = activityDayCsv;
      slideByDayRowCount = activityDayLines.length - 1;
    }
    const interceptedMap = stored[KEY_INTERCEPT] ?? {};
    const bhaGridRowsByJob = stored[KEY_BHA_GRID] ?? {};
    const cachedBhaGridMap = bhaGridRowsByJob[String(jobId)] ?? {};
    const liveBhaGridMap = toBhaGridMap(liveBhaRows);
    const liveActivityMap = toActivityMetresMap(mergedActivityRows);
    const itemAsmMap = mapAssemblyPatchesFromItems(items);
    const bhaGridMap = {};
    for (const [k, v] of Object.entries(cachedBhaGridMap)) bhaGridMap[k] = mergeNonEmpty(bhaGridMap[k], v);
    for (const [k, v] of Object.entries(liveBhaGridMap)) bhaGridMap[k] = mergeNonEmpty(bhaGridMap[k], v);
    for (const [k, v] of Object.entries(liveActivityMap)) bhaGridMap[k] = mergeNonEmpty(bhaGridMap[k], v);
    // Highest-fidelity source: ActivityLogs ⨯ SurveySheetEntries time-window join.
    for (const [k, v] of Object.entries(activityLogMap)) bhaGridMap[k] = mergeNonEmpty(bhaGridMap[k], v);
    prog(92, "Building CSV…");
    results.bhaCsv      = buildBhaCsv(assemblies, items, tools, interceptedMap, bhaGridMap, customAsmMap, itemAsmMap);
    results.bhaRowCount = results.bhaCsv.split("\r\n").length - 1;
    results.bhaCount    = assemblies.length;
    results.slideByDayCsv = slideByDayCsv;
    results.slideByDayRowCount = slideByDayRowCount;
  }

  return results;
};

// ── Port handler — streams progress back to popup during fetch ────────────────
chrome.runtime.onConnect.addListener((port) => {
  if (port.name !== "fieldcap-fetch") return;

  port.onMessage.addListener(async (msg) => {
    if (msg.type !== "FETCH_ALL") return;

    const { jobId, flags = { jobDetails: true, crew: true, bha: true } } = msg;
    const send = (obj) => { try { port.postMessage(obj); } catch (_) {} };

    if (!jobId) {
      send({ type: "DONE", ok: false, error: "Job ID is required." });
      return;
    }

    try {
      send({ type: "PROGRESS", pct: 3, label: "Scraping live data…" });
      let liveData = flags.bha ? await scrapeNowFromFieldCapTabs(jobId) : { bhaRows: [], activityRows: [] };
      if (flags.bha && liveData.bhaRows.length === 0) {
        await sleep(1200);
        liveData = await scrapeNowFromFieldCapTabs(jobId);
      }

      if (flags.bha && (liveData.bhaRows.length > 0 || liveData.activityRows.length > 0)) {
        const liveMap         = toBhaGridMap(liveData.bhaRows);
        const liveActivityMap = toActivityMetresMap(liveData.activityRows);
        chrome.storage.local.get([KEY_BHA_GRID], (storedGrid) => {
          const byJob      = storedGrid[KEY_BHA_GRID] ?? {};
          const mergedJob  = { ...(byJob[String(jobId)] ?? {}) };
          for (const [k, v] of Object.entries(liveMap))         mergedJob[k] = mergeNonEmpty(mergedJob[k], v);
          for (const [k, v] of Object.entries(liveActivityMap)) mergedJob[k] = mergeNonEmpty(mergedJob[k], v);
          byJob[String(jobId)] = mergedJob;
          chrome.storage.local.set({ [KEY_BHA_GRID]: byJob });
        });
      }

      const results = await fetchAll(
        jobId, flags,
        liveData.bhaRows, liveData.activityRows,
        (pct, label) => send({ type: "PROGRESS", pct, label })
      );

      const meta = {
        jobId,
        builtAt:          new Date().toISOString(),
        crewRows:         results.crew?.length          ?? 0,
        bhaRows:          results.bhaRowCount           ?? 0,
        bhaCount:         results.bhaCount              ?? 0,
        slideByDayRows:   results.slideByDayRowCount    ?? 0,
        liveBhaRows:      liveData.bhaRows.length,
        liveActivityRows: liveData.activityRows.length,
        flags,
      };

      send({ type: "PROGRESS", pct: 96, label: "Storing to cache…" });

      const stored = {};
      if (results.jobDetailsCsv) stored[KEY_CSV_JOB]       = results.jobDetailsCsv;
      if (results.crewCsv)       stored[KEY_CSV_CREW]       = results.crewCsv;
      if (results.bhaCsv)        stored[KEY_CSV_BHA]        = results.bhaCsv;
      if (results.slideByDayCsv) stored[KEY_CSV_SLIDE_DAY]  = results.slideByDayCsv;
      stored[KEY_META] = meta;

      chrome.storage.local.set(stored, () => {
        send({ type: "DONE", ok: true, ...meta });
      });
    } catch (err) {
      send({ type: "DONE", ok: false, error: err.message });
    }
  });
});

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
          slideByDayRows: results.slideByDayRowCount ?? 0,
          liveBhaRows:  liveData.bhaRows.length,
          liveActivityRows: liveData.activityRows.length,
          flags,
        };

        const stored = {};
        if (results.jobDetailsCsv) stored[KEY_CSV_JOB]  = results.jobDetailsCsv;
        if (results.crewCsv)       stored[KEY_CSV_CREW] = results.crewCsv;
        if (results.bhaCsv)        stored[KEY_CSV_BHA]  = results.bhaCsv;
        if (results.slideByDayCsv) stored[KEY_CSV_SLIDE_DAY] = results.slideByDayCsv;
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
    const keyMap = { job: KEY_CSV_JOB, crew: KEY_CSV_CREW, bha: KEY_CSV_BHA, slideDay: KEY_CSV_SLIDE_DAY };
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
        slideDay: `fieldcap-job-${jid}-slide-rotate-metres-by-day.csv`,
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
      [KEY_META, KEY_CSV_JOB, KEY_CSV_CREW, KEY_CSV_BHA, KEY_CSV_SLIDE_DAY, KEY_INTERCEPT, KEY_BHA_GRID],
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
