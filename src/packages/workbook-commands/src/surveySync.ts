import { csvParseLine } from "./planImport";
import { numeric, SheetBook, type CellArea } from "./sheetBook";
import { SS_SHEET } from "./slidesheet";

// Port of MDL_SurveySync: audit & sync Slidesheet E/F/G (MD, Inc, Azi) against
// a Well Seeker X Survey Report (.xlsx or .csv).

const SURV_FIRST = 13;
const SURV_LAST = 520;
const TOL_EXACT_MD = 0.05;
const TOL_NEAR_MD = 0.5;
const TOL_DIFF = 0.005;

export interface WellSeekerStation {
  md: number;
  inc: number;
  azi: number;
}

export interface WellSeekerParseResult {
  stations: WellSeekerStation[];
  interpSkipped: number;
  tieMd: number | null;
}

export interface SurveyDiffItem {
  row: number;
  ssMd: number;
  ssInc: number;
  ssAzi: number;
  wsMd: number;
  wsInc: number;
  wsAzi: number;
  shifted: boolean;
}

export interface SurveyAuditResult {
  ok: boolean;
  blocked: boolean;
  reason: string | null;
  totalSlidesheet: number;
  matched: number;
  diffs: SurveyDiffItem[];
  shiftHits: number;
  shiftRange: [number, number] | null;
  deeperInWellSeeker: number;
  interpSkipped: number;
  tieNote: string | null;
  maxSsMd: number;
}

function normalizeHeader(value: unknown): string {
  const raw = String(value ?? "").trim().toUpperCase();
  const paren = raw.indexOf("(");
  const base = paren >= 0 ? raw.slice(0, paren) : raw;
  return base.replace(/[^A-Z0-9]/g, "");
}

function isSummaryRow(book: SheetBook, r: number): boolean {
  return [4, 5, 6].some((c) => (book.formula(SS_SHEET, r, c) ?? "").toUpperCase().includes("LOOKUP"));
}

/** Parse a 2D cell grid (e.g. from Well Seeker X Survey Report worksheet or CSV). */
export function parseWellSeekerGrid(rows: unknown[][], tieThreshold = 10): WellSeekerParseResult {
  let hdrRow = -1;
  let colMd = -1;
  let colInc = -1;
  let colAzi = -1;
  let colDls = -1;
  const limit = Math.min(rows.length, 150);

  for (let r = 0; r < limit; r += 1) {
    const row = rows[r] ?? [];
    colMd = -1;
    colInc = -1;
    colAzi = -1;
    colDls = -1;
    for (let c = 0; c < row.length; c += 1) {
      const tok = normalizeHeader(row[c]);
      if (colMd < 0 && (tok === "MD" || tok === "MEASUREDDEPTH")) colMd = c;
      if (colInc < 0 && (tok === "INC" || tok === "INCLINATION")) colInc = c;
      if (colAzi < 0 && (tok === "AZI" || tok === "AZM" || tok === "AZIMUTH")) colAzi = c;
      if (colDls < 0 && (tok === "DLS" || tok === "DOGLEGRATE")) colDls = c;
    }
    if (colMd >= 0 && colInc >= 0 && colAzi >= 0) {
      hdrRow = r;
      break;
    }
  }
  if (hdrRow < 0) return { stations: [], interpSkipped: 0, tieMd: null };

  const stations: WellSeekerStation[] = [];
  let interpSkipped = 0;
  let tieMd: number | null = null;

  for (let r = hdrRow + 1; r < rows.length; r += 1) {
    const row = rows[r] ?? [];
    const md = row[colMd] === "" || row[colMd] == null ? null : numeric(row[colMd]);
    const inc = row[colInc] === "" || row[colInc] == null ? null : numeric(row[colInc]);
    const azi = row[colAzi] === "" || row[colAzi] == null ? null : numeric(row[colAzi]);
    if (md == null || inc == null || azi == null) continue;
    if (inc < -0.01 || inc > 180.01) continue;

    const hasPtb = row.some((cell) => {
      if (typeof cell !== "string") return false;
      const u = cell.toUpperCase();
      return u.includes("PTB") || u.includes("PROJECTED") || u.includes("PROJECTION");
    });
    if (hasPtb) continue;

    if (md > 0.01 && md <= 15 && inc <= 0.01) tieMd = md;
    if (md <= tieThreshold) continue;

    if (colDls >= 0) {
      const dls = row[colDls] === "" || row[colDls] == null ? null : numeric(row[colDls]);
      if (dls == null) {
        interpSkipped += 1;
        continue;
      }
    }
    stations.push({ md, inc, azi });
  }
  return { stations, interpSkipped, tieMd };
}

/** Parse CSV text into a Well Seeker X survey station list. */
export function parseWellSeekerCsv(text: string, tieThreshold = 10): WellSeekerParseResult {
  const rows = text.split(/\r?\n/).map((line) => csvParseLine(line));
  return parseWellSeekerGrid(rows, tieThreshold);
}

/** Tie-on MD threshold from Slidesheet Row 12 (E12 + 0.05 m, or 10 m if blank). */
export function slidesheetTieThreshold(book: SheetBook): number {
  const e12 = book.value(SS_SHEET, SURV_FIRST - 1, 5);
  const n = e12 == null || e12 === "" ? null : numeric(e12);
  return n != null ? n + 0.05 : 10;
}

function isShiftedFromNeighbor(inc: number, azi: number, idx: number, ws: WellSeekerStation[]): boolean {
  if (Math.abs(inc - ws[idx]!.inc) <= 0.1 && Math.abs(azi - ws[idx]!.azi) <= 0.2) return false;
  if (idx > 0 && Math.abs(inc - ws[idx - 1]!.inc) <= 0.02 && Math.abs(azi - ws[idx - 1]!.azi) <= 0.05) return true;
  if (idx + 1 < ws.length && Math.abs(inc - ws[idx + 1]!.inc) <= 0.02 && Math.abs(azi - ws[idx + 1]!.azi) <= 0.05) return true;
  return false;
}

/** Audit Slidesheet rows 13..520 against parsed Well Seeker X stations. */
export function auditSlidesheetSurveys(book: SheetBook, parsed: WellSeekerParseResult): SurveyAuditResult {
  if (!book.hasSheet(SS_SHEET)) {
    return {
      ok: false, blocked: false, reason: "Slidesheet not found.",
      totalSlidesheet: 0, matched: 0, diffs: [], shiftHits: 0, shiftRange: null,
      deeperInWellSeeker: 0, interpSkipped: parsed.interpSkipped, tieNote: null, maxSsMd: 0,
    };
  }
  const ss: Array<{ row: number; md: number; inc: number; azi: number }> = [];
  for (let r = SURV_FIRST; r <= SURV_LAST; r += 1) {
    if (isSummaryRow(book, r)) continue;
    const vM = book.value(SS_SHEET, r, 5);
    const vI = book.value(SS_SHEET, r, 6);
    const vA = book.value(SS_SHEET, r, 7);
    const md = vM == null || vM === "" ? null : numeric(vM);
    const inc = vI == null || vI === "" ? null : numeric(vI);
    const azi = vA == null || vA === "" ? null : numeric(vA);
    if (md != null && inc != null && azi != null) ss.push({ row: r, md, inc, azi });
  }
  if (ss.length === 0) {
    return {
      ok: false, blocked: false, reason: `Slidesheet has no entered surveys yet (${parsed.stations.length} stations in file).`,
      totalSlidesheet: 0, matched: 0, diffs: [], shiftHits: 0, shiftRange: null,
      deeperInWellSeeker: parsed.stations.length, interpSkipped: parsed.interpSkipped, tieNote: null, maxSsMd: 0,
    };
  }

  const ws = parsed.stations;
  const match = ss.map(() => -1);
  const used = ws.map(() => false);
  let matched = 0;

  for (const tol of [TOL_EXACT_MD, TOL_NEAR_MD]) {
    for (let i = 0; i < ss.length; i += 1) {
      if (match[i]! >= 0) continue;
      let bestJ = -1;
      let bestD = tol + 1e-4;
      for (let j = 0; j < ws.length; j += 1) {
        if (used[j]) continue;
        const d = Math.abs(ss[i]!.md - ws[j]!.md);
        if (d < bestD) {
          bestD = d;
          bestJ = j;
        }
      }
      if (bestJ >= 0) {
        match[i] = bestJ;
        used[bestJ] = true;
        matched += 1;
      }
    }
  }

  if (matched * 2 < ss.length) {
    return {
      ok: false,
      blocked: true,
      reason: `Only ${matched} of ${ss.length} Slidesheet survey depths matched this file. Check that you selected the right well's Well Seeker X export.`,
      totalSlidesheet: ss.length,
      matched,
      diffs: [],
      shiftHits: 0,
      shiftRange: null,
      deeperInWellSeeker: 0,
      interpSkipped: parsed.interpSkipped,
      tieNote: null,
      maxSsMd: ss[ss.length - 1]!.md,
    };
  }

  const diffs: SurveyDiffItem[] = [];
  let shiftHits = 0;
  let shiftFirst = 0;
  let shiftLast = 0;

  for (let i = 0; i < ss.length; i += 1) {
    const j = match[i]!;
    if (j < 0) continue;
    const s = ss[i]!;
    const w = ws[j]!;
    if (Math.abs(s.md - w.md) > TOL_DIFF || Math.abs(s.inc - w.inc) > TOL_DIFF || Math.abs(s.azi - w.azi) > TOL_DIFF) {
      const shifted = isShiftedFromNeighbor(s.inc, s.azi, j, ws);
      if (shifted) {
        shiftHits += 1;
        if (shiftFirst === 0) shiftFirst = s.row;
        shiftLast = s.row;
      }
      diffs.push({ row: s.row, ssMd: s.md, ssInc: s.inc, ssAzi: s.azi, wsMd: w.md, wsInc: w.inc, wsAzi: w.azi, shifted });
    }
  }

  const maxSsMd = ss[ss.length - 1]!.md;
  const deeperInWellSeeker = ws.filter((w, j) => !used[j] && w.md > maxSsMd + TOL_EXACT_MD).length;

  let tieNote: string | null = null;
  if (parsed.tieMd != null) {
    const e12 = book.value(SS_SHEET, SURV_FIRST - 1, 5);
    const ssTie = e12 == null || e12 === "" ? 0 : (numeric(e12) ?? 0);
    if (Math.abs(ssTie - parsed.tieMd) > 0.05) {
      tieNote = `Tie-on Row 12 (E12=${ssTie.toFixed(2)}m) differs from Well Seeker KBE (${parsed.tieMd.toFixed(2)}m).`;
    }
  }

  return {
    ok: true,
    blocked: false,
    reason: null,
    totalSlidesheet: ss.length,
    matched,
    diffs,
    shiftHits,
    shiftRange: shiftFirst > 0 ? [shiftFirst, shiftLast] : null,
    deeperInWellSeeker,
    interpSkipped: parsed.interpSkipped,
    tieNote,
    maxSsMd,
  };
}

/** Write ONLY matched E (when MD differs), F, G values to Slidesheet, plus optional Setup status. */
export function applySurveySync(
  book: SheetBook,
  audit: SurveyAuditResult,
  statusText: string,
  anchorCell: CellArea | null = null,
): number {
  for (const d of audit.diffs) {
    if (Math.abs(d.ssMd - d.wsMd) > TOL_DIFF) book.setValue(SS_SHEET, d.row, 5, d.wsMd);
    book.setValue(SS_SHEET, d.row, 6, d.wsInc);
    book.setValue(SS_SHEET, d.row, 7, d.wsAzi);
  }
  if (book.hasSheet("Setup")) {
    const row = anchorCell ? anchorCell.row : 31;
    const col = anchorCell ? (anchorCell.col2 ?? anchorCell.col) + 1 : 13;
    book.setValue("Setup", row, col, statusText);
  }
  return audit.diffs.length;
}

/** Human-readable preview prompt matching MDL_SurveySync's MsgBox. */
export function formatSurveyAuditPrompt(fileName: string, audit: SurveyAuditResult): string {
  const lines = [
    `File: ${fileName}`,
    `Matched: ${audit.matched} / ${audit.totalSlidesheet} Slidesheet surveys  |  Differences: ${audit.diffs.length} row(s)`,
  ];
  if (audit.shiftHits >= 2 && audit.shiftRange) {
    lines.push(`WARNING: Paste-shift detected around rows ${audit.shiftRange[0]}-${audit.shiftRange[1]} (${audit.shiftHits} shifted rows)!`);
  }
  lines.push("");
  const maxShow = 16;
  for (let i = 0; i < Math.min(audit.diffs.length, maxShow); i += 1) {
    const d = audit.diffs[i]!;
    const mdTag = Math.abs(d.ssMd - d.wsMd) > TOL_DIFF ? ` [MD ${d.ssMd.toFixed(2)}->${d.wsMd.toFixed(2)}]` : "";
    lines.push(`Row ${d.row} (${d.wsMd.toFixed(2)}m): ${d.ssInc.toFixed(2)}/${d.ssAzi.toFixed(2)} -> ${d.wsInc.toFixed(2)}/${d.wsAzi.toFixed(2)}${mdTag}`);
  }
  if (audit.diffs.length > maxShow) lines.push(`... and ${audit.diffs.length - maxShow} more row(s).`);
  if (audit.deeperInWellSeeker > 0) {
    lines.push("", `Note: ${audit.deeperInWellSeeker} deeper survey(s) in Well Seeker X past ${audit.maxSsMd.toFixed(2)}m (not auto-appended).`);
  }
  if (audit.tieNote) lines.push(`Note: ${audit.tieNote}`);
  lines.push(
    "",
    `Only columns E/F/G (MD/Inc/Azi) on these ${audit.diffs.length} matched row(s) will be updated.`,
    "Slides, bit depths, notes, and all formulas (H/AL/AM) stay untouched.",
    "",
    "Apply these updates and refresh the proximity gauge now?",
  );
  return lines.join("\n");
}
