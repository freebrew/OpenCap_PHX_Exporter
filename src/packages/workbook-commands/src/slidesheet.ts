import { a1Of, cellText, excelRound, fixed2, isNumber, numeric, rgb, SheetBook } from "./sheetBook";

// Port of MDL_SlidesheetClear (Clear Ranges, target window, Y/Z comments, newest-survey bar)
// and the MDL_ContDI good-survey rule it shares with the gauge and the DL Tracker.

export const SS_SHEET = "Slidesheet";
export const PLANSEC_SHEET = "_OC_PlanSec";
const SS_FIRST = 13;
const Y_FIRST = 12;
const Y_LAST = 505;
const Y_DATA_FIRST = 13;
const Y_LABEL = "RKB";
const TGT_FIRST = 2;
const TGT_LAST = 5;
const FULL_TGT_MD_COL = 14;
const FULL_TGT_FIRST = 3;
const FULL_TGT_ROWS = 80;
export const OVR_MD = 19;
export const OVR_INC = 20;
export const OVR_AZM = 21;
export const OVR_TVD = 22;
const SLIDE_IGNORE_M = 1;

export const TARGET_NAMES = ["NUDGE", "VERTICAL", "KOP", "TANGENT", "SOT", "EOT", "ICP", "BUILD", "HEEL", "TD"] as const;
export const GOOD_SURVEY_FILL = rgb(204, 255, 204);
export const COMMENT_YELLOW = rgb(255, 255, 204);
const TARGET_HIGHLIGHT = rgb(198, 239, 206);
const NEWEST_FILL = GOOD_SURVEY_FILL;

const COL = { C: 3, D: 4, E: 5, F: 6, G: 7, H: 8, J: 10, T: 20, U: 21, V: 22, W: 23, X: 24, Y: 25, Z: 26, AJ: 36, AO: 41, AR: 44, AS: 45, AT: 46 } as const;

export const CLEAR_SUMMARY = [
  "Cleared:",
  `  D/F/G/T/U 13:${Y_LAST}`,
  "  Sail waypoints AC14:AD33",
  "  Targets U2:X5",
  `  Comments Y12:Y${Y_LAST} / ROT Z13:Z${Y_LAST} (auto text restored where inputs exist)`,
].join("\n");

/** Width in the same units as the sheet's column widths (px) for one string in a cell font. */
export type MeasureText = (text: string, font: { family: string; sizePt: number; bold: boolean }) => number;

export const roughMeasure: MeasureText = (text, font) => text.length * font.sizePt * 0.5 * (96 / 72);

/** VBA Val(): the leading number of the text, 0 when there is none. */
export function vbaVal(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const match = /^\s*[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(String(value ?? ""));
  return match ? Number(match[0]) : 0;
}

// ---------------------------------------------------------------- good survey rows

/** E is a number and F/G are numbers painted the survey green (colour index 35). */
export function isGoodSurveyRow(book: SheetBook, r: number): boolean {
  if (r < SS_FIRST || r > Y_LAST) return false;
  if (!isNumber(book.value(SS_SHEET, r, COL.E))) return false;
  for (const c of [COL.F, COL.G]) {
    if (!isNumber(book.value(SS_SHEET, r, c))) return false;
    if (book.fill(SS_SHEET, r, c) !== GOOD_SURVEY_FILL) return false;
  }
  return true;
}

/** Deepest good-survey row on the Slidesheet (13:505); 0 = none. */
export function lastSurveyRow(book: SheetBook): number {
  for (let r = Y_LAST; r >= SS_FIRST; r -= 1) if (isGoodSurveyRow(book, r)) return r;
  return 0;
}

const STRIP_ROW = 506;
const STRIP_COLS = [COL.E, COL.F, COL.G, COL.H, COL.J] as const;
const DEPTH_METRES = '#####.00"m"';

function sameFormula(have: string | null, want: string): boolean {
  const compact = (value: string) => value.replace(/^=/, "").replace(/\s+/g, "").toUpperCase();
  return compact(have ?? "") === compact(want);
}

/**
 * Deepest station someone entered: survey MD, inclination and azimuth are numbers,
 * and inclination and azimuth are not formulas. Projection rows calculate those
 * angles, so they are not a survey. The green bar stays on lastSurveyRow.
 */
export function lastEnteredSurveyRow(book: SheetBook): number {
  for (let r = Y_LAST; r >= SS_FIRST; r -= 1) {
    if (!isNumber(book.value(SS_SHEET, r, COL.E))) continue;
    if (!isNumber(book.value(SS_SHEET, r, COL.F)) || book.hasFormula(SS_SHEET, r, COL.F)) continue;
    if (!isNumber(book.value(SS_SHEET, r, COL.G)) || book.hasFormula(SS_SHEET, r, COL.G)) continue;
    return r;
  }
  return 0;
}

/** Data E7:E11 read Slidesheet row 506. Point that row at the latest entered survey. */
export function syncLastSurveyStrip(book: SheetBook): boolean {
  if (!book.hasSheet(SS_SHEET)) return false;
  const before = book.writes.length;
  const r = lastEnteredSurveyRow(book);
  if (r < Y_DATA_FIRST) {
    for (const col of STRIP_COLS) {
      if (book.value(SS_SHEET, STRIP_ROW, col) != null || book.hasFormula(SS_SHEET, STRIP_ROW, col)) {
        book.clear(SS_SHEET, { row: STRIP_ROW, col, row2: STRIP_ROW, col2: col });
      }
    }
  } else {
    for (const col of STRIP_COLS) {
      const want = `=${a1Of(r, col)}`;
      if (!sameFormula(book.formula(SS_SHEET, STRIP_ROW, col), want)) book.setFormula(SS_SHEET, STRIP_ROW, col, want);
    }
  }
  return book.writes.length > before;
}

/** Data!E7 is a depth in metres. A degree format belongs on inclination and azimuth. */
export function bindSurveyDepthMetres(book: SheetBook): boolean {
  if (!book.hasSheet("Data")) return false;
  const pattern = book.style("Data", 7, 5)?.n?.pattern ?? "";
  if (pattern.includes("30")) return false;
  if (/m/i.test(pattern) && !pattern.includes("°")) return false;
  if (pattern !== "" && !pattern.includes("°")) return false;
  if (book.writes.some((write) => write.kind === "format" && write.sheet === "Data" && write.row === 7 && write.col === 5)) return false;
  book.setFormat("Data", 7, 5, DEPTH_METRES);
  return true;
}

function lastSurveyBitMd(book: SheetBook): number {
  const r = lastSurveyRow(book);
  if (r < Y_DATA_FIRST) return 0;
  return numeric(book.value(SS_SHEET, r, COL.D)) ?? 0;
}

// ---------------------------------------------------------------- Clear Ranges

/** ClearSlidesheetRanges up to the comment refresh (which needs a recalculated sheet). */
export function clearSlidesheetRanges(book: SheetBook): void {
  for (const col of ["D", "F", "G", "T", "U"]) book.clear(SS_SHEET, `${col}13:${col}${Y_LAST}`);
  book.clear(SS_SHEET, "AC14:AD33");
  book.clear(SS_SHEET, "U2:X5");
  book.clear(SS_SHEET, `Y12:Y${Y_LAST}`);
  book.clear(SS_SHEET, `Z13:Z${Y_LAST}`);
  syncPlanTargetWindow(book);
  book.setValue(SS_SHEET, Y_FIRST, COL.Y, Y_LABEL);
}

// ---------------------------------------------------------------- target window

export function seedNameFromText(text: string): string {
  const s = text.trim().toLowerCase();
  if (s === "") return "";
  if (s.includes("toe") || s.indexOf("td") === 0 || s.includes(" td") || s.startsWith("td ") || s === "td") return "TD";
  if (s.includes("heel")) return "HEEL";
  if (s.includes("kop") || s.includes("kick")) return "KOP";
  if (s.includes("sot") || s.includes("start of turn") || s.includes("b & t") || s.includes("b&t")) return "SOT";
  if (s.includes("eot") || s.includes("end of turn")) return "EOT";
  if (s.includes("tang")) return "TANGENT";
  if (s.includes("nudge")) return "NUDGE";
  if (s.includes("back to vert") || s.includes("btv")) return "VERTICAL";
  if (s.includes("icp")) return "ICP";
  if (s.includes("build")) return "BUILD";
  return "";
}

export function isKnownTargetName(name: string): boolean {
  return (TARGET_NAMES as readonly string[]).includes(name.trim().toUpperCase());
}

function pickedNum(override: unknown, planValue: number): number {
  const n = override == null || override === "" ? null : numeric(override);
  return n ?? planValue;
}

export interface NamedTarget {
  planRow: number;
  name: string;
  md: number;
  inc: number;
  azm: number;
  tvd: number;
}

function planLastRow(book: SheetBook): number {
  let last = 1;
  for (let r = 3; r <= 2000; r += 1) if (cellText(book.value(PLANSEC_SHEET, r, 1)) !== "") last = r;
  return last;
}

/** LoadNamedPlanTargets / LoadTargetEditor: named plan stations with any saved overrides. */
export function loadNamedTargets(book: SheetBook, limit = 80): NamedTarget[] {
  if (!book.hasSheet(PLANSEC_SHEET)) return [];
  const out: NamedTarget[] = [];
  const lastR = planLastRow(book);
  for (let r = 3; r <= lastR; r += 1) {
    const vMd = book.value(PLANSEC_SHEET, r, 1);
    const md = vMd == null || vMd === "" ? null : numeric(vMd);
    if (md == null) continue;
    const autoName = cellText(book.value(PLANSEC_SHEET, r, 11)).toUpperCase();
    const userName = cellText(book.value(PLANSEC_SHEET, r, 12)).toUpperCase();
    let name = userName !== "" ? userName : autoName;
    if (name === "") name = seedNameFromText(cellText(book.value(PLANSEC_SHEET, r, 10)));
    if (!isKnownTargetName(name)) continue;
    if (out.length >= limit) break;
    out.push({
      planRow: r,
      name,
      md: pickedNum(book.value(PLANSEC_SHEET, r, OVR_MD), md),
      inc: pickedNum(book.value(PLANSEC_SHEET, r, OVR_INC), vbaVal(book.value(PLANSEC_SHEET, r, 2))),
      azm: pickedNum(book.value(PLANSEC_SHEET, r, OVR_AZM), vbaVal(book.value(PLANSEC_SHEET, r, 3))),
      tvd: pickedNum(book.value(PLANSEC_SHEET, r, OVR_TVD), vbaVal(book.value(PLANSEC_SHEET, r, 4))),
    });
  }
  return out;
}

function writeFullTargetTable(book: SheetBook, targets: NamedTarget[]): void {
  const start = targets.find((target) => target.name === "KOP" || target.name === "TANGENT")?.md ?? 0;
  book.setValue(PLANSEC_SHEET, 1, FULL_TGT_MD_COL, start);
  const lastR = FULL_TGT_FIRST + FULL_TGT_ROWS - 1;
  book.clear(PLANSEC_SHEET, { row: FULL_TGT_FIRST, col: 14, row2: lastR, col2: 18 });
  for (let i = 0; i < targets.length && i < FULL_TGT_ROWS; i += 1) {
    const target = targets[i]!;
    const r = FULL_TGT_FIRST + i;
    book.setValue(PLANSEC_SHEET, r, 14, target.md);
    book.setValue(PLANSEC_SHEET, r, 15, target.inc);
    book.setValue(PLANSEC_SHEET, r, 16, target.azm);
    book.setValue(PLANSEC_SHEET, r, 17, target.tvd);
    book.setValue(PLANSEC_SHEET, r, 18, target.name);
    if (target.name === "HEEL") break;
  }
}

function sameTargetCell(current: unknown, next: number | null): boolean {
  const blank = current == null || cellText(current) === "";
  if (next == null) return blank;
  if (blank) return false;
  const n = numeric(current);
  return n != null && Math.abs(n - next) < 0.005;
}

function writeTargetNumber(book: SheetBook, r: number, c: number, next: number | null): void {
  if (!sameTargetCell(book.value(SS_SHEET, r, c), next)) {
    if (next == null) book.clear(SS_SHEET, { row: r, col: c, row2: r, col2: c });
    else book.setValue(SS_SHEET, r, c, next);
    book.setFormat(SS_SHEET, r, c, "0.00");
  }
}

function writeTargetLabel(book: SheetBook, r: number, c: number, next: string): void {
  if (cellText(book.value(SS_SHEET, r, c)) !== next) {
    if (next === "") book.clear(SS_SHEET, { row: r, col: c, row2: r, col2: c });
    else book.setValue(SS_SHEET, r, c, next);
  }
}

/** SyncPlanTargetWindowOnSheet: four named plan targets on T2:Y5, sliding past the bit. */
export function syncPlanTargetWindow(book: SheetBook): void {
  const targets = loadNamedTargets(book);
  if (targets.length < 1) return;
  writeFullTargetTable(book, targets);
  const bitMd = lastSurveyBitMd(book);
  const n = targets.length;
  let aim = n;
  for (let i = 1; i <= n; i += 1) {
    if (targets[i - 1]!.md > bitMd + 0.005) {
      aim = i;
      break;
    }
  }
  let start = aim;
  if (start + 3 > n) start = n - 3;
  if (start < 1) start = 1;
  for (let vis = 0; vis <= 3; vis += 1) {
    const r = TGT_FIRST + vis;
    const target = start + vis <= n ? targets[start + vis - 1]! : null;
    writeTargetLabel(book, r, COL.T, target ? `TAR${start + vis}` : "");
    writeTargetNumber(book, r, COL.U, target?.md ?? null);
    writeTargetNumber(book, r, COL.V, target?.inc ?? null);
    writeTargetNumber(book, r, COL.W, target?.azm ?? null);
    writeTargetNumber(book, r, COL.X, target?.tvd ?? null);
    writeTargetLabel(book, r, COL.Y, target?.name ?? "");
  }
}

/** HighlightActiveTargetOnSheet: the T:Y row of the target the last good survey is aiming at. */
export function highlightActiveTargetOnSheet(book: SheetBook): void {
  let lastSurv = 0;
  let bitMd = 0;
  let aimMd = 0;
  for (let r = Y_FIRST; r <= Y_LAST; r += 1) {
    const vF = book.value(SS_SHEET, r, COL.F);
    if (vF == null || cellText(vF) === "" || numeric(vF) == null) continue;
    if (!isGoodSurveyRow(book, r)) continue;
    lastSurv = r;
    const vD = numeric(book.value(SS_SHEET, r, COL.D));
    if (vD != null) bitMd = vD;
    const vAo = book.value(SS_SHEET, r, COL.AO);
    aimMd = vAo == null || vAo === "" ? 0 : numeric(vAo) ?? 0;
  }
  let active = 0;
  if (lastSurv > 0) {
    if (aimMd > 0) {
      for (let r = TGT_FIRST; r <= TGT_LAST; r += 1) {
        const md = numeric(book.value(SS_SHEET, r, COL.U));
        if (md != null && Math.abs(md - aimMd) < 0.005) {
          active = r;
          break;
        }
      }
    }
    if (active === 0 && bitMd > 0) {
      for (let r = TGT_FIRST; r <= TGT_LAST; r += 1) {
        const md = numeric(book.value(SS_SHEET, r, COL.U));
        if (md != null && md > bitMd) {
          active = r;
          break;
        }
      }
    }
  }
  book.setFill(SS_SHEET, { row: TGT_FIRST, col: COL.T, row2: TGT_LAST, col2: COL.Y }, null);
  if (active >= TGT_FIRST && active <= TGT_LAST) book.setFill(SS_SHEET, { row: active, col: COL.T, row2: active, col2: COL.Y }, TARGET_HIGHLIGHT);
}

/** HighlightActiveTarget / SyncPlanTargetWindow / TargetEditorCommit. */
export function highlightActiveTarget(book: SheetBook): void {
  syncPlanTargetWindow(book);
  highlightActiveTargetOnSheet(book);
}

/** TargetEditorStore: one named station. Returns a message when the row cannot be saved. */
export function storeTarget(book: SheetBook, planRow: number, name: string, md: number, inc: number, azm: number, tvd: number): string {
  const show = name.trim().toUpperCase();
  if (!isKnownTargetName(show)) return `Name must be one of: ${TARGET_NAMES.join(",")}`;
  let azmD = azm;
  if (azmD >= 360 && azmD < 360.001) azmD = 0;
  if (md <= 0 || inc < 0 || inc > 180 || azmD < 0 || azmD >= 360 || tvd < 0) return `${show}: MD, INC (0-180), AZM (0-360) and TVD have to be in range.`;
  if (!book.hasSheet(PLANSEC_SHEET)) return "Plan sections are not loaded.";
  if (planRow < 3) return "That target row is not on the plan.";
  book.setValue(PLANSEC_SHEET, planRow, 12, show);
  writeOverride(book, planRow, OVR_MD, 1, md);
  writeOverride(book, planRow, OVR_INC, 2, inc);
  writeOverride(book, planRow, OVR_AZM, 3, azmD);
  writeOverride(book, planRow, OVR_TVD, 4, tvd);
  return "";
}

function writeOverride(book: SheetBook, row: number, overrideCol: number, planCol: number, entered: number): void {
  const plan = book.value(PLANSEC_SHEET, row, planCol);
  const planValue = plan == null || plan === "" ? null : numeric(plan);
  if (planValue != null && Math.abs(entered - planValue) < 0.005) {
    book.clear(PLANSEC_SHEET, { row, col: overrideCol, row2: row, col2: overrideCol });
    return;
  }
  book.setValue(PLANSEC_SHEET, row, overrideCol, entered);
}

// ---------------------------------------------------------------- Y comments and Z rotate

function isAutoSlideCommentText(text: string): boolean {
  const s = text.trim();
  if (s === "") return true;
  if (/^Sliding .*\| BURR /.test(s) || /^Sliding .*\| DLR /.test(s)) return true;
  if (/^Sliding \d.*m @ /.test(s)) return true;
  if (s.startsWith("Sliding ") && (/BURR /i.test(s) || /DLR /i.test(s))) return true;
  return /m rot/i.test(s);
}

function isAutoRotText(text: string): boolean {
  const s = text.trim();
  if (s === "") return true;
  if (s.endsWith("m ROT") || s.endsWith("m rot")) return true;
  return numeric(s) != null;
}

/** ProjInstructedSlideM: AS capped at this stand's C, two decimals. */
export function instructedSlideM(metresToSlide: unknown, maxSlide: unknown): number {
  let m = Math.max(0, numeric(metresToSlide) ?? 0);
  let cap = Math.max(0, numeric(maxSlide) ?? 0);
  if (cap > 0 && m > cap) m = cap;
  m = excelRound(m, 2);
  if (cap > 0) {
    cap = excelRound(cap, 2);
    if (m > cap) m = cap;
  }
  return m;
}

/** ProjSlideComment: "Sliding <m> @ <TF>" and "DLR <burr>", joined by Chr(1). */
export function slideComment(metresToSlide: unknown, tfText: unknown, burr: unknown, maxSlide: unknown): string {
  const b = burr == null || burr === "" ? null : numeric(burr);
  if (b == null) return "";
  const m = instructedSlideM(metresToSlide, maxSlide);
  const tf = cellText(tfText) || "-";
  if (m < SLIDE_IGNORE_M) return "";
  return `Sliding ${fixed2(m)}m @ ${tf}\u0001DLR ${fixed2(b)}`;
}

function fitLeftRight(left: string, right: string, widthPx: number, measure: (text: string) => number, spacePx: number): string {
  const marginPx = 2 * (96 / 72);
  const avail = widthPx - marginPx;
  const pad = Math.max(2, Math.round((avail - measure(left) - measure(right)) / (spacePx > 0 ? spacePx : 1)));
  return `${left}${" ".repeat(pad)}${right}`;
}

/** WriteRotateMetresOnSheet: Z[r] = round(C[r+1]) - round(T[r]), at least 0. A typed Z is left alone. */
export function writeRotateMetres(book: SheetBook, r: number): void {
  const zText = cellText(book.value(SS_SHEET, r, COL.Z));
  if (!isAutoRotText(zText)) return;
  const courseRow = r + 1;
  const course = courseRow <= Y_LAST ? book.value(SS_SHEET, courseRow, COL.C) : null;
  if (!isNumber(course)) {
    if (zText !== "") book.clear(SS_SHEET, { row: r, col: COL.Z, row2: r, col2: COL.Z });
    return;
  }
  const slidValue = book.value(SS_SHEET, r, COL.T);
  const slid = isNumber(slidValue) ? slidValue : 0;
  const remain = Math.max(0, excelRound(course, 2) - excelRound(slid, 2));
  const next = excelRound(remain, 2);
  const current = book.value(SS_SHEET, r, COL.Z);
  if (isNumber(current) && Math.abs(current - next) < 1e-9) return;
  book.setValue(SS_SHEET, r, COL.Z, next);
  book.setFormat(SS_SHEET, r, COL.Z, "0.00");
}

/** RefreshRotateMetres for an edited block of C, D or T. */
export function refreshRotateMetres(book: SheetBook, area: { row: number; col: number; row2: number; col2: number }): void {
  const seen = new Set<number>();
  const once = (r: number) => {
    if (r < Y_DATA_FIRST || r > Y_LAST || seen.has(r)) return;
    seen.add(r);
    writeRotateMetres(book, r);
  };
  for (let r = area.row; r <= area.row2; r += 1) {
    for (let c = area.col; c <= area.col2; c += 1) {
      if (c === COL.T) once(r);
      if (c === COL.C) once(r - 1);
      if (c === COL.D) {
        once(r - 1);
        once(r);
      }
    }
  }
}

/**
 * EnsureProjAimFormulasOnSheet + EnsureLastSurveyAzmLookup: formula fixes the comment pass
 * needs before the sheet is recalculated. Returns true when anything was written.
 */
export function prepareSlideComments(book: SheetBook): boolean {
  const before = book.writes.length;
  const uHook = `$U$${Y_DATA_FIRST}:$U$${Y_LAST}`;
  for (const col of [COL.AR, COL.AS, COL.W, COL.X, COL.AJ]) {
    for (let r = Y_DATA_FIRST; r <= Y_LAST; r += 1) {
      const original = book.formula(SS_SHEET, r, col);
      if (!original || !original.startsWith("=")) continue;
      let f = original;
      if (col === COL.W || col === COL.X) f = f.replace(`IF(AK${r}="",0,AK${r})`, `AK${r}`);
      else if (col === COL.AR || col === COL.AS) f = f.replace(`(D${r},F${r},G${r},`, `(D${r},W${r},X${r},`);
      if (!f.includes(uHook)) {
        if (col === COL.W || col === COL.X) f = f.replace(`AI${r})`, `AI${r},${uHook})`);
        else if (col === COL.AJ) f = f.replace(`$S$${Y_DATA_FIRST}:$S$${Y_LAST})`, `$S$${Y_DATA_FIRST}:$S$${Y_LAST},${uHook})`);
      }
      if (f !== original) book.setFormula(SS_SHEET, r, col, f);
    }
  }
  syncLastSurveyStrip(book);
  return book.writes.length > before;
}

function restFill(col: number): string {
  switch (col) {
    case 8: case 9: case 16: case 18: case 19: return rgb(242, 242, 242);
    case 20: case 21: return rgb(204, 255, 204);
    default: return COMMENT_YELLOW;
  }
}

/**
 * HighlightNewestSurveyOnSheet: H:Z of the newest good survey in the survey green. A caller that
 * has just written surveys names the row, because their survey depths are not calculated yet.
 */
export function highlightNewestSurvey(book: SheetBook, newest = lastSurveyRow(book)): void {
  for (let r = Y_DATA_FIRST; r <= Y_LAST; r += 1) {
    if (r === newest || book.fill(SS_SHEET, r, COL.J) !== NEWEST_FILL) continue;
    for (let c = 8; c <= 26; c += 1) book.setFill(SS_SHEET, { row: r, col: c, row2: r, col2: c }, restFill(c));
    book.setFontColor(SS_SHEET, { row: r, col: 8, row2: r, col2: 26 }, "#000000");
  }
  if (newest >= Y_DATA_FIRST && book.fill(SS_SHEET, newest, COL.J) !== NEWEST_FILL) {
    book.setFill(SS_SHEET, { row: newest, col: 8, row2: newest, col2: 26 }, NEWEST_FILL);
    book.setFontColor(SS_SHEET, { row: newest, col: 8, row2: newest, col2: 26 }, "#000000");
  }
}

/**
 * Where the green bar belongs now and whether the sheet shows it elsewhere. Reads only: the probe
 * runs on a second book over the same snapshot, so the caller's book is left alone. The bar follows
 * the last row whose Inc and Azm are exactly the survey green, whatever the reason the colours
 * changed (a pull, a typed row, a Fill Color change, Undo).
 */
export function surveyBarState(book: SheetBook): { newest: number; stale: boolean } {
  const newest = lastSurveyRow(book);
  const probe = new SheetBook(book.book);
  highlightNewestSurvey(probe, newest);
  return { newest, stale: probe.writes.length > 0 };
}

/**
 * RefreshSlideComments on a recalculated sheet. forceAll overwrites every data row (clear /
 * first convert); otherwise only blank cells and earlier generated text are rewritten.
 */
export function refreshSlideComments(book: SheetBook, forceAll: boolean, measure: MeasureText = roughMeasure): void {
  if (cellText(book.value(SS_SHEET, Y_FIRST, COL.Y)) === "") book.setValue(SS_SHEET, Y_FIRST, COL.Y, Y_LABEL);
  const style = book.style(SS_SHEET, Y_DATA_FIRST, COL.Y);
  const font = { family: style?.ff ?? "Calibri", sizePt: style?.fs && style.fs > 0 ? style.fs : 11, bold: style?.bl === 1 };
  const cache = new Map<string, number>();
  const measureText = (text: string) => {
    if (text === "") return 0;
    let width = cache.get(text);
    if (width == null) {
      width = measure(text, font);
      cache.set(text, width);
    }
    return width;
  };
  const spacePx = Math.max(measureText(" "), 0.25 * (96 / 72));
  const widthPx = book.columnWidth(SS_SHEET, COL.Y);
  for (let r = Y_DATA_FIRST; r <= Y_LAST; r += 1) {
    let auto = "";
    const arValue = book.value(SS_SHEET, r, COL.AR);
    if (isGoodSurveyRow(book, r) && isNumber(arValue)) {
      const asValue = book.value(SS_SHEET, r, COL.AS);
      const cValue = book.value(SS_SHEET, r, COL.C);
      auto = slideComment(isNumber(asValue) ? asValue : 0, book.value(SS_SHEET, r, COL.AT), arValue, isNumber(cValue) ? cValue : 0);
      const p = auto.indexOf("\u0001");
      if (p >= 0) auto = fitLeftRight(auto.slice(0, p), auto.slice(p + 1), widthPx, measureText, spacePx);
    }
    writeRotateMetres(book, r);
    const current = cellText(book.value(SS_SHEET, r, COL.Y));
    const yFormula = book.hasFormula(SS_SHEET, r, COL.Y);
    const writeIt = forceAll || isAutoSlideCommentText(current) || yFormula;
    if (!writeIt || (current === auto && !yFormula)) continue;
    if (auto === "") {
      if (current !== "" || yFormula) book.clear(SS_SHEET, { row: r, col: COL.Y, row2: r, col2: COL.Y });
    } else book.setValue(SS_SHEET, r, COL.Y, auto);
  }
  book.setFill(SS_SHEET, { row: Y_DATA_FIRST, col: COL.Y, row2: Y_LAST, col2: COL.Z }, COMMENT_YELLOW);
  syncPlanTargetWindow(book);
  highlightActiveTargetOnSheet(book);
  highlightNewestSurvey(book);
}
