import { bankersRound, cellText, isNumber, numeric, SheetBook } from "./sheetBook";
import { lastSurveyRow } from "./slidesheet";

// Port of MDL_DownlinkTracker + frmDLEntry: the DL Tracker log, its ADD ENTRY form and CLEAR.

export const DLT_SHEET = "DL Tracker";
const DLT_SHEET_OLD = "Downlink Tracker V4.6";
export const DLT_FIRST = 3;
export const DLT_LAST = 674;
const TBL_FIRST = 3;
const TBL_LAST = 66;

export const DLC = {
  RUN: 1,
  NUM: 2,
  PAGE: 3,
  CMD: 4,
  PERIOD: 5,
  TYPE: 6,
  BOTTOM: 7,
  SET: 8,
  I: 9,
  J: 10,
  DEPTH: 11,
  DATE: 12,
  TIME: 13,
  CODE: 14,
  MON: 15,
  PASS: 16,
  REF: 17,
  DROP: 18,
  LOW: 19,
  RPM: 20,
  NOTE: 21,
} as const;

export type IJResult = { kind: "keep" } | { kind: "blank" } | { kind: "set"; i: number; j: number };

export interface DlCommand {
  key: string;
  setting: string;
}

export function isTrackerSheetName(name: string): boolean {
  return name === DLT_SHEET || name === DLT_SHEET_OLD;
}

/** The tracker tab by name, falling back to the pre-rename tab. */
export function trackerSheet(book: SheetBook): string | null {
  if (book.hasSheet(DLT_SHEET)) return DLT_SHEET;
  if (book.hasSheet(DLT_SHEET_OLD)) return DLT_SHEET_OLD;
  return null;
}

/** Last log row with a command in D; DLT_FIRST - 1 when the log is empty. */
export function lastEntryRow(book: SheetBook, sheet: string): number {
  for (let r = DLT_LAST; r >= DLT_FIRST; r -= 1) {
    if (cellText(book.value(sheet, r, DLC.CMD)) !== "") return r;
  }
  return DLT_FIRST - 1;
}

/** "0", "3", "S" ... (C holds numbers for numbered pages, text for S). */
export function pageKey(value: unknown): string {
  const text = cellText(value).toUpperCase();
  const n = numeric(text);
  return text !== "" && n != null ? String(bankersRound(n)) : text;
}

function pageTableCol(key: string): number {
  switch (key) {
    case "0": return 27;
    case "1": return 29;
    case "2": return 31;
    case "3": return 33;
    case "7": return 35;
    case "S": return 37;
    default: return 0;
  }
}

/** D/L Setting text for a page + command, as the tab's H column shows it. */
export function lookupSetting(book: SheetBook, sheet: string, key: string, cmd: string): string {
  const c = pageTableCol(key);
  const wanted = cmd.trim().toLowerCase();
  if (c === 0 || wanted === "") return "";
  for (let r = TBL_FIRST; r <= TBL_LAST; r += 1) {
    if (cellText(book.value(sheet, r, c)).toLowerCase() === wanted) return cellText(book.value(sheet, r, c + 1));
  }
  return "";
}

/** Supported commands for a page (key, setting); settings that say ERROR are left out. */
export function commandList(book: SheetBook, sheet: string, key: string): DlCommand[] {
  const c = pageTableCol(key);
  if (c === 0) return [];
  const out: DlCommand[] = [];
  for (let r = TBL_FIRST; r <= TBL_LAST; r += 1) {
    const k = cellText(book.value(sheet, r, c));
    const s = cellText(book.value(sheet, r, c + 1));
    if (k && s && !/error/i.test(s)) out.push({ key: k, setting: s });
  }
  return out;
}

/** Two success codes per command, 64 apart: 1-x -> 256+x or 320+x; 2-x -> 288+x or 352+x. */
export function successCodes(cmd: string): { lo: number; hi: number } | null {
  const s = cmd.trim();
  const p = s.indexOf("-");
  if (p < 1) return null;
  const a = numeric(s.slice(0, p));
  const b = numeric(s.slice(p + 1));
  if (a == null || b == null) return null;
  const g = bankersRound(a);
  const x = bankersRound(b);
  if (g < 1 || g > 2 || x < 0 || x > 31) return null;
  const lo = 256 + (g - 1) * 32 + x;
  return { lo, hi: lo + 64 };
}

/** Bit depth (Slidesheet D) on the deepest good survey row. */
export function lastSurveyBitDepth(book: SheetBook): number | null {
  const r = lastSurveyRow(book);
  if (r <= 0) return null;
  const v = book.value("Slidesheet", r, 4);
  return isNumber(v) ? v : null;
}

/** Nearest row above on the same DL page with numeric I and J. */
export function previousIJ(book: SheetBook, sheet: string, beforeRow: number, key: string): { i: number; j: number } | null {
  for (let r = beforeRow - 1; r >= DLT_FIRST; r -= 1) {
    if (pageKey(book.value(sheet, r, DLC.PAGE)) !== key) continue;
    const i = book.value(sheet, r, DLC.I);
    const j = book.value(sheet, r, DLC.J);
    if (isNumber(i) && isNumber(j)) return { i, j };
  }
  return null;
}

function numAfter(s: string, keyText: string): number | null {
  const p = s.toLowerCase().indexOf(keyText.toLowerCase());
  if (p < 0) return null;
  let i = p + keyText.length;
  while (i < s.length) {
    const ch = s[i]!;
    if (/[0-9.]/.test(ch)) break;
    if (ch !== " ") return null;
    i += 1;
  }
  let num = "";
  while (i < s.length && /[0-9.]/.test(s[i]!)) {
    num += s[i];
    i += 1;
  }
  if (num === "" || numeric(num) == null) return null;
  return Number.parseFloat(num);
}

function wrap360(a: number): number {
  return bankersRound(a - 360 * Math.floor(a / 360), 2);
}

/** TF/INC (I) and SR%/AZI (J) from the D/L Setting. */
export function computeIJ(setting: string, prev: { i: number; j: number } | null): IJResult {
  const s = setting.trim();
  if (s === "") return { kind: "keep" };
  if (/error/i.test(s)) return { kind: "blank" };
  if (/TF =/i.test(s) && /SR =/i.test(s)) {
    const i = numAfter(s, "TF =");
    const j = numAfter(s, "SR =");
    return i != null && j != null ? { kind: "set", i, j } : { kind: "keep" };
  }
  if (s.startsWith("Go to Page") || s.startsWith("Go to Manual")) return { kind: "blank" };
  if (!prev) return { kind: "keep" };
  let i = prev.i;
  let j = prev.j;
  const dirn = /decrease/i.test(s) || /nudge down/i.test(s) ? -1 : 1;
  const incBy = numAfter(s, "Target INC by");
  const incTarget = incBy == null ? numAfter(s, "Inclination Target") : null;
  const azmBy = incBy == null && incTarget == null ? numAfter(s, "Target AZI by") : null;
  const srBy = incBy == null && incTarget == null && azmBy == null ? numAfter(s, "SR by") : null;
  const tfBy = incBy == null && incTarget == null && azmBy == null && srBy == null ? numAfter(s, "TF by") : null;
  if (incBy != null) i = bankersRound(prev.i + dirn * incBy, 2);
  else if (incTarget != null) i = bankersRound(prev.i + dirn * incTarget, 2);
  else if (azmBy != null) j = wrap360(prev.j + dirn * azmBy);
  else if (srBy != null) j = Math.min(100, Math.max(0, prev.j + dirn * srBy));
  else if (tfBy != null) i = wrap360(prev.i + dirn * tfBy);
  return { kind: "set", i, j };
}

/** DLT_RefreshIJ: recompute I/J on one log row from its page, command and Pass/Fail. */
export function refreshIJ(book: SheetBook, sheet: string, r: number): void {
  const cmd = cellText(book.value(sheet, r, DLC.CMD));
  if (cmd === "") return;
  const key = pageKey(book.value(sheet, r, DLC.PAGE));
  const prev = previousIJ(book, sheet, r, key);
  if (cellText(book.value(sheet, r, DLC.PASS)).toUpperCase() === "FAIL") {
    if (prev) writeIJ(book, sheet, r, prev.i, prev.j);
    return;
  }
  const result = computeIJ(lookupSetting(book, sheet, key, cmd), prev);
  switch (result.kind) {
    case "set":
      writeIJ(book, sheet, r, result.i, result.j);
      break;
    case "blank":
      book.clear(sheet, { row: r, col: DLC.I, row2: r, col2: DLC.J });
      break;
    case "keep":
      break;
    default: {
      const never: never = result;
      throw new Error(`Unhandled I/J result ${String(never)}`);
    }
  }
}

function writeIJ(book: SheetBook, sheet: string, r: number, i: number, j: number): void {
  book.setValue(sheet, r, DLC.I, i);
  book.setValue(sheet, r, DLC.J, j);
}

/** Workbook_SheetChange hook: C, D or P typed on the log rows recompute that row's I/J. */
export function onTrackerEdit(book: SheetBook, sheet: string, area: { row: number; col: number; row2: number; col2: number }): void {
  const rows: number[] = [];
  let cells = 0;
  for (let r = Math.max(area.row, DLT_FIRST); r <= Math.min(area.row2, DLT_LAST); r += 1) {
    let hit = false;
    for (const c of [DLC.PAGE, DLC.CMD, DLC.PASS]) {
      if (c >= area.col && c <= area.col2) {
        cells += 1;
        hit = true;
      }
    }
    if (hit) rows.push(r);
  }
  if (cells === 0 || cells > 60) return;
  for (const r of rows) refreshIJ(book, sheet, r);
}

/** DLT_ClearLog: unlocked, non-formula, non-empty cells in A3:U674. Returns the count. */
export function clearLog(book: SheetBook, sheet: string): number {
  let n = 0;
  for (let r = DLT_FIRST; r <= DLT_LAST; r += 1) {
    for (let c = DLC.RUN; c <= DLC.NOTE; c += 1) {
      if (book.locked(sheet, r, c) || book.hasFormula(sheet, r, c)) continue;
      const v = book.value(sheet, r, c);
      if (v == null || v === "") continue;
      book.clear(sheet, { row: r, col: c, row2: r, col2: c });
      n += 1;
    }
  }
  return n;
}

function counterFormula(r: number): string {
  return r === DLT_FIRST ? `=IF(D${r}="","",1)` : `=IF(D${r}="","",B${r - 1}+1)`;
}

/** DLT_WriteEntry: columns 1-21 except B and H; locked and formula cells are skipped; null clears. */
export function writeEntry(book: SheetBook, sheet: string, r: number, vals: Partial<Record<number, string | number | null>>): void {
  if (!book.hasFormula(sheet, r, DLC.NUM)) book.setFormula(sheet, r, DLC.NUM, counterFormula(r));
  for (let c: number = DLC.RUN; c <= DLC.NOTE; c += 1) {
    if (c === DLC.NUM || c === DLC.SET) continue;
    if (book.locked(sheet, r, c) || book.hasFormula(sheet, r, c)) continue;
    const v = vals[c] ?? null;
    if (v == null || v === "") {
      const current = book.value(sheet, r, c);
      if (current != null && current !== "") book.clear(sheet, { row: r, col: c, row2: r, col2: c });
    } else book.setValue(sheet, r, c, v);
  }
}

// ---------------------------------------------------------------- frmDLEntry

export interface DlEntryForm {
  run: string;
  page: string;
  cmd: string;
  period: string;
  type: string;
  bottom: string;
  i: string;
  j: string;
  depth: string;
  date: string;
  time: string;
  code: string;
  mon: string;
  pass: string;
  ref: string;
  drop: string;
  low: string;
  rpm: string;
  note: string;
}

export interface DlEntryLists {
  pages: string[];
  periods: string[];
  types: string[];
  bottoms: string[];
}

export interface DlEntryStart {
  sheet: string;
  row: number;
  source: number;
  form: DlEntryForm;
  lists: DlEntryLists;
  /** The higher DLNK_PD is the default; a source row that used the lower one keeps it. */
  high: boolean;
  status: string;
}

export function numText(v: unknown): string {
  const n = v == null || v === "" ? null : numeric(v);
  return n == null ? "" : String(bankersRound(n, 2));
}

function pctText(v: unknown): string {
  const n = v == null || v === "" ? null : numeric(v);
  return n == null ? "" : String(bankersRound(n * 100, 1));
}

export function depthText(v: unknown): string {
  const n = v == null || v === "" ? null : numeric(v);
  return n == null ? "" : String(Math.floor(n + 0.5));
}

/** "Go to Page 3: HIA Mode" -> "3". */
function goToPage(setting: string): string {
  if (!setting.startsWith("Go to Page ")) return "";
  const ch = setting.charAt(11);
  return /[0-9]/.test(ch) ? ch : "";
}

function pad(n: number): string {
  return String(n).padStart(2, "0");
}

export function dateText(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export function timeText(d: Date): string {
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Now, rounded to the minute. */
export function entryNow(now = new Date()): Date {
  const out = new Date(now.getTime());
  out.setSeconds(out.getSeconds() + 30, 0);
  out.setSeconds(0, 0);
  return out;
}

export type DlAddStart = { ok: true; start: DlEntryStart } | { ok: false; message: string };

/** DLT_AddEntry + frmDLEntry.LoadDefaults. */
export function startEntry(book: SheetBook, now = new Date()): DlAddStart {
  const sheet = trackerSheet(book);
  if (!sheet) return { ok: false, message: `The '${DLT_SHEET}' tab is not in this workbook.` };
  const last = lastEntryRow(book, sheet);
  if (last >= DLT_LAST) return { ok: false, message: `The downlink log is full (row ${DLT_LAST}).` };
  const row = last + 1;
  const source = last;
  const src = source < DLT_FIRST ? row : source;
  const text = (c: number) => cellText(book.value(sheet, src, c));
  const lists: DlEntryLists = {
    pages: book.validationList(sheet, row, DLC.PAGE),
    periods: book.validationList(sheet, row, DLC.PERIOD),
    types: book.validationList(sheet, row, DLC.TYPE),
    bottoms: book.validationList(sheet, row, DLC.BOTTOM),
  };
  let key = pageKey(book.value(sheet, src, DLC.PAGE));
  let cmd = text(DLC.CMD);
  if (source >= DLT_FIRST && text(DLC.PASS).toUpperCase() !== "FAIL") {
    const next = goToPage(lookupSetting(book, sheet, key, cmd));
    if (next && next !== key) {
      key = next;
      cmd = "";
    }
  }
  const surveyDepth = lastSurveyBitDepth(book);
  const at = entryNow(now);
  let high = true;
  const code = book.value(sheet, src, DLC.CODE);
  if (source >= DLT_FIRST && code != null && code !== "" && numeric(code) != null) {
    const codes = successCodes(text(DLC.CMD));
    if (codes && bankersRound(numeric(code)!) === codes.lo) high = false;
  }
  const form: DlEntryForm = {
    run: text(DLC.RUN),
    page: key,
    cmd,
    period: text(DLC.PERIOD),
    type: text(DLC.TYPE),
    bottom: text(DLC.BOTTOM),
    i: numText(book.value(sheet, src, DLC.I)),
    j: numText(book.value(sheet, src, DLC.J)),
    depth: surveyDepth != null ? depthText(surveyDepth) : depthText(book.value(sheet, src, DLC.DEPTH)),
    date: dateText(at),
    time: timeText(at),
    code: "",
    mon: pctText(book.value(sheet, src, DLC.MON)),
    pass: "PASS",
    ref: numText(book.value(sheet, src, DLC.REF)),
    drop: pctText(book.value(sheet, src, DLC.DROP)),
    low: numText(book.value(sheet, src, DLC.LOW)),
    rpm: numText(book.value(sheet, src, DLC.RPM)),
    note: text(DLC.NOTE),
  };
  const status = source < DLT_FIRST ? "First entry - fill in the command and settings." : `Copied from row ${source}. Change what differs, then Save.`;
  const start: DlEntryStart = { sheet, row, source, form, lists, high, status };
  return { ok: true, start: { ...start, form: refreshCommand(book, start, form) } };
}

export interface SettingLabel {
  text: string;
  tone: "hint" | "error" | "normal";
}

export function settingLabel(book: SheetBook, start: DlEntryStart, form: DlEntryForm): SettingLabel {
  const key = pageKey(form.page);
  const setting = lookupSetting(book, start.sheet, key, form.cmd.trim());
  if (setting === "") return { text: ` (pick a command for page ${key})`, tone: "hint" };
  return { text: ` ${setting}`, tone: /error/i.test(setting) ? "error" : "normal" };
}

/** RefreshIJ on the form. */
export function refreshFormIJ(book: SheetBook, start: DlEntryStart, form: DlEntryForm): DlEntryForm {
  const key = pageKey(form.page);
  const prev = previousIJ(book, start.sheet, start.row, key);
  if (form.pass.toUpperCase() === "FAIL") {
    return prev ? { ...form, i: numText(prev.i), j: numText(prev.j) } : form;
  }
  const result = computeIJ(lookupSetting(book, start.sheet, key, form.cmd.trim()), prev);
  switch (result.kind) {
    case "set": return { ...form, i: numText(result.i), j: numText(result.j) };
    case "blank": return { ...form, i: "", j: "" };
    case "keep": return form;
    default: {
      const never: never = result;
      throw new Error(`Unhandled I/J result ${String(never)}`);
    }
  }
}

/** DLNK_PD choices (hi first) and the value the form shows. */
export function codeChoices(form: DlEntryForm, high: boolean): { choices: string[]; code: string } {
  if (form.pass.toUpperCase() === "FAIL") return { choices: [], code: "" };
  const codes = successCodes(form.cmd);
  if (!codes) return { choices: [], code: "" };
  return { choices: [String(codes.hi), String(codes.lo)], code: String(high ? codes.hi : codes.lo) };
}

/** RefreshCommand: I/J and DLNK_PD follow the page, command and Pass/Fail. */
export function refreshCommand(book: SheetBook, start: DlEntryStart, form: DlEntryForm): DlEntryForm {
  const withIJ = refreshFormIJ(book, start, form);
  return { ...withIJ, code: codeChoices(withIJ, start.high).code };
}

function parseDate(text: string): { y: number; m: number; d: number } | null {
  const t = text.trim();
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (match) return { y: Number(match[1]), m: Number(match[2]), d: Number(match[3]) };
  match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (match) return { y: Number(match[3]), m: Number(match[1]), d: Number(match[2]) };
  return null;
}

function parseTime(text: string): { h: number; n: number } | null {
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(text.trim());
  if (!match) return null;
  const h = Number(match[1]);
  const n = Number(match[2]);
  return h < 24 && n < 60 ? { h, n } : null;
}

/** Date + time text as one local Date, or null when either is not valid. */
export function parseDateTime(date: string, time: string): Date | null {
  const d = parseDate(date);
  const t = parseTime(time);
  if (!d || !t) return null;
  const out = new Date(d.y, d.m - 1, d.d, t.h, t.n, 0, 0);
  if (out.getFullYear() !== d.y || out.getMonth() !== d.m - 1 || out.getDate() !== d.d) return null;
  return out;
}

/** Excel serial date for a local Date (1900 system). */
export function excelSerial(at: Date): number {
  const utc = Date.UTC(at.getFullYear(), at.getMonth(), at.getDate(), at.getHours(), at.getMinutes(), at.getSeconds());
  return utc / 86_400_000 + 25_569;
}

function numOk(text: string): boolean {
  const s = text.trim();
  return s === "" || numeric(s) != null;
}

function numOrEmpty(text: string): number | null {
  const n = numeric(text.trim());
  return text.trim() === "" || n == null ? null : n;
}

function pctOrEmpty(text: string): number | null {
  const n = numOrEmpty(text);
  return n == null ? null : n / 100;
}

function numOrText(text: string): string | number | null {
  const s = text.trim();
  if (s === "") return null;
  return numeric(s) ?? s;
}

function textOrEmpty(text: string): string | null {
  const s = text.trim();
  return s === "" ? null : s;
}

export type DlSaveCheck =
  | { ok: true; vals: Partial<Record<number, string | number | null>>; confirm: string | null }
  | { ok: false; error: string };

/** btnSave_Click validation and the values it writes. `confirm` is the "Save anyway?" question. */
export function checkEntry(form: DlEntryForm): DlSaveCheck {
  if (form.cmd.trim() === "") return { ok: false, error: "Pick a Dnlk command." };
  const at = parseDateTime(form.date, form.time);
  if (!at) return { ok: false, error: "Date or time is not valid (yyyy-mm-dd, hh:mm)." };
  if (![form.i, form.j, form.depth, form.mon, form.ref, form.drop, form.low, form.rpm, form.code].every(numOk)) {
    return { ok: false, error: "Numbers only in TF/INC, SR/AZI, depth, DLNK_PD, %, flow and RPM." };
  }
  const failed = form.pass.toUpperCase() === "FAIL";
  let confirm: string | null = null;
  if (!failed && form.code.trim() !== "") {
    const codes = successCodes(form.cmd);
    const code = bankersRound(numeric(form.code.trim()) ?? 0);
    if (codes && code !== codes.lo && code !== codes.hi) {
      confirm = `DLNK_PD ${form.code} is neither ${codes.hi} nor ${codes.lo} for command ${form.cmd}.\nSave anyway?`;
    }
  }
  const serial = excelSerial(at);
  const day = Math.floor(serial);
  const vals: Partial<Record<number, string | number | null>> = {
    [DLC.RUN]: numOrText(form.run),
    [DLC.PAGE]: numOrText(form.page),
    [DLC.CMD]: form.cmd.trim(),
    [DLC.PERIOD]: numOrText(form.period),
    [DLC.TYPE]: textOrEmpty(form.type),
    [DLC.BOTTOM]: textOrEmpty(form.bottom),
    [DLC.I]: numOrEmpty(form.i),
    [DLC.J]: numOrEmpty(form.j),
    [DLC.DEPTH]: numOrEmpty(form.depth),
    [DLC.DATE]: day,
    [DLC.TIME]: serial - day,
    [DLC.CODE]: failed ? null : numOrEmpty(form.code),
    [DLC.MON]: pctOrEmpty(form.mon),
    [DLC.PASS]: form.pass.toUpperCase(),
    [DLC.REF]: numOrEmpty(form.ref),
    [DLC.DROP]: pctOrEmpty(form.drop),
    [DLC.LOW]: numOrEmpty(form.low),
    [DLC.RPM]: numOrEmpty(form.rpm),
    [DLC.NOTE]: textOrEmpty(form.note),
  };
  return { ok: true, vals, confirm };
}
