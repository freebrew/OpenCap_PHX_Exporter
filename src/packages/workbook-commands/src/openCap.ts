import { OcTable, vbaCdbl, vbaIsNumeric, vbaStr } from "./ocTable";
import { bankersRound, cellText, rgb, type CellArea, type CellScalar, type SheetBook } from "./sheetBook";
import { toolHoursSync } from "./toolHours";

// Port of MDL_Setup's OpenCap refresh (RefreshAllCsvData / AdoptCurrentJobFromOpenCap): the six
// FieldCap CSVs load into the hidden _OC_* sheets, then Setup's job, crew and file panels, the
// Costs sheet and Data's Mud Motors are filled from them. Setup is refreshed in place rather
// than rebuilt (BuildSetupUI), so the AC table, notes and hand-tuned layout stay as they are.

export type OpenCapKind = "job" | "crew" | "bha" | "slide" | "inventory" | "costs";

export interface OpenCapFile {
  kind: OpenCapKind;
  token: string;
  sheet: string;
  label: string;
}

export const OPEN_CAP_FILES: readonly OpenCapFile[] = [
  { kind: "job", token: "job-details", sheet: "_OC_Job", label: "Job Details" },
  { kind: "crew", token: "crew", sheet: "_OC_Crew", label: "Crew / Personnel" },
  { kind: "bha", token: "bha-equipment", sheet: "_OC_BHA", label: "BHA Equipment" },
  { kind: "slide", token: "slide-rotate", sheet: "_OC_Slide", label: "Slide / Rotate by Day" },
  { kind: "inventory", token: "inventory", sheet: "_OC_Inventory", label: "Equipment Inventory" },
  { kind: "costs", token: "ticket-costs", sheet: "_OC_Costs", label: "Ticket Costs by Day" },
];

export interface OpenCapSource {
  kind: OpenCapKind;
  csv: string;
  /** File name shown in Setup's files panel. */
  name: string;
  /** Full path shown in Setup's files panel. */
  path: string;
}

export interface OpenCapOptions {
  /** OC_LastJobId before this load. */
  lastJobId: string;
  /** AdoptCurrentJobFromOpenCap: always treat the load as a new job. */
  adopt?: boolean;
  /** Path the files panel shows for a missing file (the job folder). */
  folderPath: string;
  /** OC_CrewTop / OC_CrewBot: the crew manifest's first and last data rows. */
  crewTop?: number;
  crewBottom?: number;
}

export interface OpenCapResult {
  /** _OC_Job "Job ID": FieldCap's ClientJobId. This is what OC_LastJobId stores. */
  jobId: string;
  /** _OC_Job "Job Code", falling back to the Job ID: the Job Number a person reads. */
  jobNumber: string;
  /** OC_LastJobId after the load. */
  lastJobId: string;
  newJob: boolean;
  rows: Partial<Record<OpenCapKind, number>>;
  status: string;
  alert: string | null;
}

export type OcValue = CellScalar;

const SETUP = "Setup";
const DATA = "Data";
const COSTS = "Costs";
const MAX_LINES = 5000;
const COSTS_FIRST = 3;
const COSTS_LAST = 72;
const MM_FIRST = 45;
const MM_ROWS = 11;
const FMT_MONEY = "\"$\"#,##0.00";
const FMT_COST_DATE = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
const CHECK = "\u2713";
const CROSS = "\u2717";

// ---------------------------------------------------------------- CSV

/** CsvConv: Trim, blank stays blank, numeric text becomes a Double. Range.Value then makes TRUE / FALSE booleans. */
export function csvConv(text: string): OcValue {
  const t = text.replace(/^ +| +$/g, "");
  if (t === "") return null;
  if (vbaIsNumeric(t)) return vbaCdbl(t);
  if (/^(true|false)$/i.test(t)) return t.toLowerCase() === "true";
  return t;
}

const DATE_FORMAT = "m/d/yyyy";

/** Range.Value = text on a cell that is not formatted as text: Excel types numbers, TRUE / FALSE and dates. */
function enterText(book: SheetBook, sheet: string, row: number, col: number, text: string): void {
  const t = text.trim();
  const pattern = book.style(sheet, row, col)?.n?.pattern ?? "";
  if (t === "" || pattern === "@") {
    book.setValue(sheet, row, col, t === "" ? null : text);
    return;
  }
  if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(t)) {
    book.setValue(sheet, row, col, Number(t));
    return;
  }
  if (/^(true|false)$/i.test(t)) {
    book.setValue(sheet, row, col, t.toLowerCase() === "true");
    return;
  }
  const serial = /^\d{4}-\d{1,2}-\d{1,2}$|^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t) ? costDateSerial(t) : 0;
  if (serial > 0) {
    if (pattern === "" || pattern === "General") book.setFormat(sheet, row, col, DATE_FORMAT);
    book.setValue(sheet, row, col, serial);
    return;
  }
  book.setValue(sheet, row, col, text);
}

/** CsvParseLine: quote-aware, "" is an escaped quote. */
function parseLine(line: string): string[] {
  const out: string[] = [];
  let token = "";
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === "\"") {
      if (quoted && line[i + 1] === "\"") {
        token += "\"";
        i += 1;
      } else quoted = !quoted;
    } else if (ch === "," && !quoted) {
      out.push(token);
      token = "";
    } else token += ch;
  }
  out.push(token);
  return out;
}

/**
 * The file's rows. A line break inside a quoted field belongs to that field (FieldCap's driving
 * directions run over several lines), so only a break outside quotes ends a row. MDL_Setup's
 * LoadCsv reads with Line Input and cuts such a row short; this does not copy that.
 */
function splitRecords(text: string): string[] {
  const records: string[] = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch === "\"") quoted = !quoted;
    else if (!quoted && (ch === "\n" || ch === "\r")) {
      records.push(text.slice(start, i));
      if (ch === "\r" && text[i + 1] === "\n") i += 1;
      start = i + 1;
    }
  }
  records.push(text.slice(start));
  return records;
}

/** LoadCsv phases 1-3: the first 5,000 rows as a block of CsvConv values. */
export function parseOpenCapCsv(text: string): OcValue[][] {
  const lines = splitRecords(text.replace(/^\uFEFF/, ""));
  if (lines.length > 0 && lines[lines.length - 1] === "") lines.pop();
  return lines.slice(0, MAX_LINES).map((line) => parseLine(line).map(csvConv));
}

/** Excel-style CStr of a value Excel would have typed on Range.Value = text (ISO dates become dates). */
function excelShown(value: unknown): string {
  if (typeof value !== "string") return vbaStr(value);
  const match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(value.trim());
  if (!match) return value;
  const [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3])];
  if (month < 1 || month > 12 || day < 1 || day > 31) return value;
  const date = `${month}/${day}/${year}`;
  const hours = Number(match[4] ?? 0);
  const minutes = Number(match[5] ?? 0);
  const seconds = Number(match[6] ?? 0);
  if (hours === 0 && minutes === 0 && seconds === 0) return date;
  const twelve = hours % 12 === 0 ? 12 : hours % 12;
  return `${date} ${twelve}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")} ${hours < 12 ? "AM" : "PM"}`;
}

function excelSerial(year: number, month: number, day: number): number {
  return Date.UTC(year, month - 1, day) / 86_400_000 + 25_569;
}

// ---------------------------------------------------------------- _OC_* sheets

/** LoadCsv: Cells.Clear, then one Range.Value write of the parsed block. */
export function loadOpenCapSheet(book: SheetBook, sheet: string, rows: OcValue[][]): number {
  book.resetSheet(sheet);
  if (rows.length === 0 || rows.every((line) => line.length === 0)) return 0;
  book.setTable(sheet, 1, 1, rows);
  return Math.max(0, book.lastUsedRow(sheet, 1) - 1);
}

// ---------------------------------------------------------------- job fields

export type JobFields = Map<string, string>;

/** ReadJobFields: _OC_Job row 1 header -> row 2 value, case-insensitive, first header wins. */
export function readJobFields(book: SheetBook): JobFields {
  const fields: JobFields = new Map();
  const table = new OcTable(book, "_OC_Job");
  if (!table.exists) return fields;
  for (let col = 1; col <= 512; col += 1) {
    const header = cellText(book.value(table.sheet, 1, col));
    if (header === "") continue;
    const key = header.toLowerCase();
    if (!fields.has(key)) fields.set(key, excelShown(book.value(table.sheet, 2, col)).trim());
  }
  return fields;
}

/** GF: the first key whose value is not blank. */
export function jobField(fields: JobFields, ...keys: string[]): string {
  let last = "";
  for (const key of keys) {
    const value = fields.get(key.toLowerCase());
    if (value === undefined) continue;
    if (value !== "") return value;
    last = value;
  }
  return last;
}

/** StrConv(text, vbProperCase). */
function properCase(text: string): string {
  return text.toLowerCase().replace(/(^|\s)(\S)/g, (_, lead: string, letter: string) => lead + letter.toUpperCase());
}

/** ClientNameFromEmails: the first non-PHX address's domain, without its last label. */
function clientNameFromEmails(raw: string): string {
  if (raw.trim() === "") return "";
  for (const part of raw.replaceAll(",", ";").replaceAll(" ", "").split(";")) {
    const email = part.trim().toLowerCase();
    const at = email.indexOf("@");
    if (email === "" || at < 1 || email.includes("phxtech.com")) continue;
    let domain = email.slice(at + 1);
    const dot = domain.lastIndexOf(".");
    if (dot > 0) domain = domain.slice(0, dot);
    return properCase(domain.replaceAll("-", " ").replaceAll(".", " "));
  }
  return "";
}

export function jobClientName(fields: JobFields): string {
  return jobField(fields, "Client", "Company", "Operator") || jobField(fields, "Client Name", "ClientName") || clientNameFromEmails(jobField(fields, "ClientEmail"));
}

function fmtDate(text: string): string {
  const t = text.trim();
  if (t.length >= 8 && /^\d{8}/.test(t)) return `${t.slice(0, 4)}-${t.slice(4, 6)}-${t.slice(6, 8)}`;
  return t;
}

function valUnit(text: string, unit: string): string {
  const t = text.trim();
  return t === "" ? "" : t + unit;
}

function looksLikeEmail(text: string): boolean {
  const t = text.trim();
  return t.length >= 5 && t.indexOf("@") > 0 && t.includes(".") && !t.includes(" ");
}

type FieldPick = (fields: JobFields) => string;
const pick = (...keys: string[]): FieldPick => (fields) => jobField(fields, ...keys);

/** BuildSetupUI Pair rows: label in B (value D:E) and label in F (value G:H). */
const LEFT_PAIRS: Record<string, FieldPick> = {
  "JOB ID": pick("Job Code", "Job ID"),
  "JOB NAME": pick("Job Name"),
  CLIENT: jobClientName,
  "WELL / UWI": pick("UWI", "Well (core)"),
  "SURFACE COORDS": pick("SurfaceCoordinates"),
  "SPUD DATE": (fields) => fmtDate(jobField(fields, "SpudDate")),
  "MAP SYSTEM": pick("MapSystem"),
  "GROUND MSL": (fields) => valUnit(jobField(fields, "GroundMSL"), " m"),
  "RKB HEIGHT": (fields) => valUnit(jobField(fields, "RKB"), " m"),
  LATITUDE: pick("LatitudeDMS", "Latitude"),
  VSD: pick("VSD"),
  "MAG DECL": (fields) => valUnit(jobField(fields, "MagneticDeclination"), "\u00B0"),
  DIP: (fields) => valUnit(jobField(fields, "DIP"), "\u00B0"),
  TGF: pick("TGF"),
  CONVERGENCE: (fields) => valUnit(jobField(fields, "Convergence"), "\u00B0"),
  "SURVEY CORR.": pick("SurveyCorrection"),
  "RIG TYPE": pick("RigType"),
  "DATA RECORDER": pick("DataRecorder"),
  "MWD GUIDANCE": pick("MWDGuidanceType"),
  "RIG NAME": pick("Rig Name", "RigName"),
  "COMPANY MAN": pick("CompanyMan"),
  "2ND COMP. MAN": pick("SecondCompanyMan"),
  GEOLOGIST: pick("Geologist"),
  "DD COORDINATOR": pick("DDCoordinator"),
  "MWD COORDINATOR": pick("MWDCoordinator"),
  "SALES REP": pick("SalesRep"),
};

const RIGHT_PAIRS: Record<string, FieldPick> = {
  "OPS STATUS": pick("opsStatus", "Ops Status"),
  PROFILE: pick("WellProfile", "Job Type"),
  AFE: pick("AFE", "AFE (core)"),
  "WELL LICENCE": pick("WellLicenceNumber"),
  PROVINCE: pick("Province"),
  "START DATE": pick("Start Date", "Planned Start Date"),
  "NORTH REF": pick("NorthReference"),
  TVD: (fields) => valUnit(jobField(fields, "TVD"), " m"),
  "TOTAL KB": (fields) => valUnit(jobField(fields, "TotalKB"), " m"),
  LONGITUDE: pick("LongitudeDMS", "Longitude"),
  "WELL TYPE": pick("WellProfile", "Job Type"),
  "REF MODEL": pick("MagneticReferenceModel"),
  "REF DATE": (fields) => fmtDate(jobField(fields, "MagneticReferenceDate")),
  TMF: pick("TMF"),
  "TOL DIP": pick("ToleranceDIP"),
  "DRILL MEAS.": pick("DrillMeasuredFrom"),
  "TOP DRIVE": pick("TopDrive"),
  GAMMA: pick("Gamma"),
  "PIPE ARM": pick("PipeArm"),
  LOADER: pick("Loader"),
  "CM PHONE": pick("CompanyManPhone"),
  "2ND CM PHONE": pick("SecondCompanyManPhone"),
  "GEOL. PHONE": pick("GeologistPhone"),
  "DD COORD EMAIL": pick("DDCoordinatorEmail"),
  "MWD EMAIL": pick("MWDCoordinatorEmail"),
  "SALES EMAIL": pick("SalesRepEmail"),
};

function writePairValue(book: SheetBook, row: number, col: number, value: string, emailLabel: boolean): void {
  enterText(book, SETUP, row, col, value);
  const email = emailLabel || looksLikeEmail(value);
  book.setLook(SETUP, { row, col, row2: row, col2: col }, email ? { bold: false, fontSize: 8 } : { bold: value !== "" });
}

/** The job pairs, found by their labels in B and F so a moved row still fills. */
export function refreshJobPanel(book: SheetBook, fields: JobFields): number {
  if (!book.hasSheet(SETUP)) return 0;
  let filled = 0;
  for (let row = 3; row <= 60; row += 1) {
    const left = LEFT_PAIRS[cellText(book.value(SETUP, row, 2)).toUpperCase()];
    if (left) {
      writePairValue(book, row, 4, left(fields), false);
      filled += 1;
    }
    const rightLabel = cellText(book.value(SETUP, row, 6)).toUpperCase();
    const right = RIGHT_PAIRS[rightLabel];
    if (right) {
      writePairValue(book, row, 7, right(fields), rightLabel.includes("EMAIL"));
      filled += 1;
    }
  }
  return filled;
}

// ---------------------------------------------------------------- crew manifest

export interface CrewRow {
  role: string;
  name: string;
  email: string;
  phone: string;
}

/** ReadCrewRows: up to 201 rows with a name; Role, else Work Type. */
export function readCrewRows(book: SheetBook): CrewRow[] {
  const table = new OcTable(book, "_OC_Crew");
  if (!table.exists) return [];
  const last = table.lastRow(1);
  const cName = table.col("Name", "FullName") || 2;
  const cRole = table.col("Role");
  const cWork = table.col("Work Type", "WorkType", "Position");
  const cEmail = table.col("Email");
  const cPhone = table.col("Phone", "Mobile", "MobilePhone");
  const crew: CrewRow[] = [];
  for (let row = 2; row <= last && crew.length <= 200; row += 1) {
    const name = table.str(row, cName).trim();
    if (name === "") continue;
    let role = table.str(row, cRole).trim();
    if (role === "" && cWork > 0) role = table.str(row, cWork).trim();
    crew.push({ role, name, email: table.str(row, cEmail).trim(), phone: table.str(row, cPhone).trim() });
  }
  return crew;
}

/** RolePriority: DD first, then MWD, then anyone else, and alphabetical inside a role. */
function rolePriority(role: string): number {
  const head = role.trim().toUpperCase().split(/\s+/)[0] ?? "";
  if (head === "DD") return 1;
  if (head === "MWD") return 2;
  return 3;
}

/**
 * SortCrewByRole. The crew manifest (OC_CrewTop:OC_CrewBot) reorders by role, then by name.
 * Only rows that hold a name move; an empty slot stays empty. True when the order changed.
 */
export function sortCrewByRole(book: SheetBook, top: number, bottom: number): boolean {
  if (!book.hasSheet(SETUP) || bottom - top < 1) return false;
  const rows: CrewRow[] = [];
  for (let row = top; row <= bottom; row += 1) {
    const name = cellText(book.value(SETUP, row, 11)).trim();
    if (name === "") continue;
    rows.push({ role: cellText(book.value(SETUP, row, 10)).trim(), name, email: cellText(book.value(SETUP, row, 13)).trim(), phone: cellText(book.value(SETUP, row, 16)).trim() });
  }
  const sorted = [...rows].sort((a, b) => rolePriority(a.role) - rolePriority(b.role) || a.name.toLowerCase().localeCompare(b.name.toLowerCase()));
  const changed = sorted.some((row, index) => row.name !== rows[index]?.name || row.role !== rows[index]?.role);
  if (changed) refreshCrewPanel(book, sorted, top, bottom);
  return changed;
}

/** DrawCrew values: the title count and one row per slot (role J, name K:L, email M:O, phone P:R). */
export function refreshCrewPanel(book: SheetBook, crew: CrewRow[], top = 5, bottom = 10): void {
  if (!book.hasSheet(SETUP)) return;
  const titleRow = top - 2;
  const cap = Math.max(1, bottom - titleRow - 1);
  let title = "  CREW MANIFEST";
  if (crew.length > cap) title += `   ${cap}+${crew.length - cap}`;
  else if (crew.length > 0) title += `   ${crew.length}`;
  book.setValue(SETUP, titleRow, 10, title);
  for (let row = top; row <= bottom; row += 1) {
    const member = crew[row - top];
    enterText(book, SETUP, row, 10, member?.role ?? "");
    enterText(book, SETUP, row, 11, member?.name ?? "");
    book.setLook(SETUP, { row, col: 11, row2: row, col2: 11 }, { bold: Boolean(member?.name) });
    enterText(book, SETUP, row, 13, member?.email ?? "");
    enterText(book, SETUP, row, 16, member?.phone ?? "");
  }
}

// ---------------------------------------------------------------- files panel

function filesFirstRow(book: SheetBook): number {
  for (let row = 1; row <= 80; row += 1) {
    if (cellText(book.value(SETUP, row, 10)).toUpperCase() === "STATUS") return row + 1;
  }
  return 0;
}

function paintFileCell(book: SheetBook, area: CellArea, value: string | null, fill: string, color: string, look: { bold: boolean; fontSize: number; hAlign: "left" | "center" | "right" }): void {
  if (area.col2 > area.col) book.merge(SETUP, area);
  enterText(book, SETUP, area.row, area.col, value ?? "");
  book.setFill(SETUP, area, fill);
  book.setFontColor(SETUP, area, color);
  book.setLook(SETUP, area, look);
}

/** DrawFiles rows: a green tick with the file name, row count and path, or a red cross with the hint. */
export function refreshFilesPanel(book: SheetBook, found: Partial<Record<OpenCapKind, { name: string; path: string }>>, folderPath: string): void {
  if (!book.hasSheet(SETUP)) return;
  const first = filesFirstRow(book);
  if (first === 0) return;
  const white = rgb(255, 255, 255);
  const dark = rgb(89, 89, 89);
  OPEN_CAP_FILES.forEach((file, i) => {
    const row = first + i;
    const cells = (col: number, col2 = col): CellArea => ({ row, col, row2: row, col2 });
    book.unmerge(SETUP, cells(10, 18));
    const hit = found[file.kind];
    if (hit) {
      const green = rgb(206, 239, 198);
      const text = rgb(0, 110, 60);
      const rows = Math.max(0, new OcTable(book, file.sheet).lastRow(1) - 1);
      paintFileCell(book, cells(10), CHECK, rgb(0, 176, 80), white, { bold: true, fontSize: 8, hAlign: "center" });
      paintFileCell(book, cells(11, 12), file.label, green, text, { bold: true, fontSize: 7, hAlign: "left" });
      paintFileCell(book, cells(13, 15), hit.name, green, text, { bold: false, fontSize: 7, hAlign: "left" });
      paintFileCell(book, cells(16), String(rows), green, text, { bold: false, fontSize: 7, hAlign: "right" });
      paintFileCell(book, cells(17, 18), hit.path, green, dark, { bold: false, fontSize: 6, hAlign: "left" });
    } else {
      const red = rgb(255, 240, 240);
      const text = rgb(160, 30, 30);
      paintFileCell(book, cells(10), CROSS, rgb(176, 40, 40), white, { bold: true, fontSize: 8, hAlign: "center" });
      paintFileCell(book, cells(11, 12), file.label, red, text, { bold: true, fontSize: 7, hAlign: "left" });
      paintFileCell(book, cells(13, 16), `*${file.token}*.csv`, red, text, { bold: false, fontSize: 7, hAlign: "left" });
      paintFileCell(book, cells(17, 18), folderPath || null, red, dark, { bold: false, fontSize: 6, hAlign: "left" });
    }
  });
}

// ---------------------------------------------------------------- Costs

/** ParseCostDateSerial: a date serial, or a date written as text. 0 when it is not a date. */
export function costDateSerial(value: unknown): number {
  if (typeof value === "number") return value >= 30_000 && value < 100_000 ? bankersRound(value) : 0;
  if (typeof value !== "string") return 0;
  const t = value.trim();
  if (vbaIsNumeric(t)) return costDateSerial(vbaCdbl(t));
  let match = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T].*)?$/.exec(t);
  if (match) return excelSerial(Number(match[1]), Number(match[2]), Number(match[3]));
  match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: .*)?$/.exec(t);
  if (match) return excelSerial(Number(match[3]), Number(match[1]), Number(match[2]));
  const parsed = Date.parse(t);
  if (!Number.isNaN(parsed) && /[a-z]/i.test(t)) {
    const date = new Date(parsed);
    return excelSerial(date.getFullYear(), date.getMonth() + 1, date.getDate());
  }
  return 0;
}

function costCell(book: SheetBook, row: number, col: number): { row: number; col: number } {
  const area = book.mergeAt(COSTS, row, col);
  return { row: area.row, col: area.col };
}

export interface CostDay {
  /** Excel date serial. */
  date: number;
  daily: number;
  /** The running total the sheet's formula shows, summed here so the form can display it. */
  total: number;
  /** The batch note in column H. */
  batch: string;
  row: number;
}

/** CostsGetDay for every filled day row: the date, the daily cost, the running total and the batch note. */
export function readCostDays(book: SheetBook): CostDay[] {
  if (!book.hasSheet(COSTS)) return [];
  const days: CostDay[] = [];
  let running = 0;
  for (let row = COSTS_FIRST; row <= COSTS_LAST; row += 1) {
    const daily = book.value(COSTS, row, 5);
    if (cellText(daily) === "") continue;
    const amount = typeof daily === "number" ? daily : vbaIsNumeric(String(daily)) ? vbaCdbl(String(daily)) : Number.NaN;
    if (Number.isNaN(amount)) continue;
    running += amount;
    days.push({ date: costDateSerial(book.value(COSTS, row, 3)) || days[0]?.date + (row - COSTS_FIRST) || 0, daily: amount, total: running, batch: cellText(book.value(COSTS, row, 8)), row });
  }
  return days;
}

export interface CostDayWrite {
  /** True when the day was new and a row was appended. */
  added: boolean;
  row: number;
}

/**
 * CostsSetDay: write one day's cost. An existing date is updated in place; a new date appends
 * a row with the same date formula and running-total formula the CSV sync writes. The batch
 * note lands in column H. The running total stays in column G. False when the date is not a
 * date or the day block is full.
 */
export function writeCostDay(book: SheetBook, date: number, daily: number, batch: string): CostDayWrite | null {
  if (!book.hasSheet(COSTS) || date <= 0) return null;
  let row = 0;
  let empty = 0;
  for (let at = COSTS_FIRST; at <= COSTS_LAST; at += 1) {
    if (cellText(book.value(COSTS, at, 5)) === "") {
      if (empty === 0) empty = at;
      continue;
    }
    const serial = costDateSerial(book.value(COSTS, at, 3)) || (empty === 0 ? date - (at - COSTS_FIRST) : 0);
    if (serial === date) row = at;
  }
  const added = row === 0;
  if (added) {
    if (empty === 0) return null;
    row = empty;
  }
  book.setFormat(COSTS, row, 5, FMT_MONEY);
  book.setValue(COSTS, row, 5, daily);
  book.setValue(COSTS, row, 8, batch.trim());
  if (added) {
    const dateCell = costCell(book, row, 3);
    book.setFormat(COSTS, dateCell.row, dateCell.col, FMT_COST_DATE);
    const previous = row > COSTS_FIRST ? costDateSerial(book.value(COSTS, row - 1, 3)) : 0;
    if (row === COSTS_FIRST || previous === 0 || date !== previous + 1) book.setValue(COSTS, dateCell.row, dateCell.col, date);
    else book.setFormula(COSTS, dateCell.row, dateCell.col, `=IF(E${row}<>"",C${row - 1}+1,"")`);
  }
  const totalFormula = row === COSTS_FIRST ? `=IF(ISBLANK(E${row}),"",E${row})` : `=IF(ISBLANK(E${row}),"",SUM($E$${COSTS_FIRST}:E${row}))`;
  if (!sameCostFormula(book.formula(COSTS, row, 7), totalFormula)) {
    book.setFormat(COSTS, row, 7, FMT_MONEY);
    book.setFormula(COSTS, row, 7, totalFormula);
  }
  bindCostHeadline(book);
  return { added, row };
}

function sameCostFormula(have: string | null, want: string): boolean {
  const compact = (value: string) => value.replace(/^=/, "").replace(/\s+/g, "").toUpperCase();
  return compact(have ?? "") === compact(want);
}

/**
 * Data!E3 is one day's cost: the last number in Costs!E, not the array lookup that was
 * adding the previous day onto it. Costs!E73 uses the same last-number lookup.
 */
export function bindCostHeadline(book: SheetBook): boolean {
  if (!book.hasSheet("Data") || !book.hasSheet(COSTS)) return false;
  let last = 0;
  for (let row = COSTS_LAST; row >= COSTS_FIRST; row -= 1) {
    const daily = book.value(COSTS, row, 5);
    if (cellText(daily) === "") continue;
    const amount = typeof daily === "number" ? daily : vbaIsNumeric(String(daily)) ? vbaCdbl(String(daily)) : Number.NaN;
    if (!Number.isFinite(amount)) continue;
    last = row;
    break;
  }
  if (last === 0) return false;
  const before = book.writes.length;
  const pointed = /COSTS!E(\d+)/i.exec(book.formula("Data", 3, 5) ?? "");
  if (!pointed || Number(pointed[1]) !== last) book.setFormula("Data", 3, 5, `=IFERROR(Costs!E${last},"")`);
  const lookup = book.formula(COSTS, 73, 5) ?? "";
  if (/1\s*\//.test(lookup) || !/9\.9E\+307/i.test(lookup)) book.setFormula(COSTS, 73, 5, "=LOOKUP(9.9E+307,E3:E72)");
  return book.writes.length > before;
}

/** SyncCostsSheet: the sheet's own dated costs (kept unless a new job), the CSV days laid over. */
export function syncCostsSheet(book: SheetBook, newJob: boolean): void {
  if (!book.hasSheet(COSTS)) return;
  const byDate = new Map<number, number>();
  if (!newJob) {
    const seedValue = book.value(COSTS, COSTS_FIRST, 3);
    const seed = typeof seedValue === "number" ? seedValue : 0;
    for (let row = COSTS_FIRST; row <= COSTS_LAST; row += 1) {
      const daily = book.value(COSTS, row, 5);
      if (cellText(daily) === "") continue;
      let serial = costDateSerial(book.value(COSTS, row, 3));
      if (serial === 0 && seed > 0) serial = bankersRound(seed) + (row - COSTS_FIRST);
      if (serial <= 0) continue;
      const amount = typeof daily === "number" ? daily : vbaIsNumeric(String(daily)) ? vbaCdbl(String(daily)) : Number.NaN;
      if (Number.isNaN(amount)) throw new Error(`Type mismatch in Costs!E${row}.`);
      byDate.set(serial, amount);
    }
  }
  const oc = new OcTable(book, "_OC_Costs");
  if (oc.exists) {
    const last = oc.lastRow(1);
    for (let row = 2; row <= last; row += 1) {
      const serial = costDateSerial(oc.value(row, 2));
      const cost = oc.value(row, 3);
      if (serial > 0 && cost != null && (typeof cost === "number" || vbaIsNumeric(String(cost)))) {
        byDate.set(serial, typeof cost === "number" ? cost : vbaCdbl(String(cost)));
      }
    }
  }
  if (byDate.size === 0) {
    if (newJob) book.clear(COSTS, { row: COSTS_FIRST, col: 5, row2: COSTS_LAST, col2: 5 });
    else throw new Error("No cost rows found in Costs sheet or ticket-costs CSV (_OC_Costs).");
    return;
  }
  if (byDate.size > COSTS_LAST - COSTS_FIRST + 1) throw new Error(`Too many cost days for Costs sheet rows ${COSTS_FIRST}-${COSTS_LAST}`);
  const dates = [...byDate.keys()].sort((a, b) => a - b);
  const consecutive = dates.every((date, i) => i === 0 || date === dates[i - 1]! + 1);
  book.clear(COSTS, { row: COSTS_FIRST, col: 5, row2: COSTS_LAST, col2: 5 });
  dates.forEach((date, i) => {
    const row = COSTS_FIRST + i;
    book.setFormat(COSTS, row, 5, FMT_MONEY);
    book.setValue(COSTS, row, 5, byDate.get(date)!);
    const dateCell = costCell(book, row, 3);
    book.setFormat(COSTS, dateCell.row, dateCell.col, FMT_COST_DATE);
    if (i === 0 || !consecutive) book.setValue(COSTS, dateCell.row, dateCell.col, date);
    else book.setFormula(COSTS, dateCell.row, dateCell.col, `=IF(E${row}<>"",C${row - 1}+1,"")`);
    book.setFormat(COSTS, row, 7, FMT_MONEY);
    book.setFormula(COSTS, row, 7, row === COSTS_FIRST ? "=IF(ISBLANK(E3),\"\",E3)" : `=IF(ISBLANK(E${row}),"",SUM($E$3:E${row}))`);
  });
  for (let row = COSTS_FIRST + dates.length; row <= COSTS_LAST; row += 1) {
    if (vbaStr(book.value(COSTS, row, 2)).trim() === "") break;
    const dateCell = costCell(book, row, 3);
    book.setFormat(COSTS, dateCell.row, dateCell.col, FMT_COST_DATE);
    book.setFormula(COSTS, dateCell.row, dateCell.col, `=IF(E${row}<>"",C${row - 1}+1,"")`);
    book.setFormat(COSTS, row, 7, FMT_MONEY);
    book.setFormula(COSTS, row, 7, `=IF(ISBLANK(E${row}),"",SUM($E$3:E${row}))`);
  }
  bindCostHeadline(book);
}

// ---------------------------------------------------------------- Mud Motors On Location

function steerText(text: string): boolean {
  const t = text.toLowerCase();
  return t.includes("orbit rss") || t.includes("icruise") || t.includes("i-cruise");
}

function mudMotorInventory(serial: string, itemName: string, subCat: string): boolean {
  if (serial.trim() === "") return false;
  if (itemName.toLowerCase().includes("mud motor")) return true;
  if (subCat.trim().toLowerCase() === "motor") return true;
  return steerText(itemName);
}

function mudMotorSerial(serial: string, desc: string): boolean {
  if (serial.trim() === "") return false;
  return desc.toLowerCase().includes("mud motor") || steerText(desc);
}

/** PHX-525010-PTS -> 525010; 24X-19722-PTS -> 19722; anything else sorts last. */
export function mudMotorSequence(serial: string): number {
  let s = serial.trim().toUpperCase();
  if (!s.endsWith("-PTS")) return 1e99;
  s = s.slice(0, -4);
  const dash = s.lastIndexOf("-");
  const token = dash >= 0 ? s.slice(dash + 1) : s;
  return token !== "" && vbaIsNumeric(token) ? vbaCdbl(token) : 1e99;
}

/** Range.Value2 = text on a General cell: Excel keeps a number-looking serial as a number. */
function typedSerial(serial: string): string | number {
  return /^-?\d+(\.\d+)?$/.test(serial) ? Number(serial) : serial;
}

/** WriteMudMotorsFromInventory: Data B45:B55 from inventory motors (BHA fallback), then ToolHours_Sync. */
export function syncMudMotors(book: SheetBook): number {
  if (!book.hasSheet(DATA)) return 0;
  const serials: string[] = [];
  const add = (serial: string) => {
    if (serial === "" || serials.some((s) => s.toLowerCase() === serial.toLowerCase())) return;
    serials.push(serial);
  };
  const inventory = new OcTable(book, "_OC_Inventory");
  if (inventory.exists) {
    const cSn = inventory.col("SerialNumber") || inventory.col("Serial #");
    const cName = inventory.col("ItemName") || inventory.col("Description");
    const cSub = inventory.col("SubCategory");
    if (cSn > 0) {
      const last = inventory.lastRow(cSn);
      for (let row = 2; row <= last; row += 1) {
        const serial = inventory.str(row, cSn).trim();
        if (mudMotorInventory(serial, inventory.str(row, cName), inventory.str(row, cSub))) add(serial);
      }
    }
  }
  if (serials.length === 0) {
    const bha = new OcTable(book, "_OC_BHA");
    const cSn = bha.col("Serial #") || bha.col("SerialNumber");
    const cDesc = bha.col("Description");
    if (bha.exists && cSn > 0) {
      const last = bha.lastRow(cSn);
      for (let row = 2; row <= last; row += 1) {
        const serial = bha.str(row, cSn).trim();
        if (mudMotorSerial(serial, bha.str(row, cDesc))) add(serial);
      }
    }
  }
  const seqs = serials.map(mudMotorSequence);
  for (let i = 0; i < serials.length - 1; i += 1) {
    for (let j = i + 1; j < serials.length; j += 1) {
      if (seqs[j]! < seqs[i]!) {
        [seqs[i], seqs[j]] = [seqs[j]!, seqs[i]!];
        [serials[i], serials[j]] = [serials[j]!, serials[i]!];
      }
    }
  }
  for (let row = MM_FIRST; row < MM_FIRST + MM_ROWS; row += 1) {
    const cell = book.mergeAt(DATA, row, 2);
    book.setValue(DATA, cell.row, cell.col, null);
  }
  const placed = serials.slice(0, MM_ROWS);
  placed.forEach((serial, i) => {
    const cell = book.mergeAt(DATA, MM_FIRST + i, 2);
    book.setValue(DATA, cell.row, cell.col, typedSerial(serial));
  });
  try {
    toolHoursSync(book);
  } catch {
    // ToolHours_Sync runs under On Error Resume Next here: a failure leaves the motor list in place.
  }
  return placed.length;
}

// ---------------------------------------------------------------- the refresh

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/**
 * RefreshAllCsvData without the Setup rebuild: load the CSVs that were found, detect a job
 * change against OC_LastJobId, refresh Setup's panels, then sync Costs and Mud Motors.
 */
export function applyOpenCap(book: SheetBook, sources: OpenCapSource[], options: OpenCapOptions): OpenCapResult {
  const rows: Partial<Record<OpenCapKind, number>> = {};
  const found: Partial<Record<OpenCapKind, { name: string; path: string }>> = {};
  for (const file of OPEN_CAP_FILES) {
    const source = sources.find((entry) => entry.kind === file.kind);
    if (!source) continue;
    rows[file.kind] = loadOpenCapSheet(book, file.sheet, parseOpenCapCsv(source.csv));
    found[file.kind] = { name: source.name, path: source.path };
  }

  const fields = readJobFields(book);
  const jobId = jobField(fields, "Job ID").trim();
  const previous = options.lastJobId.trim();
  const newJob = options.adopt === true || (jobId !== "" && previous !== "" && jobId.toLowerCase() !== previous.toLowerCase());
  const lastJobId = jobId || previous;

  refreshJobPanel(book, fields);
  const crew = readCrewRows(book);
  refreshCrewPanel(book, crew, options.crewTop, options.crewBottom);
  refreshFilesPanel(book, found, options.folderPath);

  const alerts: string[] = [];
  try {
    syncCostsSheet(book, newJob);
  } catch (error) {
    if (book.hasSheet(SETUP)) book.setValue(SETUP, 1, 20, `COSTS SYNC ERR: ${errorText(error)}`);
    alerts.push(`Costs sync failed:\n\n${errorText(error)}`);
  }
  let motors = 0;
  try {
    motors = syncMudMotors(book);
  } catch (error) {
    if (book.hasSheet(SETUP)) book.setValue(SETUP, 1, 20, `MOTOR SYNC ERR: ${errorText(error)}`);
  }

  const loaded = OPEN_CAP_FILES.filter((file) => found[file.kind]).length;
  const jobNumber = jobField(fields, "Job Code", "Job ID");
  const status = loaded === 0
    ? "No OpenCap CSVs found."
    : `OpenCap: ${loaded} of 6 files loaded${jobNumber ? ` for job ${jobNumber}` : ""}${newJob ? " (new job)" : ""}. ${crew.length} crew, ${motors} mud motor(s).`;
  return { jobId, jobNumber, lastJobId, newJob, rows, status, alert: alerts.length > 0 ? alerts.join("\n\n") : null };
}
