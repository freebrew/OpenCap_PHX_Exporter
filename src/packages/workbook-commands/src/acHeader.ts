import { cellText, type SheetBook } from "./sheetBook";

// The Compass report header and the Summary's site headings. The report is already on one datum
// (the active well's slot, true north, the reference elevation), so nothing is converted: the check
// only says whether this report belongs to the plan that is loaded.

export const DATUM_SHEET = "_OC_Datum";

export interface AcReportHeader {
  referenceWell: string;
  referenceWellbore: string;
  referenceDesign: string;
  referenceSite: string;
  /** "Local Co-ordinate Reference", including a slot name that wrapped onto the next line. */
  localCoordinateReference: string;
  coordinatesRelativeTo: string;
  tvdReference: string;
  mdReference: string;
  northReference: string;
  offsetTvdReference: string;
  reportDate: string;
}

export interface PlanIdentity {
  well: string;
  slot: string;
  design: string;
  northReference: string;
  tvdReference: string;
}

export interface HeaderCheck {
  ok: boolean;
  /** The one line the view shows. A failure is also the import note and the badge. */
  verdict: string;
  reasons: string[];
}

export interface AcSiteWell {
  site: string;
  well: string;
}

export interface DatumRecord {
  /** The report column has a reference well or a design. */
  reportStored: boolean;
  /** The plan column has a well or a design. */
  planStored: boolean;
  report: AcReportHeader;
  plan: PlanIdentity;
}

/** A label we read, or a label that only stops the value before it. Longest first is not required; overlaps are dropped. */
const LABELS = [
  "Local Co-ordinate Reference",
  "Coordinates are relative to",
  "Survey Calculation Method",
  "Offset TVD Reference",
  "Output errors are at",
  "Interpolation Method",
  "Reference Wellbore",
  "Results Limited by",
  "Reference Depths",
  "North Reference",
  "Reference Design",
  "Reference Site",
  "Reference Well",
  "Coordinate System",
  "Offset Depths",
  "TVD Reference",
  "MD Reference",
  "Warning Levels",
  "Depth Range",
  "Central Meridian",
  "Grid Convergence",
  "Site Error",
  "Well Error",
  "Filter type",
  "Database",
  "Company",
  "Project",
];

const WANTED = new Set([
  "Local Co-ordinate Reference",
  "Coordinates are relative to",
  "Offset TVD Reference",
  "North Reference",
  "Reference Wellbore",
  "Reference Design",
  "Reference Site",
  "Reference Well",
  "TVD Reference",
  "MD Reference",
]);

/** Column 1 is the field name, column 2 the report, column 3 the plan. Row 1 is the titles. */
const FIELDS: Array<{ row: number; field: string; report: keyof AcReportHeader; plan: keyof PlanIdentity | "slotLine" | "" }> = [
  { row: 2, field: "Reference Well", report: "referenceWell", plan: "well" },
  { row: 3, field: "Reference Wellbore", report: "referenceWellbore", plan: "" },
  { row: 4, field: "Reference Design", report: "referenceDesign", plan: "design" },
  { row: 5, field: "Reference Site", report: "referenceSite", plan: "" },
  { row: 6, field: "Local Co-ordinate Reference", report: "localCoordinateReference", plan: "slotLine" },
  { row: 7, field: "Coordinates Relative To", report: "coordinatesRelativeTo", plan: "" },
  { row: 8, field: "TVD Reference", report: "tvdReference", plan: "tvdReference" },
  { row: 9, field: "MD Reference", report: "mdReference", plan: "" },
  { row: 10, field: "North Reference", report: "northReference", plan: "northReference" },
  { row: 11, field: "Offset TVD Reference", report: "offsetTvdReference", plan: "" },
  { row: 12, field: "Report Date", report: "reportDate", plan: "" },
  { row: 13, field: "Plan Slot", report: "reportDate", plan: "slot" },
];

const SITE_HEADING = /^(?:offset\s+)?\d{1,3}-\d{1,3}-\d{2,3}-\d{1,2}(?:\s*w\s?\d[a-z]*)?$/i;

function squash(text: string): string {
  return text.toLowerCase().replace(/\s+/g, " ").trim();
}

function cleanValue(raw: string): string {
  return raw.replace(/^[:\s]+/, "").replace(/\s+/g, " ").trim();
}

function emptyReport(): AcReportHeader {
  return {
    referenceWell: "",
    referenceWellbore: "",
    referenceDesign: "",
    referenceSite: "",
    localCoordinateReference: "",
    coordinatesRelativeTo: "",
    tvdReference: "",
    mdReference: "",
    northReference: "",
    offsetTvdReference: "",
    reportDate: "",
  };
}

function emptyPlan(): PlanIdentity {
  return { well: "", slot: "", design: "", northReference: "", tvdReference: "" };
}

interface LabelAt {
  label: string;
  at: number;
  len: number;
}

function labelSpans(text: string): LabelAt[] {
  const lower = text.toLowerCase();
  const found: LabelAt[] = [];
  for (const label of LABELS) {
    const needle = label.toLowerCase();
    let from = 0;
    while (from < lower.length) {
      const at = lower.indexOf(needle, from);
      if (at < 0) break;
      found.push({ label, at, len: label.length });
      from = at + label.length;
    }
  }
  found.sort((a, b) => a.at - b.at || b.len - a.len);
  const kept: LabelAt[] = [];
  let covered = -1;
  for (const item of found) {
    if (item.at < covered) continue;
    kept.push(item);
    covered = item.at + item.len;
  }
  return kept;
}

/** The value ends at the line end. A slot name is the exception: "Slot" is printed, and the slot itself is on the next line. */
function valueAfter(text: string, start: number, nextAt: number): string {
  const raw = text.slice(start, nextAt);
  const lines = raw.split(/\r?\n/);
  let value = lines[0] ?? "";
  const tail = (lines[1] ?? "").trim();
  if (/slot\s*$/i.test(value) && tail !== "" && tail.length <= 24 && !tail.includes(":")) value = `${value} ${tail}`;
  return cleanValue(value);
}

function fieldsOf(text: string): Map<string, string> {
  const spans = labelSpans(text);
  const out = new Map<string, string>();
  spans.forEach((span, index) => {
    if (!WANTED.has(span.label) || out.has(span.label)) return;
    const next = spans[index + 1]?.at ?? text.length;
    out.set(span.label, valueAfter(text, span.at + span.len, next));
  });
  return out;
}

function reportDate(text: string): string {
  const long = text.match(/\b(\d{1,2}\s+[A-Z][a-z]+,\s+\d{4})\b/);
  if (long) return long[1]!;
  const short = text.match(/\b(\d{1,2}\/\d{1,2}\/\d{4})\b/);
  return short ? short[1]! : "";
}

/** The report's identity, datums and coordinate reference. Empty strings where a label was not printed. */
export function parseAcHeader(text: string): AcReportHeader {
  const fields = fieldsOf(text);
  return {
    referenceWell: fields.get("Reference Well") ?? "",
    referenceWellbore: fields.get("Reference Wellbore") ?? "",
    referenceDesign: fields.get("Reference Design") ?? "",
    referenceSite: fields.get("Reference Site") ?? "",
    localCoordinateReference: fields.get("Local Co-ordinate Reference") ?? "",
    coordinatesRelativeTo: fields.get("Coordinates are relative to") ?? "",
    tvdReference: fields.get("TVD Reference") ?? "",
    mdReference: fields.get("MD Reference") ?? "",
    northReference: fields.get("North Reference") ?? "",
    offsetTvdReference: fields.get("Offset TVD Reference") ?? "",
    reportDate: reportDate(text),
  };
}

function summaryWell(line: string): string {
  const at = line.search(/\d{2,3}\/\d{1,3}-/);
  if (at < 0) return "";
  const rest = line.slice(at);
  const numAt = rest.search(/\s+\d{1,3}(?:,\d{3})*\.\d{2}/);
  const oor = rest.toLowerCase().indexOf("out of range");
  let end = rest.length;
  if (numAt > 0) end = Math.min(end, numAt);
  if (oor > 0) end = Math.min(end, oor);
  return rest.slice(0, end).replace(/\s+/g, " ").trim();
}

/** Each Summary well with the site heading printed above it. A well before any heading has an empty site. */
export function parseAcSites(text: string): AcSiteWell[] {
  const lower = text.toLowerCase();
  const start = lower.indexOf("summary");
  if (start < 0) return [];
  let end = text.length;
  for (const marker of ["semi major", "offset design", "highside"]) {
    const at = lower.indexOf(marker, start + "summary".length);
    if (at >= 0 && at < end) end = at;
  }
  const out: AcSiteWell[] = [];
  let site = "";
  for (const raw of text.slice(start, end).split("\n")) {
    const line = raw.replace(/\s+/g, " ").trim();
    if (line === "" || line.toLowerCase() === "summary") continue;
    if (SITE_HEADING.test(line)) {
      site = line;
      continue;
    }
    const well = summaryWell(line);
    if (well !== "") out.push({ site, well });
  }
  return out;
}

/** The plan report's well, slot and design. North and TVD are filled when that report prints them. */
export function parsePlanHeader(text: string): PlanIdentity {
  const first = text.split("\f")[0] ?? text;
  let well = "";
  let slot = "";
  let design = "";
  for (const raw of first.split("\n")) {
    const line = raw.replace(/\s+/g, " ").trim();
    const lower = line.toLowerCase();
    const slotAt = lower.indexOf(" - slot ");
    if (slotAt >= 0 && well === "") {
      well = line.slice(0, slotAt).trim();
      slot = line.slice(slotAt + " - slot ".length).trim();
    }
    const plan = /^plan:\s*(.+)$/i.exec(line);
    if (plan && design === "") design = (plan[1]!.split(/\s{2,}/)[0] ?? "").trim();
  }
  const fields = fieldsOf(text);
  return {
    well,
    slot,
    design,
    northReference: fields.get("North Reference") ?? "",
    tvdReference: fields.get("TVD Reference") ?? "",
  };
}

function sameText(a: string, b: string): boolean {
  const left = squash(a);
  const right = squash(b);
  if (left === "" || right === "") return true;
  return left === right || left.startsWith(right) || right.startsWith(left);
}

function sameDesign(a: string, b: string): boolean {
  const left = a.toLowerCase().replace(/\s+/g, "");
  const right = b.toLowerCase().replace(/\s+/g, "");
  if (left === "" || right === "") return true;
  return left === right;
}

function sameWell(a: string, b: string): boolean {
  const left = squash(a);
  const right = squash(b);
  if (left === "" || right === "") return true;
  return left.includes(right) || right.includes(left);
}

/** The coordinate reference names the reference well's slot. A site centre or a map grid does not. An empty value was not printed, so it is not a failure. */
function slotOfWell(local: string, well: string): boolean {
  const text = squash(local);
  if (text === "") return true;
  if (text.includes("site centre") || text.includes("site center")) return false;
  const name = squash(well);
  return name !== "" && text.includes(name);
}

function referenceDatum(value: string): boolean {
  const text = squash(value);
  if (text === "") return true;
  return text.includes("reference datum");
}

/**
 * Whether the report is the plan's own report. `plan` is null when the plan's identity was never
 * stored: the centre-to-centre proof then decides on its own, as it did before this check.
 */
export function compareHeaders(report: AcReportHeader, plan: PlanIdentity | null): HeaderCheck {
  if (squash(report.referenceWell) === "" && squash(report.referenceDesign) === "") {
    return { ok: true, verdict: "Report Header Was Not Read", reasons: [] };
  }
  if (!plan || (squash(plan.well) === "" && squash(plan.design) === "")) {
    return { ok: true, verdict: "Plan Identity Not Stored", reasons: [] };
  }
  const reasons: string[] = [];
  if (!sameWell(report.referenceWell, plan.well) || !sameDesign(report.referenceDesign, plan.design)) {
    reasons.push(`This report is for ${report.referenceWell} ${report.referenceDesign}. The plan is ${plan.well} ${plan.design}.`);
  }
  const local = report.localCoordinateReference || report.coordinatesRelativeTo;
  if (!slotOfWell(local, report.referenceWell)) {
    reasons.push(`Coordinates are relative to ${local}, not ${report.referenceWell}'s slot.`);
  }
  if (!referenceDatum(report.offsetTvdReference)) {
    reasons.push(`Offset TVD Reference is ${report.offsetTvdReference}, not Reference Datum.`);
  }
  if (!sameText(report.northReference, plan.northReference)) {
    reasons.push(`North Reference is ${report.northReference}, the plan is ${plan.northReference}.`);
  }
  if (reasons.length === 0) return { ok: true, verdict: "Same Well and Design", reasons };
  return { ok: false, verdict: reasons.join(" "), reasons };
}

function ensureDatum(book: SheetBook): void {
  if (!book.hasSheet(DATUM_SHEET)) book.resetSheet(DATUM_SHEET);
  book.setValue(DATUM_SHEET, 1, 1, "Field");
  book.setValue(DATUM_SHEET, 1, 2, "Report");
  book.setValue(DATUM_SHEET, 1, 3, "Plan");
  for (const item of FIELDS) book.setValue(DATUM_SHEET, item.row, 1, item.field);
}

function planValue(plan: PlanIdentity, column: (typeof FIELDS)[number]["plan"]): string {
  if (column === "" || column === "slotLine") {
    if (column === "slotLine" && plan.slot !== "") return `${plan.well} - Slot ${plan.slot}`;
    return "";
  }
  return plan[column];
}

/** Column 2 of `_OC_Datum`. The plan column is left as it was. */
export function writeReportHeader(book: SheetBook, report: AcReportHeader): void {
  ensureDatum(book);
  for (const item of FIELDS) {
    if (item.field === "Plan Slot") continue;
    book.setValue(DATUM_SHEET, item.row, 2, report[item.report]);
  }
}

/** Column 3 of `_OC_Datum`. The report column is left as it was. */
export function writePlanIdentity(book: SheetBook, plan: PlanIdentity): void {
  ensureDatum(book);
  for (const item of FIELDS) {
    if (item.plan === "") continue;
    book.setValue(DATUM_SHEET, item.row, 3, planValue(plan, item.plan));
  }
}

/** A CSV plan has no header, so the plan column is cleared and the check waits. */
export function clearPlanIdentity(book: SheetBook): void {
  if (!book.hasSheet(DATUM_SHEET)) return;
  ensureDatum(book);
  for (const item of FIELDS) book.setValue(DATUM_SHEET, item.row, 3, "");
}

function cell(book: SheetBook, row: number, col: number): string {
  return cellText(book.value(DATUM_SHEET, row, col)).trim();
}

/** Both columns of `_OC_Datum`. Missing sheet and blank cells read as empty. */
export function readDatum(book: SheetBook): DatumRecord {
  const report = emptyReport();
  const plan = emptyPlan();
  if (book.hasSheet(DATUM_SHEET)) {
    for (const item of FIELDS) {
      if (item.field !== "Plan Slot") report[item.report] = cell(book, item.row, 2);
      if (item.plan === "slot") plan.slot = cell(book, item.row, 3);
      else if (item.plan !== "" && item.plan !== "slotLine") plan[item.plan] = cell(book, item.row, 3);
    }
  }
  return {
    reportStored: report.referenceWell !== "" || report.referenceDesign !== "",
    planStored: plan.well !== "" || plan.design !== "",
    report,
    plan,
  };
}
