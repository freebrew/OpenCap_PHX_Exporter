import { integrateLocalNE } from "../../dd-math/src/plan";
import { clearPlanIdentity, parsePlanHeader, writePlanIdentity } from "./acHeader";
import { unreadablePdfText } from "./glyphText";
import { cellText, numeric, type SheetBook } from "./sheetBook";
import { PLANSEC_SHEET, SS_SHEET, syncPlanTargetWindow, seedNameFromText, vbaVal } from "./slidesheet";

// Port of MDL_Setup Import Plan (ImportSurveyPlanFile and its PDF / CSV parsers) and
// MDL_PlanImport (Planned Survey + Formations from the same Compass Planning Report).
// The parsers take the PDF text the VBA gets from `pdftotext -layout`.

export const SURVEY_SHEET = "_OC_Survey";
export const FORMATIONS_SHEET = "_OC_Formations";
export const SETUP_SHEET = "Setup";
const DATA_SHEET = "Data";
const LITHO_SRC = "AER Table of Formations 2019 / AGS WCSB Atlas";
const MAX_SECTIONS = 80;
const MAX_PLANNED = 800;
const MAX_FORMATIONS = 120;
const TD_ROW_PLANNED = 54;
const TD_COL_FIRST = 9;

export const NO_PLAN_TABLE = "No Plan Sections / SECTION DETAILS / Plan Annotations table found in that PDF.";
export const NO_PDF_TEXT = "Could not read any text from that PDF. A scanned (image-only) plan has no text to import.";
export const NO_LETTER_MAP = "That PDF's text cannot be read: its fonts store glyph numbers instead of letters, and the glyphs did not match Arial. The previous plan was left unchanged.\n\nImport the plan's .csv instead, or ask for the plan printed to the Adobe PDF printer.";

export interface PlanStation {
  md: number;
  inc: number;
  azm: number;
  tvd: number;
  ns: number;
  ew: number;
  dls: number;
  bld: number;
  trn: number;
  ann: string;
}

export interface PlannedStation {
  md: number;
  inc: number;
  azm: number;
  tvd: number;
  ns: number;
  ew: number;
}

export interface Formation {
  md: number;
  tvd: number;
  name: string;
  lithology: string;
}

// ---------------------------------------------------------------- VBA string helpers

/** InStr(start, text, find, vbTextCompare) as a 0-based index, -1 when absent. */
function findText(text: string, find: string, from = 0): number {
  return text.toLowerCase().indexOf(find.toLowerCase(), from);
}

function hasText(text: string, find: string): boolean {
  return findText(text, find) >= 0;
}

/** VBA Trim$: spaces only. */
function vbTrim(text: string): string {
  return text.replace(/^ +| +$/g, "");
}

function collapseSpaces(text: string): string {
  let out = text;
  while (out.includes("  ")) out = out.replaceAll("  ", " ");
  return out;
}

/** CDbl of a number the regex matched. */
function cdbl(text: string): number {
  return Number(text);
}

/** CleanNum: commas and stray spaces the extractor puts inside numbers. */
export function cleanNum(text: string): number {
  return Number(text.replaceAll(",", "").replaceAll(" ", ""));
}

/** Mid$(text, startPos, endPos - startPos) between a heading and the first end marker after it. */
function regionAfter(text: string, heading: number, skip: number, markers: string[]): string {
  let end = text.length;
  for (const marker of markers) {
    const p = findText(text, marker, heading + skip);
    if (p >= 0 && p < end) end = p;
  }
  return text.slice(heading, end);
}

/** The PDF text's table region with the extractor's noise removed ("\", thousands commas, "0. 000"). */
function tidyRegion(region: string): string {
  let out = region.replaceAll("\\", "").replaceAll(",", "");
  while (out.includes(". ")) out = out.replaceAll(". ", ".");
  return out;
}

function station(md: number, inc: number, azm: number, tvd: number, ns: number, ew: number, dls: number, bld: number, trn: number, ann: string): PlanStation {
  return { md, inc, azm, tvd, ns, ew, dls, bld, trn, ann };
}

function validPlanStation(md: number, inc: number, azi: number): boolean {
  return md >= 0 && md < 20000 && inc >= 0 && inc <= 180 && azi >= 0 && azi <= 360;
}

export function cleanSectionAnnot(text: string): string {
  let s = collapseSpaces(vbTrim(text.replaceAll("\n", " ").replaceAll("\r", " ")));
  const eq = s.lastIndexOf("=");
  if (eq >= 0) {
    const m = s.indexOf("m", eq);
    if (m >= 0) s = s.slice(0, m + 1);
  } else if (s.length > 60) s = s.slice(0, 60);
  return vbTrim(s);
}

function scorePlanParse(rows: PlanStation[]): number {
  if (rows.length < 2) return -1;
  let maxMd = 0;
  for (const row of rows) if (row.md > maxMd) maxMd = row.md;
  return maxMd + rows.length * 5;
}

// ---------------------------------------------------------------- Plan Sections parsers

function parseSectionDetailsSpaced(region: string): PlanStation[] {
  const re = /(-?\d{1,5}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,2}\.\d{3})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,5}\.\d{2})/g;
  const out: PlanStation[] = [];
  for (const m of region.matchAll(re)) {
    if (out.length > MAX_SECTIONS) break;
    const md = cdbl(m[1]!);
    if (!validPlanStation(md, cdbl(m[2]!), cdbl(m[3]!))) continue;
    const prev = out[out.length - 1];
    if (prev && md + 0.05 < prev.md) continue;
    out.push(station(md, cdbl(m[2]!), cdbl(m[3]!), cdbl(m[4]!), cdbl(m[5]!), cdbl(m[6]!), cdbl(m[7]!), 0, 0, ""));
  }
  return out;
}

function parseSectionDetailsBlock(text: string): PlanStation[] {
  const start = findText(text, "SECTION DETAILS");
  if (start < 0) return [];
  const region = tidyRegion(regionAfter(text, start, 20, ["DESIGN TARGET", "LOCAL COORDINATES", "Planning Report"]));
  const re = /(-?\d{1,5}\.\d{2})(-?\d{1,3}\.\d{2})(-?\d{1,3}\.\d{2})(-?\d{1,5}\.\d{2})(-?\d{1,5}\.\d{2})(-?\d{1,5}\.\d{2})(-?\d{1,2}\.\d{3})(-?\d{1,3}\.\d{2})(-?\d{1,5}\.\d{2})/g;
  const matches = [...region.matchAll(re)];
  const out: PlanStation[] = [];
  for (let i = 0; i < matches.length; i += 1) {
    if (out.length > MAX_SECTIONS) break;
    const m = matches[i]!;
    const after = m.index + m[0].length;
    const next = matches[i + 1];
    const annLen = next ? next.index - after : region.length - after;
    const ann = annLen > 0 ? cleanSectionAnnot(region.substr(after, annLen)) : "";
    out.push(station(cdbl(m[1]!), cdbl(m[2]!), cdbl(m[3]!), cdbl(m[4]!), cdbl(m[5]!), cdbl(m[6]!), cdbl(m[7]!), 0, 0, ann));
  }
  return out.length < 2 ? parseSectionDetailsSpaced(region) : out;
}

/** Glyphs grouped by Y land a row as: EW NS TVD Azi Inc MD [Target]. */
function parsePlanSectionsYGrouped(region: string): PlanStation[] {
  const re = /(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,5}\.\d{2})(?:\s+([A-Za-z][^0-9]{0,40}?))?(?=\s+-?\d{1,5}\.\d{2}|\s*$)/g;
  const out: PlanStation[] = [];
  for (const m of region.matchAll(re)) {
    if (out.length > MAX_SECTIONS) break;
    const md = cdbl(m[6]!);
    const inc = cdbl(m[5]!);
    const azi = cdbl(m[4]!);
    if (!validPlanStation(md, inc, azi)) continue;
    const prev = out[out.length - 1];
    if (prev && (md + 0.05 < prev.md || Math.abs(md - prev.md) < 0.05)) continue;
    out.push(station(md, inc, azi, cdbl(m[3]!), cdbl(m[2]!), cdbl(m[1]!), 0, 0, 0, cleanSectionAnnot(m[7] ?? "")));
  }
  return out;
}

/** pdftotext -layout: MD Inc Azi TVD NS EW DLS BR TR [TFO] [Target]. */
function parsePlanSectionsLayout(region: string): PlanStation[] {
  const re = /(-?\d{1,5}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,2}\.\d{3})\s+(-?\d{1,3}\.\d{3})\s+(-?\d{1,3}\.\d{3})(?:\s+(-?\d{1,3}\.\d{2}))?(?:\s+([A-Za-z][^0-9]{0,40}?))?/g;
  const out: PlanStation[] = [];
  for (const m of region.matchAll(re)) {
    if (out.length > MAX_SECTIONS) break;
    const md = cdbl(m[1]!);
    const inc = cdbl(m[2]!);
    const azi = cdbl(m[3]!);
    if (!validPlanStation(md, inc, azi)) continue;
    const prev = out[out.length - 1];
    if (prev && (md + 0.05 < prev.md || Math.abs(md - prev.md) < 0.05)) continue;
    out.push(station(md, inc, azi, cdbl(m[4]!), cdbl(m[5]!), cdbl(m[6]!), cdbl(m[7]!), cdbl(m[8]!), cdbl(m[9]!), cleanSectionAnnot(m[11] ?? "")));
  }
  return out;
}

function parsePlanSectionsTable(text: string): PlanStation[] {
  let start = findText(text, "Plan Sections");
  if (start < 0) start = findText(text, "Plan Section");
  if (start < 0) return [];
  const region = tidyRegion(regionAfter(text, start, 14, ["Planning Report", "Planned Survey", "SECTION DETAILS", "DESIGN TARGET", "Plan Annotations", "Plan  Annotations"]));
  const grouped = parsePlanSectionsYGrouped(region);
  const layout = parsePlanSectionsLayout(region);
  const sLayout = scorePlanParse(layout);
  const sGrouped = scorePlanParse(grouped);
  if (sLayout >= sGrouped && sLayout >= 0) return layout;
  if (sGrouped >= 0) return grouped;
  return [];
}

function parsePlanAnnotationsBlock(text: string): PlanStation[] {
  let start = findText(text, "Plan  Annotations");
  if (start < 0) start = findText(text, "Plan Annotations");
  if (start < 0) return [];
  let end = text.length;
  const p = findText(text, "Planning Report", start + 20);
  if (p >= 0) end = p;
  const region = text.slice(start, end).replaceAll("\\", "");
  // Comments include digits ("KOP, 4°/30m = 1797.80m"), so the comment runs to the next station.
  const n2 = String.raw`\d{1,3}(?:,\d{3})*\.\d{2}`;
  const re = new RegExp(String.raw`(${n2})\s+(${n2})\s+(-?${n2})\s+(-?${n2})\s+(.+?)(?=\s+${n2}\s+${n2}|\s*$)`, "g");
  const out: PlanStation[] = [];
  for (const m of region.matchAll(re)) {
    if (out.length > MAX_SECTIONS) break;
    out.push(station(cleanNum(m[1]!), 0, 0, cleanNum(m[2]!), cleanNum(m[4]!), cleanNum(m[3]!), 0, 0, 0, vbTrim(m[5]!)));
  }
  return out;
}

function overlayPlanAnnotations(text: string, rows: PlanStation[]): void {
  const notes = parsePlanAnnotationsBlock(text);
  if (notes.length < 1) return;
  for (const row of rows) {
    if (vbTrim(row.ann) !== "") continue;
    const hit = notes.find((note) => Math.abs(row.md - note.md) < 0.25);
    if (hit) row.ann = hit.ann;
  }
}

/** ParsePlanSections: SECTION DETAILS, Plan Sections and Plan Annotations, keeping the richest. */
export function parsePlanSections(text: string): PlanStation[] {
  const safe = (parse: () => PlanStation[]) => {
    try {
      return parse();
    } catch {
      return [];
    }
  };
  const details = safe(() => parseSectionDetailsBlock(text));
  const table = safe(() => parsePlanSectionsTable(text));
  const notes = safe(() => parsePlanAnnotationsBlock(text));
  const sDet = scorePlanParse(details);
  const sTab = scorePlanParse(table);
  const sAnn = scorePlanParse(notes);
  const best = Math.max(sDet, sTab, sAnn);
  if (best < 0) return [];
  const rows = (best === sTab ? table : best === sDet ? details : notes).map((row) => ({ ...row }));
  try {
    overlayPlanAnnotations(text, rows);
  } catch {
    // On Error Resume Next around the overlay.
  }
  return rows;
}

// ---------------------------------------------------------------- target names

/** NamePlanSectionTargets: KOP, HEEL, TD, SOT, EOT, NUDGE and TANGENT from the notes and the shape. */
export function namePlanSectionTargets(rows: PlanStation[]): string[] {
  const n = rows.length;
  const md = (i: number) => rows[i]!.md;
  const inc = (i: number) => rows[i]!.inc;
  const dls = (i: number) => rows[i]!.dls;
  const trn = (i: number) => rows[i]!.trn;
  const ann = (i: number) => rows[i]!.ann;
  const auto = rows.map((row) => seedNameFromText(row.ann));

  let kopI = -1;
  let heelI = -1;
  let sotI = -1;
  let eotI = -1;
  let tangI = -1;
  for (let i = 0; i < n; i += 1) {
    if (auto[i] === "KOP" && kopI < 0) kopI = i;
    if (auto[i] === "HEEL" && heelI < 0) heelI = i;
    if (auto[i] === "SOT" && sotI < 0) sotI = i;
    if (auto[i] === "EOT") eotI = i;
    if (auto[i] === "TANGENT" && tangI < 0) tangI = i;
  }

  if (kopI < 0) {
    for (let i = 1; i < n; i += 1) {
      if (inc(i) > 1 && inc(i - 1) < 1 && md(i) > 100 && !hasText(ann(i), "Nudge")) {
        kopI = i;
        auto[i] = "KOP";
        break;
      }
      if (hasText(ann(i), "KOP")) {
        kopI = i;
        auto[i] = "KOP";
        break;
      }
    }
  }

  if (heelI < 0) {
    for (let i = 0; i < n; i += 1) {
      if (inc(i) >= 88) {
        heelI = i;
        auto[i] = "HEEL";
        break;
      }
    }
  }

  if (n > 0) {
    if (auto[n - 1] === "") auto[n - 1] = "TD";
    if (hasText(ann(n - 1), "Toe")) auto[n - 1] = "TD";
  }

  if (sotI < 0) {
    for (let i = kopI >= 0 ? kopI + 1 : 0; i < n; i += 1) {
      if (auto[i] === "HEEL" || auto[i] === "TD") break;
      if (Math.abs(trn(i)) > 0.5 || hasText(ann(i), "B & T") || hasText(ann(i), "B&T")) {
        sotI = i;
        auto[i] = "SOT";
        break;
      }
    }
  }

  if (eotI < 0 && sotI >= 0) {
    let lastTurn = sotI;
    const lim = heelI > sotI ? heelI - 1 : n - 1;
    for (let i = sotI + 1; i <= lim; i += 1) {
      if (Math.abs(trn(i)) > 0.5 || dls(i) > 0.3) lastTurn = i;
      if (auto[i] === "TANGENT" && i > sotI) eotI = i;
    }
    if (eotI < 0) {
      for (let i = lastTurn; i <= lim; i += 1) {
        if (hasText(ann(i), "Hold")) {
          eotI = i;
          break;
        }
      }
    }
    if (eotI < 0 && lastTurn > sotI) eotI = lastTurn;
    if (eotI >= 0 && auto[eotI] === "") auto[eotI] = "EOT";
  }

  // NUDGE arrival: first 3–20° station after a Nudge note, walking only from that note.
  for (let i = 0; i <= n - 2; i += 1) {
    if (!hasText(ann(i), "Nudge")) continue;
    for (let j = i + 1; j < n; j += 1) {
      if (auto[j] === "KOP" || auto[j] === "VERTICAL" || auto[j] === "TD") break;
      if (inc(j) >= 3 && inc(j) < 20) {
        if (auto[j] === "") auto[j] = "NUDGE";
        break;
      }
    }
  }

  if (tangI < 0 && kopI >= 0) {
    let tangLim = n - 1;
    if (sotI > kopI) tangLim = sotI - 1;
    if (heelI > kopI && heelI < tangLim) tangLim = heelI - 1;
    for (let i = kopI + 1; i <= tangLim; i += 1) {
      if (auto[i] !== "") continue;
      if (inc(i) > 15 && inc(i) < 85 && (dls(i) < 0.3 || hasText(ann(i), "Hold"))) {
        tangI = i;
        auto[i] = "TANGENT";
        break;
      }
    }
  }
  return auto;
}

// ---------------------------------------------------------------- MDL_PlanImport

const PLANNED_NUMBER = /-?\d{1,6}(?:,\d{3})*\.\d{2,3}/g;

function planNums(line: string): number[] {
  const out: number[] = [];
  for (const m of line.matchAll(PLANNED_NUMBER)) {
    out.push(Number(m[0].replaceAll(",", "")));
    if (out.length > 50) break;
  }
  return out;
}

/** ParsePlannedSurveys: the dense Planned Survey stations, one per text line. */
export function parsePlannedSurveys(text: string): PlannedStation[] {
  const out: PlannedStation[] = [];
  let lastMd = -1;
  let inSurvey = false;
  for (const line of text.split("\n")) {
    if (hasText(line, "Planned Survey")) {
      inSurvey = true;
      continue;
    }
    if (hasText(line, "Plan Sections") || hasText(line, "Plan Annotations") || hasText(line, "Formations") || hasText(line, "SECTION DETAILS")) {
      inSurvey = false;
      continue;
    }
    if (!inSurvey) continue;
    const nums = planNums(line);
    if (nums.length < 8) continue;
    if (nums[0]! + 0.001 < lastMd) continue;
    if (nums[1]! < -0.01 || nums[1]! > 180.01) continue;
    if (nums[3]! < -50 || nums[3]! > 12000) continue;
    if (out.length > MAX_PLANNED) break;
    out.push({ md: nums[0]!, inc: nums[1]!, azm: nums[2]!, tvd: nums[3]!, ns: 0, ew: 0 });
    lastMd = nums[0]!;
  }
  if (out.length > 0) {
    const ne = integrateLocalNE(out.map((s) => s.md), out.map((s) => s.inc), out.map((s) => s.azm));
    for (let i = 0; i < out.length; i += 1) {
      out[i]!.ns = ne.ns[i]!;
      out[i]!.ew = ne.ew[i]!;
    }
  }
  return out;
}

function nameAfterTwoNums(line: string): string {
  const s = vbTrim(line);
  const matches = [...s.matchAll(PLANNED_NUMBER)];
  if (matches.length < 2) return "";
  const second = matches[1]!;
  let name = collapseSpaces(vbTrim(s.slice(second.index + second[0].length)));
  if (hasText(name, "Lithology") || hasText(name, "Depth")) name = "";
  return vbTrim(name);
}

export function normFormName(name: string): string {
  let t = collapseSpaces(vbTrim(name).toLowerCase().replaceAll(".", "").replaceAll(",", "").replaceAll("-", " ").replaceAll("_", " "));
  t = t.replaceAll(" formation", "").replaceAll(" group", "").replaceAll(" member", "");
  if (t.endsWith(" fm")) t = t.slice(0, -3);
  t = vbTrim(t);
  if (t === "garminia" || t === "gaminie" || t === "gramina") t = "graminia";
  if (t === "wabamum") t = "wabamun";
  if (t === "sws" || t === "2ws" || t === "2nd white specks" || t === "second white speck") t = "second white specks";
  if (t === "bfs" || t === "fish scales" || t === "base fish scales" || t === "bfss") t = "base of fish scales";
  if (t === "cardium sandstone") t = "cardium ss";
  if (t === "viking sandstone") t = "viking ss";
  return t;
}

// AER Table of Formations 2019 and the AGS Atlas of the Western Canada Sedimentary Basin.
const LITHOLOGY: Record<string, string> = {
  "belly river": "ss / silt / mdst, coal",
  "lea park": "marine shale",
  colorado: "marine shale",
  cardium: "shale to ss / cgl",
  "cardium ss": "sandstone",
  "second white specks": "calc. marine shale",
  "base of fish scales": "shale / silt marker",
  viking: "ss / silt / shale",
  "viking ss": "ss / silt / shale",
  "joli fou": "marine shale",
  mannville: "ss / shale / coal",
  "upper mannville": "ss / shale / coal",
  "lower mannville": "ss / shale",
  ellerslie: "sandstone",
  glauconitic: "glauc. sandstone",
  "glauconitic ss": "glauc. sandstone",
  ostracod: "lst / shale",
  detrital: "ss / cgl / shale",
  "rock creek": "ss / siltstone",
  "poker chip": "black shale",
  nordegg: "phos. lst / shale",
  fernie: "marine shale",
  shunda: "lst / dolomite",
  pekisko: "limestone",
  banff: "lst / shale",
  exshaw: "black shale",
  wabamun: "lst / dolomite",
  graminia: "silt / dolomite",
  blueridge: "dolomite",
  "blue ridge": "dolomite",
  calmar: "siltstone",
  nisku: "dol / limestone",
  ireton: "calc. shale",
  duvernay: "org. shale / source",
  "cooking lake": "limestone",
  leduc: "reef dol / lst",
  waterways: "lst / shale",
  "slave point": "limestone",
  "fort simpson": "shale",
  muskeg: "anhydrite / dol",
  "keg river": "dol / limestone",
  "granite wash": "ss / cgl",
  wapiabi: "marine shale",
  kaskapau: "marine shale",
  dunvegan: "ss / shale",
  "spirit river": "ss / shale / coal",
  bluesky: "sandstone",
  gething: "ss / coal",
  cadomin: "conglomerate",
  nikanassin: "ss / shale",
  "charlie lake": "dol / anhydrite",
  halfway: "sandstone",
  doig: "silt / shale",
  montney: "silt / shale",
  belloy: "ss / dol",
  debolt: "lst / dolomite",
  elkton: "limestone",
  shale: "shale",
};

export function lithologyFor(name: string): string {
  const key = normFormName(name);
  const known = LITHOLOGY[key];
  if (known != null) return known;
  if (key.includes("ss") || key.includes("sand")) return "sandstone";
  if (key.includes("shale")) return "shale";
  return "";
}

/** ParseFormations: MD, TVD and the name after them, from the Formations table. */
export function parseFormations(text: string): Formation[] {
  const start = findText(text, "Formations");
  if (start < 0) return [];
  const region = regionAfter(text, start, 10, ["Plan Annotations", "Planning Report", "Planned Survey", "Plan Sections"]);
  const out: Formation[] = [];
  let lastMd = -1;
  for (const line of region.split("\n")) {
    const nums = planNums(line);
    if (nums.length < 2) continue;
    if (nums[0]! + 0.05 < lastMd) continue;
    if (nums[0]! < 1 || nums[1]! < -50) continue;
    const name = nameAfterTwoNums(line);
    if (name.length < 2) continue;
    if (out.length > MAX_FORMATIONS) break;
    out.push({ md: nums[0]!, tvd: nums[1]!, name, lithology: lithologyFor(name) });
    lastMd = nums[0]!;
  }
  return out;
}

// ---------------------------------------------------------------- sheet writes

function writePlanSecSheet(book: SheetBook, path: string, rows: PlanStation[], names: string[]): void {
  book.resetSheet(PLANSEC_SHEET);
  book.setValue(PLANSEC_SHEET, 1, 1, path);
  ["MD", "INC", "AZM", "TVD", "NS", "EW", "DLS", "BUILD", "TURN", "ANNOT", "AUTONAME", "USERNAME"].forEach((head, i) => book.setValue(PLANSEC_SHEET, 2, i + 1, head));
  rows.forEach((row, i) => {
    const r = i + 3;
    [row.md, row.inc, row.azm, row.tvd, row.ns, row.ew, row.dls, row.bld, row.trn].forEach((value, c) => book.setValue(PLANSEC_SHEET, r, c + 1, value));
    book.setValue(PLANSEC_SHEET, r, 10, row.ann === "" ? null : row.ann);
    book.setValue(PLANSEC_SHEET, r, 11, names[i] ? names[i]! : null);
  });
}

function writeSurveySheet(book: SheetBook, path: string, rows: PlannedStation[]): void {
  book.resetSheet(SURVEY_SHEET);
  book.setValue(SURVEY_SHEET, 1, 1, path);
  ["MD", "INC", "AZI", "TVD", "NS", "EW"].forEach((head, i) => book.setValue(SURVEY_SHEET, 2, i + 1, head));
  rows.forEach((row, i) => {
    [row.md, row.inc, row.azm, row.tvd, row.ns, row.ew].forEach((value, c) => book.setValue(SURVEY_SHEET, i + 3, c + 1, value));
  });
}

function writeFormationsSheet(book: SheetBook, path: string, rows: Formation[]): void {
  book.resetSheet(FORMATIONS_SHEET);
  book.setValue(FORMATIONS_SHEET, 1, 1, path);
  book.setValue(FORMATIONS_SHEET, 1, 2, LITHO_SRC);
  ["MD", "TVD", "Name", "Lithology"].forEach((head, i) => book.setValue(FORMATIONS_SHEET, 2, i + 1, head));
  rows.forEach((row, i) => {
    book.setValue(FORMATIONS_SHEET, i + 3, 1, row.md);
    book.setValue(FORMATIONS_SHEET, i + 3, 2, row.tvd);
    book.setValue(FORMATIONS_SHEET, i + 3, 3, row.name);
    book.setValue(FORMATIONS_SHEET, i + 3, 4, row.lithology === "" ? null : row.lithology);
  });
}

/** CsvParseLine: commas outside quotes split; "" inside quotes is one quote. */
export function csvParseLine(line: string): string[] {
  const out: string[] = [];
  let token = "";
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i]!;
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

/** ImportSurveyPlanCsv: the header row with an "MD" field, then every row whose first field is a number. */
export function importSurveyPlanCsv(book: SheetBook, path: string, csv: string): void {
  book.resetSheet(SURVEY_SHEET);
  book.setValue(SURVEY_SHEET, 1, 1, path);
  let colCount = 0;
  let outRow = 3;
  for (const raw of csv.split(/\r\n|\n|\r/)) {
    const line = vbTrim(raw);
    if (line === "") continue;
    const fields = csvParseLine(line);
    if (colCount === 0) {
      if (fields.some((field) => vbTrim(field).toLowerCase() === "md")) {
        colCount = fields.length;
        fields.forEach((field, c) => book.setValue(SURVEY_SHEET, 2, c + 1, vbTrim(field) === "" ? null : vbTrim(field)));
      }
      continue;
    }
    if (numeric(vbTrim(fields[0] ?? "")) == null) continue;
    fields.forEach((field, c) => {
      if (c >= colCount) return;
      const value = vbTrim(field);
      book.setValue(SURVEY_SHEET, outRow, c + 1, numeric(value) ?? (value === "" ? null : value));
    });
    outRow += 1;
  }
}

/** The VBA's Select Case over C1:T1 keeps the last column that matches any of the names. */
function columnPicker(cols: Map<string, number>): (...names: string[]) => number {
  return (...names) => Math.max(0, ...names.map((name) => cols.get(name) ?? 0));
}

function headerColumns(book: SheetBook, sheet: string, upper: boolean): Map<string, number> {
  const cols = new Map<string, number>();
  for (let c = 1; c <= 20; c += 1) {
    const head = vbTrim(cellText(book.value(sheet, 2, c)));
    const keyText = upper ? head.toUpperCase() : head.toLowerCase();
    if (keyText !== "") cols.set(keyText, c);
  }
  return cols;
}

/** BuildPlanSectionsFromSurveyComments: a CSV plan's commented stations (and its last row) become the sections. */
export function buildPlanSectionsFromSurveyComments(book: SheetBook): void {
  if (!book.hasSheet(SURVEY_SHEET)) return;
  const col = columnPicker(headerColumns(book, SURVEY_SHEET, false));
  const cmd = col("md");
  if (cmd === 0) return;
  const cInc = col("inc");
  const cAzi = col("azm", "azi");
  const cTvd = col("tvd");
  const cNs = col("ns");
  const cEw = col("ew");
  const cDls = col("dls");
  const cCom = col("annotation", "annot", "comment");
  const lastR = book.lastUsedRow(SURVEY_SHEET, cmd);
  if (lastR < 3) return;
  const rows: PlanStation[] = [];
  const num = (r: number, c: number) => (c > 0 ? vbaVal(book.value(SURVEY_SHEET, r, c)) : 0);
  for (let r = 3; r <= lastR; r += 1) {
    const mdValue = numeric(book.value(SURVEY_SHEET, r, cmd));
    if (mdValue == null) continue;
    const ann = cCom > 0 ? vbTrim(cellText(book.value(SURVEY_SHEET, r, cCom))) : "";
    const takeIt = r === lastR || (ann !== "" && (seedNameFromText(ann) !== "" || hasText(ann, "KOP") || hasText(ann, "Hold") || hasText(ann, "B & T") || hasText(ann, "Build")));
    if (!takeIt) continue;
    if (rows.length > MAX_SECTIONS) break;
    rows.push(station(mdValue, num(r, cInc), num(r, cAzi), num(r, cTvd), num(r, cNs), num(r, cEw), num(r, cDls), 0, 0, ann));
  }
  if (rows.length < 1) return;
  writePlanSecSheet(book, cellText(book.value(SURVEY_SHEET, 1, 1)), rows, namePlanSectionTargets(rows));
}

/** RefreshTdPlannedFromPlan: Data I54:M54 = the plan's last station. */
export function refreshTdPlannedFromPlan(book: SheetBook): void {
  if (!book.hasSheet(SURVEY_SHEET) || !book.hasSheet(DATA_SHEET)) return;
  const col = columnPicker(headerColumns(book, SURVEY_SHEET, true));
  const cmd = col("MD");
  const cInc = col("INC");
  const cAzi = col("AZI", "AZM");
  const cNs = col("NS");
  const cEw = col("EW");
  if (!cmd || !cInc || !cAzi || !cNs || !cEw) return;
  let lastR = book.lastUsedRow(SURVEY_SHEET, cmd);
  while (lastR >= 3 && numeric(book.value(SURVEY_SHEET, lastR, cmd)) == null) lastR -= 1;
  if (lastR < 3) return;
  [cmd, cInc, cAzi, cNs, cEw].forEach((c, i) => book.setValue(DATA_SHEET, TD_ROW_PLANNED, TD_COL_FIRST + i, numeric(book.value(SURVEY_SHEET, lastR, c))));
}

/** UpdateImportPathDisplay: the path cell right of the button host merge at OC_ImpPlanRow / OC_ImpAcRow. */
export function showImportPath(book: SheetBook, anchor: { row: number; col: number } | null, hiddenSheet: string): void {
  if (!anchor || !book.hasSheet(SETUP_SHEET)) return;
  const host = book.mergeAt(SETUP_SHEET, anchor.row, anchor.col);
  const pathArea = book.mergeAt(SETUP_SHEET, anchor.row, host.col2 + 1);
  const path = book.hasSheet(hiddenSheet) ? vbTrim(cellText(book.value(hiddenSheet, 1, 1))) : "";
  book.setValue(SETUP_SHEET, pathArea.row, pathArea.col, path === "" ? null : path);
}

// ---------------------------------------------------------------- Import Plan

export interface PlanFile {
  /** Shown on Setup and kept in A1 of the plan sheets, as the VBA keeps the file path. */
  path: string;
  kind: "pdf" | "csv";
  /** PDF text (layout order) or the CSV contents. */
  text: string;
  /** The same-name .csv beside a PDF, read when the PDF has fewer than 50 planned stations. */
  siblingCsv?: () => Promise<string | null>;
}

export interface PlanImportResult {
  ok: boolean;
  /** MsgBox text (the import stops). */
  alert?: string;
  /** Application.StatusBar text. */
  status: string;
}

/** ImportSurveyPlanPdf. False leaves the previous plan as it was. */
async function importPlanPdf(book: SheetBook, file: PlanFile): Promise<PlanImportResult> {
  if (file.text.trim() === "") return { ok: false, alert: NO_PDF_TEXT, status: "Plan import failed — previous plan left unchanged" };
  const sections = parsePlanSections(file.text);
  if (sections.length < 1) return { ok: false, alert: unreadablePdfText(file.text) ? NO_LETTER_MAP : NO_PLAN_TABLE, status: "Plan import failed — previous plan left unchanged" };
  writePlanSecSheet(book, file.path, sections, namePlanSectionTargets(sections));
  const planned = parsePlannedSurveys(file.text);
  if (planned.length >= 50) writeSurveySheet(book, file.path, planned);
  else {
    const csv = file.siblingCsv ? await file.siblingCsv() : null;
    if (csv != null) {
      importSurveyPlanCsv(book, file.path, csv);
      book.setValue(SURVEY_SHEET, 1, 1, file.path);
    } else writeSurveySheet(book, file.path, sections);
  }
  const formations = parseFormations(file.text);
  if (formations.length >= 1) writeFormationsSheet(book, file.path, formations);
  return {
    ok: true,
    status: `Plan: ${sections.length} sections, ${planned.length >= 50 ? `${planned.length} planned surveys` : "section stations"}, ${formations.length} formations`,
  };
}

/** ImportSurveyPlanFile up to RenderPlanGauge, which the host redraws from the new plan. */
export async function importSurveyPlan(book: SheetBook, file: PlanFile, planRowAnchor: { row: number; col: number } | null): Promise<PlanImportResult> {
  if (file.kind === "pdf") {
    const result = await importPlanPdf(book, file);
    if (!result.ok) return result;
    writePlanIdentity(book, parsePlanHeader(file.text));
  } else {
    importSurveyPlanCsv(book, file.path, file.text);
    buildPlanSectionsFromSurveyComments(book);
    clearPlanIdentity(book);
  }
  book.clear(SS_SHEET, "T2:Y5");
  showImportPath(book, planRowAnchor, SURVEY_SHEET);
  syncPlanTargetWindow(book);
  refreshTdPlannedFromPlan(book);
  const name = file.path.split(/[\\/]/).pop() ?? file.path;
  return { ok: true, status: `Plan imported from ${name}` };
}
