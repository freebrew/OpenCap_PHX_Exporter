import type { MacroButton } from "@phx/xlsx-bridge";
import { renderCorridor } from "../../../packages/corridor-3d/src/index";
import { morningReport as composeMorningReport, morningSummaryMini, type MorningFields } from "../../../packages/reports/src/index";
import {
  appendTallyBitDepths,
  bhaNumbers,
  bhaSummary,
  CLEAR_SUMMARY,
  clearLog,
  clearSlidesheetRanges,
  clearTallyInputs,
  contDiTargetHit,
  DLT_FIRST,
  DLT_LAST,
  highlightActiveTarget,
  importAntiCollision,
  importNote,
  followSlideTarget,
  importSurveyPlan,
  jobField,
  lastSlidesheetBitMd,
  readJobFields,
  lengthsFromCsv,
  loadNamedTargets,
  missingTallySheet,
  PT_SHEET,
  prepareSlideComments,
  pullLastSurvey,
  readNumber,
  refreshRotateMetres,
  serialHours,
  writeSerialHoursTable,
  refreshSlideComments,
  startEntry,
  TALLY_LAYOUT_HELP,
  trackerSheet,
  writeLengths,
  type DlEntryStart,
  type MeasureText,
  type NamedTarget,
  SheetBook,
  type SnapshotBook,
} from "../../../packages/workbook-commands/src/index";
import type { WorkbookHost } from "../workbook/facade";
import { sendDailyReport } from "./mail/sendDaily";
import { pdfLayoutText } from "../workbook/pdfText";
import { pickFile, type FolderEntry, type PickedFile as ImportPick } from "./pickFile";

export type ToolSection = "costs" | "report";

/** What a macro needs from the app: a fresh workbook view after each recalculation, and its forms. */
export interface MacroContext {
  host: WorkbookHost;
  /** The workbook as it is now, with the sheet protection the import recorded. */
  book: () => SheetBook | null;
  notify: (message: string) => void;
  openTool: (section: ToolSection) => void;
  openDlEntry: (start: DlEntryStart, book: SheetBook) => void;
  openTargets: (targets: NamedTarget[]) => void;
  /** frmPipeTally, with the note the CSV import left for it. */
  openPipeTally: (note: string) => void;
  /** A workbook defined name that points at one cell ("Setup!$J$29"), 1-based. */
  namedCell: (name: string) => { row: number; col: number } | null;
  /** The granted job folder, so file pickers open there. */
  jobFolder: () => Promise<FolderEntry | null>;
  /** RenderPlanGauge on the recalculated workbook. */
  redrawGauge: () => void;
  /** RefreshAllCsvData from the job folder's OpenCap CSVs. */
  refreshOpenCap: () => Promise<void>;
  /** RenderPlanGauge + the slide comments after the selection leaves the Inc/Azm cells. */
  redrawAfterSurvey: () => Promise<void>;
}

const ATTACHMENT_ROWS: Record<string, number> = {
  PickAttachmentFile_Y14: 29,
  PickAttachmentFile_Y15: 30,
  PickAttachmentFile_Y16: 31,
  PickAttachmentFile_Y17: 32,
  PickAttachmentFile_Y18: 33,
};

const DLT_CLEAR_PROMPT = `Clear every downlink entry in rows ${DLT_FIRST}-${DLT_LAST}?\n\nClient / Rig / Well / Job # and all formulas stay.\nThis cannot be undone.`;

let measureCanvas: HTMLCanvasElement | null = null;

/** Text width in px for the cell font, so padded slide comments fit column Y as in Excel. */
export const measureText: MeasureText = (text, font) => {
  measureCanvas ??= document.createElement("canvas");
  const context = measureCanvas.getContext("2d");
  if (!context) return text.length * font.sizePt * 0.5 * (96 / 72);
  context.font = `${font.bold ? "bold " : ""}${font.sizePt}pt "${font.family}"`;
  return context.measureText(text).width;
};

/** Apply a macro step's writes, then let the formulas catch up before the next step reads them. */
async function commit(context: MacroContext, book: SheetBook): Promise<SheetBook | null> {
  if (book.writes.length === 0) return book;
  context.host.applyWrites(book.writes);
  await context.host.flushCalculation();
  return context.book();
}

export async function runMacro(button: MacroButton, context: MacroContext): Promise<void> {
  const { host, notify } = context;
  const macro = button.macro;
  const attachmentRow = ATTACHMENT_ROWS[macro];
  if (attachmentRow) {
    await pickAttachment(host, attachmentRow, notify);
    return;
  }
  switch (macro) {
    case "ClearSlidesheetRanges":
      await clearSlidesheet(context);
      return;
    case "ShowTargetEditor":
      showTargetEditor(context);
      return;
    case "ShowPipeTallyForm":
      await showPipeTallyForm(context);
      return;
    case "cleartally": {
      const book = context.book();
      if (!book) return;
      clearTallyInputs(book);
      host.applyWrites(book.writes);
      notify(`Cleared ${PT_SHEET} F7:F315.`);
      return;
    }
    case "AppendTallyBitDepths":
      await appendTally(context);
      return;
    case "DLT_AddEntry": {
      const book = context.book();
      if (!book) return;
      const started = startEntry(book);
      if (!started.ok) {
        window.alert(started.message);
        return;
      }
      context.openDlEntry(started.start, book);
      return;
    }
    case "DLT_ClearEntries":
      clearTracker(context);
      return;
    case "CLEAR_TIMETABLE":
      if (!window.confirm("Clear the Data day-row inputs H4:J22, L4:N22, and P4:Q22?")) return;
      host.clearRange("Data", "H4:J22");
      host.clearRange("Data", "L4:N22");
      host.clearRange("Data", "P4:Q22");
      notify("Cleared the day-row inputs on Data.");
      return;
    case "ImportSurveyPlan":
      await importPlan(context);
      return;
    case "ImportAntiCollision":
      await importAc(context);
      return;
    case "RefreshSetup":
    case "RefreshAllCsvData":
    case "RefreshData":
      // Data REFRESH is the same OpenCap load, then the serial-hours table the dashboard shows.
      await context.refreshOpenCap();
      writeSerialHours(context);
      return;
    case "SelectBHA":
      selectBha(context, Number((button.caption.match(/\d+/) ?? [button.id.match(/\d+/) ?? []])[0]) || Number(button.id.match(/\d+/)?.[0]) || 1);
      return;
    case "ContDI_PullLastSurvey":
      pullContDiSurvey(context);
      return;
    case "saveaspdf":
      // Data's PDF is the sheet as printed. The browser's print dialog saves it.
      window.print();
      context.notify("Print the Data sheet to PDF from the dialog.");
      return;
    case "ShowCostsForm":
      context.openTool("costs");
      notify("Cost row is open.");
      return;
    case "GenerateMorningReport":
    case "GenerateMorningSummaryMini":
    case "RTP_printW_Email":
      await morningReport(macro, context);
      return;
    default:
      notify(`${button.caption} is on ${button.sheet}. ${macro} still runs from the Excel workbook.`);
  }
}

/** ClearSlidesheetRanges. The browser asks first; Excel clears straight away. */
async function clearSlidesheet(context: MacroContext): Promise<void> {
  if (!window.confirm("Clear the Slidesheet entry ranges (D/F/G/T/U, sail waypoints, targets U2:X5, comments Y/Z)?")) return;
  let book = context.book();
  if (!book) return;
  clearSlidesheetRanges(book);
  book = await commit(context, book);
  if (!book) return;
  // RefreshSlideComments(forceAll) repairs the projection formulas first, on a recalculated sheet.
  if (prepareSlideComments(book)) book = await commit(context, book);
  if (!book) return;
  refreshSlideComments(book, true, measureText);
  highlightActiveTarget(book);
  await commit(context, book);
  window.alert(CLEAR_SUMMARY);
  context.notify("Slidesheet ranges cleared. Auto slide comments were rebuilt where inputs exist.");
}

function showTargetEditor(context: MacroContext): void {
  const book = context.book();
  if (!book) return;
  context.openTargets(loadNamedTargets(book, 40));
}

/** ShowPipeTallyForm: pick the rig CSV (Cancel skips it), then the form. */
async function showPipeTallyForm(context: MacroContext): Promise<void> {
  const book = context.book();
  if (!book) return;
  const missing = missingTallySheet(book);
  if (missing) {
    window.alert(`Pipe Tally Calculator cannot open.\n\n${missing}\n\n${TALLY_LAYOUT_HELP}`);
    return;
  }
  const note = await importTallyCsv(context);
  context.openPipeTally(note);
}

/** ImportPipeTallyCsv. Returns the note the form shows ("" when the pick was cancelled). */
export async function importTallyCsv(context: MacroContext): Promise<string> {
  const picked = await pickFile("phx-import-tally", [{ description: "Pipe tally CSV (*.csv)", accept: { "text/csv": [".csv"] } }], null);
  if (!picked) return "";
  const text = await picked.file.text();
  const book = context.book();
  if (!book) return "";
  const loaded = lengthsFromCsv(text, lastSlidesheetBitMd(book));
  if (!loaded.ok) {
    window.alert(loaded.error);
    return loaded.error;
  }
  writeLengths(book, loaded.lengths);
  await commit(context, book);
  return importNote(loaded);
}

/** AppendTallyBitDepths, then the rotate metres and slide comments for the new D rows. */
export async function appendTally(context: MacroContext): Promise<boolean> {
  let book = context.book();
  if (!book) return false;
  const result = appendTallyBitDepths(book);
  if (!result.ok) {
    window.alert(result.message);
    return false;
  }
  book = await commit(context, book);
  if (book && result.added) {
    refreshRotateMetres(book, { row: result.added.row, col: 4, row2: result.added.row2, col2: 4 });
    book = await commit(context, book);
    if (book && prepareSlideComments(book)) book = await commit(context, book);
    if (book) {
      refreshSlideComments(book, false, measureText);
      await commit(context, book);
    }
  }
  if (result.warning) window.alert(result.warning);
  window.alert(result.message);
  context.notify(result.message);
  return true;
}

/** The PDF's text, or the reason PDF.js could not read it. */
async function readPdf(picked: ImportPick): Promise<{ text: string; error: string | null }> {
  try {
    return { text: await pdfLayoutText(new Uint8Array(await picked.file.arrayBuffer())), error: null };
  } catch (error) {
    return { text: "", error: error instanceof Error ? error.message : String(error) };
  }
}

/** ImportSurveyPlan: a Compass plan PDF (Plan Sections) or a survey CSV. */
async function importPlan(context: MacroContext): Promise<void> {
  const picked = await pickFile("phx-import-plan", [
    { description: "Well Plan PDF", accept: { "application/pdf": [".pdf"] } },
    { description: "CSV Files", accept: { "text/csv": [".csv"] } },
  ], context.jobFolder());
  if (!picked) return;
  context.notify("Importing survey plan...");
  try {
    const name = picked.file.name;
    const kind = name.toLowerCase().endsWith(".pdf") ? "pdf" : "csv";
    const read = kind === "pdf" ? await readPdf(picked) : { text: await picked.file.text(), error: null };
    if (read.error) {
      window.alert(`Could not read that PDF.\n\n${read.error}`);
      context.notify("Plan import failed — previous plan left unchanged");
      return;
    }
    const base = name.includes(".") ? name.slice(0, name.lastIndexOf(".")) : name;
    const book = context.book();
    if (!book) return;
    const result = await importSurveyPlan(book, {
      path: picked.path,
      kind,
      text: read.text,
      siblingCsv: async () => (await picked.sibling(`${base}.csv`))?.text() ?? null,
    }, context.namedCell("OC_ImpPlanRow"));
    if (!result.ok) {
      if (result.alert) window.alert(result.alert);
      context.notify(result.status);
      return;
    }
    await commit(context, book);
    context.redrawGauge();
    context.notify(result.status);
  } catch (error) {
    context.notify(`Survey import failed: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/** ImportAntiCollision: the Compass Anticollision Report Summary. */
async function importAc(context: MacroContext): Promise<void> {
  const picked = await pickFile("phx-import-ac", [{ description: "PDF Files", accept: { "application/pdf": [".pdf"] } }], context.jobFolder());
  if (!picked) return;
  context.notify("Reading anti-collision report...");
  try {
    const read = await readPdf(picked);
    const book = context.book();
    if (!book) return;
    const result = importAntiCollision(book, picked.path, read.text, context.namedCell("OC_ImpAcRow"));
    await commit(context, book);
    if (!result.ok) {
      window.alert(read.error ? `Could not read that PDF.\n\n${read.error}` : (result.alert ?? result.status));
      context.notify(result.status);
      return;
    }
    if (result.note) window.alert(result.note);
    context.notify(result.status);
  } catch (error) {
    context.notify(`AC import failed: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/** BuildDataSerialHoursTable, after the OpenCap load has filled _OC_BHA and _OC_Inventory. */
function writeSerialHours(context: MacroContext): void {
  const book = context.book();
  if (!book) return;
  const count = writeSerialHoursTable(book, serialHours(book));
  context.host.applyWrites(book.writes);
  context.notify(`Serial hours: ${count} DD tool(s) on Data.`);
}

/** SelectBHA: the BHA the button names, with its components and the serial-hours table beside it. */
function selectBha(context: MacroContext, bha: number): void {
  const book = context.book();
  if (!book) return;
  const summary = bhaSummary(book, bha);
  if (!summary) {
    const known = bhaNumbers(book);
    window.alert(known.length === 0
      ? "No BHA data yet. Pull from FieldCap or run REFRESH CSVs first."
      : `BHA ${bha} is not in this job. The sheet has BHA ${known.join(", ")}.`);
    return;
  }
  writeSerialHoursTable(book, serialHours(book));
  context.host.applyWrites(book.writes);
  const lines = [
    `BHA ${summary.bha}  ${summary.section}  ${summary.status}  ${summary.motor}`,
    `${summary.metres} m   ${summary.hrsTotal.toFixed(2)} h total   ${summary.hrsSlid.toFixed(2)} h slid   ${(summary.slideFraction * 100).toFixed(0)}% slide`,
    "",
    ...summary.components.map((row) => `${row.index}  ${row.serial}  ${row.itemCode}  ${row.description}`),
  ];
  context.notify(`BHA ${summary.bha}: ${summary.components.length} component(s), ${summary.metres} m.`);
  window.alert(lines.join("\n"));
}

/** ContDI_PullLastSurvey: the newest good survey becomes the tie-on, and the last typed row stays. */
function pullContDiSurvey(context: MacroContext): void {
  const book = context.book();
  if (!book) return;
  const result = pullLastSurvey(book);
  if (!result) {
    window.alert("The 'Cont DI' tab is not in this workbook.");
    return;
  }
  if (result.surveyRow === 0) {
    window.alert("No Slidesheet row has Inc and Azm in the green survey fill.");
    return;
  }
  context.host.applyWrites(book.writes);
  context.notify(`Cont DI tie-on is Slidesheet row ${result.surveyRow}.`);
}

/**
 * ContDI_ApplyTargetSelection: an edit inside the yellow target block (P2, S1:S4) rewrites the
 * aim formulas so the tab follows the Slidesheet's active station.
 */
export function contDiEdited(book: SheetBook, range: { row: number; col: number; row2: number; col2: number }): boolean {
  let hit = false;
  for (let row = range.row; row <= range.row2 && !hit; row += 1) {
    for (let col = range.col; col <= range.col2; col += 1) if (contDiTargetHit(row, col)) hit = true;
  }
  return hit && followSlideTarget(book);
}

/** DLT_ClearEntries. */
function clearTracker(context: MacroContext): void {
  const book = context.book();
  if (!book) return;
  const sheet = trackerSheet(book);
  if (!sheet) {
    window.alert("The 'DL Tracker' tab is not in this workbook.");
    return;
  }
  if (!window.confirm(DLT_CLEAR_PROMPT)) return;
  const count = clearLog(book, sheet);
  context.host.applyWrites(book.writes);
  window.alert(`${count} entries cleared.`);
  context.notify(`${count} downlink entries cleared. Formulas and the header stayed.`);
}

const REPORT_LABELS = [
  "00:00 Depth", "Start Depth", "24:00 Depth", "Midnight Depth", "Meters Drilled",
  "Sliding Percentage", "Sliding Values", "Rotating Percentage", "Rotating Values",
  "Motor S/N", "Motor SN", "Configuration", "Config", "Rev/L", "Rev / L", "RevL", "Pump Rate", "Rotary", "Total RPM",
  "Today Bit Hours", "Daily Bit Hours", "Total Bit Hours", "Today Circ Hours", "Daily Circ Hours", "Total Circ Hours", "Total Motor Hours",
  "Average Motor Output", "Motor Output", "Dogleg Needed", "Down Hole Temp", "Downhole Temp", "TVD @ Midnight", "TVD at Midnight", "TVD:",
  "R/L from Plan P1", "R/L from Plan P3", "Above/Below from Plan P1", "Above/Below from Plan P3",
  "Distance from Current Geo Target", "Distance from Current GeoTarget", "Distance from Target",
];

/** The Data cells and labels Module21 reads, taken off the recalculated snapshot. */
function morningFields(context: MacroContext): MorningFields {
  const snapshot = context.host.snapshot() as SnapshotBook | null;
  const book = context.book();
  const cells: Record<string, unknown> = {};
  const read = (address: string) => value(snapshot, address);
  for (const address of ["E7", "E8", "E9", "E13", "E14", "E15", "F13", "F14", "H3", "I23", "J23", "K23", "L23", "P23", "P24", "Q23", "Q24"]) {
    cells[address] = read(address);
  }
  // A label's value sits in the cell to its right; the direction (offset 2) one past that.
  for (const row of Object.values(snapshot?.sheets?.["Data"]?.cellData ?? {})) {
    if (!row) continue;
    for (const [colKey, cell] of Object.entries(row)) {
      const label = REPORT_LABELS.find((candidate) => String(cell?.v ?? "").trim().toLowerCase() === candidate.toLowerCase());
      if (!label) continue;
      const col = Number(colKey);
      cells[label] ??= row[String(col + 1)]?.v ?? "";
      cells[`${label} dir`] ??= row[String(col + 2)]?.v ?? "";
    }
  }
  const today = new Date();
  const reportDate = today.toLocaleDateString("en-US", { month: "long", day: "2-digit", year: "numeric" });
  return { wellNumber: book ? jobField(readJobFields(book), "Job Code", "Job ID") : "", reportDate, cells };
}

/** Module21 copies the prose report. EMAIL uses the Data table and does not send. */
async function morningReport(macro: string, context: MacroContext): Promise<void> {
  context.openTool("report");
  if (macro === "RTP_printW_Email") {
    await emailReport(context);
    return;
  }
  const fields = morningFields(context);
  const report = macro === "GenerateMorningSummaryMini" ? morningSummaryMini(fields) : composeMorningReport(fields);
  try {
    await navigator.clipboard.writeText(report.text);
  } catch {
    // The clipboard needs a gesture the button already is; a failure still leaves the text in Well tools.
  }
  context.notify(macro === "GenerateMorningSummaryMini" ? "Mini report copied." : "Morning report copied.");
}

/** RTP_printW_Email. One builder for the sheet button and Well tools. The workbook does not send. */
export async function emailReport(context: MacroContext): Promise<void> {
  const snapshot = context.host.snapshot() as SnapshotBook | null;
  await sendDailyReport({
    snapshot,
    book: context.book(),
    pngBase64: await corridorPng(snapshot),
    jobFolder: () => context.jobFolder(),
    notify: context.notify,
    openChooser: () => context.openTool("report"),
  });
}

/** RenderCorridorPng for the mail body as base64. Blank when the day has no picture to draw. */
async function corridorPng(snapshot: SnapshotBook | null): Promise<string | undefined> {
  if (!snapshot) return undefined;
  const canvas = document.createElement("canvas");
  if (!renderCorridor(canvas, new SheetBook(snapshot, []))) return undefined;
  const url = canvas.toDataURL("image/png");
  return url.slice(url.indexOf(",") + 1);
}

async function pickAttachment(host: WorkbookHost, row: number, notify: (message: string) => void): Promise<void> {
  const picker = (window as Window & { showOpenFilePicker?: () => Promise<Array<{ name: string }>> }).showOpenFilePicker;
  if (!picker) {
    notify("This Chrome window cannot pick an attachment file.");
    return;
  }
  const [file] = await picker();
  if (!file) return;
  host.setCell("Data", `I${row}`, file.name);
  notify(`Attachment ${row - 28} is ${file.name}.`);
}

function value(book: SnapshotBook | null, a1: string): string {
  const number = book ? readNumber(book, "Data", a1) : null;
  return number == null ? "" : String(number);
}
