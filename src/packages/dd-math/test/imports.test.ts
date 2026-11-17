import { describe, expect, it } from "vitest";
import {
  AC_SHEET,
  buildAcTable,
  csvParseLine,
  FORMATIONS_SHEET,
  importAntiCollision,
  importSurveyPlan,
  layoutText,
  lithologyFor,
  namePlanSectionTargets,
  parseAcSummary,
  parseFormations,
  parsePlannedSurveys,
  parsePlanSections,
  SheetBook,
  sortAcRows,
  SURVEY_SHEET,
  writeAcConcerns,
  type CellWrite,
  type PdfTextItem,
} from "../../workbook-commands/src/index";

function makeBook(sheets: Record<string, Record<number, Record<number, unknown>>>): SheetBook {
  const out: Record<string, { id: string; name: string; cellData: Record<number, Record<number, unknown>> }> = {};
  Object.entries(sheets).forEach(([name, cellData], index) => {
    out[`s${index}`] = { id: `s${index}`, name, cellData };
  });
  return new SheetBook({ sheets: out, styles: {} } as never, []);
}

function valueAt(writes: CellWrite[], sheet: string, row: number, col: number): unknown {
  const hit = [...writes].reverse().find((write) => write.kind === "value" && write.sheet === sheet && write.row === row && write.col === col);
  return hit && hit.kind === "value" ? hit.value : undefined;
}

// Compass "Plan Sections" as pdftotext -layout prints it: MD Inc Azi TVD NS EW DLS BR TR TFO Target.
const PLAN_SECTIONS = [
  "Plan Sections",
  "   Measured            Vertical                     Dogleg   Build    Turn",
  "    Depth  Inclination Azimuth   Depth  +N/-S  +E/-W  Rate  Rate  Rate  TFO  Target",
  "     0.00     0.00     0.00      0.00    0.00    0.00  0.000   0.000   0.000    0.00",
  "   600.00     0.00     0.00    600.00    0.00    0.00  0.000   0.000   0.000    0.00",
  "  1200.00    40.00    90.00   1150.00    0.00  200.00  2.000   2.000   0.000   90.00",
  "  1600.00    40.00    90.00   1460.00    0.00  460.00  0.000   0.000   0.000    0.00",
  "  2100.00    90.00   120.00   1600.00 -120.00  800.00  3.000   2.500   1.500   45.00",
  "  3500.00    90.00   120.00   1600.00 -820.00 2010.00  0.000   0.000   0.000    0.00  Toe",
  "Planned Survey",
].join("\n");

function plannedRows(count: number): string {
  const rows = ["Planned Survey"];
  for (let i = 0; i < count; i += 1) {
    const md = (i * 30).toFixed(2);
    rows.push(`  ${md}   0.00   0.00   ${md}   0.00   0.00   0.00   0.00   0.00`);
  }
  return rows.join("\n");
}

const FORMATIONS = [
  "Formations",
  "    Measured    Vertical",
  "     Depth       Depth     Name                Lithology",
  "    850.00      850.00     Belly River",
  "   1210.00     1158.00     2WS",
  "   1990.00     1590.00     Montney Fm",
  "Plan Annotations",
].join("\n");

describe("Import Plan (MDL_Setup ImportSurveyPlan)", () => {
  it("reads the Plan Sections table and names the targets from the shape", () => {
    const rows = parsePlanSections(PLAN_SECTIONS);
    expect(rows.map((row) => row.md)).toEqual([0, 600, 1200, 1600, 2100, 3500]);
    expect(rows[4]).toMatchObject({ inc: 90, azm: 120, tvd: 1600, ns: -120, ew: 800, dls: 3, bld: 2.5, trn: 1.5 });
    expect(namePlanSectionTargets(rows)).toEqual(["", "", "KOP", "TANGENT", "HEEL", "TD"]);
  });

  it("keeps planned stations in order and maps formations to lithology", () => {
    const planned = parsePlannedSurveys(plannedRows(60));
    expect(planned).toHaveLength(60);
    expect(planned[59]).toMatchObject({ md: 1770, tvd: 1770, ns: 0, ew: 0 });
    const tops = parseFormations(FORMATIONS);
    expect(tops.map((top) => [top.md, top.name])).toEqual([[850, "Belly River"], [1210, "2WS"], [1990, "Montney Fm"]]);
    expect(lithologyFor("2WS")).toBe("calc. marine shale");
    expect(lithologyFor("Montney Fm")).toBe("silt / shale");
  });

  it("rebuilds N/E from the angles when Crystal glues Subsea into EW", () => {
    const glued = [
      "Planned Survey",
      "0.00 0.00 0.00 0.00 0.00 0.000891.820.00 0.00 0.000 0.000",
      "30.00 0.00 0.00 30.00 0.00 0.000861.820.00 0.00 0.000 0.000",
      "100.00 0.00 0.00 100.00 0.00 0.000791.820.00 0.00 0.000 0.000",
      "200.00 30.00 90.00 196.59 0.00 0.000691.2325.88 0.00 0.000 0.000",
    ].join("\n");
    const planned = parsePlannedSurveys(glued);
    expect(planned).toHaveLength(4);
    expect(planned[0]).toMatchObject({ md: 0, inc: 0, ns: 0, ew: 0 });
    expect(planned[1]!.ew).toBeCloseTo(0, 5);
    expect(planned[1]!.ns).toBeCloseTo(0, 5);
    expect(planned[3]!.ew).toBeGreaterThan(20);
    expect(planned[3]!.ew).toBeLessThan(40);
    expect(Math.abs(planned[3]!.ns)).toBeLessThan(0.05);
  });

  it("splits CSV fields the way CsvParseLine does", () => {
    expect(csvParseLine("1,\"KOP, 4\"\"\",x")).toEqual(["1", "KOP, 4\"", "x"]);
  });

  it("writes the plan sheets, the path beside the button and the TD planned row", async () => {
    const book = makeBook({ Setup: {}, Data: {}, Slidesheet: { 1: { 19: { v: "old target" } } } });
    const text = `${PLAN_SECTIONS}\n${plannedRows(60).replace("Planned Survey\n", "")}\n${FORMATIONS}`;
    const result = await importSurveyPlan(book, { path: "Job\\plan.pdf", kind: "pdf", text }, { row: 29, col: 10 });
    expect(result).toEqual({ ok: true, status: "Plan imported from plan.pdf" });
    expect(book.hasSheet(SURVEY_SHEET) && book.hasSheet("_OC_PlanSec") && book.hasSheet(FORMATIONS_SHEET)).toBe(true);
    expect(valueAt(book.writes, "_OC_PlanSec", 7, 11)).toBe("HEEL");
    expect(valueAt(book.writes, SURVEY_SHEET, 62, 1)).toBe(1770);
    expect(valueAt(book.writes, FORMATIONS_SHEET, 4, 4)).toBe("calc. marine shale");
    expect(valueAt(book.writes, "Setup", 29, 11)).toBe("Job\\plan.pdf");
    expect(valueAt(book.writes, "Data", 54, 9)).toBe(1770);
    expect(book.value("Slidesheet", 2, 20)).not.toBe("old target");
  });

  it("uses the CSV beside a PDF with too few planned stations", async () => {
    const book = makeBook({ Setup: {}, Data: {}, Slidesheet: {} });
    const csv = "MD,Inc,Azm,TVD,NS,EW,Comment\n0,0,0,0,0,0,\n2500,90,120,1600,-500,1500,Toe\n";
    const result = await importSurveyPlan(book, { path: "Job\\plan.pdf", kind: "pdf", text: PLAN_SECTIONS, siblingCsv: async () => csv }, null);
    expect(result.ok).toBe(true);
    expect(book.value(SURVEY_SHEET, 1, 1)).toBe("Job\\plan.pdf");
    expect(book.value(SURVEY_SHEET, 4, 1)).toBe(2500);
  });

  it("leaves the previous plan alone when the PDF has no table or no text", async () => {
    const book = makeBook({ Setup: {}, Slidesheet: {} });
    const scanned = await importSurveyPlan(book, { path: "x.pdf", kind: "pdf", text: "  " }, null);
    expect(scanned).toMatchObject({ ok: false, status: "Plan import failed — previous plan left unchanged" });
    expect(scanned.alert).toContain("scanned");
    const noTable = await importSurveyPlan(book, { path: "x.pdf", kind: "pdf", text: "Planning Report\nnothing here" }, null);
    expect(noTable.ok).toBe(false);
    expect(book.writes).toEqual([]);
  });

  it("builds plan sections from a CSV plan's commented stations", async () => {
    const book = makeBook({ Setup: {}, Data: {}, Slidesheet: {} });
    const csv = "MD,Inc,Azm,TVD,NS,EW,DLS,Comment\n0,0,0,0,0,0,0,\n600,0,0,600,0,0,0,KOP\n900,30,90,880,0,75,3,\n2500,90,120,1600,-500,1500,0,Toe\n";
    await importSurveyPlan(book, { path: "plan.csv", kind: "csv", text: csv }, null);
    expect(book.value("_OC_PlanSec", 3, 1)).toBe(600);
    expect(book.value("_OC_PlanSec", 3, 11)).toBe("KOP");
    expect(book.value("_OC_PlanSec", 4, 11)).toBe("TD");
  });
});

// Compass Anticollision Report Summary rows: Offset, Ref MD, Offset MD, Between Centres, Between Ellipses, SF.
const AC_SUMMARY = [
  "Summary",
  "  100/01-02-003-04W5/00 - Offset A   1,110.00   1,105.00   12.40   6.10   1.950 SF",
  "  100/01-02-003-04W5/00 - Offset A   1,110.00   1,105.00   12.40   6.10   1.950 CC",
  "  102/05-06-007-08W5/02 - Offset B     850.00     848.00    9.00   7.20   1.200",
  "  103/09-10-011-12W5/00 - Offset C   2,400.00   2,390.00   40.00   35.00  4.500",
  "Semi Major Axis",
].join("\n");

describe("Import AC (MDL_Setup ImportAntiCollision)", () => {
  it("keeps one SF row per offset and skips the CC / ES depth lines", () => {
    const hits = parseAcSummary(AC_SUMMARY, 1e6);
    expect(hits.map((hit) => [hit.well, hit.refMd, hit.sf])).toEqual([
      ["100/01-02-003-04W5/00 - Offset A", 1110, 1.95],
      ["102/05-06-007-08W5/02 - Offset B", 850, 1.2],
      ["103/09-10-011-12W5/00 - Offset C", 2400, 4.5],
    ]);
    expect(parseAcSummary(AC_SUMMARY).map((hit) => hit.sf)).toEqual([1.95, 1.2]);
    expect(sortAcRows(hits, false).map((hit) => hit.sf)).toEqual([1.2, 1.95, 4.5]);
    expect(sortAcRows(hits, true).map((hit) => hit.refMd)).toEqual([850, 1110, 2400]);
  });

  it("fills the Setup table worst first and the Data concerns by depth", () => {
    const book = makeBook({ Setup: {}, Data: { 32: { 1: { v: "AC Info & Concerns" } }, 42: { 1: { v: "Motors On Location" } } } });
    const hits = sortAcRows(parseAcSummary(AC_SUMMARY, 1e6), false);
    buildAcTable(book, hits);
    expect([10, 11, 13, 15, 17].map((col) => book.value("Setup", 12, col))).toEqual([1, 850, 9, 1.2, "HIGH"]);
    expect(book.value("Setup", 13, 17)).toBe("CAUTION");
    expect(book.writes.some((write) => write.kind === "merge" && write.sheet === "Setup")).toBe(true);
    expect(writeAcConcerns(book, hits)).toBeNull();
    expect(cellAt(book, 35)).toContain("Offset B");
    expect(cellAt(book, 36)).toContain("Offset A");
    expect(book.value("Data", 35, 4)).toBe(1.2);
  });

  it("stores the path, the hits and the summary status", () => {
    const book = makeBook({ Setup: {}, Data: { 32: { 1: { v: "AC Info & Concerns" } } } });
    const result = importAntiCollision(book, "Job\\ac.pdf", AC_SUMMARY, { row: 30, col: 10 });
    expect(result).toMatchObject({ ok: true, status: "AC import complete: 3 offset well(s), 2 with SF < 2.0." });
    expect(book.value(AC_SHEET, 1, 1)).toBe("Job\\ac.pdf");
    expect(book.value("Setup", 30, 11)).toBe("Job\\ac.pdf");
  });

  it("keeps the path and says so when the PDF has no text", () => {
    const book = makeBook({ Setup: {} });
    const result = importAntiCollision(book, "scan.pdf", "", null);
    expect(result.ok).toBe(false);
    expect(result.alert).toContain("scanned");
    expect(book.value(AC_SHEET, 1, 1)).toBe("scan.pdf");
  });
});

function cellAt(book: SheetBook, row: number): string {
  return String(book.value("Data", row, 2) ?? "");
}

describe("PDF text layout (PDF.js items to pdftotext -layout lines)", () => {
  const item = (str: string, x: number, y: number, width: number): PdfTextItem => ({ str, transform: [10, 0, 0, 10, x, y], width, height: 10 });

  it("joins items on one baseline with gap-sized spaces and splits pages", () => {
    const text = layoutText([
      [item("MD", 10, 700, 12), item("Inc", 52, 700, 18), item("600.00", 10, 688, 36), item("0.00", 52, 688, 24)],
      [item("Summary", 10, 700, 42)],
    ]);
    const [first, second] = text.split("\n\f");
    const lines = first!.split("\n");
    expect(lines[0]).toMatch(/^MD {2,}Inc$/);
    expect(lines[1]).toMatch(/^600\.00 +0\.00$/);
    expect(second).toBe("Summary");
  });
});
