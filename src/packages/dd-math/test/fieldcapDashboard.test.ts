import { describe, expect, it } from "vitest";
import { surveyStations } from "../../fieldcap-sync/src/index";
import { backfillDayBreakdown, projectFieldCapDashboard, SheetBook, writeSurveyStations, type SnapshotBook } from "../../workbook-commands/src/index";

function makeBook(sheets: Record<string, Record<number, Record<number, unknown>>>): SheetBook {
  const out: Record<string, { id: string; name: string; cellData: Record<number, Record<number, unknown>> }> = {};
  Object.entries(sheets).forEach(([name, cellData], index) => {
    out[`s${index}`] = { id: `s${index}`, name, cellData };
  });
  return new SheetBook({ sheets: out, styles: {} } as never, []);
}

interface SurveyFixtureRow {
  row: number;
  d: number;
  e: number;
  inc?: number;
  azm?: number;
  tvd?: number;
  good?: boolean;
}

function surveyBook(rows: SurveyFixtureRow[]): SheetBook {
  const cellData: Record<number, Record<number, unknown>> = {};
  for (const row of rows) {
    const cells: Record<number, unknown> = {
      3: { v: row.d },
      4: { v: row.e },
    };
    if (row.inc != null) cells[5] = { v: row.inc, s: row.good ? { bg: { rgb: "#CCFFCC" } } : undefined };
    if (row.azm != null) cells[6] = { v: row.azm, s: row.good ? { bg: { rgb: "#CCFFCC" } } : undefined };
    if (row.tvd != null) cells[7] = { v: row.tvd };
    cellData[row.row - 1] = cells;
  }
  return makeBook({ Slidesheet: cellData });
}

const bhaHeaders = ["BHA #", "Status", "Motor", "Metres Drilled", "BHA Mtrs Slid", "BHA Mtrs Rot", "BHA Hrs Slid", "BHA Hrs Rot", "BHA Hrs Circ", "BHA Total Hrs", "BHA Below Rot", "Serial #", "Item Code", "Description"];

function headerRow(): Record<number, { v: string }> {
  return Object.fromEntries(bhaHeaders.map((header, index) => [index, { v: header }]));
}

function bhaRow(values: unknown[]): Record<number, { v: unknown }> {
  return Object.fromEntries(values.map((value, index) => [index, { v: value }]));
}

describe("FieldCap dashboard", () => {
  it("fills the DD Tools mirror, keeps Data formulas, and puts the active motor on C22", () => {
    const book = makeBook({
      Data: {
        2: { 7: { v: 9 } },
        35: { 7: { f: "=IFERROR('DD Tools'!O6,\"\")", v: 1 } },
      },
      "DD Tools": {},
      _OC_BHA: {
        0: headerRow(),
        1: bhaRow([1, "Active", "PHX-100", 611, 10, 601, 0.52, 7.03, 4.27, 11.82, 0, "PHX-100", "MOT", "Mud Motor"]),
        2: bhaRow([2, "Complete", "PHX-200", 1248, 262.24, 986.52, 16.34, 19.12, 5.07, 40.53, 0, "PHX-200", "MOT", "Mud Motor"]),
      },
      _TH_Hours: { 0: { 0: { v: "Serial" } } },
    });
    const dash = projectFieldCapDashboard(book);
    expect(dash).toMatchObject({ bhas: 2, activeBha: 1, motor: "PHX-100" });
    expect(book.value("DD Tools", 6, 15)).toBe(1);
    expect(book.value("DD Tools", 6, 17)).toBe(10);
    expect(book.value("DD Tools", 6, 18)).toBe(601);
    expect(book.value("DD Tools", 6, 19)).toBe(0.52);
    expect(book.value("DD Tools", 7, 15)).toBe(2);
    expect(book.value("Data", 3, 8)).toBe(1);
    expect(book.value("Data", 22, 3)).toBe("PHX-100");
    expect(book.hasFormula("Data", 36, 8)).toBe(true);
    expect(book.value("DD Tools", 6, 2)).toBe("PHX-100");
  });

  it("keeps H3 when it names a BHA in the mirror", () => {
    const book = makeBook({
      Data: { 2: { 7: { v: 2 } } },
      "DD Tools": {},
      _OC_BHA: {
        0: headerRow(),
        1: bhaRow([1, "Complete", "PHX-100", 1, 1, 0, 1, 0, 0, 1, 0, "PHX-100", "MOT", "Mud Motor"]),
        2: bhaRow([2, "Active", "PHX-200", 2, 2, 0, 2, 0, 0, 2, 0, "PHX-200", "MOT", "Mud Motor"]),
      },
    });
    expect(projectFieldCapDashboard(book)?.activeBha).toBe(2);
    expect(book.value("Data", 22, 3)).toBe("PHX-200");
  });

  it("fills only a matching pre-existing bit-depth row", () => {
    const book = surveyBook([
      { row: 13, d: 500, e: 467, inc: 3, azm: 4, good: true },
      { row: 14, d: 640, e: 607 },
    ]);
    book.setFill("Slidesheet", { row: 13, col: 10, row2: 13, col2: 10 }, "#CCFFCC");
    const result = writeSurveyStations(book, [
      { md: 500, inc: 3, azm: 4 },
      { md: 640, inc: 8, azm: 9 },
    ]);
    expect(result).toMatchObject({ coordinate: "bit-depth", changed: 1, rejectedMd: null });
    expect(book.value("Slidesheet", 14, 4)).toBe(640);
    expect(book.value("Slidesheet", 14, 6)).toBe(8);
    expect(book.value("Slidesheet", 14, 7)).toBe(9);
    expect(book.fill("Slidesheet", 13, 10)).not.toBe("#CCFFCC");
    expect(book.fill("Slidesheet", 14, 10)).toBe("#CCFFCC");
    expect(book.writes.some((write) => write.kind === "value" && write.col === 4)).toBe(false);
  });

  it("does not append complete FieldCap history after the reserved future rows", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87, inc: 2, azm: 11, good: true },
      { row: 15, d: 140, e: 107, inc: 3, azm: 12, good: true },
      { row: 16, d: 160, e: 127, inc: 4, azm: 13, good: true },
      { row: 17, d: 180, e: 147 },
      { row: 18, d: 200, e: 167 },
      { row: 19, d: 220, e: 187 },
    ]);

    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
      { md: 140, inc: 3, azm: 12 },
      { md: 160, inc: 4, azm: 13 },
    ]);

    expect(result).toMatchObject({
      coordinate: "bit-depth",
      rejectedMd: null,
      warning: null,
      changed: 0,
    });
    expect(book.writes).toEqual([]);
    expect(book.value("Slidesheet", 20, 4)).toBeNull();
  });

  it("detects survey-depth workbooks and fills only the next matched F:G row", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87, inc: 2, azm: 11, good: true },
      { row: 15, d: 140, e: 107 },
    ]);
    const result = writeSurveyStations(book, [
      { md: 67, inc: 1, azm: 10 },
      { md: 87, inc: 2, azm: 11 },
      { md: 107, inc: 3, azm: 12 },
    ]);
    expect(result).toMatchObject({ coordinate: "survey-depth", changed: 1, rejectedMd: null });
    expect(book.value("Slidesheet", 15, 4)).toBe(140);
    expect(book.value("Slidesheet", 15, 6)).toBe(3);
    expect(book.value("Slidesheet", 15, 7)).toBe(12);
  });

  it("appends only the contiguous newer survey suffix beyond reserved rows", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87 },
    ]);
    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
      { md: 140, inc: 3, azm: 12 },
      { md: 160, inc: 4, azm: 13 },
    ]);
    expect(result).toMatchObject({ coordinate: "bit-depth", changed: 3, appended: 2, rejectedMd: null });
    expect(book.value("Slidesheet", 15, 4)).toBe(140);
    expect(book.value("Slidesheet", 15, 6)).toBe(3);
    expect(book.value("Slidesheet", 16, 4)).toBe(160);
    expect(book.value("Slidesheet", 16, 7)).toBe(13);
  });

  it("uses the survey offset when appending survey-depth stations", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87 },
    ]);
    const result = writeSurveyStations(book, [
      { md: 67, inc: 1, azm: 10 },
      { md: 87, inc: 2, azm: 11 },
      { md: 107, inc: 3, azm: 12 },
    ]);
    expect(result).toMatchObject({ coordinate: "survey-depth", appended: 1, rejectedMd: null });
    expect(book.value("Slidesheet", 15, 4)).toBe(140);
    expect(book.value("Slidesheet", 15, 6)).toBe(3);
  });

  it("omits trailing zero-inc zero-azimuth FieldCap placeholders", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
    ]);
    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
      { md: 140, inc: 0, azm: 0 },
      { md: 160, inc: 0, azm: 0 },
    ]);
    expect(result).toMatchObject({ appended: 1, skippedPlaceholders: 2, rejectedMd: null });
    expect(book.value("Slidesheet", 14, 4)).toBe(120);
    expect(book.value("Slidesheet", 15, 4)).toBeNull();
  });

  it("lets FieldCap's depth replace a reserved row's planned depth, and leaves the rows below alone", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87 },
      { row: 15, d: 140, e: 107 },
    ]);
    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 140, inc: 3, azm: 12 },
    ]);
    expect(result).toMatchObject({ rejectedMd: null, warning: null, changed: 1 });
    expect(result.depthCorrections).toEqual([{ row: 14, md: 140, from: 120, to: 140 }]);
    expect(book.value("Slidesheet", 14, 4)).toBe(140);
    expect(book.value("Slidesheet", 14, 6)).toBe(3);
    expect(book.value("Slidesheet", 14, 7)).toBe(12);
    expect(book.value("Slidesheet", 15, 4)).toBe(140);
    expect(book.value("Slidesheet", 15, 6)).toBeNull();
  });

  it("catches up many surveys after a long gap even though every planned depth has drifted", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87 },
      { row: 15, d: 140, e: 107 },
      { row: 16, d: 160, e: 127 },
    ]);
    const stations = [
      { md: 100, inc: 1, azm: 10 },
      { md: 121.5, inc: 2, azm: 11 },
      { md: 142.1, inc: 3, azm: 12 },
      { md: 161, inc: 4, azm: 13 },
      { md: 181.2, inc: 5, azm: 14 },
      { md: 200.9, inc: 6, azm: 15 },
    ];
    const result = writeSurveyStations(book, stations);
    expect(result).toMatchObject({ rejectedMd: null, warning: null, changed: 5, appended: 2 });
    expect(result.depthCorrections.map((fix) => [fix.row, fix.from, fix.to])).toEqual([[14, 120, 121.5], [15, 140, 142.1], [16, 160, 161]]);
    expect([14, 15, 16, 17, 18].map((row) => book.value("Slidesheet", row, 4))).toEqual([121.5, 142.1, 161, 181.2, 200.9]);
    expect([14, 15, 16, 17, 18].map((row) => book.value("Slidesheet", row, 6))).toEqual([2, 3, 4, 5, 6]);
    expect(book.fill("Slidesheet", 18, 7)).toBe("#CCFFCC");
    expect(writeSurveyStations(book, stations)).toMatchObject({ changed: 0, appended: 0 });
  });

  it("keeps the row's own survey offset when it rewrites a survey-depth workbook's bit depth", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87 },
    ]);
    const result = writeSurveyStations(book, [
      { md: 67, inc: 1, azm: 10 },
      { md: 88.5, inc: 2, azm: 11 },
    ]);
    expect(result).toMatchObject({ coordinate: "survey-depth", warning: null });
    expect(book.value("Slidesheet", 14, 4)).toBe(121.5);
  });

  it("corrects a past good survey whose typed depth is off, and gives the row's survey back to FieldCap", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87, inc: 2, azm: 11, good: true },
      { row: 15, d: 139.6, e: 106.6, inc: 3, azm: 12, good: true },
      { row: 16, d: 160, e: 127 },
    ]);
    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
      { md: 140, inc: 3, azm: 12 },
      { md: 160, inc: 4, azm: 13 },
    ]);
    expect(result).toMatchObject({ rejectedMd: null, warning: null, changed: 2 });
    expect(result.depthCorrections).toEqual([{ row: 15, md: 140, from: 139.6, to: 140 }]);
    expect(book.value("Slidesheet", 15, 4)).toBe(140);
    expect(book.value("Slidesheet", 16, 6)).toBe(4);
  });

  it("pairs a mistyped depth between two matching surveys by order", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 119.7, e: 86.7, inc: 2, azm: 11, good: true },
      { row: 15, d: 140, e: 107, inc: 3, azm: 12, good: true },
    ]);
    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
      { md: 140, inc: 3, azm: 12 },
    ]);
    expect(result.warning).toBeNull();
    expect(result.depthCorrections).toEqual([{ row: 14, md: 120, from: 119.7, to: 120 }]);
    expect(book.value("Slidesheet", 14, 4)).toBe(120);
  });

  it("asks before it rewrites a past survey depth that moves its TVD by more than 0.10 m", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, tvd: 66, good: true },
      { row: 14, d: 119.7, e: 86.7, inc: 2, azm: 11, tvd: 86, good: true },
    ]);
    const stations = [
      { md: 100, inc: 1, azm: 10, tvd: 66 },
      { md: 120, inc: 2, azm: 11, tvd: 86.5 },
    ];
    const pending = writeSurveyStations(book, stations);
    expect(pending).toMatchObject({ needsConfirmation: true });
    expect(pending.tvdMismatches).toHaveLength(1);
    expect(book.writes).toEqual([]);
    const accepted = writeSurveyStations(book, stations, { allowTvdMismatch: true });
    expect(accepted.depthCorrections).toHaveLength(1);
    expect(book.value("Slidesheet", 14, 4)).toBe(120);
  });

  it("still refuses a FieldCap survey inside the Slidesheet that has no row to take it", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 140, e: 107, inc: 3, azm: 12, good: true },
    ]);
    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
      { md: 140, inc: 3, azm: 12 },
    ]);
    expect(result.rejectedMd).toBe(120);
    expect(book.writes).toEqual([]);
  });

  it("refuses when the surveys to add would pass row 505", () => {
    const book = surveyBook([
      { row: 504, d: 100, e: 67, inc: 1, azm: 10, good: true },
    ]);
    const result = writeSurveyStations(book, [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
      { md: 140, inc: 3, azm: 12 },
    ]);
    expect(result.warning).toContain("505");
    expect(book.writes).toEqual([]);
  });

  it("is idempotent after applying a matched survey", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, good: true },
      { row: 14, d: 120, e: 87 },
    ]);
    const stations = [
      { md: 100, inc: 1, azm: 10 },
      { md: 120, inc: 2, azm: 11 },
    ];
    expect(writeSurveyStations(book, stations).changed).toBe(1);
    const writesAfterFirstPull = book.writes.length;
    expect(writeSurveyStations(book, stations).changed).toBe(0);
    expect(book.writes).toHaveLength(writesAfterFirstPull);
  });

  it("requires confirmation before correcting a past survey with a TVD difference over 0.10 m", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, tvd: 66.8, good: true },
    ]);
    const stations = [{ md: 100, inc: 1.2, azm: 10, tvd: 67.01 }];
    const pending = writeSurveyStations(book, stations);
    expect(pending).toMatchObject({ needsConfirmation: true, changed: 0 });
    expect(pending.tvdMismatches).toHaveLength(1);
    expect(book.writes).toEqual([]);

    const accepted = writeSurveyStations(book, stations, { allowTvdMismatch: true });
    expect(accepted).toMatchObject({ needsConfirmation: false, changed: 1 });
    expect(book.value("Slidesheet", 13, 6)).toBe(1.2);
  });

  it("does not require confirmation at exactly a 0.10 m TVD difference", () => {
    const book = surveyBook([
      { row: 13, d: 100, e: 67, inc: 1, azm: 10, tvd: 66.8, good: true },
    ]);
    const result = writeSurveyStations(book, [{ md: 100, inc: 1.2, azm: 10, tvd: 66.9 }]);
    expect(result).toMatchObject({ needsConfirmation: false, changed: 1 });
  });

  it("matches a FieldCap station exactly 0.05 m from its reserved survey-depth row", () => {
    const book = surveyBook([
      { row: 332, d: 6199.68, e: 6166.68, inc: 89.9, azm: 147.84, good: true },
      { row: 333, d: 6218.88, e: 6185.88 },
      { row: 334, d: 6238.08, e: 6205.08 },
    ]);
    const result = writeSurveyStations(book, [
      { md: 6166.69, inc: 89.9, azm: 147.84 },
      { md: 6185.93, inc: 90.28, azm: 147.25 },
      { md: 6205.13, inc: 90.36, azm: 147.77 },
    ]);

    expect(result).toMatchObject({ coordinate: "survey-depth", changed: 2, warning: null });
    expect(book.value("Slidesheet", 333, 6)).toBe(90.28);
    expect(book.value("Slidesheet", 334, 7)).toBe(147.77);
  });
});

describe("day breakdown follows H3", () => {
  it("puts the selected BHA slide metres on the open Midnight row", () => {
    const book: SnapshotBook = {
      sheets: {
        data: {
          id: "data",
          name: "Data",
          cellData: {
            2: { 7: { v: 1 } },
            3: { 7: { v: "Midnight" } },
            35: { 7: { v: 1 }, 8: { v: 10 }, 10: { v: 2 } },
          },
        },
      },
    } as SnapshotBook;
    const writes = backfillDayBreakdown(book);
    expect(writes).toContainEqual({ a1: "L4", value: 10 });
    expect(writes).toContainEqual({ a1: "M4", value: 2 });
  });
});

describe("surveyStations", () => {
  it("keeps rows that have depth, inclination and azimuth", () => {
    const fields = { md: "MeasuredDepth", inc: "Inclination", azm: "Azimuth", tvd: "TrueVerticalDepth", bit: null };
    expect(surveyStations([
      { MeasuredDepth: 100, Inclination: 1, Azimuth: 2, TrueVerticalDepth: 99.8 },
      { MeasuredDepth: "bad", Inclination: 1, Azimuth: 2, TrueVerticalDepth: 40 },
      { MeasuredDepth: 50, Inclination: 0, Azimuth: 9, TrueVerticalDepth: 49.9 },
    ], fields)).toEqual([
      { md: 50, inc: 0, azm: 9, tvd: 49.9 },
      { md: 100, inc: 1, azm: 2, tvd: 99.8 },
    ]);
    expect(surveyStations([{ CourseLength: 12 }], { md: null, inc: null, azm: null, tvd: null, bit: null })).toEqual([]);
  });

  it("carries the bit depth when FieldCap publishes one", () => {
    const fields = { md: "SurveyDepth", inc: "Inclination", azm: "Azimuth", tvd: null, bit: "BitDepth" };
    expect(surveyStations([
      { SurveyDepth: 30, BitDepth: 48, Inclination: 1, Azimuth: 2 },
      { SurveyDepth: 50, BitDepth: null, Inclination: 3, Azimuth: 4 },
    ], fields)).toEqual([
      { md: 30, inc: 1, azm: 2, bit: 48 },
      { md: 50, inc: 3, azm: 4 },
    ]);
  });
});

/** A new well: the tie-in on row 12 (Survey Offset in B) and empty survey rows that carry the sheet's formulas. */
function emptySurveyBook(offset: number | null = 19, tieIn = 5.5): SheetBook {
  const green = { bg: { rgb: "#CCFFCC" } };
  const cellData: Record<number, Record<number, unknown>> = {
    11: { 1: { v: offset ?? undefined }, 3: { v: tieIn }, 4: { v: tieIn }, 5: { v: 0, s: green }, 6: { v: 0, s: green }, 7: { v: tieIn } },
  };
  for (let row = 13; row <= 20; row += 1) {
    cellData[row - 1] = {
      1: { f: `=IF(ISBLANK(D${row}),"",B${row - 1})`, v: "" },
      3: { s: green },
      4: { f: `=IF(D${row}="","",(ROUND(SUM(D${row}-B${row}),2)))`, v: "" },
      5: { s: green },
      6: { s: green },
    };
  }
  return makeBook({ Slidesheet: cellData });
}

describe("FieldCap surveys on a Slidesheet with no surveys yet", () => {
  it("fills from row 13 with FieldCap's bit depths and sets the Survey Offset from them", () => {
    const book = emptySurveyBook(19);
    const result = writeSurveyStations(book, [
      { md: 30.259999999999998, inc: 0.2, azm: 34, bit: 48.26 },
      { md: 49.62, inc: 0.3, azm: 98, bit: 67.62 },
      { md: 69.02, inc: 0.6, azm: 198, bit: 87.02 },
    ]);
    expect(result).toMatchObject({ coordinate: "survey-depth", appended: 3, changed: 3, warning: null, checkOffset: false });
    expect(result.note).toContain("Survey Offset was set to 18 from FieldCap (it was 19)");
    expect(book.value("Slidesheet", 12, 2)).toBe(18);
    expect([13, 14, 15].map((row) => book.value("Slidesheet", row, 4))).toEqual([48.26, 67.62, 87.02]);
    expect([13, 14, 15].map((row) => book.value("Slidesheet", row, 6))).toEqual([0.2, 0.3, 0.6]);
    expect([13, 14, 15].map((row) => book.value("Slidesheet", row, 7))).toEqual([34, 98, 198]);
    expect(book.fill("Slidesheet", 15, 6)).toBe("#CCFFCC");
    expect(book.hasFormula("Slidesheet", 14, 2)).toBe(true);
    expect(book.fill("Slidesheet", 15, 10)).toBe("#CCFFCC");
    expect(book.fill("Slidesheet", 14, 10)).not.toBe("#CCFFCC");
    expect(book.value("Slidesheet", 16, 4)).toBeNull();
  });

  it("leaves a Survey Offset that already agrees with FieldCap alone", () => {
    const book = emptySurveyBook(18);
    const result = writeSurveyStations(book, [{ md: 30, inc: 0.2, azm: 34, bit: 48 }]);
    expect(result.appended).toBe(1);
    expect(result.note).toContain("Survey Offset is 18, as in FieldCap");
    expect(book.writes.some((write) => write.kind === "value" && write.row === 12)).toBe(false);
  });

  it("puts a changed offset on the row where FieldCap's offset changes", () => {
    const book = emptySurveyBook(18);
    const result = writeSurveyStations(book, [
      { md: 30, inc: 0.2, azm: 34, bit: 48 },
      { md: 50, inc: 0.3, azm: 98, bit: 68 },
      { md: 70, inc: 0.6, azm: 198, bit: 91.5 },
      { md: 90, inc: 0.7, azm: 199, bit: 111.5 },
    ]);
    expect(result.appended).toBe(4);
    expect(book.hasFormula("Slidesheet", 14, 2)).toBe(true);
    expect(book.value("Slidesheet", 15, 2)).toBe(21.5);
    expect(book.hasFormula("Slidesheet", 16, 2)).toBe(true);
    expect(book.value("Slidesheet", 15, 4)).toBe(91.5);
  });

  it("uses the sheet's Survey Offset when FieldCap gives no bit depths, and says so", () => {
    const book = emptySurveyBook(19);
    const result = writeSurveyStations(book, [
      { md: 30.259999999999998, inc: 0.2, azm: 34 },
      { md: 49.62, inc: 0.3, azm: 98 },
    ]);
    expect(result).toMatchObject({ appended: 2, warning: null, checkOffset: true });
    expect(result.note).toContain("FieldCap gave no bit depths");
    expect([13, 14].map((row) => book.value("Slidesheet", row, 4))).toEqual([49.26, 68.62]);
    expect(book.writes.some((write) => write.kind === "value" && write.col === 2)).toBe(false);
  });

  it("refuses when FieldCap gives no bit depths and the sheet has no Survey Offset", () => {
    const book = emptySurveyBook(null);
    const result = writeSurveyStations(book, [{ md: 30, inc: 0.2, azm: 34 }]);
    expect(result.warning).toContain("Survey Offset");
    expect(book.writes).toEqual([]);
  });

  it("leaves out stations at or above the tie-in and trailing placeholders", () => {
    const book = emptySurveyBook(18, 5.5);
    const result = writeSurveyStations(book, [
      { md: 5.5, inc: 0, azm: 0, bit: 5.5 },
      { md: 30, inc: 0.2, azm: 34, bit: 48 },
      { md: 50, inc: 0, azm: 0, bit: 68 },
    ]);
    expect(result).toMatchObject({ appended: 1, ignored: 1, skippedPlaceholders: 1 });
    expect(book.value("Slidesheet", 13, 4)).toBe(48);
    expect(book.value("Slidesheet", 14, 4)).toBeNull();
  });

  it("still refuses when depth rows exist but none is a good survey", () => {
    const book = surveyBook([{ row: 13, d: 100, e: 67 }]);
    const result = writeSurveyStations(book, [{ md: 67, inc: 1, azm: 10, bit: 100 }]);
    expect(result.warning).toBe("FieldCap surveys were not changed at MD 67.00: no existing good survey establishes the depth coordinate.");
    expect(book.writes).toEqual([]);
  });
});

/** Rows 13.. from [d, good]; column B carries the tie-in offset down by formula, as the sheet does. */
function offsetBook(rows: Array<{ d: number; good: boolean; inc?: number; azm?: number }>, offset = 18): SheetBook {
  const cellData: Record<number, Record<number, unknown>> = { 11: { 1: { v: offset } } };
  rows.forEach((entry, index) => {
    const row = 13 + index;
    const green = { bg: { rgb: "#CCFFCC" } };
    cellData[row - 1] = {
      1: { v: offset, f: `=IF(ISBLANK(D${row}),"",B${row - 1})` },
      3: { v: entry.d },
      4: { v: Math.round((entry.d - offset) * 100) / 100 },
      ...(entry.good ? { 5: { v: entry.inc ?? 1, s: green }, 6: { v: entry.azm ?? 10, s: green } } : {}),
    };
  });
  return makeBook({ Slidesheet: cellData });
}

describe("Survey Offset follows FieldCap's bit depth (each BHA has its own)", () => {
  const history = [
    { md: 82, bit: 100, inc: 1, azm: 10 },
    { md: 102, bit: 120, inc: 2, azm: 11 },
  ];

  it("types the new offset on the row where a new BHA's surveys begin, and the rows below carry it", () => {
    const book = offsetBook([{ d: 100, good: true }, { d: 120, good: true, inc: 2, azm: 11 }, { d: 140, good: false }, { d: 160, good: false }, { d: 180, good: false }]);
    const result = writeSurveyStations(book, [...history, { md: 123.5, bit: 145, inc: 3, azm: 12 }, { md: 142.5, bit: 164, inc: 4, azm: 13 }]);
    expect(result).toMatchObject({ rejectedMd: null, warning: null, changed: 2 });
    expect(result.offsetChanges).toEqual([{ row: 15, md: 123.5, from: 18, to: 21.5 }]);
    expect([15, 16].map((row) => book.value("Slidesheet", row, 4))).toEqual([145, 164]);
    expect(book.value("Slidesheet", 15, 2)).toBe(21.5);
    expect(book.hasFormula("Slidesheet", 16, 2)).toBe(true);
    expect([15, 16].map((row) => book.fill("Slidesheet", row, 6))).toEqual(["#CCFFCC", "#CCFFCC"]);
  });

  it("leaves the offset alone when FieldCap's bit depths agree with it", () => {
    const book = offsetBook([{ d: 100, good: true }, { d: 120, good: true, inc: 2, azm: 11 }, { d: 140, good: false }]);
    const result = writeSurveyStations(book, [...history, { md: 122, bit: 140, inc: 3, azm: 12 }]);
    expect(result.offsetChanges).toEqual([]);
    expect(book.value("Slidesheet", 15, 2)).toBe(18);
    expect(book.writes.some((write) => write.kind === "value" && write.col === 2)).toBe(false);
  });

  it("fixes a past survey whose bit depth FieldCap corrects, keeping its survey depth", () => {
    const book = offsetBook([{ d: 100, good: true }, { d: 120, good: true, inc: 2, azm: 11 }, { d: 140, good: false }]);
    const result = writeSurveyStations(book, [{ md: 82, bit: 100, inc: 1, azm: 10 }, { md: 102, bit: 122, inc: 2, azm: 11 }, { md: 122, bit: 142, inc: 3, azm: 12 }]);
    expect(result.offsetChanges).toEqual([{ row: 14, md: 102, from: 18, to: 20 }]);
    expect(book.value("Slidesheet", 14, 4)).toBe(122);
    expect(book.value("Slidesheet", 14, 2)).toBe(20);
    expect(result.depthCorrections).toContainEqual({ row: 14, md: 102, from: 120, to: 122 });
    expect(book.value("Slidesheet", 15, 4)).toBe(142);
  });

  it("is quiet the second time", () => {
    const book = offsetBook([{ d: 100, good: true }, { d: 120, good: true, inc: 2, azm: 11 }, { d: 140, good: false }, { d: 160, good: false }]);
    const stations = [...history, { md: 123.5, bit: 145, inc: 3, azm: 12 }, { md: 142.5, bit: 164, inc: 4, azm: 13 }];
    writeSurveyStations(book, stations);
    const before = book.writes.length;
    const again = writeSurveyStations(book, stations);
    // The overlay does not recalculate E, so only the offset is judged here.
    expect(again.offsetChanges).toEqual([]);
    expect(book.writes.slice(before).some((write) => write.kind === "value" && write.col === 2)).toBe(false);
  });

  it("keeps the sheet's offset when FieldCap gives no bit depth", () => {
    const book = offsetBook([{ d: 100, good: true }, { d: 120, good: true, inc: 2, azm: 11 }, { d: 140, good: false }]);
    const result = writeSurveyStations(book, [{ md: 82, inc: 1, azm: 10 }, { md: 102, inc: 2, azm: 11 }, { md: 122, inc: 3, azm: 12 }]);
    expect(result.offsetChanges).toEqual([]);
    expect(book.value("Slidesheet", 15, 4)).toBe(140);
  });
});
