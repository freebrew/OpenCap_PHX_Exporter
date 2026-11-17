import { describe, expect, it } from "vitest";
import {
  applyOpenCap,
  costDateSerial,
  csvConv,
  sortCrewByRole,
  readCostDays,
  writeCostDay,
  mudMotorSequence,
  parseOpenCapCsv,
  SheetBook,
  toolHoursSync,
  vbaIsNumeric,
  type OpenCapSource,
} from "../../workbook-commands/src/index";

function makeBook(sheets: Record<string, Record<number, Record<number, unknown>>>): SheetBook {
  const out: Record<string, { id: string; name: string; cellData: Record<number, Record<number, unknown>> }> = {};
  Object.entries(sheets).forEach(([name, cellData], index) => {
    out[`s${index}`] = { id: `s${index}`, name, cellData };
  });
  return new SheetBook({ sheets: out, styles: {} } as never, []);
}

const JOB_CSV = "Job ID,Job Code,Client,ClientEmail,opsStatus,MagneticReferenceDate,NorthReference,GroundMSL\r\n4242,J-4242,,ops@acme-energy.ca,Drilling,20260716,TRUE,767.11\r\n";
const CREW_CSV = "Job ID,Name,Employee Code,Email,Phone,Role,Work Type\r\n4242,Ann Lee,1,ann@x.com,403-555-1234,Day DD,\r\n4242,Bob Roe,2,,4035550000,,Night MWD\r\n4242,,3,,,,\r\n";
const BHA_CSV = [
  "Job ID,BHA #,Description,Serial #,BHA Total Hrs,Source,Sub Description",
  "4242,1,Mud Motor 6.5,PHX-525010-PTS,40,DD Inventory,",
  "4242,1,Agitator,AG-100,40,Other Inventory,",
  "4242,2,Mud Motor 6.5,PHX-525074-PTS,12,DD Inventory,",
  "4242,2,Agitator,AG-100,12,Other Inventory,",
  "4242,2,Drill Bit,BIT-1,12,Other Inventory,",
].join("\r\n");
const INVENTORY_CSV = "ItemName,SerialNumber,Category,SubCategory\r\nMud Motor 6.5,PHX-525074-PTS,DD,Motor\r\nMud Motor 6.5,PHX-525010-PTS,DD,Motor\r\nAgitator,AG-100,Other Inventory,Agitator\r\n";
const COSTS_CSV = "Job ID,Date,Daily Cost,Ticket Count\r\n4242,2026-07-20,1000,1\r\n4242,2026-07-21,\"1,250.50\",2\r\n";

const sources: OpenCapSource[] = [
  { kind: "job", csv: JOB_CSV, name: "fieldcap-job-4242-job-details.csv", path: "Job\\OpenCap\\fieldcap-job-4242-job-details.csv" },
  { kind: "crew", csv: CREW_CSV, name: "fieldcap-job-4242-crew.csv", path: "Job\\OpenCap\\fieldcap-job-4242-crew.csv" },
  { kind: "bha", csv: BHA_CSV, name: "fieldcap-job-4242-bha-equipment.csv", path: "Job\\OpenCap\\fieldcap-job-4242-bha-equipment.csv" },
  { kind: "inventory", csv: INVENTORY_CSV, name: "fieldcap-job-4242-inventory.csv", path: "Job\\OpenCap\\fieldcap-job-4242-inventory.csv" },
  { kind: "costs", csv: COSTS_CSV, name: "fieldcap-job-4242-ticket-costs-by-day.csv", path: "Job\\OpenCap\\fieldcap-job-4242-ticket-costs-by-day.csv" },
];

// 0-based cellData rows / columns, as a Univer snapshot stores them.
function workbook(): SheetBook {
  return makeBook({
    Setup: {
      2: { 9: { v: "  CREW MANIFEST" } },
      3: { 1: { v: "JOB ID" }, 5: { v: "OPS STATUS" } },
      4: { 1: { v: "GROUND MSL" }, 5: { v: "REF DATE" } },
      5: { 1: { v: "CLIENT" } },
      6: { 5: { v: "NORTH REF" } },
      18: { 9: { v: "STATUS" } },
    },
    Costs: {
      2: { 1: { v: "Day 1" }, 2: { v: 46221 }, 4: { v: 500 } },
      3: { 1: { v: "Day 2" }, 2: { f: "=IF(E4<>\"\",C3+1,\"\")" }, 4: { v: 600 } },
    },
    Data: {
      2: { 7: { v: 2 } },
      23: { 16: { v: 5.5 } },
      27: { 14: { v: "Enter any 3rd Party Tools/hours below" } },
      29: { 14: { v: "Serial Number" } },
      30: { 14: { v: "AG-100" }, 16: { v: 100 } },
      43: { 14: { v: "Serial Number" } },
    },
    _TH_Hours: { 0: { 0: { v: "Serial" } } },
  });
}

describe("OpenCap CSV values (LoadCsv / CsvConv)", () => {
  it("types numbers like VBA IsNumeric and keeps phones and dates as text", () => {
    expect(vbaIsNumeric("1,250.50")).toBe(true);
    expect(vbaIsNumeric("$12")).toBe(true);
    expect(vbaIsNumeric("1e3")).toBe(true);
    expect(vbaIsNumeric("403-555-1234")).toBe(false);
    expect(vbaIsNumeric("2026-07-16")).toBe(false);
    expect(csvConv("  42 ")).toBe(42);
    expect(csvConv("TRUE")).toBe(true);
    expect(csvConv("   ")).toBeNull();
  });

  it("parses quoted fields with escaped quotes", () => {
    expect(parseOpenCapCsv("A,B\r\n\"x, \"\"y\"\"\",2\r\n")).toEqual([["A", "B"], ["x, \"y\"", 2]]);
  });

  it("keeps a quoted field that runs over several lines in one row", () => {
    expect(parseOpenCapCsv("A,Directions,C\r\n1,\"from town:\nEast 5 km\nLeft into.\",3\r\n4,x,6\r\n")).toEqual([
      ["A", "Directions", "C"],
      [1, "from town:\nEast 5 km\nLeft into.", 3],
      [4, "x", 6],
    ]);
  });

  it("reads cost dates the way ParseCostDateSerial does", () => {
    expect(costDateSerial("2026-07-16")).toBe(46219);
    expect(costDateSerial("7/16/2026")).toBe(46219);
    expect(costDateSerial(46219.4)).toBe(46219);
    expect(costDateSerial(12)).toBe(0);
  });

  it("sorts PTS motors by their sequence number", () => {
    expect(mudMotorSequence("PHX-525010-PTS")).toBe(525010);
    expect(mudMotorSequence("24X-19722-PTS")).toBe(19722);
    expect(mudMotorSequence("ABC-1")).toBe(1e99);
  });
});

describe("applyOpenCap (RefreshAllCsvData without the rebuild)", () => {
  it("loads the CSVs and fills the Setup job, crew and file panels", () => {
    const book = workbook();
    const result = applyOpenCap(book, sources, { lastJobId: "4242", folderPath: "Job" });
    expect(result).toMatchObject({ jobId: "4242", jobNumber: "J-4242", lastJobId: "4242", newJob: false, rows: { job: 1, crew: 3, bha: 5, inventory: 3, costs: 2 }, alert: null });
    expect(book.value("_OC_Job", 2, 8)).toBe(767.11);
    expect(book.value("Setup", 4, 4)).toBe("J-4242");
    expect(book.value("Setup", 4, 7)).toBe("Drilling");
    expect(book.value("Setup", 5, 4)).toBe("767.11 m");
    expect(book.value("Setup", 5, 7)).toBe(46219);
    expect(book.value("Setup", 6, 4)).toBe("Acme Energy");
    expect(book.value("Setup", 7, 7)).toBe(true);
    expect(book.value("Setup", 3, 10)).toBe("  CREW MANIFEST   2");
    expect([book.value("Setup", 5, 10), book.value("Setup", 5, 11), book.value("Setup", 5, 16)]).toEqual(["Day DD", "Ann Lee", "403-555-1234"]);
    expect([book.value("Setup", 6, 10), book.value("Setup", 6, 16)]).toEqual(["Night MWD", 4035550000]);
    expect(book.value("Setup", 7, 11)).toBeNull();
    expect([book.value("Setup", 20, 10), book.value("Setup", 20, 13), book.value("Setup", 20, 16)]).toEqual(["\u2713", "fieldcap-job-4242-job-details.csv", 1]);
    expect([book.value("Setup", 23, 10), book.value("Setup", 23, 13), book.value("Setup", 23, 17)]).toEqual(["\u2717", "*slide-rotate*.csv", "Job"]);
  });

  it("keeps GroundMSL and UWI when Directions contains several physical line breaks", () => {
    const csv =
      "Job ID,Job Code,Directions,GroundMSL,UWI\r\n" +
      "4242,20803,\"from town:\nEast\nSouth\nEast\nSouth\nWest\nLeft into.\",712.5,100/01-01\r\n";
    const rows = parseOpenCapCsv(csv);
    expect(rows).toHaveLength(2);
    expect(rows[1][2]).toContain("Left into.");
    expect(rows[1][3]).toBe(712.5);
    expect(rows[1][4]).toBe("100/01-01");
  });

  it("fills the fields that come after multi-line driving directions", () => {
    const book = workbook();
    const csv = "Job ID,Job Code,Directions,opsStatus,GroundMSL\r\n4242,J-4242,\"from town:\nEast 5 km\nLeft into.\",Drilling,767.11\r\n";
    const job: OpenCapSource = { kind: "job", csv, name: "fieldcap-job-4242-job-details.csv", path: "Job\\OpenCap\\fieldcap-job-4242-job-details.csv" };
    const result = applyOpenCap(book, [job], { lastJobId: "4242", folderPath: "Job" });
    expect(result.rows.job).toBe(1);
    expect(book.value("Setup", 4, 7)).toBe("Drilling");
    expect(book.value("Setup", 5, 4)).toBe("767.11 m");
  });

  it("keeps the sheet's own cost days and lays the CSV days after them", () => {
    const book = workbook();
    applyOpenCap(book, sources, { lastJobId: "4242", folderPath: "Job" });
    expect([3, 4, 5, 6].map((row) => book.value("Costs", row, 5))).toEqual([500, 600, 1000, 1250.5]);
    expect(book.value("Costs", 3, 3)).toBe(46221);
    expect(book.formula("Costs", 5, 3)).toBe("=IF(E5<>\"\",C4+1,\"\")");
    expect(book.formula("Costs", 6, 7)).toBe("=IF(ISBLANK(E6),\"\",SUM($E$3:E6))");
  });

  it("drops the previous job's costs when the job ID changes", () => {
    const book = workbook();
    const result = applyOpenCap(book, sources, { lastJobId: "4100", folderPath: "Job" });
    expect(result.newJob).toBe(true);
    expect([3, 4].map((row) => book.value("Costs", row, 5))).toEqual([1000, 1250.5]);
    expect(book.value("Costs", 3, 3)).toBe(46223);
    expect(book.value("Costs", 5, 5)).toBeNull();
  });

  it("lists inventory mud motors in sequence and syncs the hours tables", () => {
    const book = workbook();
    applyOpenCap(book, sources, { lastJobId: "", folderPath: "Job" });
    expect([book.value("Data", 45, 2), book.value("Data", 46, 2), book.value("Data", 47, 2)]).toEqual(["PHX-525010-PTS", "PHX-525074-PTS", null]);
    // AG-100 ran on BHA 1 (40 h) and is on the selected BHA 2, whose live hours are Q24.
    expect([book.value("Data", 31, 15), book.value("Data", 31, 17), book.value("Data", 31, 18)]).toEqual(["AG-100", 100, 45.5]);
    expect(book.value("Data", 32, 15)).toBeNull();
    expect([book.value("Data", 45, 15), book.value("Data", 45, 18)]).toEqual(["PHX-525074-PTS", 5.5]);
    expect(book.formula("Data", 45, 6)).toBe("=IFERROR(VLOOKUP(B45,$O$45:$S$55,5,FALSE),IFERROR(VLOOKUP(B45,_TH_Hours!$A$2:$E$80,5,FALSE),\"\"))");
    const tracker = new Map<string, unknown[]>();
    for (let row = 2; row <= book.lastUsedRow("_TH_Hours", 1); row += 1) {
      tracker.set(String(book.value("_TH_Hours", row, 1)), [2, 3, 4, 5, 6].map((col) => book.value("_TH_Hours", row, col)));
    }
    expect(tracker.get("AG-100")).toEqual(["Tool", 100, 45.5, 145.5, "1,2"]);
    expect(tracker.get("PHX-525010-PTS")).toEqual(["Motor", 0, 40, 40, "1"]);
    expect(tracker.get("PHX-525074-PTS")).toEqual(["Motor", 0, 5.5, 5.5, "2"]);
    expect(tracker.has("BIT-1")).toBe(false);
    expect([book.value("_TH_Hours", 2, 13), book.value("_TH_Hours", 3, 13)]).toEqual(["PHX-525010-PTS", "PHX-525074-PTS"]);
  });

  it("clears the Previous hours when a different serial takes the row", () => {
    const book = workbook();
    applyOpenCap(book, sources, { lastJobId: "", folderPath: "Job" });
    book.setValue("Data", 3, 8, 1);
    book.setValue("Data", 31, 15, "OLD-77");
    book.setValue("Data", 31, 17, 9);
    toolHoursSync(book);
    // BHA 1 is now the live one (Q24 = 5.5); BHA 2 adds its 12 h.
    expect([book.value("Data", 31, 15), book.value("Data", 31, 17), book.value("Data", 31, 18)]).toEqual(["AG-100", null, 17.5]);
    expect(book.value("_TH_Hours", 2, 1)).toBe("AG-100");
  });

  it("sorts the crew manifest by role, then by name", () => {
    const book = workbook();
    book.setValue("Setup", 5, 10, "MWD");
    book.setValue("Setup", 5, 11, "Zoe");
    book.setValue("Setup", 6, 10, "DD");
    book.setValue("Setup", 6, 11, "Ann");
    book.setValue("Setup", 7, 10, "DD");
    book.setValue("Setup", 7, 11, "Bob");
    expect(sortCrewByRole(book, 5, 10)).toBe(true);
    expect([5, 6, 7].map((row) => book.value("Setup", row, 11))).toEqual(["Ann", "Bob", "Zoe"]);
    expect([5, 6, 7].map((row) => book.value("Setup", row, 10))).toEqual(["DD", "DD", "MWD"]);
    // A second pass finds nothing to move.
    expect(sortCrewByRole(book, 5, 10)).toBe(false);
  });

  it("writes one cost day, appending a row with the running total", () => {
    const book = workbook();
    const existing = readCostDays(book);
    expect(existing.map((day) => day.daily)).toEqual([500, 600]);
    expect(existing[1]?.total).toBe(1100);
    const added = writeCostDay(book, 46223, 250, "night batch");
    expect(added).toEqual({ added: true, row: 5 });
    expect(book.value("Costs", 5, 5)).toBe(250);
    expect(book.value("Costs", 5, 7)).toBe("night batch");
    expect(book.formula("Costs", 5, 6)).toBe("=IF(ISBLANK(E5),\"\",SUM($E$3:E5))");
    // The same date again updates the row instead of adding another.
    expect(writeCostDay(book, 46223, 275, "")).toEqual({ added: false, row: 5 });
    expect(book.value("Costs", 5, 5)).toBe(275);
  });

  it("reports what it found when nothing loads", () => {
    const book = workbook();
    const result = applyOpenCap(book, [], { lastJobId: "", folderPath: "Job" });
    expect(result.status).toBe("No OpenCap CSVs found.");
    expect(book.value("Setup", 20, 10)).toBe("\u2717");
  });
});
