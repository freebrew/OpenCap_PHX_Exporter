import { describe, expect, it } from "vitest";
import { bhaNumbers, bhaSummary, parseOpenCapCsv, serialHours, SheetBook, writeSerialHoursTable } from "../../workbook-commands/src/index";

/** A snapshot sheet whose row 1 is the header and the rows below are the data, 0-based like Univer. */
function sheetFrom(header: string[], data: Array<Array<string | number>>): Record<number, Record<number, { v: string | number }>> {
  const rows: Record<number, Record<number, { v: string | number }>> = { 0: {} };
  header.forEach((name, col) => {
    rows[0]![col] = { v: name };
  });
  data.forEach((record, index) => {
    rows[index + 1] = {};
    record.forEach((value, col) => {
      if (value !== "") rows[index + 1]![col] = { v: value };
    });
  });
  return rows;
}

function book(): SheetBook {
  const sheets: Record<string, { id: string; name: string; cellData: ReturnType<typeof sheetFrom> }> = {
    bha: {
      id: "bha", name: "_OC_BHA", cellData: sheetFrom(
        ["BHA #", "Section", "Status", "Motor", "Metres Drilled", "BHA Total Hrs", "BHA Hrs Slid", "BHA Hrs Rot", "BHA Hrs Circ", "Serial #", "Item Code", "Description", "Sub Description", "Category", "HSLS"],
        [
          [1, "12.25", "Complete", "PHX-525010-PTS", 640, 40, 12, 28, 40, "PHX-525010-PTS", "MOT", "Mud Motor", "", "DD", 120],
          [1, "", "", "", 0, "", "", "", "", "BIT-1", "BIT", "PDC Bit", "", "DD", ""],
          [1, "", "", "", "", "", "", "", "", "MWD-9", "PRB", "Probe", "", "MWD", ""],
          [1, "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
          [2, "8.75", "Active", "PHX-525074-PTS", 210, 18, 6, 12, 18, "AG-100", "AG", "Agitator", "DD other", "", 310],
          [2, "", "", "", "", "", "", "", "", "SUB-3", "SUB", "Shock Sub", "", "", ""],
        ],
      ),
    },
    inv: { id: "inv", name: "_OC_Inventory", cellData: sheetFrom(["SerialNumber", "Category", "SubCategory", "HrsSinceService"], [["AG-100", "DD", "DD other", 50], ["SUB-3", "", "DD shock", 15]]) },
    data: { id: "data", name: "Data", cellData: sheetFrom(["Data"], []) },
  };
  return new SheetBook({ sheets, styles: {} } as never, []);
}

describe("bhaSummary (Module23 SelectBHA)", () => {
  const built = book();

  it("lists the BHA numbers and the components of the one asked for", () => {
    expect(bhaNumbers(built)).toEqual([1, 2]);
    const summary = bhaSummary(built, 1);
    expect(summary).toMatchObject({ bha: 1, section: "12.25", status: "Complete", motor: "PHX-525010-PTS", metres: 640, hrsTotal: 40, hrsSlid: 12, slideFraction: 0.3 });
    // The blank row is dropped. The component list keeps every tool on the BHA, MWD included.
    expect(summary?.components.map((row) => row.serial)).toEqual(["PHX-525010-PTS", "BIT-1", "MWD-9"]);
    expect(summary?.components[0]).toMatchObject({ hrsSlid: 12, hrsRot: 28, hrsTotal: 40 });
  });

  it("is null for a BHA the sheet does not have", () => {
    expect(bhaSummary(built, 3)).toBeNull();
  });
});

describe("serialHours (CollectBhaTabSerialRows + BuildDataSerialHoursTable)", () => {
  const built = book();

  it("keeps DD tools, drops MWD, and sorts by BHA then hours since service", () => {
    const rows = serialHours(built);
    // SUB-3 survives on its sub-description; MWD-9 does not. BHA 2 leads, and its rows fall by HSLS.
    expect(rows.map((row) => [row.bha, row.serial, row.hsls])).toEqual([[2, "AG-100", 310], [2, "SUB-3", 15], [1, "PHX-525010-PTS", 120], [1, "BIT-1", 0]]);
    expect(rows[1]).toMatchObject({ description: "Shock Sub", circHours: 18 });
  });

  it("writes Data!X2:AB and clears the rows a previous table left", () => {
    const rows = serialHours(built);
    expect(writeSerialHoursTable(built, rows)).toBe(4);
    expect([24, 25, 26, 27, 28].map((col) => built.value("Data", 2, col))).toEqual(["Serial", "Description", "Circ Hours", "HSLS", "BHA#"]);
    expect([built.value("Data", 3, 24), built.value("Data", 3, 27), built.value("Data", 3, 28)]).toEqual(["AG-100", 310, 2]);
    expect(built.value("Data", 7, 24)).toBeNull();
    writeSerialHoursTable(built, rows.slice(0, 1));
    expect(built.value("Data", 4, 24)).toBeNull();
  });

  it("bands each BHA and writes the hours as 0.00", () => {
    const rows = serialHours(built);
    writeSerialHoursTable(built, rows);
    expect(built.fill("Data", 2, 24)).toBe("#D9D9D9");
    expect(built.fill("Data", 3, 26)).toBeNull();
    expect(built.fill("Data", 5, 26)).toBe("#E2EFDA");
    expect(built.writes).toEqual(expect.arrayContaining([
      expect.objectContaining({ kind: "format", row: 3, col: 26, pattern: "0.00" }),
      expect.objectContaining({ kind: "format", row: 3, col: 27, pattern: "0.00" }),
      expect.objectContaining({ kind: "format", row: 3, col: 28, pattern: "0" }),
      expect.objectContaining({ kind: "look", row: 3, col: 25, look: expect.objectContaining({ hAlign: "left", wrap: "clip" }) }),
      expect.objectContaining({ kind: "width", col: 25, px: 360 }),
    ]));
  });

  it("reads a BHA number out of the same CSV the refresh loads", () => {
    const [header, row] = parseOpenCapCsv("BHA #,Serial #,Category\r\n1,MOT-1,DD\r\n");
    expect(header?.[0]).toBe("BHA #");
    expect(row?.[0]).toBe(1);
  });
});
