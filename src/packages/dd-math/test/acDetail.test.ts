import { describe, expect, it } from "vitest";
import { checkAcReport, interpPlan } from "../../workbook-commands/src/acCheck";
import { AC_DETAIL_SHEET, parseAcDetail, readAcDetail, wellKey, wellsMatch } from "../../workbook-commands/src/acDetail";
import { AC_SHEET, acConcernCapacity, concernWells, importAntiCollision, type AcHit } from "../../workbook-commands/src/acImport";
import { SheetBook, type CellWrite } from "../../workbook-commands/src/sheetBook";
import type { SnapshotBook } from "../../workbook-commands/src/cells";

function makeBook(sheets: Record<string, Record<number, Record<number, unknown>>>): SheetBook {
  const out: Record<string, { id: string; name: string; cellData: Record<number, Record<number, unknown>> }> = {};
  Object.entries(sheets).forEach(([name, cellData], index) => {
    out[`s${index}`] = { id: `s${index}`, name, cellData };
  });
  return new SheetBook({ sheets: out, styles: {} } as SnapshotBook, []);
}

function valueAt(writes: CellWrite[], sheet: string, row: number, col: number): unknown {
  const hit = [...writes].reverse().find((write) => write.kind === "value" && write.sheet === sheet && write.row === row && write.col === col);
  return hit && hit.kind === "value" ? hit.value : undefined;
}

const NORTH = "100/01-02-003-04W5/0 - North Offset";
const SOUTH = "102/05-06-007-08W5/2 - South Offset";

function page(well: string, row: string): string {
  return [`Offset Design ${well}`, "Offset Site Pad", row, ""].join("\n");
}

const NORTH_ROW = "  500.00  200.00  490.00  220.00  10.00  90.00  45.00  50.00  250.00  20.00  14.00  9.00  1.800 SF";
const SOUTH_ROW = "  800.00  300.00  790.00  310.00  20.00  90.00  12.00  -8.25  30.00  18.00  10.00  1. 200";

describe("AC detail pages", () => {
  const text = `${page(NORTH, NORTH_ROW)}\f${page(SOUTH, SOUTH_ROW)}`;

  it("reads a 13-number station and a 12-number station, including a stray space in the SF", () => {
    const stations = parseAcDetail(text);
    expect(stations).toHaveLength(2);
    expect(stations[0]).toMatchObject({ well: NORTH, refMd: 500, offTvd: 220, n: 50, e: 250, c2c: 20, sf: 1.8, warn: "SF" });
    expect(stations[1]).toMatchObject({ well: SOUTH, refMd: 800, n: 12, e: -8.25, c2c: 30, sf: 1.2, warn: "" });
  });

  it("matches a well by its wellbore id", () => {
    expect(wellKey(NORTH)).toBe("100/01-02-003-04W5");
    expect(wellsMatch(`Offset Design ${NORTH}`, "100/01-02-003-04W5/0")).toBe(true);
    expect(wellsMatch(NORTH, SOUTH)).toBe(false);
  });

  it("reads a short row by the header columns", () => {
    const header = `${" ".repeat(56)}+N/-S   +E/-W     Centres Ellipses Factor`;
    const line = `${"1200.00".padEnd(9)}${"1180.00".padEnd(9)}${"1190.00".padEnd(9)}${"1170.00".padEnd(9)}${"10.00".padEnd(7)}${"90.00".padEnd(13)}12.50   -8.25     20.00   15.00    1.800`;
    const stations = parseAcDetail([`Offset Design ${NORTH}`, "Offset Site Pad", header, line].join("\n"));
    expect(stations).toHaveLength(1);
    expect(stations[0]).toMatchObject({ refMd: 1200, n: 12.5, e: -8.25, c2c: 20, sf: 1.8 });
  });

  it("stores the detail sheet and the check badge on import", () => {
    const book = makeBook({
      Setup: {},
      Data: { 9: { 1: { v: "AC Info & Concerns" } }, 19: { 1: { v: "Motors On Location" } } },
      _OC_Survey: {},
    });
    book.setValue("_OC_Survey", 3, 1, 0);
    book.setValue("_OC_Survey", 3, 4, 0);
    book.setValue("_OC_Survey", 3, 5, 0);
    book.setValue("_OC_Survey", 3, 6, 0);
    book.setValue("_OC_Survey", 4, 1, 1000);
    book.setValue("_OC_Survey", 4, 4, 400);
    book.setValue("_OC_Survey", 4, 5, 100);
    book.setValue("_OC_Survey", 4, 6, 500);
    const summary = [
      "Summary",
      `  ${NORTH}   500.00   490.00   20.00   14.00   1.800 SF`,
      "Semi Major",
    ].join("\n");
    const result = importAntiCollision(book, "job\\ac.pdf", `${summary}\f${page(NORTH, NORTH_ROW)}`, null, 1e6, true);
    expect(result.badge).toBe("Checked: 1 stations, worst 0.00 m");
    expect(valueAt(book.writes, AC_DETAIL_SHEET, 3, 1)).toBe(NORTH);
    expect(valueAt(book.writes, AC_DETAIL_SHEET, 3, 7)).toBe(20);
    expect(book.value(AC_SHEET, 1, 6)).toBe(result.badge);
    expect(readAcDetail(book)[0]).toMatchObject({ well: NORTH, c2c: 20, sf: 1.8 });
  });
});

describe("AC checks and the concern rule", () => {
  const plan = [{ md: 0, n: 0, e: 0, tvd: 0 }, { md: 1000, n: 100, e: 500, tvd: 400 }];
  const hit = (well: string, refMd: number, between: number, sf: number): AcHit => ({ well, refMd, between, sf });

  it("interpolates the plan in measured depth", () => {
    expect(interpPlan(plan, 500)).toEqual({ md: 500, n: 50, e: 250, tvd: 200 });
  });

  it("passes when the 3D distance is the printed centre-to-centre", () => {
    const stations = parseAcDetail(page(NORTH, NORTH_ROW));
    const check = checkAcReport([hit(NORTH, 500, 20, 1.8)], stations, plan);
    expect(check.ok).toBe(true);
    expect(check.badge).toBe("Checked: 1 stations, worst 0.00 m");
  });

  it("fails the frame when the offset point is not where the report says", () => {
    const moved = NORTH_ROW.replace("  220.00", "  250.00");
    const check = checkAcReport([hit(NORTH, 500, 20, 1.8)], parseAcDetail(page(NORTH, moved)), plan);
    expect(check.ok).toBe(false);
    expect(check.badge).toBe("Report and plan do not line up");
  });

  it("fails when a concern well has no detail page", () => {
    const check = checkAcReport([hit(SOUTH, 800, 18, 1.2)], [], plan);
    expect(check.badge).toBe("A concern well has no path");
  });

  it("keeps the lowest separation factors, then lists them by depth", () => {
    const hits = [hit(SOUTH, 800, 18, 1.2), hit(NORTH, 500, 20, 1.8), hit("far", 100, 40, 4)];
    expect(concernWells(hits, 2).map((item) => item.refMd)).toEqual([500, 800]);
    const book = makeBook({ Data: { 9: { 1: { v: "AC Info & Concerns" } }, 13: { 1: { v: "Motors On Location" } } } });
    expect(acConcernCapacity(book)).toBe(2);
  });
});
