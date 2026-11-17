import { describe, expect, it } from "vitest";
import {
  clearPlanIdentity,
  compareHeaders,
  DATUM_SHEET,
  parseAcHeader,
  parseAcSites,
  parsePlanHeader,
  readDatum,
  writePlanIdentity,
  writeReportHeader,
  type PlanIdentity,
} from "../../workbook-commands/src/acHeader";
import { AC_DETAIL_SHEET, readAcDetail } from "../../workbook-commands/src/acDetail";
import { importAntiCollision } from "../../workbook-commands/src/acImport";
import { SheetBook } from "../../workbook-commands/src/sheetBook";
import type { SnapshotBook } from "../../workbook-commands/src/cells";

const NORTH = "100/01-02-003-04W5/0 - North Offset";

const HEADER = [
  "Company: NORTHWIND ENERGY LTD. Local Co-ordinate Reference: Well North Offset - Slot",
  "14-20",
  "Project: Demo Field TVD Reference: KBE(5.55m) @ 100.00m",
  "Reference Site: 1-2-003-04 W5M MD Reference: KBE(5.55m) @ 100.00m",
  "Site Error: 0.00 North Reference: True",
  "Reference Well: North Offset Survey Calculation Method: Minimum Curvature",
  "Well Error: 0.00 Output errors are at 2.00 sigma",
  "Reference Wellbore 100/01-02-003-04W5/00 Database: CDN Compass",
  "Reference Design: 1 P1 Offset TVD Reference: Reference Datum",
  "10 November, 2026",
  "Summary",
  "1-2-003-04 W5M",
  `${NORTH}  500.00  490.00  20.00  14.00  1.800 SF`,
  "Offset 05-06-007-08W5",
  "102/05-06-007-08W5/2 - South Offset  800.00  790.00  18.00  10.00  1.200",
  "Semi Major",
  "Coordinates are relative to: North Offset - Slot 14-20",
  "Offset Depths are relative to Offset Datum Coordinate System is Universal Transverse Mercator",
].join("\n");

const PLAN = ["Demo Field", "North Offset - Slot 14-20", "Plan: 1 P1", "North Reference: True"].join("\n");

const PLAN_IDENTITY: PlanIdentity = {
  well: "North Offset",
  slot: "14-20",
  design: "1 P1",
  northReference: "True",
  tvdReference: "",
};

function makeBook(sheets: Record<string, Record<number, Record<number, unknown>>>): SheetBook {
  const out: Record<string, { id: string; name: string; cellData: Record<number, Record<number, unknown>> }> = {};
  Object.entries(sheets).forEach(([name, cellData], index) => {
    out[`s${index}`] = { id: `s${index}`, name, cellData };
  });
  return new SheetBook({ sheets: out, styles: {} } as SnapshotBook, []);
}

describe("AC report header", () => {
  const header = parseAcHeader(HEADER);

  it("reads the two columns, including a slot name that wraps onto the next line", () => {
    expect(header).toMatchObject({
      referenceWell: "North Offset",
      referenceWellbore: "100/01-02-003-04W5/00",
      referenceDesign: "1 P1",
      referenceSite: "1-2-003-04 W5M",
      localCoordinateReference: "Well North Offset - Slot 14-20",
      coordinatesRelativeTo: "North Offset - Slot 14-20",
      tvdReference: "KBE(5.55m) @ 100.00m",
      mdReference: "KBE(5.55m) @ 100.00m",
      northReference: "True",
      offsetTvdReference: "Reference Datum",
      reportDate: "10 November, 2026",
    });
  });

  it("keeps each Summary well under the site heading above it", () => {
    expect(parseAcSites(HEADER)).toEqual([
      { site: "1-2-003-04 W5M", well: NORTH },
      { site: "Offset 05-06-007-08W5", well: "102/05-06-007-08W5/2 - South Offset" },
    ]);
  });

  it("reads the plan's well, slot and design", () => {
    expect(parsePlanHeader(PLAN)).toEqual(PLAN_IDENTITY);
  });

  it("accepts a report for the same well and design", () => {
    expect(compareHeaders(header, PLAN_IDENTITY)).toEqual({ ok: true, verdict: "Same Well and Design", reasons: [] });
  });

  it("names the report and the plan when the well or the design differs", () => {
    const check = compareHeaders(header, { ...PLAN_IDENTITY, well: "Other Well", design: "9 P9" });
    expect(check.ok).toBe(false);
    expect(check.verdict).toBe("This report is for North Offset 1 P1. The plan is Other Well 9 P9.");
  });

  it("rejects a site centre, an offset datum and a different north", () => {
    const moved = {
      ...header,
      localCoordinateReference: "Site Centre",
      coordinatesRelativeTo: "",
      offsetTvdReference: "Offset Datum",
      northReference: "Grid",
    };
    const check = compareHeaders(moved, PLAN_IDENTITY);
    expect(check.ok).toBe(false);
    expect(check.reasons).toEqual([
      "Coordinates are relative to Site Centre, not North Offset's slot.",
      "Offset TVD Reference is Offset Datum, not Reference Datum.",
      "North Reference is Grid, the plan is True.",
    ]);
  });

  it("waits when the plan's identity was never stored, and when the header was not read", () => {
    expect(compareHeaders(header, null).verdict).toBe("Plan Identity Not Stored");
    expect(compareHeaders(header, null).ok).toBe(true);
    expect(compareHeaders(parseAcHeader("no header here"), PLAN_IDENTITY).verdict).toBe("Report Header Was Not Read");
  });

  it("stores the report and the plan in separate columns, and a CSV plan clears only the plan", () => {
    const book = makeBook({});
    writeReportHeader(book, header);
    writePlanIdentity(book, PLAN_IDENTITY);
    const both = readDatum(book);
    expect(both.reportStored).toBe(true);
    expect(both.planStored).toBe(true);
    expect(both.report.referenceWell).toBe("North Offset");
    expect(both.plan).toMatchObject({ well: "North Offset", design: "1 P1", slot: "14-20" });
    expect(book.value(DATUM_SHEET, 1, 1)).toBe("Field");
    clearPlanIdentity(book);
    const cleared = readDatum(book);
    expect(cleared.planStored).toBe(false);
    expect(cleared.report.referenceDesign).toBe("1 P1");
  });

  it("writes the site onto the detail sheet at import", () => {
    const book = makeBook({
      Setup: {},
      Data: { 9: { 1: { v: "AC Info & Concerns" } }, 19: { 1: { v: "Motors On Location" } } },
      _OC_Survey: {},
    });
    const row = "  500.00  200.00  490.00  220.00  10.00  90.00  45.00  50.00  250.00  20.00  14.00  9.00  1.800 SF";
    importAntiCollision(book, "job\\ac.pdf", `${HEADER}\fOffset Design ${NORTH}\nOffset Site Pad\n${row}`, null, 1e6, true);
    expect(readAcDetail(book)[0]).toMatchObject({ well: NORTH, site: "1-2-003-04 W5M" });
    expect(readDatum(book).report.referenceDesign).toBe("1 P1");
    expect(book.value(AC_DETAIL_SHEET, 2, 10)).toBe("Site");
  });
});
