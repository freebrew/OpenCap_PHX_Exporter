import { describe, expect, it } from "vitest";
import { writePlanIdentity, writeReportHeader, type AcReportHeader } from "../../workbook-commands/src/acHeader";
import type { SnapshotBook } from "../../workbook-commands/src/cells";
import { SheetBook } from "../../workbook-commands/src/sheetBook";
import { GOOD_SURVEY_FILL } from "../../workbook-commands/src/slidesheet";
import { buildWellScene } from "../../wells-3d/src/wellScene";

function book(): SheetBook {
  const sheets = ["_OC_Survey", "_OC_AC", "_OC_AC_Detail", "_OC_PlanSec", "_OC_Formations", "Data", "Slidesheet"];
  const out: Record<string, { id: string; name: string; cellData: Record<string, never> }> = {};
  sheets.forEach((name, index) => {
    out[`s${index}`] = { id: `s${index}`, name, cellData: {} };
  });
  return new SheetBook({ sheets: out, styles: {} } as SnapshotBook, []);
}

describe("well scene", () => {
  it("places the plan, the drilled path and the concern wells in local metres", () => {
    const sheet = book();
    sheet.setValue("_OC_Survey", 3, 1, 0);
    sheet.setValue("_OC_Survey", 3, 4, 0);
    sheet.setValue("_OC_Survey", 3, 5, 0);
    sheet.setValue("_OC_Survey", 3, 6, 0);
    sheet.setValue("_OC_Survey", 4, 1, 1000);
    sheet.setValue("_OC_Survey", 4, 4, 400);
    sheet.setValue("_OC_Survey", 4, 5, 100);
    sheet.setValue("_OC_Survey", 4, 6, 500);
    sheet.setValue("_OC_PlanSec", 3, 1, 500);
    sheet.setValue("_OC_PlanSec", 3, 11, "HEEL");

    sheet.setValue("Data", 10, 2, "AC Info & Concerns");
    sheet.setValue("Data", 14, 2, "Motors On Location");
    sheet.setValue("_OC_AC", 3, 1, "103/09-10-011-12W5/0 - East Offset");
    sheet.setValue("_OC_AC", 3, 2, 100);
    sheet.setValue("_OC_AC", 3, 3, 15);
    sheet.setValue("_OC_AC", 3, 4, 1.1);
    sheet.setValue("_OC_AC", 4, 1, "100/01-02-003-04W5/0 - North Offset");
    sheet.setValue("_OC_AC", 4, 2, 500);
    sheet.setValue("_OC_AC", 4, 3, 20);
    sheet.setValue("_OC_AC", 4, 4, 1.2);
    sheet.setValue("_OC_AC", 5, 1, "102/05-06-007-08W5/2 - South Offset");
    sheet.setValue("_OC_AC", 5, 2, 800);
    sheet.setValue("_OC_AC", 5, 3, 40);
    sheet.setValue("_OC_AC", 5, 4, 4.5);

    sheet.setValue("_OC_AC_Detail", 3, 1, "100/01-02-003-04W5/0 - North Offset");
    sheet.setValue("_OC_AC_Detail", 3, 2, 500);
    sheet.setValue("_OC_AC_Detail", 3, 3, 490);
    sheet.setValue("_OC_AC_Detail", 3, 4, 220);
    sheet.setValue("_OC_AC_Detail", 3, 5, 50);
    sheet.setValue("_OC_AC_Detail", 3, 6, 250);
    sheet.setValue("_OC_AC_Detail", 3, 7, 20);
    sheet.setValue("_OC_AC_Detail", 3, 8, 1.2);
    sheet.setValue("_OC_AC_Detail", 4, 1, "102/05-06-007-08W5/2 - South Offset");
    sheet.setValue("_OC_AC_Detail", 4, 2, 800);
    sheet.setValue("_OC_AC_Detail", 4, 3, 790);
    // 40 m straight below the plan at Ref MD 800 (80 N, 400 E, 320 down), so the printed C2C is true.
    sheet.setValue("_OC_AC_Detail", 4, 4, 360);
    sheet.setValue("_OC_AC_Detail", 4, 5, 80);
    sheet.setValue("_OC_AC_Detail", 4, 6, 400);
    sheet.setValue("_OC_AC_Detail", 4, 7, 40);
    sheet.setValue("_OC_AC_Detail", 4, 8, 4.5);

    for (const row of [13, 14]) {
      sheet.setValue("Slidesheet", row, 5, row === 13 ? 100 : 200);
      sheet.setValue("Slidesheet", row, 6, 0);
      sheet.setValue("Slidesheet", row, 7, 0);
      sheet.setFill("Slidesheet", { row, col: 6, row2: row, col2: 6 }, GOOD_SURVEY_FILL);
      sheet.setFill("Slidesheet", { row, col: 7, row2: row, col2: 7 }, GOOD_SURVEY_FILL);
    }

    const scene = buildWellScene(sheet);
    expect(scene).not.toBeNull();
    if (!scene) return;
    expect(scene.plan.xyz[3]).toBeCloseTo(500);
    expect(scene.plan.xyz[4]).toBeCloseTo(100);
    expect(scene.plan.xyz[5]).toBeCloseTo(-400);
    expect(scene.drilled?.xyz[2]).toBeCloseTo(-100);
    expect(scene.drilled?.md[1]).toBeCloseTo(200);

    const north = scene.offsets.find((offset) => offset.name.includes("North"));
    const south = scene.offsets.find((offset) => offset.name.includes("South"));
    expect(north).toMatchObject({ concern: true, color: "#C62828", bandSf: 1.2 });
    const heel = scene.targets.find((target) => target.name === "HEEL");
    expect(heel).toMatchObject({ md: 500, x: 250, y: 50, z: -200 });
    const dx = (heel?.outer[0] ?? 0) - 250;
    const dy = (heel?.outer[1] ?? 0) - 50;
    const dz = (heel?.outer[2] ?? 0) + 200;
    expect(heel?.radius).toBe(20);
    expect(Math.hypot(dx, dy, dz)).toBeCloseTo(20);
    const ix = (heel?.inner[0] ?? 0) - 250;
    const iy = (heel?.inner[1] ?? 0) - 50;
    const iz = (heel?.inner[2] ?? 0) + 200;
    expect(Math.hypot(ix, iy, iz)).toBeCloseTo(10);
    expect(Math.abs(dx * 500 + dy * 100 + dz * -400)).toBeLessThan(0.05);
    expect(south?.concern).toBe(false);
    expect(south?.color).toBe("#2E7D32");
    expect(scene.missing).toEqual(["103/09-10-011-12W5/0 - East Offset"]);
    expect(scene.badgeOk).toBe(false);
    expect(scene.badge).toBe("A concern well has no path");
    expect(scene.hash).toMatch(/^[0-9a-f]{8}$/);
    expect(buildWellScene(sheet)?.hash).toBe(scene.hash);
  });

  it("draws nothing until the plan has two stations", () => {
    expect(buildWellScene(book())).toBeNull();
  });

  it("reads header-shifted _OC_Survey columns (Sub-Sea before TVD/NS/EW)", () => {
    const sheet = book();
    ["MD", "INC", "AZI", "Sub-Sea", "TVD", "NS", "EW"].forEach((head, idx) => sheet.setValue("_OC_Survey", 2, idx + 1, head));
    sheet.setValue("_OC_Survey", 3, 1, 0);
    sheet.setValue("_OC_Survey", 3, 2, 0);
    sheet.setValue("_OC_Survey", 3, 3, 0);
    sheet.setValue("_OC_Survey", 3, 4, 891.4);
    sheet.setValue("_OC_Survey", 3, 5, 0);
    sheet.setValue("_OC_Survey", 3, 6, 0);
    sheet.setValue("_OC_Survey", 3, 7, 0);
    sheet.setValue("_OC_Survey", 4, 1, 100);
    sheet.setValue("_OC_Survey", 4, 2, 90);
    sheet.setValue("_OC_Survey", 4, 3, 90);
    sheet.setValue("_OC_Survey", 4, 4, 827.74);
    sheet.setValue("_OC_Survey", 4, 5, 63.66);
    sheet.setValue("_OC_Survey", 4, 6, 0.25);
    sheet.setValue("_OC_Survey", 4, 7, 63.8);
    const scene = buildWellScene(sheet);
    expect(scene).not.toBeNull();
    // Exact Compass NS/EW within 2 m of min-curvature are preserved (not overwritten).
    expect(scene!.plan.xyz[3]).toBeCloseTo(63.8, 2);
    expect(scene!.plan.xyz[4]).toBeCloseTo(0.25, 2);
    expect(scene!.plan.xyz[5]).toBeCloseTo(-63.66, 2);
  });

  it("rebuilds NS/EW from MD/INC/AZI when a legacy Crystal import stored Sub-Sea in EW", () => {
    const sheet = book();
    ["MD", "INC", "AZI", "TVD", "NS", "EW"].forEach((head, idx) => sheet.setValue("_OC_Survey", 2, idx + 1, head));
    // Legacy glued-Subsea corruption: NS = 0 and EW = Sub-Sea (891.82 down to 791.82)
    sheet.setValue("_OC_Survey", 3, 1, 0);
    sheet.setValue("_OC_Survey", 3, 2, 0);
    sheet.setValue("_OC_Survey", 3, 3, 0);
    sheet.setValue("_OC_Survey", 3, 4, 0);
    sheet.setValue("_OC_Survey", 3, 5, 0);
    sheet.setValue("_OC_Survey", 3, 6, 891.82);
    sheet.setValue("_OC_Survey", 4, 1, 1000);
    sheet.setValue("_OC_Survey", 4, 2, 0);
    sheet.setValue("_OC_Survey", 4, 3, 0);
    sheet.setValue("_OC_Survey", 4, 4, 1000);
    sheet.setValue("_OC_Survey", 4, 5, 0);
    sheet.setValue("_OC_Survey", 4, 6, -108.18);
    sheet.setValue("_OC_Survey", 5, 1, 1100);
    sheet.setValue("_OC_Survey", 5, 2, 90);
    sheet.setValue("_OC_Survey", 5, 3, 180);
    sheet.setValue("_OC_Survey", 5, 4, 1063.66);
    sheet.setValue("_OC_Survey", 5, 5, 0);
    sheet.setValue("_OC_Survey", 5, 6, -171.84);
    const scene = buildWellScene(sheet);
    expect(scene).not.toBeNull();
    // Origin is (0, 0, 0), and station 3 (100 m build 0->90 deg due South) is ~63.66 m South, 0 m East.
    expect(scene!.plan.xyz[0]).toBeCloseTo(0, 3);
    expect(scene!.plan.xyz[1]).toBeCloseTo(0, 3);
    expect(scene!.plan.xyz[6]).toBeCloseTo(0, 1);
    expect(scene!.plan.xyz[7]).toBeCloseTo(-63.66, 1);
    expect(scene!.plan.xyz[8]).toBeCloseTo(-1063.66, 1);
  });

  it("integrates the drilled path from Slidesheet row 12 tie-on", () => {
    const sheet = book();
    sheet.setValue("_OC_Survey", 3, 1, 0);
    sheet.setValue("_OC_Survey", 3, 4, 0);
    sheet.setValue("_OC_Survey", 3, 5, 0);
    sheet.setValue("_OC_Survey", 3, 6, 0);
    sheet.setValue("_OC_Survey", 4, 1, 1000);
    sheet.setValue("_OC_Survey", 4, 4, 1000);
    sheet.setValue("_OC_Survey", 4, 5, 0);
    sheet.setValue("_OC_Survey", 4, 6, 0);
    // Row 12 tie-on: MD 500, Inc 90, Azi 90 (due East), TVD 495, NS 25 (AL12), EW 40 (AM12)
    sheet.setValue("Slidesheet", 12, 5, 500);
    sheet.setValue("Slidesheet", 12, 6, 90);
    sheet.setValue("Slidesheet", 12, 7, 90);
    sheet.setValue("Slidesheet", 12, 8, 495);
    sheet.setValue("Slidesheet", 12, 38, 25);
    sheet.setValue("Slidesheet", 12, 39, 40);
    // First survey at MD 600 holding 90 deg Inc, 90 deg Azi -> +100 m East
    sheet.setValue("Slidesheet", 13, 5, 600);
    sheet.setValue("Slidesheet", 13, 6, 90);
    sheet.setValue("Slidesheet", 13, 7, 90);
    sheet.setFill("Slidesheet", { row: 13, col: 6, row2: 13, col2: 6 }, GOOD_SURVEY_FILL);
    sheet.setFill("Slidesheet", { row: 13, col: 7, row2: 13, col2: 7 }, GOOD_SURVEY_FILL);

    const scene = buildWellScene(sheet);
    expect(scene?.drilled).not.toBeNull();
    // Point 0 is the row 12 tie-on (E=40, N=25, z=-495); Point 1 is row 13 (E=140, N=25, z=-495).
    expect(scene!.drilled!.md[0]).toBeCloseTo(500);
    expect(scene!.drilled!.xyz[0]).toBeCloseTo(40, 2);
    expect(scene!.drilled!.xyz[1]).toBeCloseTo(25, 2);
    expect(scene!.drilled!.xyz[2]).toBeCloseTo(-495, 2);
    expect(scene!.drilled!.md[1]).toBeCloseTo(600);
    expect(scene!.drilled!.xyz[3]).toBeCloseTo(140, 2);
    expect(scene!.drilled!.xyz[4]).toBeCloseTo(25, 2);
    expect(scene!.drilled!.xyz[5]).toBeCloseTo(-495, 2);
  });

  it("puts a formation top on the scene at its TVD above the first plan station", () => {
    const sheet = book();
    sheet.setValue("_OC_Survey", 3, 1, 0);
    sheet.setValue("_OC_Survey", 3, 4, 0);
    sheet.setValue("_OC_Survey", 3, 5, 0);
    sheet.setValue("_OC_Survey", 3, 6, 0);
    sheet.setValue("_OC_Survey", 4, 1, 1000);
    sheet.setValue("_OC_Survey", 4, 4, 400);
    sheet.setValue("_OC_Survey", 4, 5, 100);
    sheet.setValue("_OC_Survey", 4, 6, 500);
    sheet.setValue("_OC_Formations", 3, 1, 850);
    sheet.setValue("_OC_Formations", 3, 2, 200);
    sheet.setValue("_OC_Formations", 3, 3, "Belly River");
    sheet.setValue("_OC_Formations", 3, 4, "sandstone");
    sheet.setValue("_OC_Formations", 4, 1, 900);
    sheet.setValue("_OC_Formations", 4, 2, 210);
    sheet.setValue("_OC_Formations", 4, 3, "  ");
    const scene = buildWellScene(sheet);
    expect(scene?.formations).toEqual([{ name: "Belly River", lith: "sandstone", tvd: 200, z: -200 }]);
    sheet.setValue("_OC_Formations", 4, 3, "Second White Specks");
    expect(buildWellScene(sheet)?.hash).not.toBe(scene?.hash);
  });
});

const NORTH = "100/01-02-003-04W5/0 - North Offset";
const SOUTH = "102/05-06-007-08W5/2 - South Offset";
/** Metres of plan per metre of measured depth on the invented plan below: |(0.1, 0.5, 0.4)|. */
const PLAN_RATE = Math.sqrt(0.42);

interface Station {
  well: string;
  refMd: number;
  offMd: number;
  /** Printed C2C; the point is placed 20 m north of the plan unless this says otherwise. */
  c2c?: number;
  n?: number;
  e?: number;
  offTvd?: number;
  sf?: number;
  site?: string;
}

/** Plan from the origin to (100 N, 500 E, 400 down) over 1000 m. Offset points sit 20 m north of it. */
function sceneWith(stations: Station[], concern: boolean) {
  const sheet = book();
  sheet.setValue("_OC_Survey", 3, 1, 0);
  sheet.setValue("_OC_Survey", 3, 4, 0);
  sheet.setValue("_OC_Survey", 3, 5, 0);
  sheet.setValue("_OC_Survey", 3, 6, 0);
  sheet.setValue("_OC_Survey", 4, 1, 1000);
  sheet.setValue("_OC_Survey", 4, 4, 400);
  sheet.setValue("_OC_Survey", 4, 5, 100);
  sheet.setValue("_OC_Survey", 4, 6, 500);
  if (concern) {
    sheet.setValue("Data", 10, 2, "AC Info & Concerns");
    sheet.setValue("Data", 14, 2, "Motors On Location");
    sheet.setValue("_OC_AC", 3, 1, NORTH);
    sheet.setValue("_OC_AC", 3, 2, 300);
    sheet.setValue("_OC_AC", 3, 3, 20);
    sheet.setValue("_OC_AC", 3, 4, 1.5);
  }
  stations.forEach((station, i) => {
    const row = 3 + i;
    sheet.setValue("_OC_AC_Detail", row, 1, station.well);
    sheet.setValue("_OC_AC_Detail", row, 2, station.refMd);
    sheet.setValue("_OC_AC_Detail", row, 3, station.offMd);
    sheet.setValue("_OC_AC_Detail", row, 4, station.offTvd ?? station.refMd * 0.4);
    sheet.setValue("_OC_AC_Detail", row, 5, station.n ?? station.refMd * 0.1 + 20);
    sheet.setValue("_OC_AC_Detail", row, 6, station.e ?? station.refMd * 0.5);
    sheet.setValue("_OC_AC_Detail", row, 7, station.c2c ?? 20);
    sheet.setValue("_OC_AC_Detail", row, 8, station.sf ?? 1.5);
    sheet.setValue("_OC_AC_Detail", row, 10, station.site ?? "");
  });
  return buildWellScene(sheet);
}

const REPORT: AcReportHeader = {
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
};

function sceneWithHeader(stations: Station[], planWell: string, planDesign: string) {
  const sheet = book();
  sheet.setValue("_OC_Survey", 3, 1, 0);
  sheet.setValue("_OC_Survey", 3, 4, 0);
  sheet.setValue("_OC_Survey", 3, 5, 0);
  sheet.setValue("_OC_Survey", 3, 6, 0);
  sheet.setValue("_OC_Survey", 4, 1, 1000);
  sheet.setValue("_OC_Survey", 4, 4, 400);
  sheet.setValue("_OC_Survey", 4, 5, 100);
  sheet.setValue("_OC_Survey", 4, 6, 500);
  stations.forEach((station, i) => {
    const row = 3 + i;
    sheet.setValue("_OC_AC_Detail", row, 1, station.well);
    sheet.setValue("_OC_AC_Detail", row, 2, station.refMd);
    sheet.setValue("_OC_AC_Detail", row, 3, station.offMd);
    sheet.setValue("_OC_AC_Detail", row, 4, station.offTvd ?? station.refMd * 0.4);
    sheet.setValue("_OC_AC_Detail", row, 5, station.n ?? station.refMd * 0.1 + 20);
    sheet.setValue("_OC_AC_Detail", row, 6, station.e ?? station.refMd * 0.5);
    sheet.setValue("_OC_AC_Detail", row, 7, station.c2c ?? 20);
    sheet.setValue("_OC_AC_Detail", row, 8, station.sf ?? 1.5);
    sheet.setValue("_OC_AC_Detail", row, 10, station.site ?? "");
  });
  writeReportHeader(sheet, REPORT);
  if (planWell !== "" || planDesign !== "") {
    writePlanIdentity(sheet, { well: planWell, slot: "14-20", design: planDesign, northReference: "True", tvdReference: "" });
  }
  return buildWellScene(sheet);
}

describe("offset paths", () => {
  it("run along the offset's own depth, not the reference's, so a line never doubles back", () => {
    const stations = [100, 200, 300, 400, 500].map((refMd, i) => ({ well: NORTH, refMd, offMd: 500 - i * 100 }));
    const north = sceneWith(stations, false)?.offsets[0];
    expect(Array.from(north?.md ?? [])).toEqual([500, 400, 300, 200, 100]);
  });

  it("break where the report gives nothing for a stretch of the offset, and join a straight run", () => {
    const stations: Station[] = [
      { well: NORTH, refMd: 100, offMd: 100 },
      { well: NORTH, refMd: 200, offMd: 150 },
      { well: NORTH, refMd: 500, offMd: 400 },
      { well: SOUTH, refMd: 100, offMd: 100 * PLAN_RATE },
      { well: SOUTH, refMd: 300, offMd: 300 * PLAN_RATE },
      { well: SOUTH, refMd: 700, offMd: 700 * PLAN_RATE },
    ];
    const scene = sceneWith(stations, false);
    const north = scene?.offsets.find((offset) => offset.name.includes("North"));
    const south = scene?.offsets.find((offset) => offset.name.includes("South"));
    expect(north?.breaks).toEqual([2]);
    expect(south?.breaks).toEqual([]);
  });

  it("joins a curved survey leg and leaves a jump across a missing stretch open", () => {
    const scene = sceneWith([
      { well: NORTH, refMd: 100, offMd: 100 },
      { well: NORTH, refMd: 390, offMd: 300 },
      { well: SOUTH, refMd: 100, offMd: 100 },
      { well: SOUTH, refMd: 180, offMd: 400 },
    ], false);
    const north = scene?.offsets.find((offset) => offset.name.includes("North"));
    const south = scene?.offsets.find((offset) => offset.name.includes("South"));
    expect(north?.breaks).toEqual([]);
    expect(south?.breaks).toEqual([1]);
  });

  it("leave out a row that is not where its printed C2C says, and say so", () => {
    const stations: Station[] = [
      // Read wrong: the point sits on the active well's surface, with no C2C and no SF.
      { well: NORTH, refMd: 0, offMd: 0.1, n: 0.1, e: 0, offTvd: 0, c2c: 0, sf: 0 },
      { well: NORTH, refMd: 100, offMd: 30 },
      { well: NORTH, refMd: 200, offMd: 60 },
      { well: NORTH, refMd: 300, offMd: 90 },
      // 12 m off the plan frame the C2C describes.
      { well: NORTH, refMd: 400, offMd: 120, n: 400 * 0.1 + 32 },
    ];
    const scene = sceneWith(stations, true);
    const north = scene?.offsets[0];
    expect(scene?.leftOut).toBe(2);
    expect(Array.from(north?.md ?? [])).toEqual([100, 200, 300]);
    // The import check still fails on the row that was off the frame, and the note says what was left out.
    expect(scene?.badge).toBe("Report and plan do not line up, 2 rows left out");
    expect(scene?.badgeOk).toBe(false);
  });

  it("lists the tightest SF last where several reference stations share one offset point", () => {
    // The point (30 N, 50 E, 40 down) is 20 m from the plan at Ref MD 100 and 64.8 m from it at Ref MD 200.
    const stations: Station[] = [
      { well: NORTH, refMd: 200, offMd: 40, n: 30, e: 50, offTvd: 40, c2c: Math.hypot(10, 50, 40), sf: 1.4 },
      { well: NORTH, refMd: 100, offMd: 40, n: 30, e: 50, offTvd: 40, c2c: 20, sf: 3 },
    ];
    const north = sceneWith(stations, false)?.offsets[0];
    expect(Array.from(north?.md ?? [])).toEqual([100, 200]);
    expect(Array.from(north?.sf ?? [])).toEqual([Math.fround(3), Math.fround(1.4)]);
  });

  it("reads the wellhead from the surface row and whether the well is on the pad", () => {
    const scene = sceneWithHeader([
      { well: NORTH, refMd: 0, offMd: 0, n: 0, e: 20, offTvd: 0, c2c: 20, sf: 0, site: "1-2-003-04 W5M" },
      { well: NORTH, refMd: 100, offMd: 30, site: "1-2-003-04 W5M" },
      { well: SOUTH, refMd: 200, offMd: 100, site: "Offset 05-06-007-08W5" },
      { well: SOUTH, refMd: 300, offMd: 160, site: "Offset 05-06-007-08W5" },
    ], "North Offset", "1 P1");
    const north = scene?.offsets.find((offset) => offset.name.includes("North"));
    const south = scene?.offsets.find((offset) => offset.name.includes("South"));
    expect(north).toMatchObject({ site: "1-2-003-04 W5M", onPad: true, minC2c: 20 });
    expect(north?.wellhead).toEqual({ north: 0, east: 20, distance: 20, bearing: 90 });
    expect(south).toMatchObject({ onPad: false, wellhead: null });
    expect(scene?.report).toMatchObject({
      stored: true,
      referenceWell: "North Offset",
      design: "1 P1",
      verdict: "Same Well and Design",
      verdictOk: true,
    });
  });

  it("asks for the header when the report was stored before the header was, and names a different plan", () => {
    const rows: Station[] = [
      { well: NORTH, refMd: 100, offMd: 30 },
      { well: NORTH, refMd: 200, offMd: 60 },
    ];
    expect(sceneWith(rows, false)?.report).toMatchObject({
      stored: false,
      verdict: "Run Import AC again to store the report header",
    });
    const other = sceneWithHeader(rows, "Other Well", "9 P9");
    expect(other?.badge).toBe("This report is for North Offset 1 P1. The plan is Other Well 9 P9.");
    expect(other?.badgeOk).toBe(false);
    expect(sceneWithHeader(rows, "", "")?.report.verdict).toBe("Plan Identity Not Stored");
    expect(sceneWithHeader(rows, "", "")?.badgeOk).toBe(true);
  });
});
