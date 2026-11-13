import { describe, expect, it } from "vitest";
import type { SnapshotBook } from "../../workbook-commands/src/cells";
import { SheetBook } from "../../workbook-commands/src/sheetBook";
import { GOOD_SURVEY_FILL } from "../../workbook-commands/src/slidesheet";
import { buildWellScene } from "../../wells-3d/src/wellScene";

function book(): SheetBook {
  const sheets = ["_OC_Survey", "_OC_AC", "_OC_AC_Detail", "Data", "Slidesheet"];
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
    sheet.setValue("_OC_AC_Detail", 4, 4, 300);
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
    expect(north?.marker).toEqual({ x: 250, y: 50, z: -220 });
    expect(north?.join?.x).toBeCloseTo(250);
    expect(north?.join?.y).toBeCloseTo(50);
    expect(north?.join?.z).toBeCloseTo(-200);
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
});
