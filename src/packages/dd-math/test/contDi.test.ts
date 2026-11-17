import { describe, expect, it } from "vitest";
import { contDiTargetHit, followSlideTarget, pullLastSurvey, SheetBook } from "../../workbook-commands/src/index";

function sheet(rows: Record<number, Record<number, { v?: number | string; f?: string; s?: string }>>): { id: string; name: string; cellData: typeof rows } {
  return { id: "", name: "", cellData: rows };
}

function book(): SheetBook {
  const green = { s: "green" };
  const slidesheet = sheet({
    // Rows are 0-based: Slidesheet row 13 is index 12.
    12: { 4: { v: 1000 }, 5: { v: 10, ...green }, 6: { v: 90, ...green }, 7: { v: 800 }, 2: { v: 30 } },
    13: { 4: { v: 1090 }, 5: { v: 12, ...green }, 6: { v: 92, ...green }, 7: { v: 860 }, 2: { v: 28 } },
    14: { 4: { v: 1180 }, 5: { v: 40, s: "yellow" }, 6: { v: 100, s: "yellow" }, 7: { v: 900 } },
  });
  const contDi = sheet({
    9: { 0: { v: "typed A" }, 2: { v: 5 }, 5: { v: 1 } },
    19: { 0: { v: "later A" }, 3: { v: 7 } },
  });
  const data = sheet({});
  return new SheetBook({
    sheets: { ss: { ...slidesheet, id: "ss", name: "Slidesheet" }, cd: { ...contDi, id: "cd", name: "Cont DI" }, h: { ...data, id: "h", name: "ContDI_Data" } },
    styles: { green: { bg: { rgb: "CCFFCC" } }, yellow: { bg: { rgb: "FFFF00" } } },
  } as never, []);
}

describe("pullLastSurvey (ContDI_PullLastSurvey)", () => {
  it("ties on to the newest green survey and keeps the last typed grid row", () => {
    const built = book();
    const result = pullLastSurvey(built);
    // Row 15 is yellow, not green, so the tie-on is row 14.
    expect(result).toEqual({ surveyRow: 14, keptRow: 20 });
    expect([1, 2, 3, 4].map((col) => built.value("Cont DI", 2, col))).toEqual([1090, 12, 92, 860]);
    // The deepest typed row moved to row 8; the rows below it cleared.
    expect(built.value("Cont DI", 8, 1)).toBe("later A");
    expect(built.value("Cont DI", 8, 4)).toBe(7);
    expect(built.value("Cont DI", 10, 1)).toBeNull();
    expect(built.value("Cont DI", 20, 1)).toBeNull();
  });

  it("reports no survey when nothing is green", () => {
    const built = new SheetBook({ sheets: { cd: { id: "cd", name: "Cont DI", cellData: {} } }, styles: {} } as never, []);
    expect(pullLastSurvey(built)?.surveyRow).toBe(0);
  });
});

describe("followSlideTarget (ContDI_ApplyTargetSelection)", () => {
  it("only reacts to the yellow target block", () => {
    expect(contDiTargetHit(2, 16)).toBe(true);
    expect(contDiTargetHit(3, 19)).toBe(true);
    expect(contDiTargetHit(5, 19)).toBe(false);
    expect(contDiTargetHit(2, 15)).toBe(false);
  });

  it("writes the aim formulas once and leaves them on a second pass", () => {
    const built = book();
    expect(followSlideTarget(built)).toBe(true);
    expect(built.formula("Cont DI", 2, 16)).toContain("ProjTargets_NAME");
    expect(built.formula("Cont DI", 1, 21)).toContain('ProjContDiAim("BRR"');
    expect(built.formula("ContDI_Data", 2, 24)).toContain('ProjContDiAim("TRRS"');
    expect(followSlideTarget(built)).toBe(false);
  });

});
