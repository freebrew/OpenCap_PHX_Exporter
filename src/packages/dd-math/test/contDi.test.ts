import { describe, expect, it } from "vitest";
import { minCurveStep } from "../src/angles";
import { contDiTvdAtBit } from "../src/project";
import { applyContDiHeldReadings, contDiCalcFormula, contDiTargetHit, followSlideTarget, paintContDiColumnA, pullLastSurvey, SheetBook } from "../../workbook-commands/src/index";

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

describe("paintContDiColumnA", () => {
  it("paints Bit Depth yellow, centred, and 0.00", () => {
    const built = new SheetBook({ sheets: { cd: { id: "cd", name: "Cont DI", cellData: { 7: { 0: { v: 3768.8 } } } } }, styles: {} } as never, []);
    expect(paintContDiColumnA(built)).toBe(true);
    const fills = built.writes.filter((write) => write.kind === "fill");
    expect(fills).toEqual([
      expect.objectContaining({ row: 2, col: 1, fill: "#FFFFCC" }),
      expect.objectContaining({ row: 8, row2: 44, col: 1, fill: "#FFFFCC" }),
    ]);
    expect(built.writes.some((write) => write.kind === "format" && write.row === 8 && write.pattern === "0.00")).toBe(true);
  });
});

describe("held inclination and azimuth", () => {
  function balanced(formula: string): boolean {
    let depth = 0;
    for (const ch of formula) {
      if (ch === "(") depth += 1;
      if (ch === ")") depth -= 1;
      if (depth < 0) return false;
    }
    return depth === 0;
  }

  it("carries the last MWD inclination and leaves a never-typed PD inclination blank", () => {
    const mwd = contDiCalcFormula(8, 8) ?? "";
    const pd = contDiCalcFormula(8, 15) ?? "";
    const next = contDiCalcFormula(9, 11) ?? "";
    expect(mwd).toContain("IF(ISNUMBER(C8),C8,$B$2)");
    expect(mwd).toContain("IF(ISNUMBER(D8),D8,$C$2)");
    expect(pd).toContain('IF(ISNUMBER(F8),F8,"")');
    expect(pd).not.toContain("IF(ISNUMBER(F8),F8,$B$2)");
    expect(next).toContain('IFERROR(LOOKUP(2,1/(C$8:C8<>""),C$8:C8),$B$2)');
    expect(contDiCalcFormula(8, 9)).toBeNull();
    expect(contDiCalcFormula(8, 22)).toBeNull();
    for (const row of [8, 9, 44]) {
      for (const col of [8, 9, 11, 13, 14, 15, 16, 19, 20, 21]) {
        const formula = contDiCalcFormula(row, col);
        if (formula) {
          expect(formula.length).toBeLessThan(8000);
          expect(balanced(formula)).toBe(true);
        }
      }
    }
  });

  it("rewrites the rate formulas once", () => {
    const built = new SheetBook({
      sheets: { cd: { id: "cd", name: "Cont DI", cellData: { 7: { 7: { f: '=IF(ISNUMBER(C8),"old","")' } } } } },
      styles: {},
    } as never, []);
    expect(applyContDiHeldReadings(built)).toBe(true);
    expect(built.formula("Cont DI", 8, 8)).toContain("IF(ISNUMBER(C8),C8,$B$2)");
    expect(built.formula("Cont DI", 9, 15)).toContain('IF(ISNUMBER(F9),F9,IFERROR(LOOKUP(2,1/(F$8:F8<>""),F$8:F8),""))');
    expect(built.formula("Cont DI", 8, 9)).toBeNull();
    expect(applyContDiHeldReadings(built)).toBe(false);
  });

  it("keeps the previous inclination through a blank station", () => {
    const first = minCurveStep(0, 0, 0, 100, 10, 0);
    const turned = minCurveStep(100, 10, 0, 200, 10, 90);
    const last = minCurveStep(200, 10, 90, 300, 10, 90);
    const tvd = contDiTvdAtBit(300, 10, 90, 0, 0, 0, 0, null, [[100], [200]], [[10], [null]], [[0], [90]], null, null);
    expect(tvd).toBeCloseTo(first.tvd + turned.tvd + last.tvd, 6);
    const skipped = first.tvd + minCurveStep(100, 10, 0, 300, 10, 90).tvd;
    expect(tvd).not.toBeCloseTo(skipped, 3);
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
