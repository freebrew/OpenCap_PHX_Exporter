import { describe, expect, it } from "vitest";
import { engineFormula, normalizeSnapshotFormulas } from "../src/readWorkbook.js";

describe("engine formula text", () => {
  it("unwraps the @ operator and drops Excel storage prefixes", () => {
    expect(engineFormula('_xlfn.SINGLE(IF(E13<_xlfn.SINGLE(ProjBuildStartMd),"",1))')).toBe('(IF(E13<(ProjBuildStartMd),"",1))');
    expect(engineFormula("_xlfn.AGGREGATE(14,6,A1:A9,1)")).toBe("AGGREGATE(14,6,A1:A9,1)");
    expect(engineFormula("_xlfn._xlws.SORT(A1:A3)")).toBe("SORT(A1:A3)");
  });

  it("leaves string literals and plain formulas alone", () => {
    expect(engineFormula('IF(A1="_xlfn.SINGLE(","x","y")')).toBe('IF(A1="_xlfn.SINGLE(","x","y")');
    expect(engineFormula("SUM(A1:A3)")).toBe("SUM(A1:A3)");
  });

  it("normalises an older snapshot in place and types text", () => {
    const snapshot = {
      sheets: {
        s1: {
          cellData: {
            0: { 0: { f: "=_xlfn.SINGLE(A2)" }, 1: { v: "00123" }, 2: { v: 5 }, 3: { v: "x", t: 1 } },
          },
        },
      },
    };
    normalizeSnapshotFormulas(snapshot);
    const row = snapshot.sheets.s1.cellData[0] as Record<number, { f?: string; v?: unknown; t?: number }>;
    expect(row[0]?.f).toBe("=(A2)");
    expect(row[1]?.t).toBe(1);
    expect(row[2]?.t).toBeUndefined();
    expect(row[3]?.t).toBe(1);
  });
});
