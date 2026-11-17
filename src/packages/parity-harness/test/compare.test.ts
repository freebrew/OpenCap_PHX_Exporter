import { describe, expect, it } from "vitest";
import { valuesMatch } from "../src/compare.js";
import { classifyFormulas } from "../src/classify.js";
import type { CellFormulaRecord, DefinedNameRecord } from "../../xlsx-bridge/src/index.ts";

describe("valuesMatch", () => {
  it("accepts absolute or relative numeric tolerance and compares errors by type", () => {
    expect(valuesMatch({ t: "n", v: 1 }, { t: "n", v: 1 + 5e-7 })).toBe(true);
    expect(valuesMatch({ t: "n", v: 1_000_000 }, { t: "n", v: 1_000_000 + 0.0005 })).toBe(true);
    expect(valuesMatch({ t: "n", v: 10 }, { t: "n", v: 10.01 })).toBe(false);
    expect(valuesMatch({ t: "e", v: "#DIV/0!" }, { t: "e", v: "#DIV/0!" })).toBe(true);
    expect(valuesMatch({ t: "e", v: "#DIV/0!" }, { t: "e", v: "#VALUE!" })).toBe(false);
    expect(valuesMatch({ t: "s", v: "Midnight" }, { t: "s", v: "Midnight" })).toBe(true);
    expect(valuesMatch({ t: "s", v: "Midnight" }, { t: "s", v: "midnight" })).toBe(false);
    expect(valuesMatch({ t: "s", v: null }, { t: "s", v: "" })).toBe(true);
  });
});

describe("classifyFormulas", () => {
  it("marks direct calls and cells that reference them", () => {
    const formulas: CellFormulaRecord[] = [
      { sheet: "S", a1: "A1", row: 0, col: 0, formula: "PROJBURR(1)" },
      { sheet: "S", a1: "A2", row: 1, col: 0, formula: "A1+1" },
      { sheet: "S", a1: "A3", row: 2, col: 0, formula: "SUM(B1)" },
    ];
    const names: DefinedNameRecord[] = [];
    const classes = classifyFormulas(formulas, names);
    expect(classes.get("S\u00000\u00000")).toBe("udf-dependent");
    expect(classes.get("S\u00001\u00000")).toBe("udf-dependent");
    expect(classes.get("S\u00002\u00000")).toBe("evaluable");
  });
});
