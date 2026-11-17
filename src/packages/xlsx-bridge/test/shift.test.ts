import { describe, expect, it } from "vitest";
import { extractFunctionCalls, extractRefs, shiftFormula } from "../src/a1.js";
import { unknownFunctions } from "../src/excelFunctions.js";

describe("shiftFormula", () => {
  it("shifts relative cells and keeps absolute cells", () => {
    expect(shiftFormula('IF(T12="", "", D12)', 1, 0)).toBe('IF(T13="", "", D13)');
    expect(shiftFormula("SUM($F$7+H12)", 2, 1)).toBe("SUM($F$7+I14)");
  });

  it("shifts the slidesheet dogleg formula down one row", () => {
    const master = "IF(ISNUMBER(F13),COS(((F13+F12)/2)*(PI()/180))*(E13-E12)+H12,\"\")";
    const next = shiftFormula(master, 1, 0);
    expect(next).toBe("IF(ISNUMBER(F14),COS(((F14+F13)/2)*(PI()/180))*(E14-E13)+H13,\"\")");
  });

  it("does not shift text inside strings", () => {
    expect(shiftFormula('IF(A1="B2",A1,"")', 1, 0)).toBe('IF(A2="B2",A2,"")');
  });

  it("shifts only the relative edge of a mixed range", () => {
    expect(shiftFormula("SUM($F$7:F12)", 1, 0)).toBe("SUM($F$7:F13)");
  });

  it("keeps sheet-qualified references", () => {
    expect(shiftFormula("SUM('Pipe Tally'!A1:A2)", 1, 1)).toBe("SUM('Pipe Tally'!B2:B3)");
  });
});

describe("formula tokens", () => {
  it("lists calls and refs without treating functions as cells", () => {
    expect(extractFunctionCalls('IF(ISNUMBER(F13),COS(PI()/180),"")')).toEqual(["IF", "ISNUMBER", "COS", "PI"]);
    expect(extractRefs("SUM(R12+T12)")).toHaveLength(2);
    expect(unknownFunctions(["IF", "PROJBURR", "SUM"])).toEqual(["PROJBURR"]);
  });
});
