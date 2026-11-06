import { describe, expect, it } from "vitest";
import { ENTRY_FAMILIES } from "../src/families.gen";
import {
  addMark,
  moveMark,
  nextLetter,
  normalizeCustomCells,
  parseMarks,
  removeMark,
  snapMarkX,
  tubePath,
} from "../src/custom";
import { findFamily, normalizeEntry, validateEntry } from "../src/index";
import { renderCards } from "../src/pdf";

describe("custom marks", () => {
  it("reuses the first free letter and stops at twelve", () => {
    let marks = addMark([], "length");
    expect(Array.isArray(marks) && marks[0]?.letter).toBe("A");
    if (!Array.isArray(marks)) throw new Error(String(marks));
    marks = removeMark(marks, "A");
    const again = addMark(marks, "od");
    expect(Array.isArray(again) && again[0]?.letter).toBe("A");
    let full = [] as typeof marks;
    for (let count = 0; count < 12; count += 1) {
      const next = addMark(full, "length");
      if (!Array.isArray(next)) throw new Error(next);
      full = next;
    }
    expect(nextLetter(full)).toBe("");
    expect(addMark(full, "length")).toBe("A tube holds 12 marks.");
  });

  it("allows one pin at each end and snaps a pin to that end", () => {
    const left = addMark([], "pin");
    if (!Array.isArray(left)) throw new Error(left);
    expect(left[0]?.x).toBe(0.04);
    const both = addMark(left, "pin");
    if (!Array.isArray(both)) throw new Error(both);
    expect(both[1]?.x).toBe(0.96);
    expect(addMark(both, "pin")).toBe("Only one pin fits at each end.");
    const moved = moveMark(both, "A", 0.8);
    expect(moved.find((mark) => mark.letter === "A")?.x).toBe(0.04);
    expect(snapMarkX("length", -1, [], "C")).toBe(0.04);
    expect(snapMarkX("length", 2, [], "C")).toBe(0.96);
  });

  it("reads legacy mandrel JSON and ignores a broken blob", () => {
    const marks = parseMarks(JSON.stringify({ marks: [{ letter: "B", kind: "mandrel", x: 0.3, value: "1" }] }));
    expect(marks[0]?.kind).toBe("spline");
    expect(parseMarks("{")).toEqual([]);
  });

  it("requires total length and keeps a bad measurement visible", () => {
    const family = findFamily(ENTRY_FAMILIES, "Custom");
    expect(family?.steps.map((step) => step.kind)).toContain("customBuilder");
    if (!family) throw new Error("missing Custom");
    const empty = normalizeEntry(family, { which: "Custom", cells: { serial: "abc" } });
    expect(validateEntry(family, empty).some((problem) => problem.message.includes("Total Length"))).toBe(true);
    const bad = normalizeCustomCells({ totalLength: "12", internalId: "", marksJson: JSON.stringify({ marks: [{ letter: "A", kind: "od", x: 0.4, value: "nope" }] }) });
    expect(validateEntry(family, { which: "Custom", cells: { ...bad, serial: "ABC" } }).some((problem) => problem.message.includes("A needs"))).toBe(true);
  });

  it("draws the custom card after a standard page and names the serial", () => {
    const bytes = renderCards([
      { which: "Subs", cells: { 0: "FLOAT", 1: "Steel", 2: "SUB1" } },
      {
        which: "Custom",
        cells: {
          serial: "CUST1",
          material: "Steel",
          boxUpper: "PIN",
          boxLower: "BOX",
          threadUpper: "DS 38",
          threadLower: "DS 40",
          totalLength: "10.5",
          internalId: "2.25",
          marksJson: JSON.stringify({ marks: [{ letter: "A", kind: "pin", x: 0.04, value: "" }, { letter: "B", kind: "length", x: 0.4, value: "1.2" }] }),
        },
      },
    ]);
    const text = new TextDecoder().decode(bytes);
    expect(text.indexOf("SUBS")).toBeGreaterThan(-1);
    expect(text.indexOf("CUSTOM")).toBeGreaterThan(text.indexOf("SUBS"));
    expect(text).toContain("CUST1");
    expect(text).toContain("10.5");
    expect(tubePath(parseMarks(JSON.stringify({ marks: [{ letter: "A", kind: "pin", x: 0.04, value: "" }] }))).startsWith("M")).toBe(true);
  });
});
