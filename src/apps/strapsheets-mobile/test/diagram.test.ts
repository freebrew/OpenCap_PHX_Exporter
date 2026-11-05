import { describe, expect, it } from "vitest";
import { ENTRY_FAMILIES } from "../../../packages/strapsheets-core/src/index";
import { letterForStep } from "../src/diagramData";
import { TOOL_DIAGRAMS } from "../src/diagrams.gen";

describe("tool diagrams", () => {
  it("has a drawing for every tool family", () => {
    expect(ENTRY_FAMILIES.filter((family) => !TOOL_DIAGRAMS[family.which]).map((family) => family.which)).toEqual([]);
  });

  it("rings a mark for every lettered measure", () => {
    const missing: string[] = [];
    for (const family of ENTRY_FAMILIES) {
      const diagram = TOOL_DIAGRAMS[family.which]!;
      for (const step of family.steps) {
        const letter = letterForStep(step, diagram);
        if (letter && !diagram.marks.some((mark) => mark.letter === letter)) missing.push(`${family.which}: ${letter}`);
      }
    }
    expect(missing).toEqual([]);
  });

  it("uses the drawing's own letter where it differs from the label", () => {
    const exciter = TOOL_DIAGRAMS["Exciter / Agitator / Sidewinder"]!;
    expect(letterForStep({ kind: "decimal", col: 9, id: "m", header: "", label: "E - Mandrill OD" }, exciter)).toBe("F");
    expect(letterForStep({ kind: "decimal", col: 9, id: "m", header: "", label: "A - BitBox" }, TOOL_DIAGRAMS.Motor)).toBe("A");
  });
});
