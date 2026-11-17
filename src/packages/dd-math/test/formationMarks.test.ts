import { describe, expect, it } from "vitest";
import { formationMarks, formationTvdLabel } from "../../wells-3d/src/formationMarks";
import type { WellFormation } from "../../wells-3d/src/wellScene";

function top(name: string, tvd: number, z: number): WellFormation {
  return { name, lith: "", tvd, z };
}

describe("formation marks", () => {
  it("writes the TVD the way the email picture does", () => {
    expect(formationTvdLabel(200)).toBe("TVD 200 m");
    expect(formationTvdLabel(1590)).toBe("TVD 1,590 m");
  });

  it("keeps tops on the box and steps a close pair sideways", () => {
    const marks = formationMarks(-400, 0, [
      top("Montney", 300, -300),
      top("Below", 900, -900),
      top("", 150, -150),
      top("Belly River", 100, -100),
      top("2WS", 105, -105),
    ]);
    expect(marks.map((mark) => mark.name)).toEqual(["Belly River", "2WS", "Montney"]);
    expect(marks[1]!.along).toBeGreaterThan(marks[0]!.along);
    expect(marks[2]!.along).toBe(marks[0]!.along);
    expect(marks.every((mark) => mark.z >= -400 && mark.z <= 0)).toBe(true);
  });

  it("draws nothing when the box has no height", () => {
    expect(formationMarks(0, 0, [top("Belly River", 10, -10)])).toEqual([]);
  });
});
