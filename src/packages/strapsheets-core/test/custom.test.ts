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
  type Mark,
} from "../src/custom";
import { BLUE, HATCH, INK, NAVY, customScene } from "../src/customScene";
import { BODY_HALF, FLAG_GAP, FLAG_R, HIT_R, LENGTH_FLAG_Y, PIN_W, SPIRAL_HALF, STATION_GAP, TUBE_CY, flagPoint, lengthSpots, pickMark, tubeLayout } from "../src/customShape";
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
      const next = addMark(full, count % 2 === 0 ? "od" : "upset");
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
    expect(snapMarkX("length", -1, [], "C")).toBe(0);
    expect(snapMarkX("length", 2, [], "C")).toBe(1);
    expect(snapMarkX("od", -1, [], "C")).toBe(0.04);
  });

  it("snaps a Length station to the tube's own features only", () => {
    const upset: Mark[] = [{ letter: "A", kind: "upset", x: 0.5, value: "" }];
    expect(lengthSpots(upset).map((spot) => Math.round(spot * 1000) / 1000)).toEqual([0, 0.445, 0.555, 1]);
    expect(snapMarkX("length", 0.47, upset, "B")).toBe(0.445);
    expect(snapMarkX("length", 0.53, upset, "B")).toBe(0.555);
    const spiral: Mark[] = [{ letter: "A", kind: "spiral", x: 0.5, value: "" }];
    const half = 39 / 584;
    expect(lengthSpots(spiral).some((spot) => Math.abs(spot - (0.5 - half)) < 1e-9)).toBe(true);
    expect(lengthSpots(spiral).some((spot) => Math.abs(spot - (0.5 + half)) < 1e-9)).toBe(true);
    const pinned: Mark[] = [{ letter: "A", kind: "pin", x: 0.04, value: "" }];
    expect(lengthSpots(pinned)[0]).toBe(0);
    expect(snapMarkX("length", -0.2, pinned, "B")).toBe(0);
  });

  it("does not spawn a flag on or near another flag", () => {
    let marks: Mark[] = [];
    for (const kind of ["od", "od", "od", "upset", "upset"] as const) {
      const next = addMark(marks, kind);
      if (!Array.isArray(next)) throw new Error(next);
      marks = next;
    }
    for (const mark of marks) {
      for (const other of marks) {
        if (mark.letter >= other.letter || mark.kind !== other.kind) continue;
        expect(Math.abs(flagPoint(mark, marks).x - flagPoint(other, marks).x)).toBeGreaterThanOrEqual(FLAG_GAP);
      }
    }
    const ends: Mark[] = [
      { letter: "A", kind: "length", x: 0, value: "" },
      { letter: "B", kind: "length", x: 1, value: "" },
    ];
    expect(addMark(ends, "length")).toBe("No free edge on the tube for another Length station.");
    const crowded = addMark([{ letter: "A", kind: "upset", x: 0.5, value: "" }], "length");
    if (!Array.isArray(crowded)) throw new Error(crowded);
    expect(Math.abs((crowded[1]?.x ?? 0) - 0.5)).toBeGreaterThan(0.05);
  });

  it("draws pin threads, a plain larger spiral and the two-part Length line", () => {
    const pin: Mark[] = [{ letter: "A", kind: "pin", x: 0.04, value: "" }];
    const threads = customScene(pin).filter((prim) => prim.t === "line" && prim.w === 1.15);
    expect(threads).toHaveLength(4);

    const both: Mark[] = [
      { letter: "A", kind: "pin", x: 0.04, value: "" },
      { letter: "B", kind: "pin", x: 0.96, value: "" },
    ];
    const layout = tubeLayout(both);
    const xs = customScene(both)
      .filter((prim) => prim.t === "line" && prim.w === 1.15)
      .map((prim) => (prim.t === "line" ? prim.x1 : 0));
    expect(xs).toHaveLength(8);
    expect(xs.filter((x) => x < layout.bodyX && x > layout.bodyX - PIN_W)).toHaveLength(4);
    expect(xs.filter((x) => x > layout.bodyX + layout.bodyW && x < layout.bodyX + layout.bodyW + PIN_W)).toHaveLength(4);

    const spiral: Mark[] = [{ letter: "A", kind: "spiral", x: 0.5, value: "" }];
    const scene = customScene(spiral, "A");
    const box = scene.find((prim) => prim.t === "poly" && prim.stroke === INK && prim.fill === "none");
    if (!box || box.t !== "poly") throw new Error("spiral outline missing");
    const ys = box.pts.map((point) => point[1]);
    expect(Math.max(...ys) - Math.min(...ys)).toBe(SPIRAL_HALF * 2);
    expect(SPIRAL_HALF).toBeGreaterThan(BODY_HALF);
    const hatch = scene.filter((prim) => prim.t === "line" && prim.stroke === HATCH);
    expect(hatch.length).toBeGreaterThan(4);
    for (const prim of scene) {
      if (prim.t === "circle" || prim.t === "text") continue;
      expect(prim.stroke).not.toBe(BLUE);
      expect(prim.stroke).not.toBe(NAVY);
    }

    const marks: Mark[] = [{ letter: "A", kind: "length", x: 0, value: "" }];
    const vertical = customScene(marks).filter((prim) => prim.t === "line" && prim.x1 === prim.x2);
    expect(vertical).toHaveLength(2);
    const [across, up] = vertical;
    if (!across || !up || across.t !== "line" || up.t !== "line") throw new Error("Length line missing");
    expect(across.y1).toBe(TUBE_CY - BODY_HALF);
    expect(across.y2).toBe(TUBE_CY + BODY_HALF);
    expect(up.y1).toBe(TUBE_CY - BODY_HALF - STATION_GAP);
    expect(up.y2).toBe(LENGTH_FLAG_Y + FLAG_R);

    const wide: Mark[] = [{ letter: "A", kind: "spiral", x: 0.5, value: "" }, { letter: "B", kind: "length", x: 0.5 - 39 / 584, value: "" }];
    const spanned = customScene(wide).find((prim) => prim.t === "line" && prim.x1 === prim.x2 && prim.stroke === NAVY && prim.y2 - prim.y1 > 40);
    if (!spanned || spanned.t !== "line") throw new Error("Length line should span the spiral");
    expect(spanned.y2 - spanned.y1).toBe(SPIRAL_HALF * 2);
  });

  it("points an OD station at the tube from below with a gap", () => {
    const marks: Mark[] = [{ letter: "A", kind: "od", x: 0.3, value: "" }];
    const arrow = customScene(marks).find((prim) => prim.t === "poly" && prim.pts.length === 3);
    if (!arrow || arrow.t !== "poly") throw new Error("arrow missing");
    expect(arrow.pts[0]?.[1]).toBe(TUBE_CY + BODY_HALF + STATION_GAP);
    const under: Mark[] = [{ letter: "A", kind: "spiral", x: 0.3, value: "" }, { letter: "B", kind: "od", x: 0.3, value: "" }];
    const below = customScene(under).find((prim) => prim.t === "poly" && prim.pts.length === 3);
    if (!below || below.t !== "poly") throw new Error("arrow missing");
    expect(below.pts[0]?.[1]).toBe(TUBE_CY + SPIRAL_HALF + STATION_GAP);
  });

  it("picks the flag under a finger and ignores empty space", () => {
    const marks: Mark[] = [{ letter: "A", kind: "upset", x: 0.5, value: "" }, { letter: "B", kind: "length", x: 0, value: "" }];
    const upsetFlag = flagPoint(marks[0]!, marks);
    const lengthFlag = flagPoint(marks[1]!, marks);
    expect(pickMark(upsetFlag.x + HIT_R - 2, upsetFlag.y, marks)).toBe("A");
    expect(pickMark(lengthFlag.x, lengthFlag.y + 2, marks)).toBe("B");
    expect(pickMark(300, 250, marks)).toBe("");
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
