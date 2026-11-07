import { describe, expect, it } from "vitest";
import { ENTRY_FAMILIES } from "../src/families.gen";
import {
  addMark,
  moveMark,
  nextLetter,
  MENU_KINDS,
  activeMarks,
  normalizeCustomCells,
  markMenuLabel,
  parseMarks,
  removeMark,
  snapMarkX,
  tubePath,
  customFlagCells,
  customLines,
  type Mark,
} from "../src/custom";
import { BLUE, HATCH, INK, NAVY, customScene } from "../src/customScene";
import {
  BODY_HALF,
  BUILDER_H,
  BUILDER_W,
  FLAG_GAP,
  FLAG_R,
  HIT_R,
  LENGTH_KNEE_Y,
  OD_KNEE_Y,
  PIN_W,
  SPIRAL_HALF,
  SPIRAL_TAPER,
  SPIRAL_W,
  STATION_GAP,
  TOUCH_GAP,
  TOUCH_R,
  TUBE_CY,
  TUBE_W,
  TUBE_X,
  dragX,
  featureHalfAt,
  flagPoint,
  lengthFlagY,
  lengthSpots,
  odFlagY,
  pickMark,
  spiralHalfAt,
  stemPx,
  tubeLayout,
} from "../src/customShape";
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
    expect(lengthSpots(upset).map((spot) => Math.round(spot * 1000) / 1000)).toEqual([0, 0.445, 0.475, 0.525, 0.555, 1]);
    expect(snapMarkX("length", 0.47, upset, "B")).toBe(0.445);
    expect(snapMarkX("length", 0.53, upset, "B")).toBe(0.555);
    const spiral: Mark[] = [{ letter: "A", kind: "spiral", x: 0.5, value: "" }];
    const spots = lengthSpots(spiral);
    for (const offset of [SPIRAL_W / 2, SPIRAL_W / 2 - SPIRAL_TAPER]) {
      expect(spots.some((spot) => Math.abs(spot - (0.5 - offset / TUBE_W)) < 1e-9)).toBe(true);
      expect(spots.some((spot) => Math.abs(spot - (0.5 + offset / TUBE_W)) < 1e-9)).toBe(true);
    }
    const panels: Mark[] = [{ letter: "A", kind: "panels", x: 0.5, value: "" }];
    expect(lengthSpots(panels).map((spot) => Math.round(spot * 1000) / 1000)).toEqual([0, 0.446, 0.554, 1]);
    const spline: Mark[] = [{ letter: "A", kind: "spline", x: 0.5, value: "" }];
    expect(lengthSpots(spline)).toHaveLength(5);
    const pinned: Mark[] = [{ letter: "A", kind: "pin", x: 0.04, value: "" }];
    expect(lengthSpots(pinned)[0]).toBe(0);
    expect(snapMarkX("length", -0.2, pinned, "B")).toBe(0);
  });

  it("lets a Length station reach both very ends of the tube, with or without pins", () => {
    expect(snapMarkX("length", -1, [], "C")).toBe(0);
    expect(snapMarkX("length", 2, [], "C")).toBe(1);
    const pins: Mark[] = [
      { letter: "A", kind: "pin", x: 0.04, value: "" },
      { letter: "B", kind: "pin", x: 0.96, value: "" },
    ];
    expect(snapMarkX("length", -1, pins, "C")).toBe(0);
    expect(snapMarkX("length", 2, pins, "C")).toBe(1);
    const start = [...pins, { letter: "C", kind: "length", x: 0.5, value: "" } as Mark];
    expect(moveMark(start, "C", -3).find((mark) => mark.letter === "C")?.x).toBe(0);
    expect(moveMark(start, "C", 4).find((mark) => mark.letter === "C")?.x).toBe(1);
    const at = (x: number) => flagPoint({ letter: "C", kind: "length", x, value: "" }, pins, true);
    expect(at(0).x - TOUCH_R).toBeGreaterThanOrEqual(0);
    expect(at(1).x + TOUCH_R).toBeLessThanOrEqual(BUILDER_W);
  });

  it("does not spawn a flag on or near another flag", () => {
    let marks: Mark[] = [];
    for (const kind of ["od", "od", "od", "length", "length", "upset", "upset"] as const) {
      const next = addMark(marks, kind);
      if (!Array.isArray(next)) throw new Error(next);
      marks = next;
    }
    for (const mark of marks) {
      const own = flagPoint(mark, marks, true);
      expect(own.x - TOUCH_R).toBeGreaterThanOrEqual(-1e-9);
      for (const other of marks) {
        if (mark.letter >= other.letter) continue;
        const gap = Math.abs(flagPoint(mark, marks, true).x - flagPoint(other, marks, true).x);
        if (mark.kind === "upset" && other.kind === "upset") expect(gap).toBeGreaterThanOrEqual(FLAG_GAP);
        if (mark.kind === other.kind && mark.kind !== "upset") expect(gap).toBeGreaterThanOrEqual(TOUCH_GAP - 1e-9);
      }
    }
    const ends: Mark[] = [
      { letter: "A", kind: "length", x: 0, value: "" },
      { letter: "B", kind: "length", x: 1, value: "" },
    ];
    expect(addMark(ends, "length")).toBe("No free edge on the tube for another Length station.");
    const crowded = addMark([{ letter: "A", kind: "upset", x: 0.5, value: "" }], "length");
    if (!Array.isArray(crowded)) throw new Error(crowded);
    expect(Math.abs((crowded[1]?.x ?? 0) - 0.5)).toBeGreaterThan(0.03);
  });

  it("keeps the big flags apart and on the canvas when Length stations sit close together", () => {
    const panels: Mark[] = [
      { letter: "A", kind: "panels", x: 0.5, value: "" },
      { letter: "B", kind: "length", x: 0.5 - 29 / TUBE_W, value: "" },
      { letter: "C", kind: "length", x: 0.5 + 29 / TUBE_W, value: "" },
    ];
    const left = flagPoint(panels[1] as Mark, panels, true);
    const right = flagPoint(panels[2] as Mark, panels, true);
    expect(right.x - left.x).toBeGreaterThanOrEqual(TOUCH_GAP - 1e-9);
    expect(left.x).toBeLessThan(stemPx(panels[1] as Mark, panels));
    expect(right.x).toBeGreaterThan(stemPx(panels[2] as Mark, panels));
    expect((left.x + right.x) / 2).toBeCloseTo((stemPx(panels[1] as Mark, panels) + stemPx(panels[2] as Mark, panels)) / 2, 6);
    const scene = customScene(panels, "", true);
    const leaders = scene.filter((prim) => prim.t === "line" && prim.y1 === LENGTH_KNEE_Y && prim.x1 !== prim.x2);
    expect(leaders).toHaveLength(2);
    const six: Mark[] = Array.from({ length: 6 }, (_, index) => ({ letter: "ABCDEF"[index] ?? "A", kind: "length", x: 0.4 + index * 0.02, value: "" }));
    const xs = six.map((mark) => flagPoint(mark, six, true).x);
    expect(Math.min(...xs) - TOUCH_R).toBeGreaterThanOrEqual(-1e-9);
    expect(Math.max(...xs) + TOUCH_R).toBeLessThanOrEqual(BUILDER_W + 1e-9);
    for (let index = 1; index < xs.length; index += 1) expect((xs[index] ?? 0) - (xs[index - 1] ?? 0)).toBeGreaterThanOrEqual(TOUCH_GAP - 1e-9);
  });

  it("draws pin threads, a plain tapered spiral and the two-part Length line", () => {
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
    expect(box.pts).toHaveLength(8);
    const ys = box.pts.map((point) => point[1]);
    expect(Math.max(...ys) - Math.min(...ys)).toBe(SPIRAL_HALF * 2);
    expect(SPIRAL_HALF).toBeGreaterThan(BODY_HALF);
    const toes = box.pts.filter((point) => Math.abs(point[1] - TUBE_CY) === BODY_HALF);
    const shoulders = box.pts.filter((point) => Math.abs(point[1] - TUBE_CY) === SPIRAL_HALF);
    expect(toes).toHaveLength(4);
    expect(shoulders).toHaveLength(4);
    const centre = stemPx(spiral[0] as Mark, spiral);
    expect(Math.min(...toes.map((point) => point[0]))).toBeCloseTo(centre - SPIRAL_W / 2, 9);
    expect(Math.min(...shoulders.map((point) => point[0]))).toBeCloseTo(centre - SPIRAL_W / 2 + SPIRAL_TAPER, 9);
    const hatch = scene.filter((prim) => prim.t === "line" && prim.stroke === HATCH);
    expect(hatch.length).toBeGreaterThan(4);
    for (const prim of hatch) {
      if (prim.t !== "line") continue;
      for (const [x, y] of [[prim.x1, prim.y1], [prim.x2, prim.y2]] as const) {
        expect(Math.abs(y - TUBE_CY)).toBeLessThanOrEqual(spiralHalfAt(x - centre) + 1e-6);
      }
    }
    for (const prim of scene) {
      if (prim.t === "circle" || prim.t === "text") continue;
      expect(prim.stroke).not.toBe(BLUE);
      expect(prim.stroke).not.toBe(NAVY);
    }
    expect(spiralHalfAt(0)).toBe(SPIRAL_HALF);
    expect(spiralHalfAt(SPIRAL_W / 2)).toBe(BODY_HALF);
    expect(spiralHalfAt(SPIRAL_W / 2 + 40)).toBe(BODY_HALF);

    const marks: Mark[] = [{ letter: "A", kind: "length", x: 0, value: "" }];
    const vertical = customScene(marks).filter((prim) => prim.t === "line" && prim.x1 === prim.x2);
    expect(vertical).toHaveLength(3);
    const [across, up, leader] = vertical;
    if (!across || !up || !leader || across.t !== "line" || up.t !== "line" || leader.t !== "line") throw new Error("Length line missing");
    expect(across.y1).toBe(TUBE_CY - BODY_HALF);
    expect(across.y2).toBe(TUBE_CY + BODY_HALF);
    expect(up.y1).toBe(TUBE_CY - BODY_HALF - STATION_GAP);
    expect(up.y2).toBe(LENGTH_KNEE_Y);
    expect(leader.y1).toBe(LENGTH_KNEE_Y);
    expect(leader.y2).toBe(lengthFlagY() + FLAG_R);
  });

  it("spans the spiral at its shoulder and the tube at its toe", () => {
    const shoulder = 0.5 - (SPIRAL_W / 2 - SPIRAL_TAPER) / TUBE_W;
    const toe = 0.5 - SPIRAL_W / 2 / TUBE_W;
    const marks: Mark[] = [
      { letter: "A", kind: "spiral", x: 0.5, value: "" },
      { letter: "B", kind: "length", x: shoulder, value: "" },
      { letter: "C", kind: "length", x: toe, value: "" },
    ];
    expect(featureHalfAt(shoulder, marks)).toBe(SPIRAL_HALF);
    expect(featureHalfAt(toe, marks)).toBe(BODY_HALF);
    expect(featureHalfAt(0.2, marks)).toBe(BODY_HALF);
    const span = (letter: string) => {
      const mark = marks.find((candidate) => candidate.letter === letter) as Mark;
      const x = stemPx(mark, marks);
      const line = customScene(marks).find((prim) => prim.t === "line" && prim.x1 === x && prim.x2 === x && prim.y2 - prim.y1 > 20 && prim.y2 > TUBE_CY);
      if (!line || line.t !== "line") throw new Error(`Length line missing for ${letter}`);
      return line.y2 - line.y1;
    };
    expect(span("B")).toBe(SPIRAL_HALF * 2);
    expect(span("C")).toBe(BODY_HALF * 2);
  });

  it("makes the builder's Length and OD flags 3x the compact flag, and no other", () => {
    expect(TOUCH_R).toBe(3 * FLAG_R);
    const marks: Mark[] = [
      { letter: "A", kind: "length", x: 0.3, value: "" },
      { letter: "B", kind: "od", x: 0.6, value: "" },
      { letter: "C", kind: "upset", x: 0.8, value: "" },
    ];
    const radii = (touch: boolean) => customScene(marks, "", touch).flatMap((prim) => (prim.t === "circle" ? [prim.r] : []));
    expect(radii(true)).toEqual([TOUCH_R, TOUCH_R, FLAG_R]);
    expect(radii(false)).toEqual([FLAG_R, FLAG_R, FLAG_R]);
    expect(lengthFlagY(true) - TOUCH_R).toBeGreaterThanOrEqual(0);
    expect(odFlagY(true) + TOUCH_R).toBeLessThanOrEqual(BUILDER_H);
    expect(TUBE_X - TOUCH_R).toBeGreaterThanOrEqual(0);
    expect(TUBE_X + TUBE_W + TOUCH_R).toBeLessThanOrEqual(BUILDER_W);
  });

  it("points an OD station at the tube from below with a gap", () => {
    const marks: Mark[] = [{ letter: "A", kind: "od", x: 0.3, value: "" }];
    const arrow = customScene(marks).find((prim) => prim.t === "poly" && prim.pts.length === 3);
    if (!arrow || arrow.t !== "poly") throw new Error("arrow missing");
    expect(arrow.pts[0]?.[1]).toBe(TUBE_CY + BODY_HALF + STATION_GAP);
    const stem = customScene(marks).find((prim) => prim.t === "line" && prim.x1 === prim.x2);
    if (!stem || stem.t !== "line") throw new Error("OD stem missing");
    expect(stem.y2).toBe(OD_KNEE_Y);
    const under: Mark[] = [{ letter: "A", kind: "spiral", x: 0.3, value: "" }, { letter: "B", kind: "od", x: 0.3, value: "" }];
    const below = customScene(under).find((prim) => prim.t === "poly" && prim.pts.length === 3);
    if (!below || below.t !== "poly") throw new Error("arrow missing");
    expect(below.pts[0]?.[1]).toBe(TUBE_CY + SPIRAL_HALF + STATION_GAP);
  });

  it("picks the flag under a finger and ignores empty space", () => {
    const marks: Mark[] = [{ letter: "A", kind: "upset", x: 0.5, value: "" }, { letter: "B", kind: "length", x: 0, value: "" }];
    const upsetFlag = flagPoint(marks[0] as Mark, marks, true);
    const lengthFlag = flagPoint(marks[1] as Mark, marks, true);
    expect(pickMark(upsetFlag.x + HIT_R - 2, upsetFlag.y, marks)).toBe("A");
    expect(pickMark(lengthFlag.x + TOUCH_R + 4, lengthFlag.y, marks)).toBe("B");
    expect(pickMark(lengthFlag.x, lengthFlag.y + 2, marks)).toBe("B");
    expect(pickMark(300, 380, marks)).toBe("");
  });

  it("moves a mark by as far as the pointer travels", () => {
    const marks: Mark[] = [{ letter: "A", kind: "upset", x: 0.5, value: "" }];
    expect(dragX(0.5, 100, 100 + 0.1 * TUBE_W, marks)).toBeCloseTo(0.6, 9);
    expect(dragX(0.5, 100, 100, marks)).toBe(0.5);
  });

  it("names the Add menu in Proper Case and gives an add-on's letter its OD", () => {
    const labels = ["pin", "upset", "spline", "spiral", "panels", "length", "od"].map((kind) => markMenuLabel(kind as Mark["kind"]));
    expect(labels).toEqual(["Pin", "Upset / Inset", "Spline", "Spiral (Stab / Turbolizer)", "Panels", "Length Station", "OD Station"]);
    expect(markMenuLabel("pin")).toBe("Pin");
    const cells = {
      marksJson: JSON.stringify({
        marks: [
          { letter: "A", kind: "spiral", x: 0.5, value: "5.25" },
          { letter: "B", kind: "length", x: 0.1, value: "0.6" },
          { letter: "C", kind: "od", x: 0.3, value: "4.75" },
          { letter: "D", kind: "pin", x: 0.04, value: "" },
        ],
      }),
    };
    expect(customLines(cells)).toEqual([
      { label: "A Spiral OD", value: "5.25" },
      { label: "B Length", value: "0.6" },
      { label: "C OD", value: "4.75" },
    ]);
    const marks = parseMarks(cells.marksJson);
    // The card's second row follows the sheet: Length reads LEN, OD and add-ons read OD, pins are blank.
    expect(customFlagCells(marks)).toEqual([
      { slot: 0, label: "A OD", value: "5.25" },
      { slot: 1, label: "B LEN", value: "0.6" },
      { slot: 2, label: "C OD", value: "4.75" },
    ]);
    expect(removeMark(marks, "A").some((mark) => mark.letter === "A")).toBe(false);
    expect(customLines({ marksJson: JSON.stringify({ marks: removeMark(marks, "A") }) }).some((line) => line.label.startsWith("A "))).toBe(false);
    expect(MENU_KINDS).not.toContain("pin");
    const kept = normalizeCustomCells(cells);
    const stored = parseMarks(kept.marksJson);
    expect(stored.map((mark) => mark.letter)).toEqual(["A", "B", "C"]);
    expect(activeMarks(marks).map((mark) => mark.letter)).toEqual(["A", "B", "C"]);
    const configured = customScene(marks, "", false, { lower: "PIN", upper: "BOX", threadLower: "NC 50", threadUpper: "4 1/2 IF" });
    const threads = configured.filter((prim) => prim.t === "line" && prim.w === 1.15);
    expect(threads.length).toBe(4);
    expect(configured.some((prim) => prim.t === "text" && prim.s === "NC 50")).toBe(true);
    expect(configured.some((prim) => prim.t === "text" && prim.s === "4 1/2 IF")).toBe(true);
    expect(configured.some((prim) => prim.t === "text" && prim.s === "D")).toBe(false);
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
