import { describe, expect, it } from "vitest";
import { flattenOnto, gridValues, niceStep, shadowBoxOf, wallsFor } from "../../wells-3d/src/shadowBox";
import type { PathArrays, WellScene, WellTarget } from "../../wells-3d/src/wellScene";

function path(points: Array<[number, number, number]>): PathArrays {
  const count = points.length;
  return {
    xyz: new Float32Array(points.flat()),
    md: new Float32Array(count),
    tvd: new Float32Array(count),
    c2c: new Float32Array(count),
    sf: new Float32Array(count),
  };
}

function sceneOf(plan: PathArrays, drilled: PathArrays | null = null, targets: WellTarget[] = []): WellScene {
  return {
    plan,
    drilled,
    offsets: [],
    targets,
    formations: [],
    missing: [],
    leftOut: 0,
    report: {
      stored: false,
      referenceWell: "",
      design: "",
      referenceSite: "",
      coordinatesRelativeTo: "",
      north: "",
      tvdReference: "",
      offsetTvdReference: "",
      reportDate: "",
      verdict: "",
      verdictOk: true,
    },
    badge: "",
    badgeOk: true,
    hash: "",
  };
}

/** A plan 500 m east, 100 m north and 400 m down: wider than it is tall, and much longer than it is wide. */
const PLAN = path([[0, 0, 0], [500, 100, -400]]);

describe("shadow box", () => {
  it("pads the active well the way the email picture does, and widens a thin side to the depth", () => {
    const box = shadowBoxOf(sceneOf(PLAN));
    expect(box).not.toBeNull();
    if (!box) return;
    // 12% of 500 east, 8% of 400 down.
    expect(box.minX).toBeCloseTo(-60);
    expect(box.maxX).toBeCloseTo(560);
    expect(box.minZ).toBeCloseTo(-432);
    expect(box.maxZ).toBeCloseTo(32);
    // North is 100 m plus 20 m padding each side, so it is widened to the 464 m of depth, round its middle.
    expect(box.maxY - box.minY).toBeCloseTo(464);
    expect((box.minY + box.maxY) / 2).toBeCloseTo(50);
    expect(box.step).toBe(100);
  });

  it("holds the drilled path and the target rings too", () => {
    const outer = new Float32Array([900, 0, -100, 900, 10, -100]);
    const target: WellTarget = { name: "HEEL", md: 500, tvd: 100, x: 900, y: 0, z: -100, radius: 10, outer, inner: outer };
    const drilled = path([[0, 0, 0], [0, -300, -700]]);
    const box = shadowBoxOf(sceneOf(PLAN, drilled, [target]));
    expect(box?.maxX).toBeGreaterThan(900);
    expect(box?.minY).toBeLessThan(-300);
    expect(box?.minZ).toBeLessThan(-700);
  });

  it("has no box without a plan", () => {
    expect(shadowBoxOf(sceneOf(path([])))).toBeNull();
  });

  it("shows the walls behind the well from where the camera is", () => {
    const box = shadowBoxOf(sceneOf(PLAN));
    if (!box) throw new Error("no box");
    expect(wallsFor(box, { x: 900, y: 400, z: 300 })).toEqual({ x: "min", y: "min", floor: true });
    expect(wallsFor(box, { x: -400, y: -300, z: 300 })).toEqual({ x: "max", y: "max", floor: true });
    expect(wallsFor(box, { x: 900, y: -300, z: -900 })).toEqual({ x: "min", y: "max", floor: false });
  });

  it("lays grid lines on whole steps", () => {
    expect(gridValues(-60, 560, 100)).toEqual([0, 100, 200, 300, 400, 500]);
    expect(gridValues(-182, 282, 100)).toEqual([-100, 0, 100, 200]);
    expect(gridValues(10, 10, 100)).toEqual([]);
  });

  it("chooses a 1, 2 or 5 step", () => {
    expect(niceStep(620, 6)).toBe(100);
    expect(niceStep(2400, 6)).toBe(500);
    expect(niceStep(90, 6)).toBe(20);
    expect(niceStep(0, 6)).toBe(1);
  });

  it("flattens a path onto a wall without moving the other two axes", () => {
    const flat = flattenOnto(PLAN.xyz, "z", -432);
    expect(Array.from(flat)).toEqual([0, 0, -432, 500, 100, -432]);
    const wall = flattenOnto(PLAN.xyz, "x", -60);
    expect(Array.from(wall)).toEqual([-60, 0, 0, -60, 100, -400]);
  });
});
