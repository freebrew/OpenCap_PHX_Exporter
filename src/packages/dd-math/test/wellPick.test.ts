import { describe, expect, it } from "vitest";
import { cameraFromOrbit, dragOrbit, fitOrbit, wheelOrbit, type Bounds } from "../../wells-3d/src/orbit";
import { closestOnPath, pickHover, type PickPath, type PickPoint } from "../../wells-3d/src/wellPick";

function point(x: number, y: number, z: number, md: number, tvd = 0, c2c: number | null = null, sf: number | null = null): PickPoint {
  return { x, y, z, md, tvd, c2c, sf };
}

const plan: PickPath = {
  kind: "plan",
  name: "Plan",
  points: [point(0, 0, 0, 0, 0), point(100, 0, 0, 100, 10)],
};

const offset: PickPath = {
  kind: "offset",
  name: "North Offset",
  points: [point(0, 20, 0, 0, 0, 20, 1.2), point(100, 20, 0, 100, 10, 20, 1.4)],
};

function flat(x: number, y: number): { x: number; y: number; behind: boolean } {
  return { x, y, behind: false };
}

describe("hover", () => {
  it("reports the nearest station's printed numbers and the live distance to the plan", () => {
    const hit = pickHover([plan, offset], { x: 40, y: 20 }, flat, 12);
    expect(hit).toMatchObject({ kind: "offset", name: "North Offset", c2c: 20, sf: 1.2, measuredM: 20 });
    expect(hit?.refMd).toBe(0);
  });

  it("reports the tightest printed SF on the plan, and never a centre-to-centre", () => {
    const hit = pickHover([plan, offset], { x: 100, y: 0 }, flat, 8);
    expect(hit).toMatchObject({ kind: "plan", c2c: null, sf: 1.4, refMd: 100, tvd: 10 });
    expect(hit?.measuredM).toBeCloseTo(0);
  });

  it("measures the closest point of a straight run", () => {
    const closest = closestOnPath(plan.points, 40, 15, 0);
    expect(closest?.distance).toBeCloseTo(15);
    expect(closest?.md).toBeCloseTo(40);
  });
});

describe("orbit", () => {
  const bounds: Bounds = { minX: 0, maxX: 100, minY: -20, maxY: 20, minZ: -50, maxZ: 0 };

  it("fits the paths and keeps the camera off the target", () => {
    const orbit = fitOrbit(bounds);
    expect(orbit.radius).toBeCloseTo(180);
    const eye = cameraFromOrbit(orbit);
    expect(Math.hypot(eye.x - orbit.targetX, eye.y - orbit.targetY, eye.z - orbit.targetZ)).toBeCloseTo(orbit.radius);
  });

  it("orbits, pans and zooms", () => {
    const orbit = fitOrbit(bounds);
    const turned = dragOrbit(orbit, 10, 0, false);
    expect(turned.theta).toBeLessThan(orbit.theta);
    const panned = dragOrbit(orbit, 10, 0, true);
    expect(panned.targetX).not.toBeCloseTo(orbit.targetX);
    expect(wheelOrbit(orbit, 100).radius).toBeGreaterThan(orbit.radius);
    expect(wheelOrbit(orbit, -100).radius).toBeLessThan(orbit.radius);
  });
});
