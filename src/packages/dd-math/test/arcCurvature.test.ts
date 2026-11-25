import { describe, expect, it } from "vitest";
import { arcTvdGain, rad, solveArcCurvature } from "../../corridor-3d/src/corridorScene";

describe("required arc to a far target", () => {
  it("stays near a hold when the target is 1,368 m ahead and 6.77 m shallower", () => {
    const toGo = 1368.7;
    const kap = solveArcCurvature(90.05, toGo, -6.77);
    // The 40 deg/30 m clamp used to come back here and the drawn arc climbed 29 m in 53 m of hole.
    expect(Math.abs(kap)).toBeLessThan(rad(1 / 30));
    expect(arcTvdGain(rad(90.05), kap, toGo)).toBeCloseTo(-6.77, 2);
    expect(Math.abs(arcTvdGain(rad(90.05), kap, 53))).toBeLessThan(1);
  });

  it("still builds at the clamp when a close target needs it", () => {
    const kap = solveArcCurvature(60, 80, 5);
    expect(kap).toBeLessThanOrEqual(rad(40 / 30) + 1e-12);
    expect(arcTvdGain(rad(60), kap, 80)).toBeLessThan(arcTvdGain(rad(60), 0, 80));
  });
});
