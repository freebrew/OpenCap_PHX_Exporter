import { describe, expect, it } from "vitest";
import { doglegAngleDeg, minCurveStep } from "../src/angles";
import { callUdf } from "../src/call";
import { closestOffset, integrateLocalNE } from "../src/plan";
import { azmWalkDeg, doglegBelow, parseToolface } from "../src/project";

describe("directional cores", () => {
  it("keeps a vertical course on TVD", () => {
    const step = minCurveStep(0, 0, 0, 100, 0, 0);
    expect(step.tvd).toBeCloseTo(100, 6);
    expect(step.n).toBeCloseTo(0, 6);
    expect(step.e).toBeCloseTo(0, 6);
  });

  it("measures a pure build as the inclination change", () => {
    expect(doglegAngleDeg(0, 0, 10, 0)).toBeCloseTo(10, 6);
    expect(doglegAngleDeg(10, 20, 10, 20)).toBeCloseTo(0, 6);
  });

  it("parses the sheet toolface spellings", () => {
    expect(parseToolface("40R")).toBe(40);
    expect(parseToolface("L20")).toBe(-20);
    expect(parseToolface("90M")).toBe(90);
    expect(parseToolface("HSL")).toBe(180);
    expect(parseToolface("")).toBeNull();
    expect(callUdf("PROJPARSETF", ["nope"]).error).toBe("VALUE");
  });

  it("builds the unseen dogleg from the seen yield", () => {
    expect(doglegBelow(3, 10, 10, 5)).toBeCloseTo(0.5, 6);
    expect(callUdf("PROJINCATBIT", [10, 3, 10, 10, 5, 0]).value).toBeCloseTo(10.5, 6);
  });

  it("walks azimuth by the ISCWSA sine term above 5 degrees", () => {
    const walk = azmWalkDeg(10, 0.5, 90);
    expect(callUdf("PROJAZMATBIT", [10, 90, 3, 10, 10, 5, 90]).value).toBeCloseTo(90 + walk, 5);
  });

  it("sums only the slide overlap inside the window", () => {
    expect(callUdf("PROJSLIDEMETERSBETWEEN", [100, 110, [95, 108], [105, 120]]).value).toBeCloseTo(7, 6);
  });

  it("weights the last three motor samples by metres seen", () => {
    expect(callUdf("PROJROLLINGMOTOROUT", [[1, 2, 3, 4], [10, 10, 10, 10]]).value).toBeCloseTo(3, 6);
  });

  it("aims at the next target and turns that dogleg into a slide", () => {
    const targets = { md: [100, 200], inc: [10, 20], azm: [0, 0], tvd: [50, 140] };
    expect(callUdf("PROJACTIVETARGETMDEX", [150, 12, 80, targets.md, targets.inc, targets.azm, targets.tvd]).value).toBe(200);
    expect(callUdf("PROJMETERSTOSLIDE", [100, 10, 0, 3, 10, [200], [20], [0], [100]]).value).toBeCloseTo(10, 4);
    expect(callUdf("PROJBURR", [100, 10, 0, 3, 10, [200], [20], [0], [100]]).value).toBeCloseTo(3, 4);
  });

  it("returns #NUM! like VBA once the bit is past every target MD", () => {
    const targets = { md: [100, 200], inc: [10, 20], azm: [0, 0], tvd: [50, 140] };
    expect(callUdf("PROJACTIVETARGETMDEX", [250, 12, 80, targets.md, targets.inc, targets.azm, targets.tvd]).error).toBe("NUM");
    expect(callUdf("PROJBURR", [250, 10, 0, 3, 10, targets.md, targets.inc, targets.azm, targets.tvd]).error).toBe("NUM");
  });

  it("formats a target toolface with Excel rounding", () => {
    expect(callUdf("PROJTFTOTARGET", [30, 0, 30, 0, 5]).value).toBe("0");
    expect(callUdf("PROJTFTOTARGET", [30, 0, 40, 0, 5]).value).toBe("0");
    expect(callUdf("PROJTFTOTARGET", [30, 0, 20, 0, 5]).value).toBe("L180");
    expect(callUdf("PROJTFTOTARGET", [30, 0, 30, 10, 5]).value).toMatch(/^R\d+$/);
  });

  it("lands on a landing station's TVD with a circular arc", () => {
    const value = callUdf("PROJLANDINGBURR", [1000, 60, 0, 900, 30, [1200], [90], [0], [1000]]).value as number;
    const radius = 100 / (Math.sin(Math.PI / 2) - Math.sin(Math.PI / 3));
    expect(value).toBeCloseTo((30 / radius) * 180 / Math.PI, 6);
  });

  it("estimates TD from the plan offset at the bit", () => {
    const plan = [
      { md: 0, inc: 0, azm: 0, tvd: 0, n: 0, e: 0 },
      { md: 1000, inc: 0, azm: 0, tvd: 1000, n: 0, e: 0 },
    ];
    const surveys = [[500, 500, 0, 0, 500]];
    const onPlan = callUdf("TDESTFINAL", ["MD", surveys, null, [[1000, 0, 0, 0, 0]]], { plan }).value as number;
    expect(onPlan).toBeCloseTo(1000, 6);
    expect(callUdf("TDESTFINAL", ["TVD", surveys, null, [[1000, 0, 0, 0, 0]]], { plan }).value).toBeCloseTo(1000, 6);
  });

  it("rebuilds wellhead N/E from the survey angles", () => {
    const ne = integrateLocalNE([0, 100, 200], [0, 0, 30], [0, 0, 90]);
    expect(ne.ns[0]).toBeCloseTo(0, 6);
    expect(ne.ew[0]).toBeCloseTo(0, 6);
    expect(ne.ew[1]).toBeCloseTo(0, 6);
    expect(ne.ew[2]!).toBeGreaterThan(20);
    expect(Math.abs(ne.ns[2]!)).toBeLessThan(0.05);
  });

  it("signs the proximity offset up and right", () => {
    const plan = [
      { md: 0, inc: 0, azm: 0, tvd: 0, n: 0, e: 0 },
      { md: 100, inc: 0, azm: 0, tvd: 100, n: 0, e: 0 },
    ];
    const offset = closestOffset(50, 0, 4, 50, plan);
    expect(offset?.side).toBeCloseTo(4, 2);
    expect(offset?.vertical).toBeCloseTo(0, 2);
  });

  it("matches Well Seeker X 3D normal-plane Up/Dn, Left/Right, and Dist across 40°, 65°, and 90°", () => {
    const plan40 = [
      { md: 2850.0, inc: 39.45, azm: 148.08, tvd: 2762.62, n: 66.94, e: -146.65 },
      { md: 2880.0, inc: 43.95, azm: 148.08, tvd: 2785.01, n: 50.01, e: -136.10 },
      { md: 2910.0, inc: 48.45, azm: 148.08, tvd: 2805.77, n: 31.64, e: -124.65 },
      { md: 2940.0, inc: 52.95, azm: 148.08, tvd: 2824.77, n: 11.94, e: -112.38 },
      { md: 2970.0, inc: 57.45, azm: 148.08, tvd: 2841.89, n: -8.96, e: -99.36 },
    ];
    const off40 = closestOffset(2902.35, 44.79, -134.02, 2815.80, plan40)!;
    expect(off40.vertical).toBeCloseTo(-18.23, 1);
    expect(off40.side).toBeCloseTo(1.00, 1);
    expect(off40.total).toBeCloseTo(18.26, 1);

    const plan65 = [
      { md: 2940.0, inc: 52.95, azm: 148.08, tvd: 2824.77, n: 11.94, e: -112.38 },
      { md: 2970.0, inc: 57.45, azm: 148.08, tvd: 2841.89, n: -8.96, e: -99.36 },
      { md: 3000.0, inc: 61.95, azm: 148.08, tvd: 2857.02, n: -30.94, e: -85.66 },
      { md: 3030.0, inc: 66.45, azm: 148.08, tvd: 2870.08, n: -53.86, e: -71.39 },
      { md: 3060.0, inc: 70.95, azm: 148.08, tvd: 2880.97, n: -77.57, e: -56.61 },
    ];
    const off65 = closestOffset(2998.41, -22.94, -101.52, 2874.34, plan65)!;
    expect(off65.vertical).toBeCloseTo(-22.46, 1);
    expect(off65.side).toBeCloseTo(9.23, 1);
    expect(off65.total).toBeCloseTo(24.28, 1);

    const plan90 = [
      { md: 3122.01, inc: 80.25, azm: 148.08, tvd: 2896.38, n: -128.49, e: -24.89 },
      { md: 3150.0, inc: 84.45, azm: 148.08, tvd: 2900.11, n: -152.03, e: -10.22 },
      { md: 3180.0, inc: 88.95, azm: 148.08, tvd: 2901.84, n: -177.45, e: 5.62 },
      { md: 3187.02, inc: 90.0, azm: 148.08, tvd: 2901.90, n: -183.41, e: 9.33 },
    ];
    const off90 = closestOffset(3171.13, -166.14, -11.30, 2903.57, plan90)!;
    expect(off90.vertical).toBeCloseTo(-2.52, 1);
    expect(off90.side).toBeCloseTo(8.37, 1);
    expect(off90.total).toBeCloseTo(8.75, 1);
  });

  it("integrates the first survey from Row 12 tie-on in PROJCUMN / PROJCUME", () => {
    const step = minCurveStep(5.55, 0, 0, 100, 1.2, 90);
    expect(callUdf("PROJCUMN", [0, 5.55, 0, 0, 100, 1.2, 90, true]).value).toBeCloseTo(step.n, 6);
    expect(callUdf("PROJCUME", [0, 5.55, 0, 0, 100, 1.2, 90, true]).value).toBeCloseTo(step.e, 6);
    // Blank Row 12 falls back to surface (0, 0, 0)
    const fromZero = minCurveStep(0, 0, 0, 100, 1.2, 90);
    expect(callUdf("PROJCUME", [0, "", "", "", 100, 1.2, 90, true]).value).toBeCloseTo(fromZero.e, 6);
  });
});
