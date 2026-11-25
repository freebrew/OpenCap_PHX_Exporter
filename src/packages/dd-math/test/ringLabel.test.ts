import { describe, expect, it } from "vitest";
import { ringLabelPoint, ringMetresLabel, ringRadii } from "../../wells-3d/src/ringLabel";

describe("ring distance labels", () => {
  it("writes a whole number of metres", () => {
    expect(ringMetresLabel(20)).toBe("20 m");
    expect(ringMetresLabel(10)).toBe("10 m");
  });

  it("uses a 10 m ring inside a 20 m ring", () => {
    expect(ringRadii()).toEqual({ outer: 20, inner: 10 });
  });

  it("sits the words just inside the ring", () => {
    const at = ringLabelPoint({ x: 0, y: 0, z: 0 }, { x: 0, y: -10, z: 0 }, 1.2);
    expect(at.x).toBeCloseTo(0);
    expect(at.y).toBeCloseTo(-8.8);
    expect(at.z).toBeCloseTo(0);
  });

  it("keeps a large label on its own ring", () => {
    const at = ringLabelPoint({ x: 0, y: 0, z: 0 }, { x: 10, y: 0, z: 0 }, 100);
    expect(at.x).toBeCloseTo(1.5);
    expect(at.y).toBeCloseTo(0);
  });
});
