import { describe, expect, it } from "vitest";
import { shownOffsets, type WellFilter } from "../../wells-3d/src/wellFilter";
import type { OffsetPath } from "../../wells-3d/src/wellScene";

function offset(name: string, over: Partial<OffsetPath>): OffsetPath {
  return {
    xyz: new Float32Array(),
    md: new Float32Array(),
    tvd: new Float32Array(),
    c2c: new Float32Array(),
    sf: new Float32Array(),
    name,
    concern: false,
    bandSf: 4,
    color: "#2E7D32",
    site: "Far Site",
    onPad: false,
    minC2c: 200,
    wellhead: null,
    breaks: [],
    ...over,
  };
}

const OFFSETS = [
  offset("Near", { concern: true, bandSf: 1.2, site: "Pad Site", onPad: true, minC2c: 12 }),
  offset("Pad Far", { concern: false, bandSf: 3, site: "Pad Site", onPad: true, minC2c: 80 }),
  offset("Away", { concern: true, bandSf: 6, site: "Far Site", onPad: false, minC2c: 400 }),
];

const OPEN: WellFilter = { allOffsets: true, site: "all", maxSf: null, maxC2c: null };

describe("offset filter", () => {
  it("draws the concern wells until Show All is on", () => {
    expect(shownOffsets(OFFSETS, { ...OPEN, allOffsets: false })).toEqual(["Near", "Away"]);
    expect(shownOffsets(OFFSETS, OPEN)).toEqual(["Near", "Pad Far", "Away"]);
  });

  it("keeps a site, the pad, a separation factor and a distance", () => {
    expect(shownOffsets(OFFSETS, { ...OPEN, site: "pad" })).toEqual(["Near", "Pad Far"]);
    expect(shownOffsets(OFFSETS, { ...OPEN, site: "Far Site" })).toEqual(["Away"]);
    expect(shownOffsets(OFFSETS, { ...OPEN, maxSf: 2 })).toEqual(["Near"]);
    expect(shownOffsets(OFFSETS, { ...OPEN, maxSf: 5 })).toEqual(["Near", "Pad Far"]);
    expect(shownOffsets(OFFSETS, { ...OPEN, maxC2c: 50 })).toEqual(["Near"]);
    expect(shownOffsets(OFFSETS, { ...OPEN, maxC2c: 100 })).toEqual(["Near", "Pad Far"]);
  });

  it("applies every chosen filter together", () => {
    expect(shownOffsets(OFFSETS, { allOffsets: false, site: "pad", maxSf: 5, maxC2c: 100 })).toEqual(["Near"]);
    expect(shownOffsets(OFFSETS, { ...OPEN, site: "missing" })).toEqual([]);
  });
});
