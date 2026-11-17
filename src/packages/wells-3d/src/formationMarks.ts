import type { WellFormation } from "./wellScene";

/** Where a formation's name sits along a wall. Close tops step sideways so the names stay apart. */
export interface FormationMark {
  name: string;
  /** True vertical depth on the plan datum, metres. */
  tvd: number;
  /** Scene up, metres. The dashed top is drawn here. */
  z: number;
  /** Fraction from the left of the wall as the wall is read. */
  along: number;
}

/** The same wording as the email picture: "TVD 1,590 m". */
export function formationTvdLabel(tvd: number): string {
  const rounded = tvd.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  return `TVD ${rounded} m`;
}

// Six places across the wall. A top closer than the gap takes the next place.
const ALONG = [0.06, 0.22, 0.38, 0.54, 0.7, 0.86];
const GAP_FRACTION = 0.03;

/**
 * The formations that fall on the shadow box, shallowest first. A top outside the box is left off,
 * the same clip as the email picture's shaft. Close tops step along the wall.
 */
export function formationMarks(minZ: number, maxZ: number, formations: readonly WellFormation[]): FormationMark[] {
  if (!(maxZ > minZ)) return [];
  const gap = Math.max(12, (maxZ - minZ) * GAP_FRACTION);
  const inside = formations
    .filter((fm) => fm.name.trim() !== "" && fm.z >= minZ - 0.5 && fm.z <= maxZ + 0.5)
    .slice()
    .sort((a, b) => b.z - a.z);
  const marks: FormationMark[] = [];
  let lastZ = Number.POSITIVE_INFINITY;
  let slot = 0;
  for (const fm of inside) {
    if (lastZ - fm.z < gap) slot = (slot + 1) % ALONG.length;
    else slot = 0;
    marks.push({ name: fm.name, tvd: fm.tvd, z: fm.z, along: ALONG[slot]! });
    lastZ = fm.z;
  }
  return marks;
}
