import type { OffsetPath } from "./wellScene";

// Which offset wells are drawn. A concern well is drawn unless a filter drops it. Any other well
// is drawn only with Show All Offset Wells, and then only when it passes the same filters.

export interface WellFilter {
  /** Draw wells that are not of concern. */
  allOffsets: boolean;
  /** "all", "pad", or one site heading from the report. */
  site: string;
  /** Keep a well whose lowest separation factor is below this. Null keeps every well. */
  maxSf: number | null;
  /** Keep a well whose closest centre-to-centre is within this, in metres. Null keeps every well. */
  maxC2c: number | null;
}

/** The offset names to draw, in the scene's order. */
export function shownOffsets(offsets: OffsetPath[], filter: WellFilter): string[] {
  const out: string[] = [];
  for (const offset of offsets) {
    if (!offset.concern && !filter.allOffsets) continue;
    if (filter.site === "pad") {
      if (!offset.onPad) continue;
    } else if (filter.site !== "all" && offset.site !== filter.site) continue;
    if (filter.maxSf != null && !(offset.bandSf < filter.maxSf)) continue;
    if (filter.maxC2c != null && !(offset.minC2c > 0 && offset.minC2c <= filter.maxC2c)) continue;
    out.push(offset.name);
  }
  return out;
}
