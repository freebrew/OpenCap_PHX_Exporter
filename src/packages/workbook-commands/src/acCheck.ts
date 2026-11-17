import type { AcDetailStation } from "./acDetail";
import { wellsMatch } from "./acDetail";
import type { AcHit } from "./acImport";

// The three self-checks from MDL_AcPoc_Check: every concern well has a detail
// page, a Summary row agrees with its detail station, and the 3D distance from
// the plan to that station is the printed centre-to-centre distance.

export interface AcPlanPoint {
  md: number;
  n: number;
  e: number;
  tvd: number;
}

export interface AcCheck {
  ok: boolean;
  stations: number;
  worstM: number;
  badge: string;
  notes: string[];
}

const SF_TOLERANCE = 0.02;
const C2C_TOLERANCE = 0.05;
const FRAME_TOLERANCE = 0.5;

/** AcPoc_InterpPlan: linear in measured depth. */
export function interpPlan(plan: AcPlanPoint[], md: number): AcPlanPoint | null {
  if (plan.length < 2) return null;
  const first = plan[0]!;
  const last = plan[plan.length - 1]!;
  if (md <= first.md) return { ...first };
  if (md >= last.md) return { ...last };
  for (let i = 0; i < plan.length - 1; i += 1) {
    const left = plan[i]!;
    const right = plan[i + 1]!;
    if (md < left.md || md > right.md) continue;
    const span = right.md - left.md;
    const t = Math.abs(span) < 0.0001 ? 0 : (md - left.md) / span;
    return {
      md,
      n: left.n + t * (right.n - left.n),
      e: left.e + t * (right.e - left.e),
      tvd: left.tvd + t * (right.tvd - left.tvd),
    };
  }
  return null;
}

/**
 * How far a station is from the printed centre-to-centre, in metres: the 3D distance from the plan at
 * the station's Ref MD to the offset point, less the printed C2C. Null when the row has no C2C to test.
 */
export function frameResidual(plan: AcPlanPoint[], station: AcDetailStation): number | null {
  if (station.c2c <= 0) return null;
  const at = interpPlan(plan, station.refMd);
  if (!at) return null;
  const distance = Math.hypot(at.n - station.n, at.e - station.e, at.tvd - station.offTvd);
  return Math.abs(distance - station.c2c);
}

/** A station sits where the report says it does. */
export function stationInFrame(plan: AcPlanPoint[], station: AcDetailStation): boolean {
  const residual = frameResidual(plan, station);
  return residual != null && residual <= FRAME_TOLERANCE;
}

function fail(notes: string[], badge: string, stations: number, worstM: number): AcCheck {
  return { ok: false, stations, worstM, badge, notes };
}

/** Coverage, Summary-against-detail, and the frame proof. `shown` is the concern set, already chosen. */
export function checkAcReport(shown: AcHit[], stations: AcDetailStation[], plan: AcPlanPoint[]): AcCheck {
  const notes: string[] = [];
  for (const hit of shown) {
    if (!stations.some((station) => wellsMatch(hit.well, station.well))) {
      notes.push(`No detail page for ${hit.well}`);
    }
  }
  const missing = notes.length;

  for (const hit of shown) {
    const row = stations.find((station) => wellsMatch(hit.well, station.well) && Math.abs(station.refMd - hit.refMd) <= C2C_TOLERANCE);
    if (!row) {
      notes.push(`No detail station at Ref MD ${hit.refMd.toFixed(2)} for ${hit.well}`);
      continue;
    }
    if (row.sf > 0 && Math.abs(row.sf - hit.sf) > SF_TOLERANCE) {
      notes.push(`${hit.well} Ref MD ${hit.refMd.toFixed(2)} detail SF ${row.sf.toFixed(3)} summary SF ${hit.sf.toFixed(3)}`);
    }
    if (row.c2c > 0 && Math.abs(row.c2c - hit.between) > C2C_TOLERANCE) {
      notes.push(`${hit.well} C2C detail ${row.c2c.toFixed(2)} summary ${hit.between.toFixed(2)}`);
    }
  }
  const disagreed = notes.slice(missing).some((note) => note.includes("detail SF"));

  if (stations.length === 0) {
    return missing > 0
      ? fail(notes, "A concern well has no path", 0, 0)
      : { ok: true, stations: 0, worstM: 0, badge: "No offset paths in the report", notes };
  }
  if (plan.length < 2) return fail(notes, "No plan to check against", 0, 0);

  let checked = 0;
  let worst = 0;
  let bad = 0;
  for (const station of stations) {
    if (!shown.some((hit) => wellsMatch(hit.well, station.well))) continue;
    const residual = frameResidual(plan, station);
    if (residual == null) continue;
    checked += 1;
    if (residual > worst) worst = residual;
    if (residual > FRAME_TOLERANCE) {
      bad += 1;
      if (bad <= 8) notes.push(`${station.well} Ref MD ${station.refMd.toFixed(2)} is ${residual.toFixed(2)} m off the printed C2C`);
    }
  }
  if (plan.length >= 2 && checked === 0 && missing === 0) return fail(notes, "No plan to check against", 0, 0);
  if (bad > 0) return fail(notes, "Report and plan do not line up", checked, worst);
  if (disagreed) return fail(notes, "Report rows do not agree", checked, worst);
  if (missing > 0) return fail(notes, "A concern well has no path", checked, worst);
  if (checked === 0) return { ok: true, stations: 0, worstM: 0, badge: "No offset paths in the report", notes };
  return {
    ok: true,
    stations: checked,
    worstM: worst,
    badge: `Checked: ${checked} stations, worst ${worst.toFixed(2)} m`,
    notes,
  };
}
