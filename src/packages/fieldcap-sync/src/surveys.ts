import { type StationFields } from "./job";

export interface SurveyStation {
  md: number;
  inc: number;
  azm: number;
  tvd?: number;
  /** FieldCap's bit depth for the station, when it publishes one. */
  bit?: number;
}

function stationNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value !== "string" || value.trim() === "") return null;
  const parsed = Number(value.trim());
  return Number.isFinite(parsed) ? parsed : null;
}

/**
 * Directional stations from a SurveySheetEntries payload. Rows missing measured depth,
 * inclination or azimuth are skipped. The same measured depth keeps the later row.
 */
export function surveyStations(rows: ReadonlyArray<Record<string, unknown>>, fields: StationFields): SurveyStation[] {
  if (!fields.md || !fields.inc || !fields.azm) return [];
  const byMd = new Map<number, SurveyStation>();
  for (const row of rows) {
    const md = stationNumber(row[fields.md]);
    const inc = stationNumber(row[fields.inc]);
    const azm = stationNumber(row[fields.azm]);
    if (md == null || inc == null || azm == null) continue;
    const tvd = fields.tvd ? stationNumber(row[fields.tvd]) : null;
    const bit = fields.bit ? stationNumber(row[fields.bit]) : null;
    byMd.set(md, { md, inc, azm, ...(tvd == null ? {} : { tvd }), ...(bit == null ? {} : { bit }) });
  }
  return [...byMd.values()].sort((a, b) => a.md - b.md);
}
