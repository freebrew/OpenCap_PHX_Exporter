// The min and max a measure step accepts. The sheet holds them (FormOptions, Ranges block) and
// the user edits them in Settings. A value outside the range warns; it never blocks the entry.

import { z } from "zod";

export interface MeasureRange {
  family: string;
  stepId: string;
  min: number | null;
  max: number | null;
}

/** The wire shape of a Settings → Ranges save. */
export const measureRangesSchema = z.array(z.object({
  family: z.string().min(1),
  stepId: z.string().min(1),
  min: z.number().finite().nullable(),
  max: z.number().finite().nullable(),
}));

export type RangeWarning = "below" | "above" | null;

/** Where a number sits against its range. No range, or a blank value, never warns. */
export function rangeWarning(range: MeasureRange | undefined, value: number | null): RangeWarning {
  if (!range || value == null) return null;
  if (range.min != null && value < range.min) return "below";
  if (range.max != null && value > range.max) return "above";
  return null;
}

/**
 * Replace one family's step range. A step not listed is added; min above max is swapped so the
 * range always reads low to high. The other families are untouched.
 */
export function setRange(ranges: MeasureRange[], next: MeasureRange): MeasureRange[] {
  const min = next.min != null && next.max != null && next.min > next.max ? next.max : next.min;
  const max = next.min != null && next.max != null && next.min > next.max ? next.min : next.max;
  const stored: MeasureRange = { family: next.family, stepId: next.stepId, min, max };
  const rest = ranges.filter((range) => !(range.family === next.family && range.stepId === next.stepId));
  return [...rest, stored];
}

/** The range for one step, when the user has set one. */
export function rangeFor(ranges: MeasureRange[], family: string, stepId: string): MeasureRange | undefined {
  return ranges.find((range) => range.family === family && range.stepId === stepId);
}
