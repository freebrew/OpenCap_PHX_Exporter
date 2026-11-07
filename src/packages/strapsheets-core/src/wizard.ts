// The entry wizard's rules, with no UI. One family's steps run in schema order; text is forced to
// capitals as it is typed, decimals keep digits and one dot, a blank measure skips, and a range
// only warns. The phone renders these; the tests pin them.

import { decimalText, roundProblem, upperText, type EntryPayload } from "./entry";
import type { InventoryItem } from "./inventory";
import { rangeFor, rangeWarning, type MeasureRange, type RangeWarning } from "./ranges";
import { stepColumns, type CellKey, type EntryFamily, type EntryStep } from "./schema";
import { customLines } from "./custom";

export type Cells = Record<string, string>;

/** The family a tool button names, or null when the schema has none. */
export function findFamily(families: EntryFamily[], which: string): EntryFamily | null {
  return families.find((family) => family.which === which) ?? null;
}

/** The Serial step, found by id. Subs opens on Sub Type; the other families open on Serial. */
export function serialStep(family: EntryFamily): EntryStep | null {
  return family.steps.find((step) => step.id === "serial") ?? null;
}

/** The serial an entry was taken for, upper-cased; "" when its family or serial is missing. */
export function entrySerial(families: EntryFamily[], entry: EntryPayload): string {
  const family = findFamily(families, entry.which);
  const step = family ? serialStep(family) : null;
  return step && step.kind === "text" ? (entry.cells[step.col] ?? "").trim().toUpperCase() : "";
}

/** Normalised serials (letters and digits only) that already have a strap sheet, for matching inventory. */
export function measuredSerials(families: EntryFamily[], entries: EntryPayload[]): Set<string> {
  const out = new Set<string>();
  for (const entry of entries) {
    const key = entrySerial(families, entry).replace(/[^A-Z0-9]/g, "");
    if (key !== "") out.add(key);
  }
  return out;
}

/**
 * What a keystroke leaves in the field. Text goes to capitals at once (the deployed form does the
 * same on input); a decimal drops anything but digits and keeps the first dot.
 */
export function liveValue(step: EntryStep, raw: string): string {
  switch (step.kind) {
    case "text":
      return raw.toUpperCase();
    case "decimal":
      return decimalDigits(raw);
    case "round":
      return step.decimal ? decimalDigits(raw) : raw.toUpperCase();
    case "choice":
    case "boxPin":
    case "threads":
    case "customBuilder":
      return raw;
    default: {
      const never: never = step;
      return never;
    }
  }
}

function decimalDigits(raw: string): string {
  const digits = raw.replace(/[^0-9.]/g, "");
  const dot = digits.indexOf(".");
  return dot < 0 ? digits : digits.slice(0, dot + 1) + digits.slice(dot + 1).replaceAll(".", "");
}

/** The cells with one column set; a blank value clears the column. */
export function setCell(cells: Cells, col: CellKey, value: string): Cells {
  const next = { ...cells };
  const key = String(col);
  if (value === "") delete next[key];
  else next[key] = value;
  return next;
}

/** Why the user cannot leave this step, or null. Only a required blank or a bad number stops them. */
export function stepProblem(step: EntryStep, cells: Cells): string | null {
  const values = stepColumns(step).map((col) => cells[col] ?? "");
  const requiredValues = step.kind === "text" ? values.slice(0, 1) : values;
  if (step.required && requiredValues.every((value) => value.trim() === "")) return `${step.label} is required.`;
  if (
    step.kind === "text" &&
    step.companionChoice?.required &&
    (cells[step.companionChoice.col] ?? "").trim() === ""
  ) {
    return `${step.companionChoice.label} is required.`;
  }
  if (step.kind === "decimal" && decimalText(values[0] ?? "") == null) return `${step.label} needs a number.`;
  if (step.kind === "round") return roundProblem(step, cells);
  return null;
}

/** Whether a measure sits outside the user's range. Never a reason to stop. */
export function stepWarning(family: EntryFamily, step: EntryStep, cells: Cells, ranges: MeasureRange[]): RangeWarning {
  if (step.kind !== "decimal") return null;
  const text = decimalText(cells[step.col] ?? "");
  if (text == null || text === "") return null;
  return rangeWarning(rangeFor(ranges, family.which, step.id), Number(text));
}

/** The steps a user can set a range on: the family's measures. */
export function measureSteps(family: EntryFamily): EntryStep[] {
  return family.steps.filter((step) => step.kind === "decimal");
}

/**
 * Prefill from a picked inventory item. The serial is set in capitals; measures the same serial
 * had on its newest earlier entry of this family fill only columns the user left blank.
 */
export function prefillFromItem(family: EntryFamily, cells: Cells, item: InventoryItem, history: EntryPayload[]): Cells {
  const serial = serialStep(family);
  if (!serial || serial.kind !== "text") return cells;
  let next = setCell(cells, serial.col, upperText(item.serialNumber));
  const earlier = [...history].reverse().find((entry) => entry.which === family.which && upperText(entry.cells[serial.col] ?? "") === upperText(item.serialNumber));
  if (!earlier) return next;
  for (const step of family.steps) {
    if (step.id === "serial") continue;
    for (const col of stepColumns(step)) {
      const known = earlier.cells[col];
      if (known && !next[col]) next = setCell(next, col, known);
    }
  }
  return next;
}

/**
 * The confirm screen: each filled step as label and value. A family whose every step writes a
 * numbered sheet column lists those lines in column order, the same way the web confirm does, so
 * Subs still shows Serial above Sub Type. Custom and MWD keep the order their steps are written in.
 */
export function entryLines(family: EntryFamily, cells: Cells): { label: string; value: string }[] {
  const lines: { label: string; value: string; col?: number }[] = [];
  for (const step of family.steps) {
    switch (step.kind) {
      case "text":
        if (cells[step.col]) lines.push({ label: step.label, value: cells[step.col] ?? "", col: typeof step.col === "number" ? step.col : undefined });
        if (step.companionChoice && cells[step.companionChoice.col]) {
          lines.push({
            label: step.companionChoice.label,
            value: cells[step.companionChoice.col] ?? "",
          });
        }
        break;
      case "decimal":
      case "choice":
        if (cells[step.col]) lines.push({ label: step.label, value: cells[step.col] ?? "", col: typeof step.col === "number" ? step.col : undefined });
        break;
      case "boxPin":
      case "threads":
        if (cells[step.cols.upper]) lines.push({ label: `${step.label} Upper`, value: cells[step.cols.upper] ?? "", col: typeof step.cols.upper === "number" ? step.cols.upper : undefined });
        if (cells[step.cols.lower]) lines.push({ label: `${step.label} Lower`, value: cells[step.cols.lower] ?? "", col: typeof step.cols.lower === "number" ? step.cols.lower : undefined });
        break;
      case "customBuilder":
        lines.push(...customLines(cells));
        break;
      case "round":
        for (const sub of step.cols) {
          if (cells[sub.col]) lines.push({ label: `${sub.letter} ${sub.short} ${step.field}`, value: cells[sub.col] ?? "", col: typeof sub.col === "number" ? sub.col : undefined });
        }
        break;
      default: {
        const never: never = step;
        return never;
      }
    }
  }
  const byColumn = family.steps.every((step) => {
    const cols = stepColumns(step);
    return cols.length > 0 && cols.every((col) => typeof col === "number");
  });
  if (byColumn) lines.sort((a, b) => (a.col ?? 0) - (b.col ?? 0));
  return lines.map(({ label, value }) => ({ label, value }));
}
