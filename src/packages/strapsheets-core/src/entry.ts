// One strap entry, the payload submitEntry sends. Cells are keyed by their 0-based column in
// Form Responses 1. A serial is always stored in capitals; a decimal may be blank (skip).

import { z } from "zod";
import { customProblems, normalizeCustomCells } from "./custom";
import { choiceOptions, stepColumns, type EntryFamily, type EntryStep } from "./schema";

/** The wire shape of one entry, before it is checked against a family. */
export const entryPayloadSchema = z.object({
  which: z.string().min(1),
  cells: z.record(z.string(), z.string()),
});

export interface EntryPayload {
  which: string;
  /** Column index, or a Custom field name, to the text the user entered. */
  cells: Record<string, string>;
}

export interface EntryProblem {
  stepId: string;
  message: string;
}

/** A serial or any text step, forced to capitals with the surrounding space removed. */
export function upperText(value: string): string {
  return value.trim().toUpperCase();
}

/** A decimal step: digits with one dot, or blank when the user skips it. Null when it is not a number. */
export function decimalText(value: string): string | null {
  const trimmed = value.trim();
  if (trimmed === "") return "";
  return /^\d*\.?\d+$/.test(trimmed) && trimmed !== "." ? trimmed : null;
}

function cellValue(step: EntryStep, raw: string): string | null {
  switch (step.kind) {
    case "text":
      return upperText(raw);
    case "decimal":
      return decimalText(raw);
    case "choice":
    case "boxPin":
    case "threads":
    case "customBuilder":
      return raw.trim();
    default: {
      const never: never = step;
      return never;
    }
  }
}

/**
 * Check one entry against its family. A required step must be filled, a decimal must be a
 * number or blank, and a choice must be one of the options offered. The problems name the step.
 */
export function validateEntry(family: EntryFamily, payload: EntryPayload): EntryProblem[] {
  const problems: EntryProblem[] = [];
  if (payload.which !== family.which) problems.push({ stepId: "", message: `${payload.which} is not a ${family.which} entry.` });
  for (const step of family.steps) {
    const columns = stepColumns(step);
    const values = columns.map((col) => cellValue(step, payload.cells[col] ?? ""));
    if (step.required && values.every((value) => value === "")) {
      problems.push({ stepId: step.id, message: `${step.label} is required.` });
      continue;
    }
    if (step.kind === "decimal" && values.some((value) => value == null)) {
      problems.push({ stepId: step.id, message: `${step.label} needs a number.` });
    }
    if (step.kind === "choice") {
      const allowed = choiceOptions(family, step);
      if (values[0] !== "" && values[0] != null && !allowed.includes(values[0])) {
        problems.push({ stepId: step.id, message: `${step.label} must be one of ${allowed.join(", ")}.` });
      }
    }
    if (step.kind === "customBuilder") problems.push(...customProblems(payload.cells));
  }
  return problems;
}

/** The entry with its text columns in capitals and its blanks dropped, ready to store. */
export function normalizeEntry(family: EntryFamily, payload: EntryPayload): EntryPayload {
  const cells: Record<string, string> = {};
  for (const step of family.steps) {
    for (const col of stepColumns(step)) {
      const value = cellValue(step, payload.cells[col] ?? "");
      if (value) cells[String(col)] = value;
    }
    if (step.kind === "customBuilder") Object.assign(cells, normalizeCustomCells(payload.cells));
  }
  return { which: family.which, cells };
}
