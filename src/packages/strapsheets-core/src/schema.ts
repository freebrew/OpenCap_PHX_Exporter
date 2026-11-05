// The entry wizard, as data. The same shape apps-script/entry_schema.js generates, so the
// phone, the API and the extension render one wizard. A step's `kind` decides the control:
// text is ALL CAPS, decimal opens the number pad, choice / boxPin / threads are big buttons.

export type StepKind = "text" | "decimal" | "choice" | "boxPin" | "threads";

interface StepBase {
  id: string;
  label: string;
  required?: boolean;
}

/** A single typed value. `col` is the 0-based column in Form Responses 1. */
export interface TextStep extends StepBase {
  kind: "text" | "decimal";
  header: string;
  col: number;
}

/** One option from the family's list (`optionsKey` names which list). */
export interface ChoiceStep extends StepBase {
  kind: "choice";
  header: string;
  col: number;
  optionsKey: string;
}

/** A BOX/PIN pair: one column for the upper end, one for the lower. */
export interface BoxPinStep extends StepBase {
  kind: "boxPin";
  cols: { upper: number; lower: number };
  headers: { upper: string; lower: string };
}

/** A thread type for each end, from the family's thread list. */
export interface ThreadsStep extends StepBase {
  kind: "threads";
  cols: { upper: number; lower: number };
  headers: { upper: string; lower: string };
}

export type EntryStep = TextStep | ChoiceStep | BoxPinStep | ThreadsStep;

export interface EntryFamily {
  which: string;
  materials: string[];
  boxPinOptions: string[];
  /** Extra choice lists, keyed by the name a choice step's `optionsKey` points at. */
  extraOptions: Record<string, string[]>;
  steps: EntryStep[];
}

/** What one step writes. A pair step fills two columns; the rest fill one. */
export function stepColumns(step: EntryStep): number[] {
  switch (step.kind) {
    case "text":
    case "decimal":
    case "choice":
      return [step.col];
    case "boxPin":
    case "threads":
      return [step.cols.upper, step.cols.lower];
    default: {
      const never: never = step;
      return never;
    }
  }
}

/** The keyboard a step asks for. Text forces capitals; a decimal opens the number pad. */
export function stepInput(step: EntryStep): "text" | "decimal" | "buttons" {
  switch (step.kind) {
    case "text":
      return "text";
    case "decimal":
      return "decimal";
    case "choice":
    case "boxPin":
    case "threads":
      return "buttons";
    default: {
      const never: never = step;
      return never;
    }
  }
}

/**
 * The options a choice step offers. `materials` and `boxPinOptions` are the two named lists;
 * anything else comes from `extraOptions`. Unknown keys offer nothing.
 */
export function choiceOptions(family: EntryFamily, step: ChoiceStep): string[] {
  switch (step.optionsKey) {
    case "materials":
      return family.materials;
    case "boxPinOptions":
      return family.boxPinOptions;
    default:
      return family.extraOptions[step.optionsKey] ?? [];
  }
}
