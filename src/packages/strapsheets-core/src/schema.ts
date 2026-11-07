// The entry wizard, as data. The same shape apps-script/entry_schema.js generates, so the
// phone, the API and the extension render one wizard. A step's `kind` decides the control:
// text is ALL CAPS, decimal opens the number pad, choice / boxPin / threads are big buttons.

export type StepKind = "text" | "decimal" | "choice" | "boxPin" | "threads" | "customBuilder" | "round";

/** A form column index, or a named Custom field such as `serial` or `marksJson`. */
export type CellKey = number | string;

interface StepBase {
  id: string;
  label: string;
  required?: boolean;
}

/** A required button choice rendered on the same screen as a text field. */
export interface CompanionChoice {
  label: string;
  header?: string;
  col: CellKey;
  optionsKey: string;
  required?: boolean;
}

/** A single typed value. `col` is the 0-based column in Form Responses 1. */
export interface TextStep extends StepBase {
  kind: "text" | "decimal";
  /** Form header. Custom's named fields do not have one. */
  header?: string;
  col: CellKey;
  /** Motor Serial also asks whether the bend housing is fixed or adjustable. */
  companionChoice?: CompanionChoice;
}

/** One option from the family's list (`optionsKey` names which list). */
export interface ChoiceStep extends StepBase {
  kind: "choice";
  header?: string;
  col: CellKey;
  optionsKey: string;
}

/** A BOX/PIN pair: one column for the upper end, one for the lower. */
export interface BoxPinStep extends StepBase {
  kind: "boxPin";
  cols: { upper: CellKey; lower: CellKey };
  headers: { upper: string; lower: string };
}

/** A thread type for each end, from the family's thread list. */
export interface ThreadsStep extends StepBase {
  kind: "threads";
  cols: { upper: CellKey; lower: CellKey };
  headers: { upper: string; lower: string };
}

/** The Custom tube: length, ID, and the A–L marks. Values live in named cells, not form columns. */
export interface CustomBuilderStep extends StepBase {
  kind: "customBuilder";
}

/** One sub of a round: its letter on the drawing, its short name, and the cell it fills. */
export interface RoundCol {
  letter: string;
  short: string;
  col: CellKey;
}

/**
 * One screen holding a value for every sub of an assembly (the MWD Carrier's SN, Length and OD
 * rounds). `cumulative` lengths are running totals from the pin; the sheet stores each sub's own.
 */
export interface RoundStep extends StepBase {
  kind: "round";
  /** What the round measures: "SN", "Length", "OD". */
  field: string;
  cumulative: boolean;
  unit: string;
  /** A number each (Length, OD); otherwise text in capitals (SN). */
  decimal: boolean;
  cols: RoundCol[];
}

export type EntryStep = TextStep | ChoiceStep | BoxPinStep | ThreadsStep | CustomBuilderStep | RoundStep;

export interface EntryFamily {
  which: string;
  /** The Custom family draws its own tube instead of a stock diagram. */
  builder?: boolean;
  /** The MWD Carrier: four subs entered in rounds, printed on one card with computed totals. */
  mwd?: boolean;
  materials: string[];
  boxPinOptions: string[];
  /** Extra choice lists, keyed by the name a choice step's `optionsKey` points at. */
  extraOptions: Record<string, string[]>;
  steps: EntryStep[];
}

/** What one step writes. A pair step fills two columns; the rest fill one. */
export function stepColumns(step: EntryStep): CellKey[] {
  switch (step.kind) {
    case "text":
      return step.companionChoice
        ? [step.col, step.companionChoice.col]
        : [step.col];
    case "decimal":
    case "choice":
      return [step.col];
    case "boxPin":
    case "threads":
      return [step.cols.upper, step.cols.lower];
    case "customBuilder":
      return [];
    case "round":
      return step.cols.map((sub) => sub.col);
    default: {
      const never: never = step;
      return never;
    }
  }
}

/** The keyboard a step asks for. Text forces capitals; a decimal opens the number pad. */
export function stepInput(step: EntryStep): "text" | "decimal" | "buttons" | "builder" | "round" {
  switch (step.kind) {
    case "text":
      return "text";
    case "decimal":
      return "decimal";
    case "choice":
    case "boxPin":
    case "threads":
      return "buttons";
    case "customBuilder":
      return "builder";
    case "round":
      return "round";
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
export function choiceOptions(
  family: EntryFamily,
  step: Pick<ChoiceStep, "optionsKey">,
): string[] {
  switch (step.optionsKey) {
    case "materials":
      return family.materials;
    case "boxPinOptions":
      return family.boxPinOptions;
    default:
      return family.extraOptions[step.optionsKey] ?? [];
  }
}
