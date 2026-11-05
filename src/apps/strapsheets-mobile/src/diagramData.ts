// The tool drawing above each step, as Entry.html shows it: the callout letter for the measure
// being entered is ringed and the rest of the drawing is dimmed.

import type { EntryStep } from "../../../packages/strapsheets-core/src/index";

export interface DiagramMark {
  letter: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface ToolDiagram {
  /** Inner SVG markup from the shared vector kit. The generator rejects scripts. */
  svg: string;
  size: { w: number; h: number };
  marks: DiagramMark[];
  /** Step labels whose letter on the drawing differs from the label's own letter. */
  labelLetters: Record<string, string>;
}

/** The drawing's letter for a step: "A - BitBox" is A, unless the drawing letters it otherwise. */
export function letterForStep(step: EntryStep | null, diagram: ToolDiagram | undefined): string {
  if (!step) return "";
  switch (step.kind) {
    case "text":
    case "decimal": {
      const override = diagram?.labelLetters[step.label];
      if (override) return override;
      return /^([A-Z])\s*-/.exec(step.label)?.[1] ?? "";
    }
    case "choice":
    case "boxPin":
    case "threads":
      return "";
    default: {
      const never: never = step;
      return never;
    }
  }
}
