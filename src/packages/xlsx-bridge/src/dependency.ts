import { extractRefs, sheetTitle } from "./a1.js";
import type { CellFormulaRecord } from "./types.js";

const SMALL_RANGE = 32;

function idOf(sheet: string, row: number, col: number): string {
  return `${sheet}\u0000${row}\u0000${col}`;
}

/**
 * Longest formula-to-formula chain.
 * Ranges larger than 32 cells count as a single step so a full-column
 * reference cannot turn the safety check into an accidental cartesian product.
 */
export function observedDependencyDepth(formulas: CellFormulaRecord[]): number {
  const formulaAt = new Set<string>();
  for (const formula of formulas) formulaAt.add(idOf(formula.sheet, formula.row, formula.col));

  const deps = new Map<string, string[]>();
  for (const formula of formulas) {
    const edges: string[] = [];
    const self = idOf(formula.sheet, formula.row, formula.col);
    for (const ref of extractRefs(formula.formula)) {
      const sheet = sheetTitle(ref.sheet) ?? formula.sheet;
      if (ref.kind === "cols" || ref.kind === "rows") {
        edges.push(`range:${sheet}`);
        continue;
      }
      const r1 = Math.min(ref.row1, ref.row2);
      const r2 = Math.max(ref.row1, ref.row2);
      const c1 = Math.min(ref.col1, ref.col2);
      const c2 = Math.max(ref.col1, ref.col2);
      const area = (r2 - r1 + 1) * (c2 - c1 + 1);
      if (area > SMALL_RANGE) {
        edges.push(`range:${sheet}`);
        continue;
      }
      for (let row = r1; row <= r2; row += 1) {
        for (let col = c1; col <= c2; col += 1) {
          const otherId = idOf(sheet, row, col);
          if (otherId !== self && formulaAt.has(otherId)) edges.push(otherId);
        }
      }
    }
    deps.set(self, edges);
  }

  const memo = new Map<string, number>();
  const visiting = new Set<string>();
  const depthOf = (id: string): number => {
    const known = memo.get(id);
    if (known !== undefined) return known;
    if (visiting.has(id)) return 0;
    visiting.add(id);
    let max = 0;
    for (const edge of deps.get(id) ?? []) {
      const next = edge.startsWith("range:") ? 1 : 1 + depthOf(edge);
      if (next > max) max = next;
    }
    visiting.delete(id);
    memo.set(id, max);
    return max;
  };

  let max = 0;
  for (const id of formulaAt) {
    const depth = depthOf(id);
    if (depth > max) max = depth;
  }
  return max;
}
