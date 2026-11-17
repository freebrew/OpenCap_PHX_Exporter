import { shiftFormula } from "../../xlsx-bridge/src/a1";
import { SheetBook } from "./sheetBook";
import { SS_SHEET } from "./slidesheet";

// Slidesheet survey rows carry one formula per column, filled down. In Excel a deleted or cut
// cell in the Inc column (F) leaves #REF! in the formulas that read it, and the two rows below
// read one row too high. The cumulative North/East columns (AL/AM) add each row to the one above,
// so a single damaged row moves every row beneath it, and the proximity gauge, the aim columns and
// the slide advice all start from the wrong place. Job 35784 arrived from Excel with three such
// clusters, 104 m off in the gauge.

const FIRST_ROW = 14;
const LAST_ROW = 505;
const LAST_COL = 60;
/** Any row number beyond the sheet: formulas are compared as "this row's pattern" at this row. */
const BASE_ROW = 600;
/** A column needs this many formulas, and this share on one pattern, before its odd rows are treated as damage. */
const MIN_FORMULAS = 20;
const SHARE = 0.9;

export interface RepairedCell {
  row: number;
  col: number;
  before: string;
  after: string;
}

/**
 * Rewrites a survey-row formula that has left its column's pattern, from the pattern the rest of
 * the column follows. Typed values, sparse columns and mixed columns are never touched.
 */
export function repairSurveyFormulas(book: SheetBook): RepairedCell[] {
  const fixed: RepairedCell[] = [];
  for (let col = 1; col <= LAST_COL; col += 1) {
    const found: Array<{ row: number; formula: string; key: string }> = [];
    for (let row = FIRST_ROW; row <= LAST_ROW; row += 1) {
      const formula = book.formula(SS_SHEET, row, col);
      if (formula && formula.startsWith("=")) found.push({ row, formula, key: shiftFormula(formula, BASE_ROW - row, 0) });
    }
    if (found.length < MIN_FORMULAS) continue;
    const counts = new Map<string, number>();
    for (const entry of found) counts.set(entry.key, (counts.get(entry.key) ?? 0) + 1);
    let pattern = "";
    let best = 0;
    for (const [key, count] of counts) {
      if (count > best) {
        pattern = key;
        best = count;
      }
    }
    if (best / found.length < SHARE || pattern.includes("#REF!")) continue;
    for (const entry of found) {
      if (entry.key === pattern) continue;
      const after = shiftFormula(pattern, entry.row - BASE_ROW, 0);
      book.setFormula(SS_SHEET, entry.row, col, after);
      fixed.push({ row: entry.row, col, before: entry.formula, after });
    }
  }
  return fixed;
}
