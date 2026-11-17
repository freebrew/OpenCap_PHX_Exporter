// Cont DI's two actions that touch cells. The button pulls the newest good survey into the
// tie-on row and keeps the last typed grid row; an edit in the yellow target block rewrites
// the aim formulas so the tab follows the Slidesheet's active station.

import { cellText, type CellScalar, type SheetBook } from "./sheetBook";
import { isGoodSurveyRow, lastSurveyRow, SS_SHEET } from "./slidesheet";

const CONT_DI = "Cont DI";
const CONT_DI_DATA = "ContDI_Data";
const FIRST_DATA = 8;
const LAST_DATA = 44;
const ENTRY_COLS = [1, 3, 4, 6, 7, 22];

const TARGETS = "ProjTargets_MD,ProjTargets_INC,ProjTargets_AZM,ProjTargets_TVD";
const AIM = `A2,B2,C2,ContDI_Data!T1,ContDI_Data!T4,ContDI_Data!U4,ContDI_Data!V1,${TARGETS},S1`;

/** ContDI_FollowSlideTarget's nine formulas, in P2, S1:S4, U1:U4 order. */
export const CONT_DI_AIM: ReadonlyArray<readonly [string, number, number]> = [
  ["P2", 2, 16],
  ["S1", 1, 19],
  ["S2", 2, 19],
  ["S3", 3, 19],
  ["S4", 4, 19],
  ["U1", 1, 21],
  ["U2", 2, 21],
  ["U3", 3, 21],
  ["U4", 4, 21],
];

function aimFormula(address: string): string {
  switch (address) {
    case "P2": return `=IFERROR(INDEX(ProjTargets_NAME,MATCH(S1,ProjTargets_MD,0)),"")`;
    case "S1": return `=IFERROR(INDEX(Slidesheet!$AO$13:$AO$505,MATCH(A2,Slidesheet!$E$13:$E$505,0)),"")`;
    case "S2": return `=IFERROR(INDEX(ProjTargets_INC,MATCH(S1,ProjTargets_MD,0)),"")`;
    case "S3": return `=IFERROR(INDEX(ProjTargets_AZM,MATCH(S1,ProjTargets_MD,0)),"")`;
    case "S4": return `=IFERROR(INDEX(ProjTargets_TVD,MATCH(S1,ProjTargets_MD,0)),"")`;
    case "U1": return `=IF(S1="","",IFERROR(ROUND(ProjContDiAim("BRR",${AIM}),2),""))`;
    case "U2": return `=IF(S1="","",IFERROR(ROUND(ProjContDiAim("TRR",${AIM}),2),""))`;
    case "U3": return `=IF(S1="","",IFERROR(ROUND(ProjContDiAim("DLR",${AIM}),2),""))`;
    case "U4": return `=IF(S1="","",IFERROR(ProjContDiAim("TF",${AIM}),""))`;
    default: return "";
  }
}

export interface ContDiPull {
  /** The Slidesheet row the tie-on came from, 0 when no good survey exists. */
  surveyRow: number;
  /** The grid row whose typed cells were kept as row 8, 0 when the grid was empty. */
  keptRow: number;
}

function scalar(value: unknown): CellScalar {
  return typeof value === "string" || typeof value === "number" || typeof value === "boolean" ? value : null;
}

function hasEntry(value: CellScalar): boolean {
  if (value == null) return false;
  return typeof value === "string" ? value.trim() !== "" : true;
}

/** PromoteLastGridRow: the deepest typed row in A/C/D/F/G becomes row 8; the block below clears. */
function promoteLastGridRow(book: SheetBook): number {
  let last = 0;
  for (let row = LAST_DATA; row >= FIRST_DATA; row -= 1) {
    if ([1, 3, 4, 6, 7].some((col) => hasEntry(scalar(book.value(CONT_DI, row, col))))) {
      last = row;
      break;
    }
  }
  const kept = ENTRY_COLS.map((col) => (last > 0 ? scalar(book.value(CONT_DI, last, col)) : null));
  for (let row = FIRST_DATA; row <= LAST_DATA; row += 1) {
    for (const col of ENTRY_COLS) book.setValue(CONT_DI, row, col, null);
  }
  if (last > 0) ENTRY_COLS.forEach((col, index) => book.setValue(CONT_DI, FIRST_DATA, col, kept[index] ?? null));
  return last;
}

/**
 * ContDI_PullLastSurvey. The newest good survey (green Inc and Azm) becomes the tie-on in
 * A2:D2, and the last typed grid row is kept as row 8. Returns null when the tab is absent.
 */
export function pullLastSurvey(book: SheetBook): ContDiPull | null {
  if (!book.hasSheet(CONT_DI)) return null;
  const keptRow = promoteLastGridRow(book);
  const surveyRow = lastSurveyRow(book);
  if (surveyRow > 0) {
    const columns: Array<[number, number]> = [[1, 5], [2, 6], [3, 7], [4, 8]];
    for (const [to, from] of columns) book.setValue(CONT_DI, 2, to, scalar(book.value(SS_SHEET, surveyRow, from)));
    const course = scalar(book.value(SS_SHEET, surveyRow, 3));
    if (book.hasSheet(CONT_DI_DATA)) book.setValue(CONT_DI_DATA, 1, 24, course);
  }
  return { surveyRow, keptRow };
}

/** The yellow target block: P2 and S1:S4. An edit here is what ContDI_ApplyTargetSelection watches. */
export function contDiTargetHit(row: number, col: number): boolean {
  return (row === 2 && col === 16) || (col === 19 && row >= 1 && row <= 4);
}

/**
 * ContDI_FollowSlideTarget, which is all ContDI_ApplyTargetSelection does: rewrite the aim
 * formulas that differ, so the tab's target follows the Slidesheet. True when a formula changed.
 */
export function followSlideTarget(book: SheetBook): boolean {
  if (!book.hasSheet(CONT_DI)) return false;
  let changed = false;
  for (const [address, row, col] of CONT_DI_AIM) {
    const formula = aimFormula(address);
    if (cellText(book.formula(CONT_DI, row, col)) === formula) continue;
    book.setFormula(CONT_DI, row, col, formula);
    changed = true;
  }
  if (book.hasSheet(CONT_DI_DATA)) {
    const aimH = `'Cont DI'!A2,'Cont DI'!B2,'Cont DI'!C2,T1,T4,U4,V1,${TARGETS},'Cont DI'!S1`;
    const formula = `=IF('Cont DI'!S1="","",IFERROR(ROUND(ProjContDiAim("TRRS",${aimH}),2),""))`;
    if (cellText(book.formula(CONT_DI_DATA, 2, 24)) !== formula) {
      book.setFormula(CONT_DI_DATA, 2, 24, formula);
      changed = true;
    }
  }
  return changed;
}
