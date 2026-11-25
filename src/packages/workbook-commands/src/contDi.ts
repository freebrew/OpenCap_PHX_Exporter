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
const INPUT_YELLOW = "#FFFFCC";
const GRID = "#969696";

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

/**
 * Bit Depth is a typed column (A2 and A8:A44). A frozen column paints that yellow one cell to the
 * right, so the preview writes the input fill, centre, and 0.00 onto column A itself.
 */
export function paintContDiColumnA(book: SheetBook): boolean {
  if (!book.hasSheet(CONT_DI)) return false;
  const ready = (row: number) => {
    const style = book.style(CONT_DI, row, 1);
    return (book.fill(CONT_DI, row, 1) ?? "").toUpperCase() === INPUT_YELLOW
      && style?.n?.pattern === "0.00"
      && style.ht === 2
      && style.bl !== 1;
  };
  if (ready(2) && ready(8) && ready(44)) return false;
  const look = { bold: false, hAlign: "center" as const, vAlign: "middle" as const, fontFamily: "Calibri", fontSize: 11, wrap: "clip" as const };
  for (const ref of ["A2", "A8:A44"]) {
    book.setFill(CONT_DI, ref, INPUT_YELLOW);
    book.setLook(CONT_DI, ref, look);
    book.setBorder(CONT_DI, ref, "outside", "thin", GRID);
  }
  book.setFormat(CONT_DI, 2, 1, "0.00");
  for (let row = FIRST_DATA; row <= LAST_DATA; row += 1) book.setFormat(CONT_DI, row, 1, "0.00");
  return true;
}

/**
 * Same formula text as MDL_ContDI.ApplyRowFormulas. A blank MWD inclination
 * (C) or azimuth (D, G) uses the last number above it, then the tie-on.
 * A blank PD inclination (F) uses the last number in that column only, so a
 * column that has never been typed does not invent a PD reading. ROP (V) has
 * no equation on this sheet. Returns null for cells that stay empty (the
 * first grid row has no station above it).
 */
export function contDiCalcFormula(row: number, col: number): string | null {
  if (row < FIRST_DATA || row > LAST_DATA) return null;
  const previous = row - 1;
  const rowOn = `ISNUMBER(A${row})`;
  const incC = heldReading("C", row, "$B$2");
  const incF = heldReading("F", row, `""`);
  const azD = heldReading("D", row, "$C$2");
  const azG = heldReading("G", row, "$C$2");
  const seenG = heldReading("G", row, `""`);
  const hasInc = `ISNUMBER(${incC})`;
  const hasPdInc = `ISNUMBER(${incF})`;
  const tie = "ISNUMBER($A$2)";
  const mwdDepth = `ISNUMBER(B${row}),B${row}>$A$2`;
  const pdDepth = `ISNUMBER(E${row}),E${row}>$A$2`;
  const mwdStep = `ISNUMBER(B${row}),ISNUMBER(B${previous}),B${row}<>B${previous}`;
  const pdStep = `ISNUMBER(E${row}),ISNUMBER(E${previous}),E${row}<>E${previous}`;
  const offsets = "ISNUMBER($O$2),ISNUMBER($M$2),$O$2<>$M$2";
  if (col === 8) {
    return `=IF(AND(${tie},${hasInc},${mwdDepth},ISNUMBER(${azD})),(30/(B${row}-$A$2))*DEGREES(ACOS(MIN(1,MAX(-1,COS(RADIANS($B$2))*COS(RADIANS(${incC}))+SIN(RADIANS($B$2))*SIN(RADIANS(${incC}))*COS(RADIANS(${azD}-$C$2)))))),"")`;
  }
  if (col === 15) {
    return `=IF(AND(${tie},${hasPdInc},${pdDepth},ISNUMBER(${azG})),(30/(E${row}-$A$2))*DEGREES(ACOS(MIN(1,MAX(-1,COS(RADIANS($B$2))*COS(RADIANS(${incF}))+SIN(RADIANS($B$2))*SIN(RADIANS(${incF}))*COS(RADIANS(${azG}-$C$2)))))),"")`;
  }
  if (col === 13) {
    return `=IF(AND(${rowOn},${hasInc},${hasPdInc},ISNUMBER(${azD}),ISNUMBER(${azG}),ISNUMBER($O$2),ISNUMBER($M$2),$O$2<>$M$2),ProjDoglegDeg(${incC},${azD},${incF},${azG})/ABS($O$2-$M$2)*30,"")`;
  }
  if (col === 14) {
    return `=IF(AND(${rowOn},${hasInc},${hasPdInc},ISNUMBER($O$2),ISNUMBER($M$2),$O$2<>$M$2),(${incF}-${incC})/($O$2-$M$2)*30,"")`;
  }
  if (col === 11) {
    return `=IF(AND(${tie},${hasInc},${mwdDepth}),(${incC}-$B$2)/(B${row}-$A$2)*30,"")`;
  }
  if (row >= 9 && col === 9) {
    const incPrev = heldReading("C", previous, "$B$2");
    return `=IF(AND(${hasInc},ISNUMBER(${incPrev}),${mwdStep}),(${incC}-${incPrev})/(B${row}-B${previous})*30,"")`;
  }
  if (row >= 9 && col === 16) {
    const incPrev = heldReading("F", previous, `""`);
    return `=IF(AND(${hasPdInc},ISNUMBER(${incPrev}),${pdStep}),(${incF}-${incPrev})/(E${row}-E${previous})*30,"")`;
  }
  if (col === 19) {
    return `=IF(NOT(${rowOn}),"",IF(${hasPdInc},IF(ISNUMBER(N${row}),${incF}+(N${row}/30)*$M$2,${incF}),IF(NOT(${hasInc}),"",IF(ISNUMBER(K${row}),(K${row}/30)*O$2+${incC},${incC}))))`;
  }
  if (col === 20) {
    return `=IF(NOT(${rowOn}),"",IF(OR(${hasPdInc},ISNUMBER(${seenG})),IF(AND(${offsets},ISNUMBER(${azD})),MOD(${azG}+(MOD(${azG}-${azD}+180,360)-180)/($O$2-$M$2)*$M$2,360),MOD(${azG},360)),IF(NOT(${hasInc}),"",IF(AND(${tie},${mwdDepth},ISNUMBER(${azD})),MOD(((MOD(${azD}-$C$2+180,360)-180)/(B${row}-$A$2))*O$2+${azD},360),IF(ISNUMBER(${azD}),MOD(${azD},360),"")))))`;
  }
  if (col === 21) {
    return `=IF(AND(ISNUMBER($R${row}),ISNUMBER($S${row}),ISNUMBER($T${row}),ISNUMBER($A$2),ISNUMBER($B$2),ISNUMBER($C$2),ISNUMBER($D$2)),ContDITvdAtBit($R${row},$S${row},$T${row},$A$2,$B$2,$C$2,$D$2,$M$2,$B$${FIRST_DATA}:$B${row},$C$${FIRST_DATA}:$C${row},$D$${FIRST_DATA}:$D${row},${incF},${azG}),"")`;
  }
  return null;
}

const HELD_COLS = [8, 9, 11, 13, 14, 15, 16, 19, 20, 21];

function sameFormula(current: string | null, next: string): boolean {
  const compact = (value: string) => value.trim().replace(/^=/, "").replace(/\s+/g, "").toUpperCase();
  return current != null && compact(current) === compact(next);
}

/** Rewrite the Cont DI rate and bit-projection formulas so a blank inc or azimuth carries forward. */
export function applyContDiHeldReadings(book: SheetBook): boolean {
  if (!book.hasSheet(CONT_DI)) return false;
  let changed = false;
  for (let row = FIRST_DATA; row <= LAST_DATA; row += 1) {
    for (const col of HELD_COLS) {
      const formula = contDiCalcFormula(row, col);
      if (!formula || sameFormula(book.formula(CONT_DI, row, col), formula)) continue;
      book.setFormula(CONT_DI, row, col, formula);
      changed = true;
    }
  }
  return changed;
}

function heldReading(col: string, row: number, tie: string): string {
  const here = `${col}${row}`;
  const previous = row > FIRST_DATA
    ? `IFERROR(LOOKUP(2,1/(${col}$${FIRST_DATA}:${col}${row - 1}<>""),${col}$${FIRST_DATA}:${col}${row - 1}),${tie})`
    : tie;
  return `IF(ISNUMBER(${here}),${here},${previous})`;
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
