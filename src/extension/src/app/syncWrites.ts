import { parseA1, type CellOp, type ScalarCell } from "@phx/xlsx-bridge";
import type { CellWrite } from "../../../packages/workbook-commands/src/index";

/** One scalar from a visible-tab merge, as a Univer write. Text stays text, including "00" and "2-4". */
export function scalarWrites(sheet: string, a1: string, cell: ScalarCell): CellWrite[] {
  let row = 0;
  let col = 0;
  try {
    const pos = parseA1(a1);
    row = pos.row + 1;
    col = pos.col + 1;
  } catch {
    return [];
  }
  if (cell.kind === "blank") return [{ kind: "clear", sheet, row, col, row2: row, col2: col }];
  if (cell.kind === "number") {
    const n = Number(cell.value);
    if (!Number.isFinite(n)) return [];
    return [{ kind: "value", sheet, row, col, value: n }];
  }
  if (cell.kind === "bool") return [{ kind: "value", sheet, row, col, value: cell.value.toUpperCase() === "TRUE" }];
  return [{ kind: "value", sheet, row, col, value: cell.value }];
}

export function opWrites(op: CellOp): CellWrite[] {
  if (op.nextKind === "formula") return [];
  return scalarWrites(op.sheet, op.a1, { kind: op.nextKind, value: op.nextValue });
}
