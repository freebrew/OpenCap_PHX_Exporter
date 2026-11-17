import { toRangeA1 } from "./a1.js";

export interface GridPoint {
  row: number;
  col: number;
}

/** Merge a set of cells into the smallest list of rectangular A1 ranges. */
export function coalesceCells(cells: GridPoint[]): string[] {
  const present = new Set(cells.map((cell) => `${cell.row},${cell.col}`));
  const used = new Set<string>();
  const ranges: string[] = [];
  const ordered = [...cells].sort((a, b) => a.row - b.row || a.col - b.col);
  for (const cell of ordered) {
    const origin = `${cell.row},${cell.col}`;
    if (used.has(origin)) continue;
    let endCol = cell.col;
    while (present.has(`${cell.row},${endCol + 1}`) && !used.has(`${cell.row},${endCol + 1}`)) {
      endCol += 1;
    }
    let endRow = cell.row;
    let grew = true;
    while (grew) {
      for (let col = cell.col; col <= endCol; col += 1) {
        const key = `${endRow + 1},${col}`;
        if (!present.has(key) || used.has(key)) {
          grew = false;
          break;
        }
      }
      if (grew) endRow += 1;
    }
    for (let row = cell.row; row <= endRow; row += 1) {
      for (let col = cell.col; col <= endCol; col += 1) used.add(`${row},${col}`);
    }
    ranges.push(toRangeA1(cell.row, cell.col, endRow, endCol));
  }
  return ranges;
}
