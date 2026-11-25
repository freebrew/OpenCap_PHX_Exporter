/** Excel freeze, as Univer stores it. */
export interface SheetFreeze {
  xSplit: number;
  ySplit: number;
  startRow: number;
  startColumn: number;
}

interface FreezeSheet {
  name?: string;
  freeze?: SheetFreeze;
}

interface FreezeBook {
  sheets?: Record<string, FreezeSheet>;
}

/** The column freeze taken off a sheet before it is shown. */
export interface KeptColumnFreeze {
  id: string;
  name: string;
  freeze: SheetFreeze;
}

/**
 * Univer 1.4 paints a frozen column with the default number format and alignment,
 * and draws that column's fill one cell to the right. Cont DI is the sheet that
 * freezes column A, so Bit Depth lost its yellow, centre, and 0.00. Row freeze stays.
 * The real freeze is put back when the snapshot is saved, so Excel is unchanged.
 */
export function presentColumnFreeze<T>(snapshot: T): { view: T; kept: KeptColumnFreeze[] } {
  const book = snapshot as FreezeBook;
  if (!book?.sheets) return { view: snapshot, kept: [] };
  const kept: KeptColumnFreeze[] = [];
  let changed = false;
  const sheets: Record<string, FreezeSheet> = {};
  for (const [id, sheet] of Object.entries(book.sheets)) {
    const freeze = sheet?.freeze;
    if (freeze && freeze.xSplit > 0) {
      kept.push({ id, name: sheet.name ?? "", freeze: { ...freeze } });
      sheets[id] = { ...sheet, freeze: { ...freeze, xSplit: 0, startColumn: -1 } };
      changed = true;
    } else {
      sheets[id] = sheet;
    }
  }
  if (!changed) return { view: snapshot, kept };
  return { view: { ...(snapshot as object), sheets } as T, kept };
}

/** Put the Excel column freeze back onto a snapshot Univer saved without it. */
export function restoreColumnFreeze<T>(saved: T, kept: readonly KeptColumnFreeze[]): T {
  if (kept.length === 0) return saved;
  const book = saved as FreezeBook;
  if (!book?.sheets) return saved;
  const byId = new Map(kept.map((entry) => [entry.id, entry]));
  const byName = new Map(kept.filter((entry) => entry.name !== "").map((entry) => [entry.name, entry]));
  let changed = false;
  const sheets: Record<string, FreezeSheet> = {};
  for (const [id, sheet] of Object.entries(book.sheets)) {
    const match = byId.get(id) ?? (sheet.name ? byName.get(sheet.name) : undefined);
    if (!match) {
      sheets[id] = sheet;
      continue;
    }
    sheets[id] = { ...sheet, freeze: { ...match.freeze } };
    changed = true;
  }
  if (!changed) return saved;
  return { ...(saved as object), sheets } as T;
}
