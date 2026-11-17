interface StyledCell {
  s?: unknown;
  [key: string]: unknown;
}

interface StyledSheet {
  name?: string;
  cellData?: Record<string, Record<string, StyledCell>>;
  columnData?: unknown;
  rowData?: unknown;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
  showGridlines?: number;
  zoomRatio?: number;
  [key: string]: unknown;
}

interface Resource {
  name: string;
  data: string;
}

interface StyledBook {
  styles?: Record<string, unknown>;
  sheets?: Record<string, StyledSheet>;
  resources?: Resource[];
  [key: string]: unknown;
}

const VALIDATION = "SHEET_DATA_VALIDATION_PLUGIN";

/**
 * Copy cell appearance and dropdown rules from a fresh import onto an existing snapshot.
 * Values, formulas, merges, and sheet order in `target` are kept; only style ids, the style table,
 * row/column sizes, gridlines, zoom, and data-validation rules come from `source`. Sheets are matched by name.
 */
export function restyleWorkbook<T extends object>(target: T, source: object): T {
  const into = target as StyledBook;
  const from = source as StyledBook;
  const byName = new Map(Object.values(from.sheets ?? {}).map((sheet) => [sheet.name, sheet]));
  const sheets: Record<string, StyledSheet> = {};
  for (const [id, sheet] of Object.entries(into.sheets ?? {})) {
    const fresh = byName.get(sheet.name);
    if (!fresh) {
      sheets[id] = sheet;
      continue;
    }
    const cellData: Record<string, Record<string, StyledCell>> = {};
    for (const [row, cells] of Object.entries(sheet.cellData ?? {})) {
      const out: Record<string, StyledCell> = {};
      for (const [col, cell] of Object.entries(cells)) {
        const { s: _old, ...rest } = cell;
        const style = fresh.cellData?.[row]?.[col]?.s;
        out[col] = style === undefined ? rest : { ...rest, s: style };
      }
      cellData[row] = out;
    }
    for (const [row, cells] of Object.entries(fresh.cellData ?? {})) {
      for (const [col, cell] of Object.entries(cells)) {
        if (cell.s === undefined || cellData[row]?.[col]) continue;
        cellData[row] ??= {};
        cellData[row][col] = { s: cell.s };
      }
    }
    sheets[id] = {
      ...sheet,
      cellData,
      columnData: fresh.columnData,
      rowData: fresh.rowData,
      defaultColumnWidth: fresh.defaultColumnWidth,
      defaultRowHeight: fresh.defaultRowHeight,
      showGridlines: fresh.showGridlines,
      zoomRatio: fresh.zoomRatio,
    };
  }
  return { ...into, styles: { ...(from.styles ?? {}) }, sheets, resources: withValidation(into, from) } as T;
}

function withValidation(into: StyledBook, from: StyledBook): Resource[] {
  const kept = (into.resources ?? []).filter((resource) => resource.name !== VALIDATION && resource.name !== "phx-data-validation");
  const rules = parse(from.resources?.find((resource) => resource.name === VALIDATION)?.data);
  const idByName = new Map(Object.entries(into.sheets ?? {}).map(([id, sheet]) => [sheet.name, id]));
  const out: Record<string, unknown> = {};
  for (const [sourceId, sheetRules] of Object.entries(rules)) {
    const id = idByName.get(from.sheets?.[sourceId]?.name);
    if (id) out[id] = sheetRules;
  }
  return [...kept, { name: VALIDATION, data: JSON.stringify(out) }];
}

function parse(json: string | undefined): Record<string, unknown> {
  if (!json) return {};
  try {
    const value = JSON.parse(json) as unknown;
    return value && typeof value === "object" ? value as Record<string, unknown> : {};
  } catch {
    return {};
  }
}
