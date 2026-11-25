// Browser shift -> Excel handoff. The .xlsm stays authoritative.
// A bundle is a directory of CSV payloads plus a manifest. READY is not part of
// the payload: the caller writes it last, and Excel ignores a folder without it.

import { indexToCol } from "./a1.js";

export const HANDOFF_SCHEMA = 1;

export type CellKind = "blank" | "number" | "bool" | "text" | "formula";

export interface CanonCell {
  k: CellKind;
  v: string;
  /** 2 array, 3 data table. Absent for a normal formula. */
  ft?: 2 | 3;
  s?: Record<string, string>;
}

export interface CanonSheet {
  name: string;
  hidden: 0 | 1 | 2;
  rowCount: number;
  columnCount: number;
  freeze: string;
  merges: string[];
  rows: Record<string, string>;
  cols: Record<string, string>;
  cells: Record<string, CanonCell>;
}

export interface CanonName {
  name: string;
  text: string;
  hidden: boolean;
}

export interface CanonBook {
  sheets: CanonSheet[];
  names: CanonName[];
  resources: Record<string, string>;
}

export type HandoffOp =
  | { kind: "cell"; sheet: string; a1: string; baseKind: CellKind; baseValue: string; nextKind: CellKind; nextValue: string }
  | { kind: "style"; sheet: string; a1: string; field: string; base: string; next: string }
  | { kind: "merge"; sheet: string; ref: string; action: "merge" | "unmerge" }
  | { kind: "geometry"; axis: "row" | "col"; sheet: string; index: number; field: "h" | "w" | "hd"; base: string; next: string }
  | { kind: "sheet"; name: string; field: "create" | "hidden" | "order"; base: string; next: string }
  | { kind: "name"; name: string; base: string; next: string };

export interface HandoffDiff {
  ops: HandoffOp[];
  blocked: string[];
}

export type CellDecision = "apply" | "idempotent" | "conflict";

const STYLE_FIELDS = ["bg", "cl", "bl", "it", "ul", "st", "ff", "fs", "ht", "vt", "n", "bd", "tb"] as const;
const BORDER_SIDES = ["l", "r", "t", "b"] as const;

/** Commands a normal shift must not run: the cell delta cannot represent them safely. */
const BLOCKED_COMMAND = /insert-row|insert-col|insert-multi|remove-row|remove-col|remove-rows|remove-sheet|move-range|move-rows|move-columns|reorder-range|set-worksheet-name|thread-comment|hyper-link|hyperlink|data-validation|conditional-formatting/;

const IGNORED_RESOURCE = /(_UI_PLUGIN$)|VIEW|SCROLL|ZOOM|SELECTION|SHEETS_FORMULA_PLUGIN$|SHEET_DEFINED_NAME_PLUGIN$/i;

export function commandBlocksHandoff(id: string): boolean {
  if (!id.includes(".command.")) return false;
  if (id === "sheet.command.insert-sheet") return false;
  if (id.endsWith("-protection")) return false;
  return BLOCKED_COMMAND.test(id);
}

export function classifyChange(current: string, base: string, next: string): CellDecision {
  if (current === next) return "idempotent";
  if (current === base) return "apply";
  return "conflict";
}

export function isCanonBook(value: unknown): value is CanonBook {
  if (!value || typeof value !== "object") return false;
  const book = value as Partial<CanonBook>;
  return Array.isArray(book.sheets) && Array.isArray(book.names) && !!book.resources && typeof book.resources === "object";
}

export function canonicalize(snapshot: unknown): CanonBook {
  const book = snapshot && typeof snapshot === "object" ? snapshot as Record<string, unknown> : {};
  const sheets = book.sheets && typeof book.sheets === "object" ? book.sheets as Record<string, Record<string, unknown>> : {};
  const styles = book.styles && typeof book.styles === "object" ? book.styles as Record<string, Record<string, unknown>> : {};
  const order = Array.isArray(book.sheetOrder) ? book.sheetOrder.map((id) => String(id)) : Object.keys(sheets);
  const seen = new Set<string>();
  const canonSheets: CanonSheet[] = [];
  for (const id of order) {
    const sheet = sheets[id];
    if (!sheet || seen.has(id)) continue;
    seen.add(id);
    canonSheets.push(canonSheet(sheet, styles));
  }
  for (const id of Object.keys(sheets)) {
    if (seen.has(id)) continue;
    canonSheets.push(canonSheet(sheets[id], styles));
  }
  const names = Array.isArray(book.definedNames) ? book.definedNames : [];
  const canonNames: CanonName[] = names.flatMap((entry) => {
    if (!entry || typeof entry !== "object") return [];
    const record = entry as Record<string, unknown>;
    const name = String(record.name ?? "").trim();
    if (!name) return [];
    return [{ name, text: String(record.formulaOrRefString ?? ""), hidden: record.hidden === true }];
  }).sort((a, b) => a.name.localeCompare(b.name));
  return {
    sheets: canonSheets,
    names: canonNames,
    resources: resourceMap(book.resources),
  };
}

export function diffHandoff(base: CanonBook, live: CanonBook, blocks: readonly string[] = []): HandoffDiff {
  const blocked = new Set<string>(blocks.map((id) => `command blocked: ${id}`));
  const ops: HandoffOp[] = [];
  const baseSheets = new Map(base.sheets.map((sheet) => [sheet.name, sheet]));
  const liveSheets = new Map(live.sheets.map((sheet) => [sheet.name, sheet]));

  for (const name of baseSheets.keys()) {
    if (!liveSheets.has(name)) blocked.add(`sheet removed: ${name}`);
  }
  for (const [name, live] of liveSheets) {
    const prior = baseSheets.get(name);
    if (!prior) {
      ops.push({ kind: "sheet", name, field: "create", base: "", next: String(live.hidden) });
      pushSheetBody(ops, blocked, emptySheet(name), live);
      continue;
    }
    if (prior.hidden !== live.hidden) {
      ops.push({ kind: "sheet", name, field: "hidden", base: String(prior.hidden), next: String(live.hidden) });
    }
    if (prior.rowCount > live.rowCount) blocked.add(`rows removed on ${name}`);
    if (prior.columnCount > live.columnCount) blocked.add(`columns removed on ${name}`);
    if (prior.freeze !== live.freeze) blocked.add(`freeze changed on ${name}`);
    pushSheetBody(ops, blocked, prior, live);
  }

  const baseOrder = base.sheets.map((sheet) => sheet.name).join("|");
  const liveOrder = live.sheets.map((sheet) => sheet.name).join("|");
  if (baseOrder !== liveOrder && !blockedHas(blocked, "sheet removed")) {
    ops.push({ kind: "sheet", name: "*", field: "order", base: baseOrder, next: liveOrder });
  }

  const baseNames = new Map(base.names.map((entry) => [entry.name, entry]));
  const liveNames = new Map(live.names.map((entry) => [entry.name, entry]));
  for (const [name, prior] of baseNames) {
    const next = liveNames.get(name);
    if (!next) {
      ops.push({ kind: "name", name, base: prior.text, next: "" });
      continue;
    }
    if (prior.hidden !== next.hidden) blocked.add(`defined name visibility changed: ${name}`);
    if (prior.text !== next.text) ops.push({ kind: "name", name, base: prior.text, next: next.text });
  }
  for (const [name, next] of liveNames) {
    if (!baseNames.has(name)) ops.push({ kind: "name", name, base: "", next: next.text });
  }

  const resourceNames = new Set([...Object.keys(base.resources), ...Object.keys(live.resources)]);
  for (const name of resourceNames) {
    if ((base.resources[name] ?? "") !== (live.resources[name] ?? "")) {
      blocked.add(`unsupported resource change: ${name}`);
    }
  }

  return { ops, blocked: [...blocked].sort() };
}

function blockedHas(blocked: Set<string>, prefix: string): boolean {
  for (const item of blocked) if (item.startsWith(prefix)) return true;
  return false;
}

function pushSheetBody(ops: HandoffOp[], blocked: Set<string>, prior: CanonSheet, live: CanonSheet): void {
  const addresses = new Set([...Object.keys(prior.cells), ...Object.keys(live.cells)]);
  for (const a1 of [...addresses].sort(compareA1)) {
    const baseCell = prior.cells[a1] ?? blankCell();
    const liveCell = live.cells[a1] ?? blankCell();
    if ((baseCell.ft ?? 0) !== (liveCell.ft ?? 0)) {
      blocked.add(`array formula structure changed: ${live.name}!${a1}`);
      continue;
    }
    if (baseCell.k !== liveCell.k || baseCell.v !== liveCell.v) {
      ops.push({
        kind: "cell",
        sheet: live.name,
        a1,
        baseKind: baseCell.k,
        baseValue: baseCell.v,
        nextKind: liveCell.k,
        nextValue: liveCell.v,
      });
    }
    const fields = new Set([...Object.keys(baseCell.s ?? {}), ...Object.keys(liveCell.s ?? {})]);
    for (const field of fields) {
      if (field === "__extra") {
        if ((baseCell.s?.__extra ?? "") !== (liveCell.s?.__extra ?? "")) {
          blocked.add(`unsupported style change: ${live.name}!${a1}`);
        }
        continue;
      }
      const base = baseCell.s?.[field] ?? "";
      const next = liveCell.s?.[field] ?? "";
      if (base === next) continue;
      if (field === "bd" && (base === "unsupported" || next === "unsupported")) {
        blocked.add(`unsupported border change: ${live.name}!${a1}`);
        continue;
      }
      ops.push({ kind: "style", sheet: live.name, a1, field, base, next });
    }
  }

  const baseMerges = new Set(prior.merges);
  const liveMerges = new Set(live.merges);
  for (const ref of prior.merges) {
    if (!liveMerges.has(ref)) ops.push({ kind: "merge", sheet: live.name, ref, action: "unmerge" });
  }
  for (const ref of live.merges) {
    if (!baseMerges.has(ref)) ops.push({ kind: "merge", sheet: live.name, ref, action: "merge" });
  }

  pushGeometry(ops, "row", live.name, prior.rows, live.rows);
  pushGeometry(ops, "col", live.name, prior.cols, live.cols);
}

function pushGeometry(ops: HandoffOp[], axis: "row" | "col", sheet: string, base: Record<string, string>, live: Record<string, string>): void {
  const keys = new Set([...Object.keys(base), ...Object.keys(live)]);
  for (const key of [...keys].sort()) {
    const [indexText, field] = key.split(":");
    if (field !== "h" && field !== "w" && field !== "hd") continue;
    const baseValue = base[key] ?? "";
    const nextValue = live[key] ?? "";
    if (baseValue === nextValue) continue;
    ops.push({
      kind: "geometry",
      axis,
      sheet,
      index: Number(indexText),
      field,
      base: baseValue,
      next: nextValue,
    });
  }
}

function compareA1(a: string, b: string): number {
  const left = a.match(/^([A-Z]+)(\d+)$/);
  const right = b.match(/^([A-Z]+)(\d+)$/);
  if (!left || !right) return a.localeCompare(b);
  const row = Number(left[2]) - Number(right[2]);
  if (row !== 0) return row;
  return left[1].localeCompare(right[1]);
}

export function csvField(value: string): string {
  if (/[",\r\n]/.test(value)) return `"${value.replace(/"/g, '""')}"`;
  return value;
}

export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  const source = text.replace(/^\uFEFF/, "");
  for (let i = 0; i < source.length; i += 1) {
    const char = source[i];
    if (quoted) {
      if (char === '"') {
        if (source[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        field += char;
      }
      continue;
    }
    if (char === '"') {
      quoted = true;
      continue;
    }
    if (char === ",") {
      row.push(field);
      field = "";
      continue;
    }
    if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }
    if (char !== "\r") field += char;
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((entry) => !(entry.length === 1 && entry[0] === ""));
}

function csvTable(header: string[], rows: string[][]): string {
  return [header, ...rows].map((row) => row.map((value) => csvField(value)).join(",")).join("\n") + "\n";
}

export async function sha256Text(text: string): Promise<string> {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

export interface HandoffBundleInput {
  jobNumber: string;
  sourceWorkbookHash: string;
  createdAt: string;
  exporter: string;
  ops: HandoffOp[];
  /** `visible-data` is a visible-tab scalar bundle. Omitted on a full bundle. */
  mode?: string;
}

/** Payload files only. READY is a separate last step from commitPlan. */
export async function buildBundleFiles(input: HandoffBundleInput): Promise<Record<string, string>> {
  const cells = input.ops.filter((op) => op.kind === "cell");
  const styles = input.ops.filter((op) => op.kind === "style");
  const merges = input.ops.filter((op) => op.kind === "merge");
  const geometry = input.ops.filter((op) => op.kind === "geometry");
  const sheets = input.ops.filter((op) => op.kind === "sheet");
  const names = input.ops.filter((op) => op.kind === "name");
  const files: Record<string, string> = {
    "cells.csv": csvTable(
      ["sheet", "a1", "baseKind", "baseValue", "nextKind", "nextValue"],
      cells.map((op) => [op.sheet, op.a1, op.baseKind, op.baseValue, op.nextKind, op.nextValue]),
    ),
    "styles.csv": csvTable(
      ["sheet", "a1", "field", "base", "next"],
      styles.map((op) => [op.sheet, op.a1, op.field, op.base, op.next]),
    ),
    "merges.csv": csvTable(
      ["sheet", "ref", "action"],
      merges.map((op) => [op.sheet, op.ref, op.action]),
    ),
    "geometry.csv": csvTable(
      ["axis", "sheet", "index", "field", "base", "next"],
      geometry.map((op) => [op.axis, op.sheet, String(op.index), op.field, op.base, op.next]),
    ),
    "sheets.csv": csvTable(
      ["name", "field", "base", "next"],
      sheets.map((op) => [op.name, op.field, op.base, op.next]),
    ),
    "names.csv": csvTable(
      ["name", "base", "next"],
      names.map((op) => [op.name, op.base, op.next]),
    ),
  };
  const hashes = await Promise.all(Object.entries(files).map(async ([name, text]) => [name, await sha256Text(text), String(parseCsv(text).length - 1)] as const));
  const manifestRows = [
    ["schema", String(HANDOFF_SCHEMA)],
    ["jobNumber", input.jobNumber],
    ["sourceWorkbookHash", input.sourceWorkbookHash],
    ["createdAt", input.createdAt],
    ["exporter", input.exporter],
  ];
  if (input.mode) manifestRows.push(["mode", input.mode]);
  for (const [name, hash, rows] of hashes) {
    manifestRows.push([name, hash]);
    manifestRows.push([`${name}Rows`, rows]);
  }
  files["manifest.csv"] = csvTable(["key", "value"], manifestRows);
  return files;
}

export interface HandoffCommit {
  payloads: Array<{ name: string; text: string }>;
  ready: { name: "READY"; text: string };
}

/** Payloads first, in a fixed order, then READY. Callers write them in this order. */
export function commitPlan(files: Record<string, string>): HandoffCommit {
  const names = Object.keys(files).filter((name) => name !== "READY").sort();
  return {
    payloads: names.map((name) => ({ name, text: files[name] ?? "" })),
    ready: { name: "READY", text: "1\n" },
  };
}

export async function auditBundle(files: Record<string, string>): Promise<string[]> {
  const problems: string[] = [];
  if ("READY" in files) problems.push("READY is present before the payload audit");
  const manifest = files["manifest.csv"];
  if (!manifest) return ["manifest.csv is missing"];
  const rows = parseCsv(manifest).slice(1);
  const values = new Map(rows.map((row) => [row[0] ?? "", row[1] ?? ""]));
  if (values.get("schema") !== String(HANDOFF_SCHEMA)) problems.push("schema is not supported");
  if (!/^[A-Za-z0-9_-]+$/.test(values.get("jobNumber") ?? "")) problems.push("jobNumber is missing");
  for (const name of Object.keys(files)) {
    if (name === "manifest.csv") continue;
    const expected = values.get(name);
    const actual = await sha256Text(files[name] ?? "");
    if (expected !== actual) problems.push(`${name} hash does not match the manifest`);
  }
  return problems;
}

export function parseManifest(text: string): Record<string, string> {
  const rows = parseCsv(text).slice(1);
  return Object.fromEntries(rows.map((row) => [row[0] ?? "", row[1] ?? ""]));
}

export interface HandoffReceipt {
  status: string;
  handoffId: string;
  message: string;
  operations: string;
}

export function parseReceipt(text: string): HandoffReceipt {
  const values = parseManifest(text);
  return {
    status: values.status ?? "",
    handoffId: values.handoffId ?? "",
    message: values.message ?? "",
    operations: values.operations ?? "",
  };
}

function emptySheet(name: string): CanonSheet {
  return { name, hidden: 0, rowCount: 0, columnCount: 0, freeze: "0,0", merges: [], rows: {}, cols: {}, cells: {} };
}

function blankCell(): CanonCell {
  return { k: "blank", v: "" };
}

function canonSheet(sheet: Record<string, unknown>, styles: Record<string, Record<string, unknown>>): CanonSheet {
  const hidden = sheet.hidden === 1 || sheet.hidden === 2 ? sheet.hidden : 0;
  const freeze = sheet.freeze && typeof sheet.freeze === "object"
    ? `${Number((sheet.freeze as { xSplit?: unknown }).xSplit ?? 0)},${Number((sheet.freeze as { ySplit?: unknown }).ySplit ?? 0)}`
    : "0,0";
  const cells: Record<string, CanonCell> = {};
  const cellData = sheet.cellData && typeof sheet.cellData === "object" ? sheet.cellData as Record<string, unknown> : {};
  for (const [rowText, cols] of Object.entries(cellData)) {
    const row = Number(rowText);
    if (!Number.isFinite(row) || !cols || typeof cols !== "object") continue;
    for (const [colText, raw] of Object.entries(cols as Record<string, unknown>)) {
      const col = Number(colText);
      if (!Number.isFinite(col) || !raw || typeof raw !== "object") continue;
      const cell = canonCell(raw as Record<string, unknown>, styles);
      if (cell.k === "blank" && !cell.ft && !cell.s) continue;
      cells[`${indexToCol(col)}${row + 1}`] = cell;
    }
  }
  return {
    name: String(sheet.name ?? ""),
    hidden,
    rowCount: Number(sheet.rowCount ?? 0),
    columnCount: Number(sheet.columnCount ?? 0),
    freeze,
    merges: mergeRefs(sheet.mergeData),
    rows: axisMap(sheet.rowData, "h"),
    cols: axisMap(sheet.columnData, "w"),
    cells,
  };
}

function canonCell(cell: Record<string, unknown>, styles: Record<string, Record<string, unknown>>): CanonCell {
  const formula = typeof cell.f === "string" ? cell.f.trim() : "";
  const ft = cell.ft === 2 || cell.ft === 3 ? cell.ft : undefined;
  const style = styleParts(resolveStyle(cell, styles));
  const styled = Object.keys(style).length > 0 ? style : undefined;
  if (formula) {
    return { k: "formula", v: formula.startsWith("=") ? formula : `=${formula}`, ...(ft ? { ft } : {}), ...(styled ? { s: styled } : {}) };
  }
  const value = cell.v;
  if (value == null || value === "") return { k: "blank", v: "", ...(ft ? { ft } : {}), ...(styled ? { s: styled } : {}) };
  if (typeof value === "number" && Number.isFinite(value)) return { k: "number", v: String(value), ...(styled ? { s: styled } : {}) };
  if (typeof value === "boolean") return { k: "bool", v: value ? "TRUE" : "FALSE", ...(styled ? { s: styled } : {}) };
  if (typeof value === "string") return { k: "text", v: value, ...(styled ? { s: styled } : {}) };
  return { k: "text", v: stable(value), s: { ...(styled ?? {}), __extra: stable(value) } };
}

function resolveStyle(cell: Record<string, unknown>, styles: Record<string, Record<string, unknown>>): Record<string, unknown> | undefined {
  const ref = cell.s;
  if (typeof ref === "string") return styles[ref];
  if (ref && typeof ref === "object") return ref as Record<string, unknown>;
  return undefined;
}

function styleParts(style: Record<string, unknown> | undefined): Record<string, string> {
  if (!style) return {};
  const out: Record<string, string> = {};
  const bg = rgbOf(style.bg);
  const cl = rgbOf(style.cl);
  if (bg) out.bg = bg;
  if (cl) out.cl = cl;
  if (flag(style.bl)) out.bl = "1";
  if (flag(style.it)) out.it = "1";
  if (style.ul) out.ul = "1";
  if (style.st) out.st = "1";
  if (style.ff) out.ff = String(style.ff);
  if (style.fs) out.fs = String(style.fs);
  if (style.ht != null && Number(style.ht) !== 0) out.ht = String(style.ht);
  if (style.vt != null && Number(style.vt) !== 3) out.vt = String(style.vt);
  const pattern = style.n && typeof style.n === "object" ? (style.n as { pattern?: unknown }).pattern : undefined;
  if (typeof pattern === "string" && pattern) out.n = pattern;
  if (style.bd) out.bd = borderToken(style.bd);
  if (style.tb != null && Number(style.tb) !== 1) out.tb = String(style.tb);
  const extra: Record<string, unknown> = {};
  for (const key of Object.keys(style)) {
    if (!(STYLE_FIELDS as readonly string[]).includes(key)) extra[key] = style[key];
  }
  if (Object.keys(extra).length > 0) out.__extra = stable(extra);
  return out;
}

function flag(value: unknown): boolean {
  return value === 1 || value === true || value === "1";
}

function rgbOf(value: unknown): string {
  if (!value || typeof value !== "object") return "";
  const rgb = String((value as { rgb?: unknown }).rgb ?? "").replace("#", "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(rgb) ? rgb : "";
}

function borderToken(value: unknown): string {
  if (!value || typeof value !== "object") return "unsupported";
  const record = value as Record<string, unknown>;
  const parts: string[] = [];
  for (const side of BORDER_SIDES) {
    const edge = record[side];
    if (!edge || typeof edge !== "object") continue;
    const style = Number((edge as { s?: unknown }).s ?? 0);
    const color = rgbOf((edge as { cl?: unknown }).cl);
    if (!Number.isFinite(style) || style <= 0 || !color) return "unsupported";
    if (![1, 2, 8, 13].includes(style)) return "unsupported";
    parts.push(`${side}:${style}:${color}`);
  }
  return parts.length > 0 ? parts.join(";") : "unsupported";
}

function mergeRefs(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  const refs: string[] = [];
  for (const entry of value) {
    if (!entry || typeof entry !== "object") continue;
    const range = entry as { startRow?: unknown; endRow?: unknown; startColumn?: unknown; endColumn?: unknown };
    const r1 = Number(range.startRow);
    const r2 = Number(range.endRow);
    const c1 = Number(range.startColumn);
    const c2 = Number(range.endColumn);
    if (![r1, r2, c1, c2].every(Number.isFinite)) continue;
    refs.push(`${indexToCol(c1)}${r1 + 1}:${indexToCol(c2)}${r2 + 1}`);
  }
  return refs.sort();
}

function axisMap(value: unknown, sizeField: "h" | "w"): Record<string, string> {
  if (!value || typeof value !== "object") return {};
  const out: Record<string, string> = {};
  for (const [indexText, raw] of Object.entries(value as Record<string, unknown>)) {
    const index = Number(indexText);
    if (!Number.isFinite(index) || !raw || typeof raw !== "object") continue;
    const record = raw as { h?: unknown; w?: unknown; hd?: unknown };
    const size = record[sizeField];
    if (size != null && size !== "") out[`${index + 1}:${sizeField}`] = String(size);
    if (record.hd === 1) out[`${index + 1}:hd`] = "1";
  }
  return out;
}

function resourceMap(value: unknown): Record<string, string> {
  if (!Array.isArray(value)) return {};
  const out: Record<string, string> = {};
  for (const entry of value) {
    if (!entry || typeof entry !== "object") continue;
    const record = entry as { name?: unknown; data?: unknown };
    const name = String(record.name ?? "");
    if (!name || IGNORED_RESOURCE.test(name)) continue;
    const data = typeof record.data === "string" ? record.data : stable(record.data ?? "");
    out[name] = stableData(data);
  }
  return out;
}

function stableData(text: string): string {
  try {
    return stable(JSON.parse(text));
  } catch {
    return text;
  }
}

function stable(value: unknown): string {
  return JSON.stringify(sortValue(value));
}

function sortValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortValue);
  if (value && typeof value === "object") {
    const out: Record<string, unknown> = {};
    for (const key of Object.keys(value as Record<string, unknown>).sort()) {
      out[key] = sortValue((value as Record<string, unknown>)[key]);
    }
    return out;
  }
  return value;
}
