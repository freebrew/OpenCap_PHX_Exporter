import { strFromU8, unzipSync } from "fflate";
import { XMLParser } from "fast-xml-parser";
import ExcelJS from "exceljs";
import { colToIndex, extractFunctionCalls, extractRefs, hasExternalRef, parseA1, parseRangeRef, shiftFormula, toA1 } from "./a1.js";
import { observedDependencyDepth } from "./dependency.js";
import { unknownFunctions } from "./excelFunctions.js";
import { macroButtonsFromPackage } from "./macroButtons.js";
import { coalesceCells, type GridPoint } from "./ranges.js";
import {
  DEFAULT_LIMITS,
  ImportRejectedError,
  type CellFormulaRecord,
  type ConditionalFormatRecord,
  type DefinedNameRecord,
  type HyperlinkRecord,
  type ImportLimits,
  type ImportLog,
  type MappedWorkbook,
  type ProtectionRecord,
  type ReadOptions,
  type UniverCell,
  type UniverRange,
  type UniverSheet,
  type ValidationRecord,
} from "./types.js";

type XmlNode = Record<string, unknown>;

const ARRAY_TAGS = new Set([
  "sheet",
  "definedName",
  "sheetView",
  "col",
  "row",
  "c",
  "f",
  "mergeCell",
  "conditionalFormatting",
  "cfRule",
  "dataValidation",
  "hyperlink",
  "Relationship",
  "xf",
  "font",
  "fill",
  "border",
  "numFmt",
  "si",
  "r",
  "formula",
  "colorScale",
  "dataBar",
  "cfvo",
]);

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  isArray: (name) => ARRAY_TAGS.has(name),
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
});

const BUILTIN_NUMFMT: Record<string, string> = {
  "1": "0",
  "2": "0.00",
  "3": "#,##0",
  "4": "#,##0.00",
  "9": "0%",
  "10": "0.00%",
  "11": "0.00E+00",
  "14": "m/d/yyyy",
  "15": "d-mmm-yy",
  "16": "d-mmm",
  "17": "mmm-yy",
  "18": "h:mm AM/PM",
  "19": "h:mm:ss AM/PM",
  "20": "h:mm",
  "21": "h:mm:ss",
  "22": "m/d/yyyy h:mm",
  "37": "#,##0 ;(#,##0)",
  "38": "#,##0 ;[Red](#,##0)",
  "39": "#,##0.00;(#,##0.00)",
  "40": "#,##0.00;[Red](#,##0.00)",
  "45": "mm:ss",
  "46": "[h]:mm:ss",
  "47": "mmss.0",
  "48": "##0.0E+0",
  "49": "@",
};

const INDEXED = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

function asArray(value: unknown): XmlNode[] {
  if (value == null) return [];
  if (Array.isArray(value)) return value.filter((item) => item && typeof item === "object") as XmlNode[];
  if (typeof value === "object") return [value as XmlNode];
  return [];
}

function attr(node: XmlNode | undefined, name: string): string | undefined {
  if (!node) return undefined;
  const value = node[`@_${name}`];
  return value == null || value === "" ? undefined : String(value);
}

function textOf(value: unknown): string {
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (!value || typeof value !== "object") return "";
  const node = value as XmlNode;
  if (typeof node["#text"] === "string") return node["#text"];
  if (Array.isArray(value)) return value.map((item) => textOf(item)).join("");
  return "";
}

function child(node: XmlNode | undefined, name: string): unknown {
  return node?.[name];
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
  );
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function unzip(bytes: Uint8Array, limits: ImportLimits): Map<string, Uint8Array> {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch (error) {
    throw new ImportRejectedError("zip", `Workbook is not a readable xlsx package: ${error instanceof Error ? error.message : String(error)}`);
  }
  const entries = Object.keys(files);
  if (entries.length > limits.maxEntries) {
    throw new ImportRejectedError("entries", `Workbook has ${entries.length} package entries; limit is ${limits.maxEntries}.`);
  }
  let total = 0;
  const out = new Map<string, Uint8Array>();
  for (const name of entries) {
    const data = files[name];
    total += data.length;
    if (total > limits.maxUncompressedBytes) {
      throw new ImportRejectedError("expansion", `Uncompressed workbook exceeds ${limits.maxUncompressedBytes} bytes.`);
    }
    out.set(name.replace(/\\/g, "/"), data);
  }
  return out;
}

function xml(files: Map<string, Uint8Array>, path: string): XmlNode {
  const data = files.get(path);
  if (!data) throw new ImportRejectedError("ooxml", `Missing package part ${path}.`);
  const parsed = parser.parse(strFromU8(data)) as XmlNode;
  const entry = Object.entries(parsed).find(([key, value]) => !key.startsWith("?") && value && typeof value === "object");
  return (entry?.[1] ?? {}) as XmlNode;
}

function relTarget(target: string, base: string): string {
  if (target.startsWith("/")) return target.replace(/^\//, "");
  const stack = base.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (part === "..") stack.pop();
    else if (part !== ".") stack.push(part);
  }
  return stack.join("/");
}

interface BorderSide {
  s: number;
  cl: { rgb: string };
}

type BorderInfo = Partial<Record<"t" | "b" | "l" | "r", BorderSide>>;

interface StyleInfo {
  numFmt?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontSize?: number;
  fontName?: string;
  color?: string;
  fill?: string;
  border?: BorderInfo;
  locked: boolean;
  hidden: boolean;
  align?: string;
  valign?: string;
  wrap?: boolean;
}

interface ThemeColors {
  colors: Array<string | undefined>;
}

/** Bumped whenever cell appearance mapping changes, so older snapshots can be restyled. */
export const STYLE_VERSION = 4;

/** Excel numbers theme colours lt1, dk1, lt2, dk2 before the accents; the theme part lists dk1 first. */
const THEME_SLOT = [1, 0, 3, 2];

const BORDER_STYLES: Record<string, number> = {
  thin: 1,
  hair: 2,
  dotted: 3,
  dashed: 4,
  dashDot: 5,
  dashDotDot: 6,
  double: 7,
  medium: 8,
  mediumDashed: 9,
  mediumDashDot: 10,
  mediumDashDotDot: 11,
  slantDashDot: 12,
  thick: 13,
};

function themeColors(files: Map<string, Uint8Array>): ThemeColors {
  const path = [...files.keys()].find((name) => /^xl\/theme\/theme\d*\.xml$/.test(name));
  if (!path) return { colors: [] };
  const root = xml(files, path);
  const scheme = findFirst(root, "clrScheme");
  if (!scheme) return { colors: [] };
  const order = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  return { colors: order.map((name) => colorOf(scheme[name])) };
}

function findFirst(node: unknown, name: string): XmlNode | undefined {
  if (!node || typeof node !== "object") return undefined;
  const record = node as XmlNode;
  if (name in record) {
    const value = record[name];
    if (Array.isArray(value)) return value[0] as XmlNode;
    if (value && typeof value === "object") return value as XmlNode;
  }
  for (const value of Object.values(record)) {
    if (!value || typeof value !== "object") continue;
    const found = findFirst(value, name);
    if (found) return found;
  }
  return undefined;
}

function colorOf(node: unknown): string | undefined {
  if (!node || typeof node !== "object") return undefined;
  const record = node as XmlNode;
  const rgb = attr(record, "rgb") ?? attr(asArray(record.srgbClr)[0], "val");
  if (rgb) return normalizeRgb(rgb);
  const sys = asArray(record.sysClr)[0];
  const last = attr(sys, "lastClr");
  if (last) return normalizeRgb(last);
  const indexed = attr(record, "indexed");
  if (indexed && INDEXED[Number(indexed)]) return `#${INDEXED[Number(indexed)]}`;
  return undefined;
}

function normalizeRgb(value: string): string {
  const hex = value.replace("#", "");
  const rgb = hex.length === 8 ? hex.slice(2) : hex;
  return `#${rgb.toUpperCase()}`;
}

function themeRgb(theme: ThemeColors, node: unknown): string | undefined {
  if (!node || typeof node !== "object") return undefined;
  const record = node as XmlNode;
  let base = colorOf(record);
  const index = attr(record, "theme");
  if (!base && index != null) {
    const slot = Number(index);
    base = theme.colors[slot < THEME_SLOT.length ? THEME_SLOT[slot] : slot];
  }
  if (!base) return undefined;
  const tint = Number(attr(record, "tint") ?? "0");
  return tint ? applyTint(base, tint) : base;
}

/** Excel tint: lighten or darken the HLS luminance of the base colour. */
function applyTint(hex: string, tint: number): string {
  const value = hex.replace("#", "");
  const r = parseInt(value.slice(0, 2), 16) / 255;
  const g = parseInt(value.slice(2, 4), 16) / 255;
  const b = parseInt(value.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  const d = max - min;
  if (d > 0) {
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p: number, q: number, t: number) => {
    const u = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (u < 1 / 6) return p + (q - p) * 6 * u;
    if (u < 1 / 2) return q;
    if (u < 2 / 3) return p + (q - p) * (2 / 3 - u) * 6;
    return p;
  };
  let out: number[];
  if (s === 0) out = [l, l, l];
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    out = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  }
  return `#${out.map((channel) => Math.round(channel * 255).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

function borderOf(theme: ThemeColors, node: XmlNode): BorderInfo | undefined {
  const sides: Array<[keyof BorderInfo, string]> = [["l", "left"], ["r", "right"], ["t", "top"], ["b", "bottom"]];
  const border: BorderInfo = {};
  for (const [key, name] of sides) {
    const side = asArray(node[name])[0];
    const style = BORDER_STYLES[attr(side, "style") ?? ""];
    if (!side || !style) continue;
    border[key] = { s: style, cl: { rgb: themeRgb(theme, asArray(side.color)[0]) ?? "#000000" } };
  }
  return Object.keys(border).length > 0 ? border : undefined;
}

function parseStyles(files: Map<string, Uint8Array>, theme: ThemeColors): StyleInfo[] {
  if (!files.has("xl/styles.xml")) return [defaultStyle()];
  const root = xml(files, "xl/styles.xml");
  const numFmts = new Map<string, string>(Object.entries(BUILTIN_NUMFMT));
  for (const fmt of asArray(child(asArray(root.numFmts)[0], "numFmt"))) {
    const id = attr(fmt, "numFmtId");
    const code = attr(fmt, "formatCode");
    if (id && code) numFmts.set(id, code);
  }
  const fonts = asArray(child(asArray(root.fonts)[0], "font")).map((font) => ({
    bold: font.b != null && attr(asArray(font.b)[0], "val") !== "0",
    italic: font.i != null && attr(asArray(font.i)[0], "val") !== "0",
    underline: font.u != null && attr(asArray(font.u)[0], "val") !== "none",
    strike: font.strike != null && attr(asArray(font.strike)[0], "val") !== "0",
    size: Number(attr(asArray(font.sz)[0], "val") ?? "11"),
    name: attr(asArray(font.name)[0], "val"),
    color: themeRgb(theme, asArray(font.color)[0]),
  }));
  const fills = asArray(child(asArray(root.fills)[0], "fill")).map((fill) => {
    const pattern = asArray(fill.patternFill)[0];
    const type = attr(pattern, "patternType");
    if (!pattern || !type || type === "none") return undefined;
    return themeRgb(theme, asArray(pattern.fgColor)[0]) ?? (type === "solid" ? undefined : themeRgb(theme, asArray(pattern.bgColor)[0]));
  });
  const borders = asArray(child(asArray(root.borders)[0], "border")).map((border) => borderOf(theme, border));
  const xfs = asArray(child(asArray(root.cellXfs)[0], "xf"));
  if (xfs.length === 0) return [defaultStyle()];
  return xfs.map((xf) => {
    const font = fonts[Number(attr(xf, "fontId") ?? "0")] ?? fonts[0];
    const fill = fills[Number(attr(xf, "fillId") ?? "0")];
    const border = borders[Number(attr(xf, "borderId") ?? "0")];
    const protection = asArray(xf.protection)[0];
    const alignment = asArray(xf.alignment)[0];
    const numFmtId = attr(xf, "numFmtId") ?? "0";
    return {
      numFmt: numFmtId === "0" ? undefined : numFmts.get(numFmtId),
      bold: font?.bold,
      italic: font?.italic,
      underline: font?.underline,
      strike: font?.strike,
      fontSize: font?.size,
      fontName: font?.name,
      color: font?.color,
      fill,
      border,
      locked: attr(protection, "locked") !== "0",
      hidden: attr(protection, "hidden") === "1",
      align: attr(alignment, "horizontal"),
      valign: attr(alignment, "vertical"),
      wrap: attr(alignment, "wrapText") === "1",
    } satisfies StyleInfo;
  });
}

function defaultStyle(): StyleInfo {
  return { locked: true, hidden: false };
}

function sharedStrings(files: Map<string, Uint8Array>): string[] {
  if (!files.has("xl/sharedStrings.xml")) return [];
  const root = xml(files, "xl/sharedStrings.xml");
  return asArray(root.si).map((si) => {
    if (si.t != null && !Array.isArray(si.r)) {
      const direct = textOf(si.t);
      if (direct) return direct;
    }
    const runs = asArray(si.r);
    if (runs.length > 0) return runs.map((run) => textOf(run.t)).join("");
    return textOf(si.t);
  });
}

function parseRelationships(files: Map<string, Uint8Array>, path: string): Map<string, string> {
  const map = new Map<string, string>();
  if (!files.has(path)) return map;
  const root = xml(files, path);
  for (const rel of asArray(root.Relationship)) {
    const id = attr(rel, "Id");
    const target = attr(rel, "Target");
    if (id && target) map.set(id, target);
  }
  return map;
}

interface SheetSource {
  name: string;
  sheetId: string;
  state: "visible" | "hidden" | "veryHidden";
  path: string;
}

function workbookSheets(files: Map<string, Uint8Array>): { sheets: SheetSource[]; definedNames: XmlNode[] } {
  const root = xml(files, "xl/workbook.xml");
  const rels = parseRelationships(files, "xl/_rels/workbook.xml.rels");
  const sheets = asArray(child(asArray(root.sheets)[0], "sheet")).map((sheet) => {
    const rid = attr(sheet, "id");
    const target = rid ? rels.get(rid) : undefined;
    if (!target) throw new ImportRejectedError("ooxml", `Sheet ${attr(sheet, "name")} has no worksheet part.`);
    const stateAttr = attr(sheet, "state");
    const state = stateAttr === "hidden" || stateAttr === "veryHidden" ? stateAttr : "visible";
    return {
      name: attr(sheet, "name") ?? "Sheet",
      sheetId: attr(sheet, "sheetId") ?? "0",
      state,
      path: relTarget(target, "xl/workbook.xml"),
    } satisfies SheetSource;
  });
  return { sheets, definedNames: asArray(child(asArray(root.definedNames)[0], "definedName")) };
}

/**
 * Excel stores newer functions as `_xlfn.NAME(` and the `@` operator as `_xlfn.SINGLE(`.
 * The engine knows the bare names; `@` on a single cell or scalar is the value itself.
 */
export function engineFormula(text: string): string {
  if (!/_xl(?:fn|ws)\./i.test(text)) return text;
  return text
    .split(/("(?:[^"]|"")*")/)
    .map((part, index) => (index % 2 === 1 ? part : part
      .replace(/_xlfn\.SINGLE\(/gi, "(")
      .replace(/_xlfn\.|_xlws\./gi, "")))
    .join("");
}

/** Snapshots imported before engineFormula kept `_xlfn.` names and untyped text. Safe to repeat. */
export function normalizeSnapshotFormulas<T>(snapshot: T): T {
  const sheets = (snapshot as { sheets?: Record<string, { cellData?: Record<string, Record<string, UniverCell | null>> }> } | null)?.sheets;
  for (const sheet of Object.values(sheets ?? {})) {
    for (const row of Object.values(sheet?.cellData ?? {})) {
      for (const cell of Object.values(row ?? {})) {
        if (!cell) continue;
        if (typeof cell.f === "string") cell.f = engineFormula(cell.f);
        else if (typeof cell.v === "string" && cell.t == null) cell.t = 1;
      }
    }
  }
  return snapshot;
}

function supportName(text: string): { supported: boolean; reason?: string; formula: string } {
  const trimmed = engineFormula(text.trim());
  if (!trimmed) return { supported: false, reason: "empty", formula: trimmed };
  if (hasExternalRef(trimmed)) return { supported: false, reason: "external", formula: trimmed };
  if (trimmed.startsWith("{")) return { supported: false, reason: "array-constant", formula: trimmed };
  if (trimmed.includes("#REF!")) return { supported: false, reason: "ref-error", formula: trimmed };
  if (trimmed.startsWith("=") || parseRangeRef(trimmed)) return { supported: true, formula: trimmed.startsWith("=") ? trimmed : trimmed };
  return { supported: true, formula: `=${trimmed}` };
}

/** Excel stores widths in Calibri 11 character units (7 px per digit, padding included). */
function columnPixels(width: string): number {
  return Math.max(0, Math.round(Number(width) * 7));
}

function ht(align: string | undefined): number | undefined {
  if (align === "left") return 1;
  if (align === "center") return 2;
  if (align === "right") return 3;
  return undefined;
}

function vt(align: string | undefined): number | undefined {
  if (align === "top") return 1;
  if (align === "center") return 2;
  if (align === "bottom") return 3;
  return undefined;
}

function styleData(style: StyleInfo): Record<string, unknown> | undefined {
  const data: Record<string, unknown> = {};
  if (style.fontName) data.ff = style.fontName;
  if (style.fontSize) data.fs = style.fontSize;
  if (style.bold) data.bl = 1;
  if (style.italic) data.it = 1;
  if (style.underline) data.ul = { s: 1 };
  if (style.strike) data.st = { s: 1 };
  if (style.color) data.cl = { rgb: style.color };
  if (style.fill) data.bg = { rgb: style.fill };
  if (style.border) data.bd = style.border;
  const horizontal = ht(style.align);
  // Excel bottom-aligns by default; Univer's dropdown renderer otherwise draws from the top.
  const vertical = vt(style.valign) ?? 3;
  if (horizontal) data.ht = horizontal;
  if (vertical) data.vt = vertical;
  // Univer: 1 overflow, 3 wrap. Excel only breaks lines when Wrap Text is on.
  data.tb = style.wrap ? 3 : 1;
  if (style.numFmt) data.n = { pattern: style.numFmt };
  return Object.keys(data).length > 0 ? data : undefined;
}

function cellCached(type: string | undefined, raw: string | undefined, strings: string[]): { value?: string | number | boolean; type?: string } {
  if (raw == null || raw === "") return {};
  if (type === "s") return { value: strings[Number(raw)] ?? "", type: "s" };
  if (type === "b") return { value: raw === "1", type: "b" };
  if (type === "e" || type === "str") return { value: raw, type };
  if (type === "inlineStr") return { value: raw, type: "str" };
  const numeric = Number(raw);
  if (Number.isFinite(numeric)) return { value: numeric, type: type ?? "n" };
  return { value: raw, type: type ?? "str" };
}

/**
 * Text stays text: without `t` the engine reads number-like text as a number. Numbers stay
 * untyped, because a typed number in an `@` cell is read back as text, which Excel never does.
 */
function setValue(cell: UniverCell, value: string | number | boolean): void {
  cell.v = value;
  if (typeof value === "string") cell.t = 1;
  else if (typeof value === "boolean") cell.t = 3;
}

function inlineText(cell: XmlNode): string {
  const inline = asArray(cell.is)[0];
  if (!inline) return "";
  const runs = asArray(inline.r);
  if (runs.length > 0) return runs.map((run) => textOf(run.t)).join("");
  return textOf(inline.t);
}

interface MasterFormula {
  formula: string;
  row: number;
  col: number;
}

async function excelFormulaMap(bytes: Uint8Array, warnings: string[]): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(Buffer.from(bytes) as unknown as Parameters<typeof workbook.xlsx.load>[0]);
    workbook.eachSheet((sheet) => {
      sheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          const value = cell.value as { formula?: unknown } | null;
          const formula = value && typeof value === "object" && typeof value.formula === "string" ? value.formula : undefined;
          if (!formula) return;
          map.set(`${sheet.name}!${cell.address}`, formula.replace(/^=/, ""));
        });
      });
    });
  } catch (error: unknown) {
    warnings.push(`ExcelJS could not read this package (${error instanceof Error ? error.message : String(error)}). OOXML mapping continued.`);
  }
  return map;
}

function countPart(files: Map<string, Uint8Array>, test: (path: string) => boolean): number {
  let count = 0;
  for (const path of files.keys()) if (test(path)) count += 1;
  return count;
}

/** Univer's data-validation plugin reads this resource: `{ [sheetId]: rule[] }`. */
export const DATA_VALIDATION_RESOURCE = "SHEET_DATA_VALIDATION_PLUGIN";

interface UniverValidationRule {
  uid: string;
  type: string;
  ranges: UniverRange[];
  formula1?: string;
  formula2?: string;
  operator?: string;
  allowBlank: boolean;
  showDropDown?: boolean;
  renderMode?: number;
  showErrorMessage: boolean;
  errorStyle: number;
  error?: string;
  errorTitle?: string;
  showInputMessage?: boolean;
  prompt?: string;
  promptTitle?: string;
}

const VALIDATION_OPERATORS = new Set(["between", "notBetween", "equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"]);

/** Plain rules hold the formula as text; Excel 2010 rules (x14) wrap it in `<xm:f>`. */
function validationFormula(node: unknown): string {
  const first = Array.isArray(node) ? node[0] : node;
  if (first && typeof first === "object" && "f" in (first as XmlNode)) return textOf((first as XmlNode).f);
  return textOf(first);
}

function univerValidation(
  rule: XmlNode,
  type: string,
  sqref: string,
  formula1: string,
  formula2: string,
  uid: string,
): { rule?: UniverValidationRule; reason?: string } {
  const ranges = sqrefRanges(sqref);
  if (ranges.length === 0) return { reason: "no usable range" };
  if (formula1.includes("#REF!")) return { reason: "list source is #REF! in Excel" };
  if (hasExternalRef(formula1)) return { reason: "external list source" };
  const errorStyle = attr(rule, "errorStyle");
  const base = {
    uid,
    ranges,
    allowBlank: attr(rule, "allowBlank") === "1",
    showErrorMessage: attr(rule, "showErrorMessage") === "1",
    // Univer: 0 info, 1 stop, 2 warning. Excel only blocks entry when the error message is on.
    errorStyle: attr(rule, "showErrorMessage") !== "1" ? 0 : errorStyle === "warning" ? 2 : errorStyle === "information" ? 0 : 1,
    ...(attr(rule, "error") ? { error: attr(rule, "error") } : {}),
    ...(attr(rule, "errorTitle") ? { errorTitle: attr(rule, "errorTitle") } : {}),
    ...(attr(rule, "showInputMessage") === "1" ? { showInputMessage: true } : {}),
    ...(attr(rule, "prompt") ? { prompt: attr(rule, "prompt") } : {}),
    ...(attr(rule, "promptTitle") ? { promptTitle: attr(rule, "promptTitle") } : {}),
  };
  if (type === "list") {
    if (!formula1) return { reason: "empty list" };
    const literal = formula1.startsWith("\"") && formula1.endsWith("\"");
    const source = literal
      ? JSON.stringify(formula1.slice(1, -1).split(",").map((item) => item.trim()).filter(Boolean))
      : `=${formula1.replace(/^=/, "")}`;
    // Excel's showDropDown="1" means hide the in-cell arrow.
    return { rule: { ...base, type: "list", formula1: source, showDropDown: attr(rule, "showDropDown") !== "1", renderMode: 1 } };
  }
  if (type === "whole" || type === "decimal" || type === "textLength") {
    const operator = attr(rule, "operator") ?? "between";
    if (!VALIDATION_OPERATORS.has(operator)) return { reason: `unsupported operator ${operator}` };
    return { rule: { ...base, type, operator, formula1, ...(formula2 ? { formula2 } : {}) } };
  }
  if (type === "custom") {
    if (!formula1) return { reason: "empty custom formula" };
    return { rule: { ...base, type: "custom", formula1: `=${formula1.replace(/^=/, "")}` } };
  }
  return { reason: `unsupported validation type ${type || "none"}` };
}

function sqrefRanges(sqref: string): UniverRange[] {
  const ranges: UniverRange[] = [];
  for (const part of sqref.split(/\s+/)) {
    if (!part) continue;
    const [start, end] = part.split(":");
    try {
      const a = parseA1(start);
      const b = end ? parseA1(end) : a;
      ranges.push({
        startRow: Math.min(a.row, b.row),
        endRow: Math.max(a.row, b.row),
        startColumn: Math.min(a.col, b.col),
        endColumn: Math.max(a.col, b.col),
      });
    } catch {
      // Logged by the caller when a rule cannot be mapped.
    }
  }
  return ranges;
}

export async function readWorkbook(bytes: Uint8Array, options: ReadOptions = {}): Promise<MappedWorkbook> {
  const limits: ImportLimits = { ...DEFAULT_LIMITS, ...options.limits };
  const files = unzip(bytes, limits);
  const warnings: string[] = [];
  const theme = themeColors(files);
  const styles = parseStyles(files, theme);
  const strings = sharedStrings(files);
  const { sheets, definedNames } = workbookSheets(files);
  const excelFormulas = await excelFormulaMap(bytes, warnings);
  const calls = new Set<string>();
  const formulas: CellFormulaRecord[] = [];
  const rejectedExternalRefs: ImportLog["rejectedExternalRefs"] = [];
  let sharedFormulaCount = 0;
  let arrayFormulaCount = 0;
  let disagreements = 0;
  const univerSheets: Record<string, UniverSheet> = {};
  const sheetOrder: string[] = [];
  const styleBucket = new Map<string, string>();
  const workbookStyles: Record<string, Record<string, unknown>> = {};
  const protection: ProtectionRecord[] = [];
  const dataValidations: ValidationRecord[] = [];
  const validationRules: Record<string, UniverValidationRule[]> = {};
  const conditionalFormatting: ConditionalFormatRecord[] = [];
  const hyperlinks: HyperlinkRecord[] = [];
  const veryHidden: string[] = [];
  const perSheet: ImportLog["perSheet"] = [];
  let styleSerial = 1;

  const styleIdFor = (index: number): string | undefined => {
    const info = styles[index] ?? styles[0];
    if (!info) return undefined;
    const data = styleData(info);
    if (!data) return undefined;
    const key = JSON.stringify(data);
    const existing = styleBucket.get(key);
    if (existing) return existing;
    const id = `s${styleSerial}`;
    styleSerial += 1;
    styleBucket.set(key, id);
    workbookStyles[id] = data;
    return id;
  };

  for (const source of sheets) {
    const root = xml(files, source.path);
    const rels = parseRelationships(files, source.path.replace(/([^/]+)$/, "_rels/$1.rels"));
    const format = asArray(root.sheetFormatPr)[0];
    const defaultColWidth = attr(format, "defaultColWidth");
    const defaultWidth = defaultColWidth ? columnPixels(defaultColWidth) : Math.round(Number(attr(format, "baseColWidth") ?? "8") * 7 + 8);
    const defaultHeight = Math.round(Number(attr(format, "defaultRowHeight") ?? "15") * (96 / 72));
    const view = asArray(child(asArray(root.sheetViews)[0], "sheetView"))[0];
    const showGridlines = attr(view, "showGridLines") === "0" ? 0 : 1;
    const zoomScale = Number(attr(view, "zoomScale") ?? "100");
    const zoomRatio = Number.isFinite(zoomScale) && zoomScale >= 10 && zoomScale <= 400 ? zoomScale / 100 : 1;
    const pane = asArray(child(view, "pane"))[0];
    const frozen = attr(pane, "state") === "frozen" || attr(pane, "state") === "frozenSplit";
    const xSplit = frozen ? Number(attr(pane, "xSplit") ?? "0") : 0;
    const ySplit = frozen ? Number(attr(pane, "ySplit") ?? "0") : 0;
    const tab = asArray(child(asArray(root.sheetPr)[0], "tabColor"))[0];
    const tabRgb = attr(tab, "rgb");
    const merges: UniverRange[] = asArray(child(asArray(root.mergeCells)[0], "mergeCell")).flatMap((merge) => sqrefRanges(attr(merge, "ref") ?? ""));
    const rowData: UniverSheet["rowData"] = {};
    const columnData: UniverSheet["columnData"] = {};
    const cellData: UniverSheet["cellData"] = {};
    const unlocked: GridPoint[] = [];
    const hiddenCells: GridPoint[] = [];
    const masters = new Map<string, MasterFormula>();
    const followers: Array<{ row: number; col: number; si: string; a1: string; cached?: ReturnType<typeof cellCached> }> = [];
    let maxRow = 0;
    let maxCol = 0;
    let sheetFormulas = 0;
    const note = (row: number, col: number) => {
      if (row > maxRow) maxRow = row;
      if (col > maxCol) maxCol = col;
    };

    for (const col of asArray(child(asArray(root.cols)[0], "col"))) {
      const min = Number(attr(col, "min") ?? "1");
      const max = Number(attr(col, "max") ?? String(min));
      const width = attr(col, "width");
      const hidden = attr(col, "hidden") === "1";
      for (let index = min; index <= max && index <= limits.maxColumns; index += 1) {
        columnData[index - 1] = {
          ...(width ? { w: columnPixels(width) } : {}),
          ...(hidden ? { hd: 1 as const } : {}),
        };
        note(0, index - 1);
      }
    }

    for (const row of asArray(child(asArray(root.sheetData)[0], "row"))) {
      const rowNumber = Number(attr(row, "r") ?? "0");
      if (!rowNumber) continue;
      const rowIndex = rowNumber - 1;
      if (rowIndex >= limits.maxRows) {
        const hasContent = asArray(row.c).some((cell) => textOf(cell.v) !== "" || textOf(cell.f) !== "" || inlineText(cell) !== "");
        if (!hasContent) {
          warnings.push(`${source.name} row ${rowNumber} is an empty anchor past the row limit and was skipped.`);
          continue;
        }
        throw new ImportRejectedError("rows", `${source.name} has content past row ${limits.maxRows}.`);
      }
      note(rowIndex, 0);
      const height = attr(row, "ht");
      if (height || attr(row, "hidden") === "1") {
        rowData[rowIndex] = {
          ...(height ? { h: Math.round(Number(height) * (96 / 72)) } : {}),
          ...(attr(row, "hidden") === "1" ? { hd: 1 as const } : {}),
        };
      }
      for (const cell of asArray(row.c)) {
        const address = attr(cell, "r");
        if (!address) continue;
        const pos = parseA1(address);
        if (pos.col >= limits.maxColumns) {
          throw new ImportRejectedError("columns", `${source.name} exceeds ${limits.maxColumns} columns.`);
        }
        note(pos.row, pos.col);
        const styleIndex = Number(attr(cell, "s") ?? "0");
        const style = styles[styleIndex] ?? styles[0] ?? defaultStyle();
        if (!style.locked) unlocked.push(pos);
        if (style.hidden) hiddenCells.push(pos);
        const formulaNode = asArray(cell.f)[0];
        const formulaType = attr(formulaNode, "t");
        const formulaText = engineFormula(textOf(formulaNode ?? cell.f).replace(/^=/, ""));
        const cached = cellCached(attr(cell, "t"), textOf(cell.v) || inlineText(cell), strings);
        let formula = "";
        let arrayRef: string | undefined;
        if (formulaType === "shared" && formulaText) {
          masters.set(attr(formulaNode, "si") ?? "0", { formula: formulaText, row: pos.row, col: pos.col });
          formula = formulaText;
          sharedFormulaCount += 1;
        } else if (formulaType === "shared") {
          followers.push({ row: pos.row, col: pos.col, si: attr(formulaNode, "si") ?? "0", a1: address, cached });
          sharedFormulaCount += 1;
        } else if (formulaType === "array") {
          formula = formulaText;
          arrayRef = attr(formulaNode, "ref") ?? address;
          arrayFormulaCount += 1;
        } else if (formulaText) {
          formula = formulaText;
        }

        const univer: UniverCell = {};
        const sid = styleIdFor(styleIndex);
        if (sid) univer.s = sid;
        if (cached.value !== undefined && cached.type !== "e") setValue(univer, cached.value);
        if (formula) {
          if (hasExternalRef(formula)) {
            if (options.rejectExternalReferences) {
              throw new ImportRejectedError("external", `External reference in ${source.name}!${address} was rejected.`);
            }
            rejectedExternalRefs.push({ sheet: source.name, cell: address });
            if (cached.value !== undefined && cached.type !== "e") univer.v = cached.value;
            formula = "";
          }
        }
        if (formula) {
          univer.f = formula.startsWith("=") ? formula : `=${formula}`;
          if (arrayRef) {
            univer.ft = 2;
            univer.ref = arrayRef;
            univer.custom = { phxArrayRef: arrayRef };
          }
          for (const ref of extractRefs(formula)) note(Math.max(ref.row1, ref.row2), Math.max(ref.col1, ref.col2));
          for (const name of extractFunctionCalls(formula)) calls.add(name);
          formulas.push({
            sheet: source.name,
            a1: address,
            row: pos.row,
            col: pos.col,
            formula,
            ...(arrayRef ? { arrayRef } : {}),
            ...(cached.value !== undefined ? { cachedValue: cached.value, cachedType: cached.type } : {}),
          });
          sheetFormulas += 1;
        }
        if (Object.keys(univer).length > 0) {
          cellData[pos.row] ??= {};
          cellData[pos.row][pos.col] = univer;
        }
      }
    }

    for (const follower of followers) {
      const master = masters.get(follower.si);
      const expanded = master ? shiftFormula(master.formula, follower.row - master.row, follower.col - master.col) : "";
      const excel = excelFormulas.get(`${source.name}!${follower.a1}`);
      const chosen = expanded || excel || "";
      if (!chosen) {
        warnings.push(`Shared formula ${source.name}!${follower.a1} has no master.`);
        continue;
      }
      if (hasExternalRef(chosen)) {
        rejectedExternalRefs.push({ sheet: source.name, cell: follower.a1 });
        continue;
      }
      for (const ref of extractRefs(chosen)) note(Math.max(ref.row1, ref.row2), Math.max(ref.col1, ref.col2));
      for (const name of extractFunctionCalls(chosen)) calls.add(name);
      cellData[follower.row] ??= {};
      const existing = cellData[follower.row][follower.col] ?? {};
      existing.f = chosen.startsWith("=") ? chosen : `=${chosen}`;
      if (follower.cached?.value !== undefined && follower.cached.type !== "e") setValue(existing, follower.cached.value);
      cellData[follower.row][follower.col] = existing;
      formulas.push({
        sheet: source.name,
        a1: follower.a1,
        row: follower.row,
        col: follower.col,
        formula: chosen,
        ...(follower.cached?.value !== undefined ? { cachedValue: follower.cached.value, cachedType: follower.cached.type } : {}),
      });
      sheetFormulas += 1;
    }

    const sheetRules: UniverValidationRule[] = [];
    const validationNodes = [
      ...asArray(child(asArray(root.dataValidations)[0], "dataValidation")).map((rule) => ({ rule, x14: false })),
      ...asArray(root.extLst).flatMap((ext) => asArray(ext.ext)).flatMap((ext) => asArray(child(asArray(ext.dataValidations)[0], "dataValidation"))).map((rule) => ({ rule, x14: true })),
    ];
    for (const { rule, x14 } of validationNodes) {
      const type = attr(rule, "type") ?? "";
      const sqref = x14 ? textOf(rule.sqref) : attr(rule, "sqref") ?? "";
      const formula1 = validationFormula(rule.formula1);
      const formula2 = validationFormula(rule.formula2);
      const univer = univerValidation(rule, type, sqref, formula1, formula2, `dv-${source.sheetId}-${sheetRules.length + 1}`);
      if (univer.rule) sheetRules.push(univer.rule);
      dataValidations.push({
        sheet: source.name,
        sqref,
        type: type || "unknown",
        operator: attr(rule, "operator"),
        formula1,
        formula2,
        allowBlank: attr(rule, "allowBlank") === "1",
        mapped: Boolean(univer.rule),
        ...(univer.reason ? { reason: univer.reason } : {}),
      });
    }
    if (sheetRules.length > 0) validationRules[`sheet-${source.sheetId}`] = sheetRules;

    for (const block of asArray(root.conditionalFormatting)) {
      const sqref = attr(block, "sqref") ?? "";
      for (const rule of asArray(block.cfRule)) {
        const type = attr(rule, "type") ?? "unknown";
        const mapped = type === "cellIs" || type === "expression" || type === "colorScale" || type === "dataBar";
        conditionalFormatting.push({
          sheet: source.name,
          sqref,
          type,
          operator: attr(rule, "operator"),
          formula: textOf(asArray(rule.formula)[0] ?? rule.formula),
          dxfId: attr(rule, "dxfId"),
          mapped,
          ...(mapped ? {} : { reason: "unsupported conditional format" }),
        });
      }
    }

    for (const link of asArray(child(asArray(root.hyperlinks)[0], "hyperlink"))) {
      const ref = attr(link, "ref") ?? "";
      const location = attr(link, "location");
      const target = location ?? (attr(link, "id") ? rels.get(attr(link, "id") ?? "") ?? "" : "");
      const lower = target.toLowerCase();
      const rejected = lower.startsWith("file:") || lower.startsWith("javascript:") || lower.includes("\\\\");
      hyperlinks.push({
        sheet: source.name,
        ref,
        target: rejected ? "" : target,
        confirmed: false,
        rejected,
        ...(rejected ? { reason: "non-https link rejected" } : {}),
      });
    }

    const protectionNode = asArray(root.sheetProtection)[0];
    if (protectionNode) {
      const options: Record<string, string> = {};
      let passwordPresent = false;
      for (const [key, value] of Object.entries(protectionNode)) {
        if (!key.startsWith("@_")) continue;
        const name = key.slice(2);
        if (["password", "hashValue", "saltValue", "spinCount", "algorithmName"].includes(name)) {
          passwordPresent = true;
          continue;
        }
        options[name] = String(value);
      }
      protection.push({
        sheet: source.name,
        protected: options.sheet === "1",
        options,
        passwordPresent,
        unlockedRanges: coalesceCells(unlocked),
        hiddenFormulaRanges: coalesceCells(hiddenCells),
      });
    }

    if (maxRow + 1 > limits.maxRows || maxCol + 1 > limits.maxColumns) {
      throw new ImportRejectedError("dimensions", `${source.name} dimensions ${maxRow + 1}x${maxCol + 1} exceed the import limit.`);
    }
    const id = `sheet-${source.sheetId}`;
    sheetOrder.push(id);
    if (source.state === "veryHidden") veryHidden.push(source.name);
    univerSheets[id] = {
      id,
      name: source.name,
      ...(tabRgb ? { tabColor: normalizeRgb(tabRgb) } : {}),
      hidden: source.state === "veryHidden" ? 2 : source.state === "hidden" ? 1 : 0,
      rowCount: Math.max(maxRow + 1, 1),
      columnCount: Math.max(maxCol + 1, 1),
      defaultColumnWidth: defaultWidth,
      defaultRowHeight: defaultHeight,
      freeze: { xSplit, ySplit, startRow: ySplit, startColumn: xSplit },
      mergeData: merges,
      cellData,
      rowData,
      columnData,
      showGridlines,
      rowHeader: { width: 46 },
      columnHeader: { height: 20 },
      zoomRatio,
      scrollTop: 0,
      scrollLeft: 0,
      rightToLeft: 0,
    };
    perSheet.push({ name: source.name, state: source.state, formulas: sheetFormulas, rows: maxRow + 1, cols: maxCol + 1 });
  }

  if (formulas.length > limits.maxFormulas) {
    throw new ImportRejectedError("formulas", `Workbook has ${formulas.length} formulas; limit is ${limits.maxFormulas}.`);
  }
  const depth = observedDependencyDepth(formulas);
  if (depth > limits.maxDependencyDepth) {
    throw new ImportRejectedError("depth", `Formula dependency depth ${depth} exceeds ${limits.maxDependencyDepth}.`);
  }

  const names: DefinedNameRecord[] = [];
  const univerNames = definedNames.flatMap((node, index) => {
    const text = textOf(node);
    const decision = supportName(text);
    const localSheetId = attr(node, "localSheetId");
    const record: DefinedNameRecord = {
      name: attr(node, "name") ?? "",
      text,
      ...(localSheetId != null ? { localSheetId: Number(localSheetId) } : {}),
      hidden: attr(node, "hidden") === "1",
      builtin: (attr(node, "name") ?? "").startsWith("_xlnm."),
      supported: decision.supported,
      ...(decision.reason ? { unsupportedReason: decision.reason } : {}),
    };
    names.push(record);
    if (!decision.supported) return [];
    const sheet = localSheetId != null ? sheetOrder[Number(localSheetId)] : undefined;
    return [{
      id: `name-${index + 1}`,
      name: record.name,
      formulaOrRefString: decision.formula,
      ...(sheet ? { localSheetId: sheet } : {}),
      ...(record.hidden ? { hidden: true } : {}),
    }];
  });

  const oursByAddress = new Map(formulas.map((formula) => [`${formula.sheet}!${formula.a1}`, formula.formula]));
  for (const [address, excelFormula] of excelFormulas) {
    const ours = oursByAddress.get(address);
    if (ours === undefined) continue;
    if (ours.replace(/\s+/g, "") !== excelFormula.replace(/\s+/g, "")) disagreements += 1;
  }

  const macroButtons = macroButtonsFromPackage(files);
  const unknown = unknownFunctions(calls);
  const importLog: ImportLog = {
    sheetCount: sheets.length,
    formulaCount: formulas.length,
    sharedFormulaCount,
    arrayFormulaCount,
    definedNameCount: names.length,
    supportedNameCount: univerNames.length,
    unknownFunctions: unknown,
    rejectedExternalRefs,
    dropped: {
      vba: files.has("xl/vbaProject.bin"),
      drawings: countPart(files, (path) => path.startsWith("xl/drawings/")),
      charts: countPart(files, (path) => path.startsWith("xl/charts/")),
      comments: countPart(files, (path) => /xl\/comments\d+\.xml$/.test(path)),
      formControls: countPart(files, (path) => path.includes("ctrlProp") || path.includes("vmlDrawing")),
    },
    unsupportedConditionalFormats: conditionalFormatting.filter((rule) => !rule.mapped).length,
    unsupportedValidations: dataValidations.filter((rule) => !rule.mapped).length,
    excelJsFormulaDisagreements: disagreements,
    observedDependencyDepth: depth,
    warnings,
    perSheet,
    macroButtonCount: macroButtons.length,
  };

  return {
    sourceSha256: await sha256(bytes),
    workbook: {
      id: "phx-workbook",
      name: "Slide Sheet",
      appVersion: "0.1.0",
      locale: "enUS",
      styles: workbookStyles,
      sheetOrder,
      sheets: univerSheets,
      definedNames: univerNames,
      resources: [
        {
          name: "SHEET_DEFINED_NAME_PLUGIN",
          data: JSON.stringify(Object.fromEntries(univerNames.map((name) => [name.id, name]))),
        },
        { name: "phx-protection", data: JSON.stringify(protection) },
        { name: DATA_VALIDATION_RESOURCE, data: JSON.stringify(validationRules) },
        { name: "phx-conditional-formatting", data: JSON.stringify(conditionalFormatting.filter((rule) => rule.mapped)) },
        { name: "phx-hyperlinks", data: JSON.stringify(hyperlinks.filter((link) => !link.rejected)) },
      ],
    },
    sidecar: { veryHiddenSheets: veryHidden, names, protection, dataValidations, conditionalFormatting, hyperlinks, macroButtons, styleVersion: STYLE_VERSION },
    formulas,
    importLog,
  };
}
