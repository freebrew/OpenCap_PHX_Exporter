import type { SnapshotBook } from "../../../../packages/workbook-commands/src/index";
import { escapeHtml } from "../../../../packages/reports/src/index";

interface MailCell {
  v?: unknown;
  s?: unknown;
}

interface MailMerge {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

interface MailSheet {
  name?: string;
  cellData?: Record<string, Record<string, MailCell | null> | null> | null;
  mergeData?: MailMerge[];
  columnData?: Record<string, { w?: number } | null>;
}

interface MailStyle {
  bg?: { rgb?: string | null } | null;
  bl?: number;
  n?: { pattern?: string } | null;
}

interface MailBook extends SnapshotBook {
  styles?: Record<string, MailStyle | null>;
}

const FIRST_ROW = 1;
const LAST_ROW = 54;
const FIRST_COL = 1;
const LAST_COL = 5;

function dataSheet(book: MailBook): MailSheet | null {
  const sheet = Object.values(book.sheets ?? {}).find((item) => item?.name === "Data");
  return (sheet as MailSheet | null | undefined) ?? null;
}

function generalNumber(value: number): string {
  if (Object.is(value, -0) || Number.isInteger(value)) return Object.is(value, -0) ? "0" : String(value);
  return value.toFixed(6).replace(/\.?0+$/, "");
}

/** Excel cell.Text for the number formats on Data!B2:F55 (units, currency, hours). */
export function formatExcelNumber(value: number, pattern: string): string {
  const sections = splitFormat(pattern);
  const section = value > 0 ? sections[0] ?? "" : value < 0 ? sections[1] ?? sections[0] ?? "" : sections[2] ?? sections[0] ?? "";
  if (section.trim() === "") return "";
  if (/general/i.test(section) || section === "@") return generalNumber(value);
  if (isDateSection(section)) return formatExcelDate(value, section);
  const tokens = tokenize(section);
  const number = tokens.find((token) => token.kind === "num");
  if (!number) return generalNumber(value);
  const formatted = formatMask(value, number.text);
  const body = tokens.map((token) => token.kind === "num" ? formatted : token.text).join("");
  if (value < 0 && !body.includes("-") && !body.includes("(")) return `-${body}`;
  return body;
}

function splitFormat(pattern: string): string[] {
  const sections: string[] = [];
  let current = "";
  let quoted = false;
  for (let i = 0; i < pattern.length; i += 1) {
    const ch = pattern[i] ?? "";
    if (ch === '"') {
      quoted = !quoted;
      current += ch;
      continue;
    }
    if (ch === "\\" && !quoted) {
      current += ch + (pattern[i + 1] ?? "");
      i += 1;
      continue;
    }
    if (ch === ";" && !quoted) {
      sections.push(current);
      current = "";
      continue;
    }
    current += ch;
  }
  sections.push(current);
  return sections;
}

function isDateSection(section: string): boolean {
  const bare = section.replace(/"[^"]*"/g, "").replace(/\\./g, "").replace(/\[[^\]]*\]/g, "");
  return /[dyh]/i.test(bare) || /(?<![#0.])m(?![#0])/i.test(bare);
}

function tokenize(section: string): Array<{ kind: "lit" | "num"; text: string }> {
  const tokens: Array<{ kind: "lit" | "num"; text: string }> = [];
  let i = 0;
  while (i < section.length) {
    const ch = section[i] ?? "";
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      tokens.push({ kind: "lit", text: section.slice(i + 1, end === -1 ? section.length : end) });
      i = end === -1 ? section.length : end + 1;
      continue;
    }
    if (ch === "\\") {
      tokens.push({ kind: "lit", text: section[i + 1] ?? "" });
      i += 2;
      continue;
    }
    if (ch === "_" || ch === "*") {
      i += 2;
      continue;
    }
    if (ch === "[") {
      const end = section.indexOf("]", i);
      i = end === -1 ? section.length : end + 1;
      continue;
    }
    if ("#0,.%".includes(ch)) {
      let end = i;
      while (end < section.length && "#0,.%".includes(section[end] ?? "")) end += 1;
      tokens.push({ kind: "num", text: section.slice(i, end) });
      i = end;
      continue;
    }
    tokens.push({ kind: "lit", text: ch });
    i += 1;
  }
  return tokens;
}

function formatMask(value: number, mask: string): string {
  const percent = mask.includes("%");
  const grouped = mask.includes(",");
  const decMask = mask.replaceAll("%", "").replaceAll(",", "").split(".")[1] ?? "";
  const digits = decMask.length;
  const fixed = Math.abs(percent ? value * 100 : value).toFixed(digits);
  const [rawInt = "0", rawDec = ""] = digits > 0 ? fixed.split(".") : [fixed, ""];
  let decimals = rawDec;
  let keep = decimals.length;
  while (keep > 0 && decMask[keep - 1] === "#" && decimals[keep - 1] === "0") keep -= 1;
  decimals = decimals.slice(0, keep);
  const integer = grouped ? rawInt.replace(/\B(?=(\d{3})+(?!\d))/g, ",") : rawInt;
  const suffix = percent ? "%" : "";
  return decimals === "" ? `${integer}${suffix}` : `${integer}.${decimals}${suffix}`;
}

function formatExcelDate(serial: number, section: string): string {
  const adjusted = serial >= 60 ? serial - 1 : serial;
  const date = new Date(Date.UTC(1899, 11, 31) + Math.round(adjusted * 86400000));
  if (Number.isNaN(date.getTime())) return generalNumber(serial);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const day = date.getUTCDate();
  const month = months[date.getUTCMonth()] ?? "";
  const year = date.getUTCFullYear();
  if (/d-mmm-yy/i.test(section)) return `${day}-${month}-${String(year).slice(2)}`;
  if (/d-mmm/i.test(section)) return `${day}-${month}`;
  if (/m\/d\/yyyy/i.test(section)) return `${date.getUTCMonth() + 1}/${day}/${year}`;
  return `${day}-${month}-${year}`;
}

function shown(value: unknown, pattern: string): string {
  if (value == null || value === "") return "";
  if (typeof value === "number" && Number.isFinite(value)) return pattern ? formatExcelNumber(value, pattern) : generalNumber(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function lookOf(book: MailBook, cell: MailCell | null): { fill: string; bold: boolean; pattern: string } {
  const id = typeof cell?.s === "string" ? cell.s : "";
  const style = id ? book.styles?.[id] : null;
  const hex = (style?.bg?.rgb ?? "").replace("#", "");
  const pattern = style?.n?.pattern;
  return {
    fill: /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex}` : "#ffffff",
    bold: style?.bl === 1,
    pattern: typeof pattern === "string" ? pattern : "",
  };
}

function columnWidth(sheet: MailSheet, column: number): number {
  const width = sheet.columnData?.[String(column)]?.w;
  return typeof width === "number" && width >= 20 ? Math.round(width) : 96;
}

/** Data!B2:F55 as a mail table. Text is escaped and forced black, the way Module11 hands it to Outlook. */
export function rangeToHtml(snapshot: SnapshotBook, withPicture: boolean): string {
  const book = snapshot as MailBook;
  const sheet = dataSheet(book);
  const merges = sheet?.mergeData ?? [];
  const widths = [1, 2, 3, 4, 5].map((column) => columnWidth(sheet ?? {}, column));
  const total = widths.reduce((sum, width) => sum + width, 0);
  const pictureWidth = Math.max(100, total - 2);
  const rows: string[] = [];
  for (let row = FIRST_ROW; row <= LAST_ROW; row += 1) {
    const cells: string[] = [];
    for (let column = FIRST_COL; column <= LAST_COL; column += 1) {
      const merge = merges.find((item) => row >= item.startRow && row <= item.endRow && column >= item.startColumn && column <= item.endColumn);
      if (merge) {
        const masterRow = Math.max(merge.startRow, FIRST_ROW);
        const masterColumn = Math.max(merge.startColumn, FIRST_COL);
        if (row !== masterRow || column !== masterColumn) continue;
      }
      const cell = sheet?.cellData?.[String(row)]?.[String(column)] ?? null;
      const look = lookOf(book, cell);
      const display = shown(cell?.v, look.pattern);
      const colspan = merge ? Math.min(merge.endColumn, LAST_COL) - column + 1 : 1;
      const rowspan = merge ? Math.min(merge.endRow, LAST_ROW) - row + 1 : 1;
      const text = display === "" ? "&nbsp;" : escapeHtml(display).replaceAll("\n", "<br>");
      const span = `${colspan > 1 ? ` colspan="${colspan}"` : ""}${rowspan > 1 ? ` rowspan="${rowspan}"` : ""}`;
      cells.push(
        `<td${span} align="left" bgcolor="${look.fill}" style="background-color:${look.fill};color:#000000;${look.bold ? "font-weight:bold;" : ""}text-align:left;border:1px solid #000000;"><font color="#000000">${text}</font></td>`,
      );
    }
    rows.push(`<tr>${cells.join("")}</tr>`);
  }
  if (withPicture) {
    rows.push(
      `<tr><td colspan="5" bgcolor="#262626" style="padding:0;background-color:#262626;border:1px solid #000000;"><img src="cid:daily_report.png" width="${pictureWidth}" alt="Daily wellbore plot" style="display:block;width:${pictureWidth}px;max-width:100%;height:auto;border:0;"></td></tr>`,
    );
  }
  const columns = widths.map((width) => `<col style="width:${width}px;" width="${width}">`).join("");
  return [
    "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">",
    "<meta name=\"color-scheme\" content=\"light only\">",
    "<meta name=\"supported-color-schemes\" content=\"light\">",
    "<style type=\"text/css\">",
    ":root{color-scheme:light only;}",
    "body{margin:0;padding:8px;background:#ffffff;color:#000000;}",
    `table.xlmail{border-collapse:collapse;table-layout:fixed;width:${total}px;max-width:100%;font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;background:#ffffff;}`,
    "table.xlmail td{border:1px solid #000000;padding:2px 4px;vertical-align:middle;overflow:hidden;word-wrap:break-word;color:#000000;text-align:left;}",
    "</style></head><body>",
    `<table class="xlmail" cellspacing="0" cellpadding="0" width="${total}" bgcolor="#ffffff" style="color:#000000;background:#ffffff;">`,
    `<colgroup>${columns}</colgroup>`,
    rows.join(""),
    "</table></body></html>",
  ].join("");
}
