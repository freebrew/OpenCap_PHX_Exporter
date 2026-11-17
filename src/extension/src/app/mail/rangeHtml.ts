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

interface MailBook extends SnapshotBook {
  styles?: Record<string, { bg?: { rgb?: string | null } | null; bl?: number } | null>;
}

const FIRST_ROW = 1;
const LAST_ROW = 54;
const FIRST_COL = 1;
const LAST_COL = 5;

function dataSheet(book: MailBook): MailSheet | null {
  const sheet = Object.values(book.sheets ?? {}).find((item) => item?.name === "Data");
  return (sheet as MailSheet | null | undefined) ?? null;
}

function shown(value: unknown): { text: string; numeric: boolean } {
  if (value == null || value === "") return { text: "", numeric: false };
  if (typeof value === "number" && Number.isFinite(value)) {
    if (Object.is(value, -0) || Number.isInteger(value)) return { text: Object.is(value, -0) ? "0" : String(value), numeric: true };
    return { text: value.toFixed(6).replace(/\.?0+$/, ""), numeric: true };
  }
  if (typeof value === "boolean") return { text: value ? "TRUE" : "FALSE", numeric: false };
  return { text: String(value), numeric: false };
}

function fillOf(book: MailBook, cell: MailCell | null): { fill: string; bold: boolean } {
  const id = typeof cell?.s === "string" ? cell.s : "";
  const style = id ? book.styles?.[id] : null;
  const hex = (style?.bg?.rgb ?? "").replace("#", "");
  return {
    fill: /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex}` : "#ffffff",
    bold: style?.bl === 1,
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
      const display = shown(cell?.v);
      const look = fillOf(book, cell);
      const colspan = merge ? Math.min(merge.endColumn, LAST_COL) - column + 1 : 1;
      const rowspan = merge ? Math.min(merge.endRow, LAST_ROW) - row + 1 : 1;
      const text = display.text === "" ? "&nbsp;" : escapeHtml(display.text).replaceAll("\n", "<br>");
      const align = display.numeric ? "right" : "left";
      const span = `${colspan > 1 ? ` colspan="${colspan}"` : ""}${rowspan > 1 ? ` rowspan="${rowspan}"` : ""}`;
      cells.push(
        `<td${span} bgcolor="${look.fill}" style="background-color:${look.fill};color:#000000;${look.bold ? "font-weight:bold;" : ""}text-align:${align};border:1px solid #000000;"><font color="#000000">${text}</font></td>`,
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
    "table.xlmail td{border:1px solid #000000;padding:2px 4px;vertical-align:middle;overflow:hidden;word-wrap:break-word;color:#000000;}",
    "</style></head><body>",
    `<table class="xlmail" cellspacing="0" cellpadding="0" width="${total}" bgcolor="#ffffff" style="color:#000000;background:#ffffff;">`,
    `<colgroup>${columns}</colgroup>`,
    rows.join(""),
    "</table></body></html>",
  ].join("");
}
