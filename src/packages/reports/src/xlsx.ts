import { strToU8, zipSync } from "fflate";
import { sheetOf, type SnapshotBook } from "../../workbook-commands/src/cells";

function xml(value: unknown): string {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function columnName(index: number): string {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function sheetXml(book: SnapshotBook, sheetName: string): string {
  const sheet = sheetOf(book, sheetName);
  const rows = Object.entries(sheet?.cellData ?? {}).slice(0, 2000);
  const body = rows.map(([rowKey, columns]) => {
    const cells = Object.entries(columns ?? {}).map(([columnKey, cell]) => {
      const value = cell?.v;
      if (value == null || value === "") return "";
      const ref = `${columnName(Number(columnKey))}${Number(rowKey) + 1}`;
      if (typeof value === "number") return `<c r="${ref}"><v>${value}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t>${xml(value)}</t></is></c>`;
    }).join("");
    return cells ? `<row r="${Number(rowKey) + 1}">${cells}</row>` : "";
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
}

/** Office export of calculated values. This is not the runtime workbook. */
export function exportXlsx(book: SnapshotBook, sheetNames: string[]): Uint8Array {
  const names = sheetNames.slice(0, 10);
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${names.map((_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((name, index) => `<sheet name="${xml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("")}</Relationships>`),
  };
  names.forEach((name, index) => {
    files[`xl/worksheets/sheet${index + 1}.xml`] = strToU8(sheetXml(book, name));
  });
  return zipSync(files);
}
