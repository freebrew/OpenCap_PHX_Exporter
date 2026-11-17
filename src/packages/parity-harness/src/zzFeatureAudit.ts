// Read-only feature census of Slide Sheet - 35783.xlsm. Prints counts, names, and macros. No cell values.
import { readFileSync, writeFileSync } from "node:fs";
import { unzipSync } from "fflate";
import { macroButtonsFromWorkbook } from "../../xlsx-bridge/src/macroButtons.ts";

const xlsm = "Z:/PHX_FieldCap/Slide Sheet - 35783.xlsm";
const bytes = readFileSync(xlsm);
const files = unzipSync(new Uint8Array(bytes));
const text = (name: string) => new TextDecoder().decode(files[name] ?? new Uint8Array());

const workbook = text("xl/workbook.xml");
const sheets = [...workbook.matchAll(/<sheet\b([^>]*)\/>/g)].map((match) => {
  const attrs = match[1] ?? "";
  const attr = (key: string) => attrs.match(new RegExp(`${key}="([^"]*)"`))?.[1] ?? "";
  return { name: attr("name"), state: attr("state") || "visible", id: attr("sheetId") };
});

const formulaCounts: Record<string, number> = {};
const udfCounts = new Map<string, number>();
const excelFns = new Set("SUM,IF,AND,OR,NOT,IFERROR,IFNA,ROUND,ROUNDUP,ROUNDDOWN,INT,MOD,ABS,MAX,MIN,AVERAGE,COUNT,COUNTA,COUNTIF,COUNTIFS,SUMIF,SUMIFS,VLOOKUP,HLOOKUP,INDEX,MATCH,XLOOKUP,OFFSET,INDIRECT,CHOOSE,TEXT,VALUE,LEFT,RIGHT,MID,LEN,TRIM,UPPER,LOWER,FIND,SEARCH,SUBSTITUTE,REPLACE,CONCAT,CONCATENATE,TEXTJOIN,DATE,TIME,NOW,TODAY,YEAR,MONTH,DAY,HOUR,MINUTE,SECOND,DATEDIF,EOMONTH,EDATE,NETWORKDAYS,WEEKDAY,ROW,COLUMN,ROWS,COLUMNS,ADDRESS,CELL,ISBLANK,ISNUMBER,ISTEXT,ISERROR,ISNA,NA,TRUE,FALSE,AND,OR,NOT,SQRT,POWER,LOG,LN,EXP,PI,SIN,COS,TAN,ASIN,ACOS,ATAN,ATAN2,RADIANS,DEGREES,SIGN,FLOOR,CEILING,TRUNC,MROUND,SUMPRODUCT,LOOKUP,IFS,SWITCH,LET,LAMBDA,FILTER,SORT,UNIQUE,SEQUENCE,RAND,RANDBETWEEN,HYPERLINK,N,T,TYPE,ERROR.TYPE,ISODD,ISEVEN,AGGREGATE,SUBTOTAL,AVERAGEIF,AVERAGEIFS,LARGE,SMALL,RANK,MEDIAN,STDEV,VAR,IFERROR,NUMBERVALUE,FIXED,DOLLAR,PROPER,EXACT,REPT,CHAR,CODE,CLEAN,VALUE,DATEVALUE,TIMEVALUE,DAYS,YEARFRAC,WORKDAY,PMT,PV,FV,RATE,NPV,IRR,AND,OR".split(","));

function tallyFormula(xml: string) {
  let n = 0;
  for (const match of xml.matchAll(/<f\b[^>]*>([\s\S]*?)<\/f>/g)) {
    n += 1;
    const body = (match[1] ?? "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/"[^"]*"/g, "\"\"");
    for (const call of body.matchAll(/(?:^|[^A-Z0-9_.])([A-Z][A-Z0-9_.]*)\s*\(/gi)) {
      const name = (call[1] ?? "").toUpperCase();
      if (!name || excelFns.has(name)) continue;
      udfCounts.set(name, (udfCounts.get(name) ?? 0) + 1);
    }
  }
  return n;
}

let formulas = 0;
let shared = 0;
let arrays = 0;
for (const [path, data] of Object.entries(files)) {
  if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) continue;
  const xml = new TextDecoder().decode(data);
  const n = tallyFormula(xml);
  formulas += n;
  shared += xml.match(/<f\b[^>]*t="shared"/g)?.length ?? 0;
  arrays += xml.match(/<f\b[^>]*t="array"/g)?.length ?? 0;
  formulaCounts[path] = n;
}

const rels = text("xl/_rels/workbook.xml.rels");
const relMap = new Map<string, string>();
for (const match of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)) relMap.set(match[1] ?? "", match[2] ?? "");
const sheetFormula: Array<{ name: string; state: string; formulas: number }> = [];
for (const match of workbook.matchAll(/<sheet\b([^>]*)\/>/g)) {
  const attrs = match[1] ?? "";
  const name = attrs.match(/name="([^"]*)"/)?.[1] ?? "";
  const state = attrs.match(/state="([^"]*)"/)?.[1] ?? "visible";
  const rid = attrs.match(/r:id="([^"]*)"/)?.[1] ?? "";
  const target = (relMap.get(rid) ?? "").replace(/^\//, "").replace(/^xl\//, "xl/");
  const path = target.startsWith("xl/") ? target : `xl/${target}`;
  sheetFormula.push({ name, state, formulas: formulaCounts[path] ?? 0 });
}

const names = [...workbook.matchAll(/<definedName\b/g)].length;
const buttons = macroButtonsFromWorkbook(new Uint8Array(bytes)).map((button) => ({
  sheet: button.sheet,
  caption: button.caption,
  macro: button.macro,
}));

const vba = files["xl/vbaProject.bin"] ?? new Uint8Array();
const ascii = new TextDecoder("latin1").decode(vba);
const moduleHits = [...ascii.matchAll(/\b((?:MDL|frm|cls|Module|Sheet|ThisWorkbook)[A-Za-z0-9_]{0,40})\b/g)].map((match) => match[1] ?? "");
const modules = [...new Set(moduleHits)].sort();

const out = {
  bytes: bytes.length,
  vbaBytes: vba.length,
  sheets: sheetFormula,
  definedNames: names,
  formulas,
  sharedFormulaNodes: shared,
  arrayFormulaNodes: arrays,
  buttons,
  udfs: [...udfCounts.entries()].sort((a, b) => b[1] - a[1]).map(([name, count]) => ({ name, count })),
  vbaNameHits: modules,
};
const dest = "C:/Users/User/AppData/Local/Temp/phx-feature-audit.json";
writeFileSync(dest, JSON.stringify(out, null, 2));
console.log(`wrote ${dest}`);
console.log(`sheets=${sheetFormula.length} names=${names} formulas=${formulas} buttons=${buttons.length} udfs=${udfCounts.size} vbaHits=${modules.length}`);
