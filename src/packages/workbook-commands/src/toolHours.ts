import { bhaNumber, numOrZero, OcTable, vbaStr, vbTrim } from "./ocTable";
import { cellText, isNumber, rgb, type CellArea, type SheetBook } from "./sheetBook";
import { looksLikeSerial, motorLike, normSerial, thirdPartySeed } from "../../strapsheets-core/src/toolKinds";

// Port of MDL_ToolHours.ToolHours_Sync: Data's 3rd-party tool table (O31:S41) and motor table
// (O45:S55) show the selected BHA's serials, and the very hidden _TH_Hours tracker keeps every
// tool and motor seen on the job with its carry-in (Previous) and this-job hours.
// The tracker is read once, changed in memory the way the VBA changes its rows, and written back whole.

const DATA = "Data";
const TRACK = "_TH_Hours";
const KIND_TOOL = "Tool";
const KIND_MOTOR = "Motor";
const COL_SERIAL = 15;
const COL_PREV = 17;
const COL_CURRENT = 18;
const COL_TOTAL = 19;
const TOOLS_TITLE = 28;
const TOOLS_HEADER = 30;
const TOOLS_FIRST = 31;
const TOOLS_LAST = 41;
const MOTOR_TITLE = 42;
const MOTOR_HEADER = 44;
const MOTOR_FIRST = 45;
const MOTOR_LAST = 55;
const TOOL_TABLE = `$O$${TOOLS_FIRST}:$S$${TOOLS_LAST}`;
const MOTOR_TABLE = `$O$${MOTOR_FIRST}:$S$${MOTOR_LAST}`;
/** FMT_HRS as Excel stores it; the VBA's "-#####" comes back as "\-#####". */
const FMT_HRS = "#####.00\"Hrs\";\\-#####.00\"Hrs\";";
const FMT_HRS_TYPED = "#####.00\"Hrs\";-#####.00\"Hrs\";";
const TRACK_HEADERS = ["Serial", "Kind", "PreJobHours", "JobHours", "Total", "BhaList"];

interface TrackRow {
  serial: string;
  kind: string;
  pre: number;
  job: number;
  total: number;
  bhas: string;
}

export interface ToolHoursResult {
  tools: number;
  motors: number;
  tracked: number;
}

export { looksLikeSerial, normSerial };

/** Digit-only key with leading zeros stripped, so 0261 and 261 match. */
function digitKey(text: string): string {
  let s = normSerial(text);
  if (s === "" || /[^0-9]/.test(s)) return "";
  while (s.length > 1 && s.startsWith("0")) s = s.slice(1);
  return s;
}

function has(text: string, find: string): boolean {
  return text.toLowerCase().includes(find.toLowerCase());
}

function same(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

interface InventoryMeta {
  cat: string;
  subCat: string;
  itemName: string;
}

class ToolHours {
  private readonly inv: OcTable;
  private readonly bha: OcTable;
  private readonly invSn: number;
  private readonly canonicalMemo = new Map<string, string>();
  private metaIndex: Map<string, InventoryMeta> | null = null;
  readonly rows: TrackRow[] = [];
  private trackerLast = 1;
  private bhaHours = new Map<number, number>();
  private curBha = 0;

  constructor(private readonly book: SheetBook) {
    this.inv = new OcTable(book, "_OC_Inventory");
    this.bha = new OcTable(book, "_OC_BHA");
    this.invSn = this.inv.col("SerialNumber") || this.inv.col("Serial #");
  }

  // ------------------------------------------------------------ serial spelling

  /** CanonicalSerial: recover the inventory spelling (0261) of a serial Excel made a number. */
  canonical(serial: string): string {
    const memo = this.canonicalMemo.get(serial);
    if (memo !== undefined) return memo;
    let out = vbTrim(serial);
    const want = digitKey(serial);
    if (want !== "" && this.inv.exists && this.invSn > 0) {
      const last = this.inv.lastRow(this.invSn);
      for (let row = 2; row <= last; row += 1) {
        const other = vbTrim(this.inv.str(row, this.invSn));
        if (digitKey(other) === want && other.length >= serial.length) {
          out = other;
          break;
        }
      }
    }
    this.canonicalMemo.set(serial, out);
    return out;
  }

  /** CellSerial: the cell's text, canonical. */
  cellSerial(value: unknown): string {
    return this.canonical(vbTrim(vbaStr(value)));
  }

  serialAt(row: number): string {
    const cell = this.book.mergeAt(DATA, row, COL_SERIAL);
    return this.cellSerial(this.book.value(DATA, cell.row, cell.col));
  }

  meta(serial: string): InventoryMeta {
    if (!this.metaIndex) {
      this.metaIndex = new Map();
      if (this.inv.exists && this.invSn > 0) {
        const cCat = this.inv.col("Category");
        const cSub = this.inv.col("SubCategory");
        const cName = this.inv.col("ItemName") || this.inv.col("Description");
        const last = this.inv.lastRow(this.invSn);
        for (let row = 2; row <= last; row += 1) {
          const key = normSerial(this.cellSerial(this.inv.value(row, this.invSn)));
          if (key === "" || this.metaIndex.has(key)) continue;
          this.metaIndex.set(key, { cat: this.inv.str(row, cCat), subCat: this.inv.str(row, cSub), itemName: this.inv.str(row, cName) });
        }
      }
    }
    return this.metaIndex.get(normSerial(serial)) ?? { cat: "", subCat: "", itemName: "" };
  }

  // ------------------------------------------------------------ tracker

  loadTracker(): void {
    if (!this.book.hasSheet(TRACK)) this.book.resetSheet(TRACK);
    if (!same(vbaStr(this.book.value(TRACK, 1, 1)), "Serial")) {
      TRACK_HEADERS.forEach((header, i) => this.book.setValue(TRACK, 1, i + 1, header));
      this.book.setValue(TRACK, 1, 13, "MotorList");
    }
    this.trackerLast = this.book.lastUsedRow(TRACK, 1);
    for (let row = 2; row <= this.trackerLast; row += 1) {
      const value = (col: number) => this.book.value(TRACK, row, col);
      this.rows.push({ serial: vbaStr(value(1)), kind: vbaStr(value(2)), pre: numOrZero(value(3)), job: numOrZero(value(4)), total: numOrZero(value(5)), bhas: vbaStr(value(6)) });
    }
  }

  /** TrackerUpsert: the row index, or -1 when the text is not a serial. */
  upsert(serial: string, kind: string): number {
    if (!looksLikeSerial(serial)) return -1;
    const key = normSerial(serial);
    let index = this.rows.findIndex((row) => normSerial(row.serial) === key);
    if (index < 0) {
      this.rows.push({ serial, kind: "", pre: 0, job: 0, total: 0, bhas: "" });
      index = this.rows.length - 1;
    }
    if (kind !== "") this.rows[index]!.kind = kind;
    return index;
  }

  purgeAndDedup(): void {
    for (let i = this.rows.length - 1; i >= 0; i -= 1) if (!looksLikeSerial(this.rows[i]!.serial)) this.rows.splice(i, 1);
    for (let i = this.rows.length - 1; i >= 0; i -= 1) {
      const row = this.rows[i]!;
      const serial = this.canonical(row.serial);
      if (looksLikeSerial(serial)) row.serial = serial;
      const key = normSerial(serial);
      if (key === "") continue;
      const other = this.rows.findIndex((entry, j) => j < i && normSerial(entry.serial) === key);
      if (other < 0) continue;
      if (row.pre > this.rows[other]!.pre) this.rows[other]!.pre = row.pre;
      this.rows.splice(i, 1);
    }
  }

  /** The tracker rows as changed so far, then WriteMotorList when the sync got that far. */
  flushTracker(motorList: boolean): void {
    const block = this.rows.map((row) => [row.serial, row.kind || null, row.pre, row.job, row.total, row.bhas || null]);
    if (block.length > 0) this.book.setTable(TRACK, 2, 1, block);
    const firstStale = 2 + block.length;
    if (this.trackerLast >= firstStale) this.book.clear(TRACK, { row: firstStale, col: 1, row2: this.trackerLast, col2: TRACK_HEADERS.length });
    if (!motorList) return;
    this.book.clear(TRACK, "M2:M80");
    let dest = 2;
    for (const row of this.rows) {
      const serial = vbTrim(row.serial);
      if (!same(row.kind, KIND_MOTOR) || serial === "") continue;
      this.book.setValue(TRACK, dest, 13, serial);
      dest += 1;
    }
  }

  // ------------------------------------------------------------ BHA and inventory

  bhaRows(): Array<{ row: number; bha: number; serial: string; src: string; desc: string; subDesc: string }> {
    const cNum = this.bha.col("BHA #");
    const cSn = this.bha.col("Serial #");
    if (!this.bha.exists || cSn === 0) return [];
    const cSrc = this.bha.col("Source");
    const cDesc = this.bha.col("Description");
    const cSub = this.bha.col("Sub Description");
    const last = Math.max(this.bha.lastRow(cSn), cNum > 0 ? this.bha.lastRow(cNum) : 1);
    const out = [];
    for (let row = 2; row <= last; row += 1) {
      out.push({
        row,
        bha: cNum > 0 ? bhaNumber(this.bha.value(row, cNum)) : 0,
        serial: this.cellSerial(this.bha.value(row, cSn)),
        src: this.bha.str(row, cSrc),
        desc: this.bha.str(row, cDesc),
        subDesc: this.bha.str(row, cSub),
      });
    }
    return out;
  }

  /** IsThirdPartySeed with InventoryMeta and the Sub Description fallback. */
  seed(entry: { serial: string; src: string; desc: string; subDesc: string }): { motor: boolean; tool: boolean } {
    const meta = this.meta(entry.serial);
    const subCat = vbTrim(meta.subCat) === "" ? entry.subDesc : meta.subCat;
    const motor = motorLike(entry.serial, entry.desc, meta.itemName);
    return { motor, tool: !motor && thirdPartySeed(entry.src, entry.serial, entry.desc, meta.cat, subCat, meta.itemName) };
  }

  upsertKnownSerials(): void {
    this.purgeAndDedup();
    for (const entry of this.bhaRows()) {
      if (!looksLikeSerial(entry.serial)) continue;
      const kind = this.seed(entry);
      if (kind.motor) this.upsert(entry.serial, KIND_MOTOR);
      else if (kind.tool) this.upsert(entry.serial, KIND_TOOL);
    }
    if (!this.inv.exists || this.invSn === 0) return;
    const cName = this.inv.col("ItemName") || this.inv.col("Description");
    const last = this.inv.lastRow(this.invSn);
    for (let row = 2; row <= last; row += 1) {
      const serial = this.cellSerial(this.inv.value(row, this.invSn));
      const itemName = this.inv.str(row, cName);
      if (serial !== "" && motorLike(serial, itemName, itemName)) this.upsert(serial, KIND_MOTOR);
    }
  }

  /** BHA # -> BHA Total Hrs from the first row of that BHA. */
  private readBhaHours(): Map<number, number> {
    const hours = new Map<number, number>();
    const cNum = this.bha.col("BHA #");
    const cHrs = this.bha.col("BHA Total Hrs");
    if (!this.bha.exists || cNum === 0 || cHrs === 0) return hours;
    const last = this.bha.lastRow(cNum);
    for (let row = 2; row <= last; row += 1) {
      const n = bhaNumber(this.bha.value(row, cNum));
      if (n > 0 && !hours.has(n)) hours.set(n, numOrZero(this.bha.value(row, cHrs)));
    }
    return hours;
  }

  /** Serial -> comma-separated BHA numbers that carried it. */
  private memberLists(): Map<string, string> {
    const members = new Map<string, string>();
    const cNum = this.bha.col("BHA #");
    const cSn = this.bha.col("Serial #");
    if (!this.bha.exists || cNum === 0 || cSn === 0) return members;
    const last = this.bha.lastRow(cSn);
    for (let row = 2; row <= last; row += 1) {
      const n = bhaNumber(this.bha.value(row, cNum));
      const key = normSerial(this.cellSerial(this.bha.value(row, cSn)));
      if (n <= 0 || key === "") continue;
      const current = members.get(key);
      if (!current) members.set(key, String(n));
      else if (!`,${current},`.includes(`,${n},`)) members.set(key, `${current},${n}`);
    }
    return members;
  }

  recomputeJobHours(curBha: number, liveHours: number): void {
    this.curBha = curBha;
    this.bhaHours = this.readBhaHours();
    const members = this.memberLists();
    for (const row of this.rows) {
      const key = normSerial(vbTrim(row.serial));
      row.bhas = key !== "" ? members.get(key) ?? "" : "";
      const seen = new Set<number>();
      let job = 0;
      for (const part of row.bhas === "" ? [] : row.bhas.split(",")) {
        const n = bhaNumber(vbTrim(part));
        if (n <= 0 || seen.has(n)) continue;
        seen.add(n);
        if (n === curBha) job += liveHours;
        else job += this.bhaHours.get(n) ?? 0;
      }
      row.job = job;
      row.total = row.pre + job;
    }
  }

  /**
   * Current Hours follows Q24, the live circulating total, plus hours from every
   * other BHA that carried the serial. A row already on the active table with
   * no BHA history still follows Q24.
   */
  circFormula(bhas: string, curBha: number, hours: Map<number, number>): string {
    const seen = new Set<number>();
    let prior = 0;
    let live = bhas === "";
    for (const part of bhas === "" ? [] : bhas.split(",")) {
      const n = bhaNumber(vbTrim(part));
      if (n <= 0 || seen.has(n)) continue;
      seen.add(n);
      if (n === curBha) live = true;
      else prior += hours.get(n) ?? 0;
    }
    const extra = Math.round(prior * 10000) / 10000;
    if (live && extra === 0) return "=Q24";
    if (live) return `=Q24+${extra}`;
    return `=${extra}`;
  }

  /** Point the painted tool and motor rows at Q24. Does not add, clear, or restamp serials. */
  bindFrontFormulas(): boolean {
    const curBha = bhaNumber(this.book.value(DATA, 3, 8));
    const hours = this.readBhaHours();
    const members = this.memberLists();
    let changed = false;
    for (const [first, last] of [[TOOLS_FIRST, TOOLS_LAST], [MOTOR_FIRST, MOTOR_LAST]] as const) {
      for (let row = first; row <= last; row += 1) {
        const serial = this.serialAt(row);
        if (!looksLikeSerial(serial)) continue;
        const before = this.book.writes.length;
        this.writeCurrentFormula(row, this.circFormula(members.get(normSerial(serial)) ?? "", curBha, hours));
        if (this.book.writes.length > before) changed = true;
      }
    }
    return changed;
  }

  activeToolSerials(bha: number): string[] {
    if (bha <= 0) return [];
    const out: string[] = [];
    for (const entry of this.bhaRows()) {
      if (entry.bha !== bha || !looksLikeSerial(entry.serial)) continue;
      if (this.seed(entry).tool) addUnique(out, entry.serial);
    }
    return out;
  }

  activeMotorSerials(bha: number): string[] {
    if (bha <= 0 || this.bha.col("BHA #") === 0 || this.bha.col("Description") === 0) return [];
    return this.bhaRows().filter((entry) => entry.bha === bha && looksLikeSerial(entry.serial) && motorLike(entry.serial, entry.desc, "")).map((entry) => entry.serial);
  }

  // ------------------------------------------------------------ front tables

  layoutIsNew(): boolean {
    return same(cellText(this.book.value(DATA, TOOLS_HEADER, COL_SERIAL)), "Serial Number") && has(vbaStr(this.book.value(DATA, TOOLS_TITLE, COL_SERIAL)), "3rd Party");
  }

  harvest(): void {
    const ranges: Array<[number, number, string]> = this.layoutIsNew()
      ? [[TOOLS_FIRST, TOOLS_LAST, KIND_TOOL], [MOTOR_FIRST, MOTOR_LAST, KIND_MOTOR]]
      : [[33, 38, KIND_TOOL], [42, 55, KIND_MOTOR]];
    for (const [first, last, kind] of ranges) {
      for (let row = first; row <= last; row += 1) {
        const serial = this.serialAt(row);
        if (!looksLikeSerial(serial)) continue;
        const prev = numOrZero(this.book.value(DATA, row, COL_PREV));
        this.rows[this.upsert(serial, kind)]!.pre = prev;
      }
    }
  }

  rebuildTables(): void {
    const area: CellArea = { row: TOOLS_TITLE, col: COL_SERIAL, row2: MOTOR_LAST, col2: COL_TOTAL };
    this.book.unmerge(DATA, area);
    this.book.clear(DATA, area, true);
    const title = (row: number, text: string, fill: string) => {
      const box = { row, col: COL_SERIAL, row2: row + 1, col2: COL_TOTAL };
      this.book.merge(DATA, box);
      this.book.setValue(DATA, row, COL_SERIAL, text);
      this.book.setLook(DATA, box, { bold: true, hAlign: "center", vAlign: "middle" });
      this.book.setFill(DATA, box, fill);
    };
    const header = (row: number) => {
      this.book.merge(DATA, { row, col: COL_SERIAL, row2: row, col2: COL_SERIAL + 1 });
      this.book.setValue(DATA, row, COL_SERIAL, "Serial Number");
      this.book.setValue(DATA, row, COL_PREV, "Previous hours");
      this.book.setValue(DATA, row, COL_CURRENT, "Current Hours");
      this.book.setValue(DATA, row, COL_TOTAL, "Total");
      const line = { row, col: COL_SERIAL, row2: row, col2: COL_TOTAL };
      this.book.setLook(DATA, line, { bold: true });
      this.book.setLook(DATA, { row, col: COL_SERIAL, row2: row, col2: COL_SERIAL + 1 }, { hAlign: "center" });
      this.book.setFill(DATA, line, rgb(217, 217, 217));
      this.borders(row);
    };
    const data = (first: number, last: number) => {
      for (let row = first; row <= last; row += 1) {
        this.book.merge(DATA, { row, col: COL_SERIAL, row2: row, col2: COL_SERIAL + 1 });
        this.book.setFormat(DATA, row, COL_SERIAL, "@");
        this.book.setLook(DATA, { row, col: COL_SERIAL, row2: row, col2: COL_SERIAL + 1 }, { hAlign: "left" });
        this.book.setFormat(DATA, row, COL_PREV, "0.00");
        this.book.setFormat(DATA, row, COL_CURRENT, "0.00");
        this.book.setFormula(DATA, row, COL_TOTAL, `=IFERROR(IF(O${row}<>"",SUM(Q${row}:R${row}),""),"")`);
        this.borders(row);
      }
    };
    title(TOOLS_TITLE, "Enter any 3rd Party Tools/hours below", rgb(221, 235, 247));
    header(TOOLS_HEADER);
    data(TOOLS_FIRST, TOOLS_LAST);
    title(MOTOR_TITLE, "Enter Motors & hours below", rgb(252, 228, 214));
    header(MOTOR_HEADER);
    data(MOTOR_FIRST, MOTOR_LAST);
  }

  /** Borders.LineStyle = xlContinuous on O:S of one row. */
  private borders(row: number): void {
    this.book.setBorder(DATA, { row, col: COL_SERIAL, row2: row, col2: COL_SERIAL + 1 }, "outside", "thin", "#000000");
    for (let col = COL_PREV; col <= COL_TOTAL; col += 1) this.book.setBorder(DATA, { row, col, row2: row, col2: col }, "outside", "thin", "#000000");
  }

  fixTitleSpelling(): void {
    const last = this.book.lastUsedRow(DATA, COL_SERIAL);
    for (let row = 1; row <= last; row += 1) {
      const text = vbaStr(this.book.value(DATA, row, COL_SERIAL));
      if (!has(text, "Toosl")) continue;
      const cell = this.book.mergeAt(DATA, row, COL_SERIAL);
      this.book.setValue(DATA, cell.row, cell.col, vbaStr(this.book.value(DATA, cell.row, cell.col)).replace(/toosl/gi, "Tools"));
      return;
    }
  }

  private ensureFormat(row: number, col: number, pattern: string, ...alike: string[]): void {
    const current = this.book.style(DATA, row, col)?.n?.pattern ?? "";
    if (current === pattern || alike.includes(current)) return;
    this.book.setFormat(DATA, row, col, pattern);
  }

  private writeFrontRow(row: number, serial: string, kind: string): void {
    const old = this.serialAt(row);
    const index = this.upsert(serial, kind);
    if (index < 0) throw new Error(`"${serial}" is not a serial number.`);
    const cell = this.book.mergeAt(DATA, row, COL_SERIAL);
    this.ensureFormat(cell.row, cell.col, "@");
    if (!same(vbaStr(this.book.value(DATA, cell.row, cell.col)), serial)) this.book.setValue(DATA, cell.row, cell.col, serial);
    if (normSerial(old) !== normSerial(serial)) this.book.clear(DATA, { row, col: COL_PREV, row2: row, col2: COL_PREV });
    const tracked = this.rows[index]!;
    this.writeCurrentFormula(row, this.circFormula(tracked.bhas, this.curBha, this.bhaHours));
  }

  private writeCurrentFormula(row: number, formula: string): void {
    const current = this.book.formula(DATA, row, COL_CURRENT);
    const compact = (value: string) => value.trim().replace(/^=/, "").replace(/\s+/g, "").toUpperCase();
    if (current != null && compact(current) === compact(formula)) return;
    this.book.setFormula(DATA, row, COL_CURRENT, formula);
  }

  private clearFrontRow(row: number): void {
    const cell = this.book.mergeAt(DATA, row, COL_SERIAL);
    const empty = (col: number) => this.book.value(DATA, row, col) == null && !this.book.hasFormula(DATA, row, col);
    if (vbTrim(vbaStr(this.book.value(DATA, cell.row, cell.col))) === "" && empty(COL_PREV) && empty(COL_CURRENT)) return;
    this.ensureFormat(cell.row, cell.col, "@");
    this.book.setValue(DATA, cell.row, cell.col, null);
    this.book.clear(DATA, { row, col: COL_PREV, row2: row, col2: COL_CURRENT });
  }

  paint(first: number, last: number, serials: string[], kind: string): number {
    let row = first;
    for (const serial of serials) {
      if (row > last) break;
      this.writeFrontRow(row, serial, kind);
      row += 1;
    }
    const painted = row - first;
    for (; row <= last; row += 1) this.clearFrontRow(row);
    return painted;
  }

  motorSerials(bha: number): string[] {
    const out: string[] = [];
    let motor = vbTrim(vbaStr(this.book.value(DATA, 22, 3)));
    if (motor === "") motor = this.activeMotorSerials(bha)[0] ?? "";
    if (motor !== "") addUnique(out, motor);
    for (const serial of this.activeMotorSerials(bha)) addUnique(out, serial);
    return out;
  }

  private setLookup(row: number, col: number, formula: string, guarded: boolean): void {
    if (guarded) {
      if (!this.book.hasFormula(DATA, row, col)) return;
      const current = this.book.formula(DATA, row, col);
      if (current != null && !has(current, "VLOOKUP")) return;
    }
    if (this.book.formula(DATA, row, col) === formula) return;
    this.book.setFormula(DATA, row, col, formula);
  }

  fixDependentFormulas(): void {
    const lookup = (key: string, col: number) => `=IFERROR(VLOOKUP(${key},${TOOL_TABLE},${col},FALSE),"")`;
    const blank = (key: string, col: number) => `=IFERROR(IF(ISBLANK(${key}),"",VLOOKUP(${key},${TOOL_TABLE},${col},0)),"")`;
    for (const row of [24, 25, 26, 27, 28, 30]) {
      this.setLookup(row, 5, lookup(`D${row}`, 3), true);
      this.setLookup(row, 6, blank(`D${row}`, 5), true);
    }
    this.setLookup(31, 5, lookup("$D31", 3), true);
    this.setLookup(31, 6, lookup("$D31", 5), true);
    this.setLookup(32, 5, lookup("D32", 3), true);
    this.setLookup(31, 3, `=IFERROR(VLOOKUP($C$22,${MOTOR_TABLE},3,0),"")`, false);
    for (let row = 45; row <= 55; row += 1) {
      this.setLookup(row, 6, `=IFERROR(VLOOKUP(B${row},${MOTOR_TABLE},5,FALSE),IFERROR(VLOOKUP(B${row},${TRACK}!$A$2:$E$80,5,FALSE),""))`, false);
    }
    for (const row of [23, 24]) {
      const formula = this.book.formula(DATA, row, 3);
      if (formula && has(formula, "B45:F52")) this.book.setFormula(DATA, row, 3, formula.replace(/B45:F52/gi, "B45:F55"));
    }
    const hourCells: Array<[number, number]> = [[31, 3]];
    for (let row = 45; row <= 55; row += 1) hourCells.push([row, 6]);
    for (let row = 24; row <= 28; row += 1) hourCells.push([row, 5], [row, 6]);
    for (let row = 30; row <= 32; row += 1) hourCells.push([row, 5], [row, 6]);
    for (const [row, col] of hourCells) this.ensureFormat(row, col, FMT_HRS, FMT_HRS_TYPED);
  }
}

function addUnique(list: string[], serial: string): void {
  const key = normSerial(serial);
  if (key !== "" && !list.some((entry) => normSerial(entry) === key)) list.push(serial);
}

/** Replace a stamped Current Hours number with a formula that follows Q24. */
export function bindLiveToolHours(book: SheetBook): boolean {
  if (!book.hasSheet(DATA)) return false;
  return new ToolHours(book).bindFrontFormulas();
}

/** ToolHours_Sync. Returns null when the workbook has no Data sheet. */
export function toolHoursSync(book: SheetBook): ToolHoursResult | null {
  if (!book.hasSheet(DATA)) return null;
  const hoursValue = book.value(DATA, 24, 17);
  const liveHours = isNumber(hoursValue) ? hoursValue : numOrZero(hoursValue);
  const bha = bhaNumber(book.value(DATA, 3, 8));
  const sync = new ToolHours(book);
  const result: ToolHoursResult = { tools: 0, motors: 0, tracked: 0 };
  sync.loadTracker();
  let painted = false;
  try {
    sync.fixTitleSpelling();
    sync.harvest();
    if (!sync.layoutIsNew()) sync.rebuildTables();
    sync.harvest();
    sync.upsertKnownSerials();
    sync.recomputeJobHours(bha, liveHours);
    result.tools = sync.paint(TOOLS_FIRST, TOOLS_LAST, sync.activeToolSerials(bha), KIND_TOOL);
    result.motors = sync.paint(MOTOR_FIRST, MOTOR_LAST, sync.motorSerials(bha), KIND_MOTOR);
    painted = true;
  } finally {
    sync.flushTracker(painted);
    result.tracked = sync.rows.length;
  }
  sync.fixDependentFormulas();
  return result;
}
