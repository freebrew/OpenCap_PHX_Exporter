// A FieldCap inventory item, the columns the exporter's inventory CSV carries. The phone lists
// these on the Serial step when a snapshot exists, and the user can still type past them.

export interface InventoryItem {
  itemCode: string;
  itemName: string;
  /** Stored in capitals, matching the serial the wizard types. */
  serialNumber: string;
  mnfNumber: string;
  category: string;
  subCategory: string;
  status: string;
  /** Run hours, when the export has them. */
  runHours: number | null;
}

export interface InventorySnapshot {
  jobNumber: string;
  pulledAt: string;
  items: InventoryItem[];
}

const COLUMNS: Record<keyof Omit<InventoryItem, "runHours">, string> = {
  itemCode: "ItemCode",
  itemName: "ItemName",
  serialNumber: "SerialNumber",
  mnfNumber: "MnfNumber",
  category: "Category",
  subCategory: "SubCategory",
  status: "Status",
};

/** The items whose serial, name or category contains the text, case-insensitive. */
export function filterItems(items: InventoryItem[], query: string): InventoryItem[] {
  const needle = query.trim().toLowerCase();
  if (needle === "") return items;
  return items.filter((item) => [item.serialNumber, item.itemName, item.category, item.subCategory].some((field) => field.toLowerCase().includes(needle)));
}

/** One CSV row into an item. A row with no serial is dropped. Run hours parse or stay null. */
export function itemFromRow(row: Record<string, string>): InventoryItem | null {
  const read = (name: string) => (row[name] ?? "").trim();
  const serial = read(COLUMNS.serialNumber).toUpperCase();
  if (serial === "") return null;
  const hours = Number(read("HrsSinceService"));
  return {
    itemCode: read(COLUMNS.itemCode),
    itemName: read(COLUMNS.itemName),
    serialNumber: serial,
    mnfNumber: read(COLUMNS.mnfNumber),
    category: read(COLUMNS.category),
    subCategory: read(COLUMNS.subCategory),
    status: read(COLUMNS.status),
    runHours: Number.isFinite(hours) && read("HrsSinceService") !== "" ? hours : null,
  };
}

/** The exporter's inventory CSV (header row, then one item a row) into items. Rows without a serial are skipped. */
export function itemsFromCsv(csv: string): InventoryItem[] {
  const lines = csv.replace(/^\uFEFF/, "").split(/\r\n|\r|\n/).filter((line) => line !== "");
  const headers = splitCsv(lines[0] ?? "");
  const items: InventoryItem[] = [];
  for (const line of lines.slice(1)) {
    const cells = splitCsv(line);
    const row: Record<string, string> = {};
    headers.forEach((header, index) => {
      row[header] = cells[index] ?? "";
    });
    const item = itemFromRow(row);
    if (item) items.push(item);
  }
  return items;
}

function splitCsv(line: string): string[] {
  const out: string[] = [];
  let token = "";
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (quoted) {
      if (ch === '"' && line[i + 1] === '"') {
        token += '"';
        i += 1;
      } else if (ch === '"') quoted = false;
      else token += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ",") {
      out.push(token);
      token = "";
    } else token += ch;
  }
  out.push(token);
  return out;
}
