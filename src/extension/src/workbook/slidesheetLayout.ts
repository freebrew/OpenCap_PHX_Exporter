interface LayoutSheet {
  name?: string;
  rowData?: Record<number, { h?: number; hd?: 0 | 1 }>;
}

interface LayoutSnapshot {
  sheets?: Record<string, LayoutSheet>;
}

const SURVEY_FIRST_INDEX = 12;
const SURVEY_LAST_INDEX = 504;

/** Excel has three accidental default-height rows inside an otherwise uniform survey table. */
export function normalizeSlidesheetRowHeights(snapshot: LayoutSnapshot): void {
  const sheet = Object.values(snapshot.sheets ?? {}).find((candidate) => candidate.name === "Slidesheet");
  if (!sheet) return;
  const rows = sheet.rowData ?? {};
  const counts = new Map<number, number>();
  for (let row = SURVEY_FIRST_INDEX; row <= SURVEY_LAST_INDEX; row += 1) {
    const height = rows[row]?.h;
    if (height == null || rows[row]?.hd === 1) continue;
    counts.set(height, (counts.get(height) ?? 0) + 1);
  }
  const standard = [...counts].sort((left, right) => right[1] - left[1])[0]?.[0];
  if (standard == null) return;
  sheet.rowData = rows;
  for (let row = SURVEY_FIRST_INDEX; row <= SURVEY_LAST_INDEX; row += 1) {
    rows[row] = { ...rows[row], h: standard };
  }
}
