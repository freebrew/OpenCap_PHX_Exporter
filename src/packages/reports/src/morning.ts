// Module21's two morning reports, as plain text. Excel reads Data by label and by cell and puts
// the result on the clipboard; the browser hands it the same fields and builds the same sentences.
// Placeholders (HOLE_SIZE, NEXT_24_TEXT) stay blank exactly as the VBA constants are.

const DASH = " \u2013 ";

export interface MorningFields {
  /** Setup!D4, the well number. */
  wellNumber?: string;
  /** Report date, "mmmm dd, yyyy". The caller supplies it; Excel uses Date. */
  reportDate?: string;
  /** Data cells and labels, keyed by the address ("E13") or the label text ("Motor S/N"). */
  cells: Record<string, unknown>;
}

export interface MorningReport {
  text: string;
  /** The same text as HTML paragraphs, escaped, for the .eml body. */
  html: string;
}

function clean(value: unknown): string {
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
  if (typeof value !== "string") return "";
  return value.replaceAll("\u00a0", " ").trim();
}

/** TryFirstNumber: a number, or the first number inside the text. */
function firstNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const text = clean(value).replaceAll(",", "");
  if (text === "") return null;
  const match = /[-+]?\d*\.?\d+/.exec(text);
  if (!match || match[0] === "" || match[0] === "." || match[0] === "+" || match[0] === "-") return null;
  const parsed = Number(match[0]);
  return Number.isFinite(parsed) ? parsed : null;
}

function countNumbers(value: unknown): number {
  return clean(value).replaceAll(",", "").match(/[-+]?\d*\.?\d+/g)?.filter((token) => /\d/.test(token)).length ?? 0;
}

function fixed(value: number, decimals: number): string {
  return value.toFixed(decimals);
}

/** FormatDepthNumber: whole numbers drop the decimals, otherwise up to three. */
function depthNumber(value: unknown): string {
  const n = firstNumber(value);
  if (n == null) return clean(value).replace(/m/gi, "");
  if (Math.abs(n - Math.round(n)) < 0.0001) return String(Math.round(n));
  return fixed(n, 3).replace(/0+$/, "").replace(/\.$/, "");
}

/** FormatPercentSmart: a % sign or a value past 1 means the number is already in percent. */
function percentSmart(value: unknown): string {
  const n = firstNumber(value);
  if (n == null) return clean(value);
  const fraction = clean(value).includes("%") || Math.abs(n) > 1 ? n / 100 : n;
  return `${fixed(fraction * 100, 2)}%`;
}

function hoursValue(value: unknown): string {
  const text = clean(value);
  if (text === "") return "";
  // A hand-typed value like "1hr8s." carries more than one number and is kept as typed.
  if (countNumbers(text) > 1) return text.endsWith(".") ? text : `${text}.`;
  const n = firstNumber(value);
  if (n == null) {
    const lowered = text.replace(/hrs/gi, "hrs").replace(/hr/gi, "hr");
    return lowered.endsWith(".") ? lowered : `${lowered}.`;
  }
  return Math.abs(n - Math.round(n)) < 0.0001 ? `${Math.round(n)}hrs.` : `${fixed(n, 1)}hrs.`;
}

function positionDistance(value: unknown): string {
  const n = firstNumber(value);
  return n == null ? clean(value).replace(/m/gi, "") : fixed(n, 2);
}

const DIRECTIONS = ["LEFT", "RIGHT", "ABOVE", "BELOW", "NORTH", "SOUTH", "EAST", "WEST"] as const;

function directionFromText(value: unknown): string {
  const text = clean(value).toUpperCase();
  return DIRECTIONS.find((direction) => text.includes(direction)) ?? "";
}

function properCaseDirection(value: unknown): string {
  const text = clean(value).toLowerCase();
  return text === "" ? "" : text.charAt(0).toUpperCase() + text.slice(1);
}

/** FormatDegreePer30m. decimals 0 prints no fraction; spaceBeforeDegree puts a space before the degree sign. */
function degreePer30(value: unknown, decimals: number, spaceBeforeDegree: boolean): string {
  const n = firstNumber(value);
  if (n == null) return clean(value);
  const spacer = spaceBeforeDegree ? " " : "";
  const printed = decimals === 0 ? String(Math.round(n)) : fixed(n, decimals);
  return `${printed}${spacer}\u00b0/30m`;
}

function cell(fields: MorningFields, key: string): unknown {
  return fields.cells[key];
}

/**
 * GetByLabelAny with its value offset. Offset 1 is the cell itself (the plain labels). Offset 2
 * is the direction cell that sits one past the value, which the caller passes as "<label> dir".
 */
function byLabel(fields: MorningFields, labels: string[], offset = 1): unknown {
  for (const label of labels) {
    const key = offset === 1 ? label : `${label} dir`;
    if (clean(cell(fields, key)) !== "") return cell(fields, key);
  }
  return "";
}

function escapeHtml(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function asHtml(text: string): string {
  return text.split("\n").map((line) => (line === "" ? "<br>" : `<p>${escapeHtml(line)}</p>`)).join("");
}

/**
 * GenerateMorningReport. Reads Data by label (with the same fallbacks) and the wellbore
 * position from E13:F14 / E15. The hole size, next-24 and narrative lines keep the VBA's
 * blank constants and its one fixed sentence.
 */
export function morningReport(fields: MorningFields): MorningReport {
  const depth24 = byLabel(fields, ["24:00 Depth", "Midnight Depth"]);
  const drilled = byLabel(fields, ["Meters Drilled"]);
  let depth00 = byLabel(fields, ["00:00 Depth", "Start Depth"]);
  if (clean(depth00) === "") {
    const end = firstNumber(depth24);
    const metres = firstNumber(drilled);
    if (end != null && metres != null) depth00 = end - metres;
  }

  const pumpRate = byLabel(fields, ["Pump Rate"]);
  const revL = byLabel(fields, ["Rev/L", "Rev / L", "RevL"]);
  const rotary = byLabel(fields, ["Rotary"]);
  const pump = firstNumber(pumpRate);
  const revs = firstNumber(revL);
  const motorRevs = pump != null && revs != null ? `${Math.round(pump * 1000 * revs)} rpm` : "";
  const rotaryRpm = firstNumber(rotary);
  const motorRpm = firstNumber(motorRevs);
  const totalFallback = byLabel(fields, ["Total RPM"]);
  const totalRpm = motorRpm != null && rotaryRpm != null
    ? `${Math.round(motorRpm + rotaryRpm)} rpm`
    : firstNumber(totalFallback) != null ? `${Math.round(firstNumber(totalFallback) ?? 0)} rpm` : clean(totalFallback);

  let lateralDir = cell(fields, "F13");
  let vertDir = cell(fields, "F14");
  const lateral = clean(cell(fields, "E13")) !== "" ? cell(fields, "E13") : byLabel(fields, ["R/L from Plan P1", "R/L from Plan P3"]);
  const vertical = clean(cell(fields, "E14")) !== "" ? cell(fields, "E14") : byLabel(fields, ["Above/Below from Plan P1", "Above/Below from Plan P3"]);
  if (clean(lateralDir) === "") lateralDir = byLabel(fields, ["R/L from Plan P1", "R/L from Plan P3"], 2);
  if (clean(vertDir) === "") vertDir = byLabel(fields, ["Above/Below from Plan P1", "Above/Below from Plan P3"], 2);
  const total = clean(cell(fields, "E15")) !== "" ? cell(fields, "E15") : byLabel(fields, ["Distance from Current Geo Target", "Distance from Current GeoTarget", "Distance from Target"]);
  if (clean(lateralDir) === "") lateralDir = directionFromText(lateral);
  if (clean(vertDir) === "") vertDir = directionFromText(vertical);

  const lines = [
    `Daily Drilling Summary \u2013 ${clean(fields.wellNumber)} ${clean(fields.reportDate)}`,
    `Drill  ${depthNumber(depth00)}m\u2013 ${depthNumber(depth24)}m`,
    `Next 24 hrs. \u2013 `,
    `Currently \u2013 Drilling ahead no issues`,
    "No issues to report.",
    "",
    `00:00 Depth: ${depthNumber(depth00)}m`,
    `24:00 Depth: ${depthNumber(depth24)}m`,
    `Meters Drilled: ${depthNumber(drilled)}m`,
    `Sliding Percentage: ${percentSmart(byLabel(fields, ["Sliding Percentage", "Sliding Values"]))}`,
    `Rotating Percentage: ${percentSmart(byLabel(fields, ["Rotating Percentage", "Rotating Values"]))}`,
    "",
    "Motor Information",
    `Motor S/N: ${clean(byLabel(fields, ["Motor S/N", "Motor SN"]))}`,
    `Configuration ${DASH.trim()} ${clean(byLabel(fields, ["Configuration", "Config"]))}`,
    `Rev/L: ${revs == null ? clean(revL) : fixed(revs, 3)}`,
    `Pump Rate: ${pump == null ? clean(pumpRate) : `${fixed(pump, 2)}m\u00b3/min`}`,
    `Motor Revs: ${motorRevs}`,
    `Rotary: ${rotaryRpm == null ? clean(rotary) : `${Math.round(rotaryRpm)} rpm`}`,
    `Total RPM: ${totalRpm}`,
    "",
    `Today Bit Hours: ${hoursValue(byLabel(fields, ["Today Bit Hours", "Daily Bit Hours"]))}`,
    `Total Bit Hours: ${hoursValue(byLabel(fields, ["Total Bit Hours"]))}`,
    `Today Circ Hours: ${hoursValue(byLabel(fields, ["Today Circ Hours", "Daily Circ Hours"]))}`,
    `Total Circ Hours: ${hoursValue(byLabel(fields, ["Total Circ Hours"]))}`,
    `Total Motor Hours:${hoursValue(byLabel(fields, ["Total Motor Hours"]))}`,
    "",
    "Position of Wellbore (as of last survey)",
    `${positionDistance(lateral)}m ${properCaseDirection(lateralDir)} of P1`,
    `${positionDistance(vertical)}m ${properCaseDirection(vertDir)} of P1`,
    `${positionDistance(total)}m Total`,
    "",
    "Motor Performance",
    `Motor Output: ${degreePer30(byLabel(fields, ["Average Motor Output", "Motor Output"]), 2, false)}`,
    `Dogleg Needed: ${degreePer30(byLabel(fields, ["Dogleg Needed"]), 0, true)}`,
    "",
    `Down Hole Temp ${temperature(byLabel(fields, ["Down Hole Temp", "Downhole Temp"]))}`,
    `TVD @ Midnight: ${tvd(byLabel(fields, ["TVD @ Midnight", "TVD at Midnight", "TVD:"]))}`,
  ];
  const text = lines.join("\r\n");
  return { text, html: asHtml(text.replaceAll("\r\n", "\n")) };
}

function temperature(value: unknown): string {
  const n = firstNumber(value);
  return n == null ? clean(value) : `${Math.round(n)}\u00b0C`;
}

function tvd(value: unknown): string {
  const n = firstNumber(value);
  if (n == null) return clean(value);
  const printed = Math.abs(n - Math.round(n)) < 0.0005 ? String(Math.round(n)) : fixed(n, 3).replace(/0+$/, "").replace(/\.$/, "");
  return `${printed}mTVD`;
}

function metres0(value: number | null, present: boolean): string {
  return present && value != null ? String(Math.round(Math.abs(value))) : "";
}

function degrees2(value: unknown, fallback: string): string {
  const n = firstNumber(value);
  return n == null ? `[${fallback}]` : fixed(n, 2);
}

function metres2(value: unknown, fallback: string): string {
  const n = firstNumber(value);
  return n == null ? `[${fallback}]` : fixed(n, 2);
}

function signedPosition(value: unknown, positive: string, negative: string, fallback: string): string {
  const n = firstNumber(value);
  if (n == null) return fallback;
  return `${fixed(Math.abs(n), 2)}m ${n < 0 ? negative : positive}`;
}

/**
 * GenerateMorningSummaryMini. One screen of the last 24 h from the Data cells the VBA reads
 * directly: H3, I23:L23, P23:Q24, E7:E9, E13:E14. A blank cell prints its address in brackets,
 * the way the mini report marks a value the sheet has not filled in yet.
 */
export function morningSummaryMini(fields: MorningFields): MorningReport {
  const bha = clean(cell(fields, "H3"));
  const start = firstNumber(cell(fields, "I23"));
  const end = firstNumber(cell(fields, "J23"));
  const totalCell = firstNumber(cell(fields, "K23"));
  const total = totalCell != null ? Math.abs(totalCell) : start != null && end != null ? Math.abs(end - start) : null;
  const slid = firstNumber(cell(fields, "L23"));
  const slidAbs = slid == null ? null : Math.abs(slid);
  const rotate = total != null && slidAbs != null ? Math.max(total - slidAbs, 0) : null;
  const slidePct = total != null && total > 0 && slidAbs != null ? (slidAbs / total) * 100 : null;
  const rotatePct = total != null && total > 0 && rotate != null ? (rotate / total) * 100 : null;
  const pct = (value: number | null) => (value == null ? "" : `${Math.round(value)}%`);

  const lines = [
    `BHA #${bha === "" ? "1" : bha}`,
    "Last 24Hrs",
    `${start == null ? "" : String(Math.round(start))} - ${end == null ? "" : String(Math.round(end))} =  Total ${metres0(total, total != null)}`,
    `Rotating ${metres0(rotate, rotate != null)}  Rotate ${pct(rotatePct)}`,
    `Slid ${slidAbs == null ? "" : String(Math.round(slidAbs))}M Slid ${pct(slidePct)}`,
    `Bit Hours = ${hoursOrAddress(cell(fields, "P23"), "P23")} Total Bit Hours = ${hoursOrAddress(cell(fields, "P24"), "P24")}`,
    `Motor Hours = ${hoursOrAddress(cell(fields, "Q23"), "Q23")} Total Motor Hours = ${hoursOrAddress(cell(fields, "Q24"), "Q24")}`,
    `Last Survey = ${metres2(cell(fields, "E7"), "E7")} Inc = ${degrees2(cell(fields, "E8"), "E8")}, Az = ${degrees2(cell(fields, "E9"), "E9")}`,
    `Well Position @ Survey = ${signedPosition(cell(fields, "E14"), "North", "South", "[E14]m North")}, ${signedPosition(cell(fields, "E13"), "East", "West", "[E13]m East")}`,
  ];
  const text = lines.join("\r\n");
  return { text, html: asHtml(text.replaceAll("\r\n", "\n")) };
}

function hoursOrAddress(value: unknown, address: string): string {
  const printed = hoursValue(value);
  return printed === "" ? `[${address}]` : printed;
}
