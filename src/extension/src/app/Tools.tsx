import { useEffect, useRef, useState, type FormEvent } from "react";
import { createPortal } from "react-dom";
import { CUTOVER_CHECKS, readyForNative } from "../../../packages/cutover/src/index";
import { renderCorridor } from "../../../packages/corridor-3d/src/picture";
import { ODATA_ENTITIES } from "../../../packages/fieldcap-sync/src/index";
import { dailyText, escapeHtml, exportXlsx } from "../../../packages/reports/src/index";
import { MailSettings } from "./mail/MailSettings";
import { costDateSerial, daySlideMetres, readCostDays, readNumber, SheetBook, writeCostDay, type CostDay, type SnapshotBook } from "../../../packages/workbook-commands/src/index";
import type { WorkbookHost } from "../workbook/facade";

/** The panel docks beside the grid so it never sits under (or over) Univer's canvas layers. */
export const TOOLS_DOCK_ID = "tools-dock";

interface ToolsProps {
  host: () => WorkbookHost | null;
  book: () => SnapshotBook | null;
  authority: string;
  onAuthority: (authority: "native") => Promise<void>;
  onStatus: (message: string) => void;
  onEmailReport: () => Promise<void>;
  toolRequest: { section: string; token: number } | null;
}

export function Tools(props: ToolsProps) {
  const [open, setOpen] = useState(false);
  const [picture, setPicture] = useState<CorridorShot | null>(null);
  const [report, setReport] = useState("");
  const [checked, setChecked] = useState<string[]>([]);
  const [dayRow, setDayRow] = useState("4");
  const [pickedRow, setPickedRow] = useState(0);
  const [costDate, setCostDate] = useState("");
  const [costDaily, setCostDaily] = useState("");
  const [costBatch, setCostBatch] = useState("");
  const drawGeneration = useRef(0);

  useEffect(() => {
    if (props.toolRequest) setOpen(true);
  }, [props.toolRequest]);

  useEffect(() => {
    if (!open) return;
    void refreshCorridor(false);
  }, [open]);

  useEffect(() => {
    if (!open || !props.toolRequest) return;
    document.getElementById(`tool-${props.toolRequest.section}`)?.scrollIntoView({ block: "nearest" });
  }, [open, props.toolRequest]);

  const toggle = (
    <button type="button" aria-pressed={open} onClick={() => setOpen(!open)}>{open ? "Hide Well Tools" : "Well Tools"}</button>
  );
  const dock = typeof document === "undefined" ? null : document.getElementById(TOOLS_DOCK_ID);
  if (!open || !dock) return toggle;

  function snapshot(): SnapshotBook | null {
    return props.book();
  }

  async function refreshCorridor(announce: boolean) {
    const generation = drawGeneration.current + 1;
    drawGeneration.current = generation;
    const host = props.host();
    try {
      await host?.flushCalculation(12_000);
    } catch {
      // The picture still uses whatever survey values are already on the sheet.
    }
    if (generation !== drawGeneration.current) return;
    const shot = captureCorridor(props.book());
    setPicture(shot);
    if (announce) props.onStatus(shot ? `Corridor picture redrawn, ${shot.caption}.` : "This day has no survey to draw.");
  }

  function writeDay() {
    const book = snapshot();
    const host = props.host();
    const row = Number(dayRow);
    if (!book || !host || !Number.isFinite(row)) return;
    host.setCell("Data", `L${row}`, daySlideMetres(book, row));
    props.onStatus(`Data L${row} is the slide overlap inside that day's depths.`);
  }

  /**
   * frmCosts: one day, with its date, daily cost and batch note. A date already on the sheet is
   * updated; a new date appends a row with the running-total formula. The hours come from the
   * Data day row the date falls in (P and Q), shown beside the cost.
   */
  function writeCost(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const book = snapshot();
    const host = props.host();
    if (!book || !host) return;
    const serial = costDateSerial(costDate);
    const daily = numberOrBlank(costDaily);
    if (serial === 0 || daily == null) {
      props.onStatus("Enter a date and a daily cost.");
      return;
    }
    const sheet = new SheetBook(book, []);
    const written = writeCostDay(sheet, serial, daily, costBatch);
    if (!written) {
      props.onStatus("The Costs sheet has no free day row left.");
      return;
    }
    host.applyWrites(sheet.writes);
    const hours = dayHours(book, serial);
    props.onStatus(`Costs ${written.added ? "added" : "updated"} row ${written.row}${hours}.`);
  }

  function costDays(): CostDay[] {
    const book = snapshot();
    return book ? readCostDays(new SheetBook(book, [])) : [];
  }

  function buildReport() {
    const book = snapshot();
    const fields = {
      "Day start": text(book, "Data", "C4"),
      "Day end": text(book, "Data", "C5"),
      "Slide m": text(book, "Data", `L${dayRow}`),
    };
    const html = `<h1>Morning report</h1><pre>${escapeHtml(dailyText(fields))}</pre>`;
    setReport(html);
    props.onStatus("Morning report is ready to review.");
  }

  function downloadXlsx() {
    const book = snapshot();
    if (!book) return;
    const names = Object.values(book.sheets ?? {}).map((sheet) => sheet?.name).filter((name): name is string => !!name && !name.startsWith("_"));
    const bytes = exportXlsx(book, names.slice(0, 5));
    const copy = new Uint8Array(bytes.byteLength);
    copy.set(bytes);
    const url = URL.createObjectURL(new Blob([copy], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const link = document.createElement("a");
    link.href = url;
    link.download = "slide-sheet-export.xlsx";
    link.click();
    props.onStatus("Exported calculated values for office use. The preview snapshot is still the working copy.");
  }

  async function pollFieldCap() {
    const response = await chrome.runtime.sendMessage({ type: "phx-fieldcap-poll", entity: "ClientJobs" });
    props.onStatus(response?.ok ? `FieldCap answered with ${response.count ?? 0} ClientJobs rows. Review them before applying.` : (response?.error ?? "Open a FieldCap tab, then poll again."));
  }

  /** Field names only: does FieldCap publish survey stations (MD / Inc / Azm) and activity hours? */
  async function checkFieldCapFields() {
    const response = await chrome.runtime.sendMessage({ type: "phx-fieldcap-fields" }) as
      | { ok: true; survey: Array<{ entity: string; fields: string[]; station: { md: string | null; inc: string | null; azm: string | null; bit: string | null } }>; activity: Array<{ entity: string; fields: string[] }> }
      | { ok: false; error: string }
      | undefined;
    if (!response?.ok) {
      props.onStatus(response?.error ?? "Open FieldCap in a tab, then check again.");
      return;
    }
    const surveys = response.survey.map(({ entity, station }) => {
      const found = [station.md && `MD=${station.md}`, station.inc && `Inc=${station.inc}`, station.azm && `Azm=${station.azm}`].filter(Boolean);
      const bit = station.bit ? `, Bit=${station.bit}` : ", no bit depth field";
      return `${entity}: ${found.length === 3 ? `stations (${found.join(", ")}${bit})` : found.length > 0 ? `partial (${found.join(", ")})` : "no MD/Inc/Azm fields"}`;
    });
    const activity = response.activity.map(({ entity, fields }) => `${entity}: ${fields.length} fields${fields.some((field) => /duration|hours/i.test(field)) ? ", has hours" : ""}`);
    props.onStatus([...surveys, ...activity].join(" | ") || "FieldCap's $metadata has no survey or activity entities.");
  }

  const days = costDays();
  const book = snapshot();

  return (
    <>
    {toggle}
    {createPortal(
    <section className="tools" aria-label="Well tools">
      <header>
        <strong>Well tools</strong>
        <button type="button" onClick={() => setOpen(false)}>Close</button>
      </header>
      <p className="tools-note">Authority: {props.authority === "native" ? "this folder snapshot" : "the Excel workbook, with this preview beside it"}.</p>
      <section className="tool-bound" aria-label="Corridor picture">
        <div className="bound-title">
          <h2>Corridor picture</h2>
          <button type="button" onClick={() => void refreshCorridor(true)}>Redraw</button>
        </div>
        <div className="corridor-frame">
          {picture ? <img className="corridor" alt={`${picture.kind} corridor`} src={picture.url} /> : <p className="corridor-empty">This day has no survey to draw.</p>}
        </div>
        {picture ? <p className="corridor-caption">{picture.caption}</p> : null}
      </section>
      <form id="tool-costs" className="tool-bound cost-form" onSubmit={writeCost}>
        <h2>Costs</h2>
        {days.length === 0 ? <p className="cost-empty">No cost days on the sheet yet.</p> : (
          <div className="cost-calendar" role="list">
            {days.map((day) => {
              const cal = calendarFromSerial(day.date);
              const hours = chipHours(book, day.date);
              return (
                <button
                  key={day.row}
                  type="button"
                  role="listitem"
                  className={pickedRow === day.row ? "cost-chip selected" : "cost-chip"}
                  title={chipTitle(cal, day, hours)}
                  aria-label={chipTitle(cal, day, hours)}
                  onClick={() => {
                    setPickedRow(day.row);
                    setCostDate(cal?.iso ?? "");
                    setCostDaily(String(day.daily));
                    setCostBatch(day.batch);
                  }}
                >
                  <span className="chip-month">{cal ? cal.month : "Date"}</span>
                  <span className="chip-day">{cal?.day ?? "—"}</span>
                  <span className="chip-week">{cal?.weekday ?? ""}</span>
                  <span className="chip-daily">{compactMoney(day.daily)}</span>
                  <span className="chip-total">Total {money(day.total)}</span>
                  {hours ? <span className="chip-hours">{hours}</span> : null}
                  {day.batch ? <span className="chip-batch">{day.batch}</span> : null}
                </button>
              );
            })}
          </div>
        )}
        <input type="date" aria-label="Cost date" value={costDate} onChange={(event) => { setPickedRow(0); setCostDate(event.target.value); }} />
        <input aria-label="Daily cost" placeholder="Daily cost" inputMode="decimal" value={costDaily} onChange={(event) => { setPickedRow(0); setCostDaily(event.target.value); }} />
        <input aria-label="Batch note" placeholder="Batch note" value={costBatch} onChange={(event) => { setPickedRow(0); setCostBatch(event.target.value); }} />
        <button type="submit">Write Cost</button>
      </form>
      <section className="tool-bound" aria-label="Day and export">
        <h2>Day and export</h2>
        <div className="tool-row">
          <label>Day row <input value={dayRow} onChange={(event) => setDayRow(event.target.value)} /></label>
          <button type="button" onClick={writeDay}>Clip Slide Metres</button>
          <button type="button" onClick={buildReport}>Morning Report</button>
          <button type="button" onClick={downloadXlsx}>Export XLSX</button>
          <button type="button" onClick={() => window.print()}>Print / PDF</button>
        </div>
        {report ? <div className="report" dangerouslySetInnerHTML={{ __html: report }} /> : null}
      </section>
      <MailSettings onEmail={props.onEmailReport} onStatus={props.onStatus} />
      <section className="tool-bound" aria-label="FieldCap">
        <h2>FieldCap</h2>
        <div className="tool-row">
          <button type="button" onClick={() => void pollFieldCap()}>Poll FieldCap</button>
          <button type="button" onClick={() => void checkFieldCapFields()}>Check FieldCap Fields</button>
        </div>
        <p className="cost-empty">Poll entities: {ODATA_ENTITIES.join(", ")}. The poll is read-only.</p>
      </section>
      <fieldset className="tool-bound">
        <legend>Make this snapshot the job record</legend>
        {CUTOVER_CHECKS.map((item) => (
          <label key={item.id}>
            <input type="checkbox" checked={checked.includes(item.id)} onChange={(event) => setChecked(event.target.checked ? [...checked, item.id] : checked.filter((id) => id !== item.id))} />
            {item.label}
          </label>
        ))}
        <button type="button" disabled={!readyForNative(checked) || props.authority === "native"} onClick={() => void props.onAuthority("native")}>Accept Snapshot as the Record</button>
      </fieldset>
    </section>,
    dock,
    )}
    </>
  );
}

interface CorridorShot {
  url: string;
  kind: string;
  caption: string;
}

function captureCorridor(book: SnapshotBook | null): CorridorShot | null {
  if (!book) return null;
  const canvas = document.createElement("canvas");
  const picture = renderCorridor(canvas, new SheetBook(book, []), { scale: 2 });
  if (!picture) return null;
  return { url: canvas.toDataURL("image/png"), kind: picture.section.toLowerCase(), caption: picture.caption };
}

function chipTitle(cal: { weekday: string; month: string; day: string; year: string } | null, day: CostDay, hours: string): string {
  const when = cal ? `${cal.weekday} ${cal.month} ${cal.day}, ${cal.year}` : "Cost day";
  return [when, money(day.daily), `Total ${money(day.total)}`, day.batch, hours].filter((part) => part !== "").join(" · ");
}

function calendarFromSerial(serial: number): { month: string; day: string; weekday: string; year: string; iso: string } | null {
  if (!Number.isFinite(serial) || serial <= 0) return null;
  const date = new Date(Math.round(serial - 25569) * 86_400_000);
  if (Number.isNaN(date.getTime())) return null;
  const iso = date.toISOString().slice(0, 10);
  return {
    month: date.toLocaleString("en-US", { month: "short", timeZone: "UTC" }).toUpperCase(),
    day: String(date.getUTCDate()),
    weekday: date.toLocaleString("en-US", { weekday: "short", timeZone: "UTC" }),
    year: String(date.getUTCFullYear()),
    iso,
  };
}

function money(amount: number): string {
  return amount.toLocaleString("en-US", { style: "currency", currency: "USD" });
}

function compactMoney(amount: number): string {
  return amount.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
}

function chipHours(book: SnapshotBook | null, serial: number): string {
  if (!book) return "";
  for (let row = 4; row <= 22; row += 1) {
    const day = readNumber(book, "Data", `C${row}`);
    if (day == null || Math.round(day) !== serial) continue;
    const bit = readNumber(book, "Data", `P${row}`);
    const motor = readNumber(book, "Data", `Q${row}`);
    return `Bit ${bit ?? 0} h · Motor ${motor ?? 0} h`;
  }
  return "";
}

function text(book: SnapshotBook | null, sheet: string, a1: string): string {
  if (!book) return "";
  const value = readNumber(book, sheet, a1);
  return value == null ? "" : String(value);
}

/** The Data day row whose depth window holds the report, and its bit and motor hours (P and Q). */
function dayHours(book: SnapshotBook, serial: number): string {
  for (let row = 4; row <= 22; row += 1) {
    const day = readNumber(book, "Data", `C${row}`);
    if (day == null || Math.round(day) !== serial) continue;
    const bit = readNumber(book, "Data", `P${row}`);
    const motor = readNumber(book, "Data", `Q${row}`);
    return ` — bit ${bit ?? 0} h, motor ${motor ?? 0} h`;
  }
  return "";
}

function numberOrBlank(value: FormDataEntryValue | null): number | null {
  const number = Number(value);
  return value == null || value === "" || !Number.isFinite(number) ? null : number;
}
