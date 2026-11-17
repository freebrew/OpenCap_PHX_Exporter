import { useState, type FormEvent, type ReactNode } from "react";
import {
  checkEntry,
  codeChoices,
  commandList,
  dateText,
  depthText,
  pageKey,
  parseDateTime,
  refreshCommand,
  refreshFormIJ,
  settingLabel,
  timeText,
  type DlEntryForm,
  type DlEntryStart,
  type SheetBook,
} from "../../../packages/workbook-commands/src/index";

interface DlEntryDialogProps {
  book: SheetBook;
  start: DlEntryStart;
  /** Writes the row; returns an error to show, or null when saved. */
  onSave: (vals: Partial<Record<number, string | number | null>>) => string | null;
  onCancel: () => void;
}

/** frmDLEntry: one new downlink row, copied from the last entry. */
export function DlEntryDialog(props: DlEntryDialogProps) {
  const { book, start } = props;
  const [form, setForm] = useState<DlEntryForm>(start.form);
  const [at, setAt] = useState<Date>(() => parseDateTime(start.form.date, start.form.time) ?? new Date());
  const [status, setStatus] = useState<{ text: string; error: boolean }>({ text: start.status, error: false });
  const setting = settingLabel(book, start, form);
  const commands = commandList(book, start.sheet, pageKey(form.page));
  const codes = codeChoices(form, start.high).choices;
  const id = (name: string) => `dl-${name}`;

  function change(field: keyof DlEntryForm, value: string) {
    const next = { ...form, [field]: value };
    if (field === "page" || field === "cmd") setForm(refreshCommand(book, start, next));
    else if (field === "pass") setForm({ ...refreshFormIJ(book, start, next), code: codeChoices(next, start.high).code });
    else {
      if (field === "date" || field === "time") {
        const parsed = parseDateTime(next.date, next.time);
        if (parsed) setAt(parsed);
      }
      setForm(next);
    }
  }

  function showAt(next: Date) {
    setAt(next);
    setForm({ ...form, date: dateText(next), time: timeText(next) });
  }

  function bumpDepth(step: number) {
    const n = Number(form.depth.trim());
    if (form.depth.trim() === "" || !Number.isFinite(n)) return;
    setForm({ ...form, depth: depthText(n + step) });
  }

  function bumpDays(days: number) {
    const next = new Date(at.getTime());
    next.setDate(next.getDate() + days);
    showAt(next);
  }

  function save(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const check = checkEntry(form);
    if (!check.ok) {
      setStatus({ text: check.error, error: true });
      return;
    }
    if (check.confirm && !window.confirm(check.confirm)) return;
    const error = props.onSave(check.vals);
    if (error) setStatus({ text: `Save failed: ${error}`, error: true });
  }

  const text = (field: keyof DlEntryForm, label: string, width: number, extra?: ReactNode) => (
    <div className="dl-field">
      <label htmlFor={id(field)}>{label}</label>
      <span>
        <input id={id(field)} style={{ width }} value={form[field]} onChange={(event) => change(field, event.target.value)} />
        {extra}
      </span>
    </div>
  );

  const combo = (field: keyof DlEntryForm, label: string, width: number, items: string[]) => (
    <div className="dl-field">
      <label htmlFor={id(field)}>{label}</label>
      <span>
        <input id={id(field)} list={id(`${field}-list`)} style={{ width }} value={form[field]} onChange={(event) => change(field, event.target.value)} />
        <datalist id={id(`${field}-list`)}>{items.map((item) => <option key={item} value={item} />)}</datalist>
      </span>
    </div>
  );

  const spin = (label: string, up: () => void, down: () => void) => (
    <span className="spin" role="group" aria-label={label}>
      <button type="button" tabIndex={-1} aria-label={`${label} up`} onClick={up}>&#9650;</button>
      <button type="button" tabIndex={-1} aria-label={`${label} down`} onClick={down}>&#9660;</button>
    </span>
  );

  return (
    <div className="modal-backdrop">
      <form className="modal dl-entry" role="dialog" aria-modal="true" aria-label={`DL Tracker - new entry (row ${start.row})`} onSubmit={save} onKeyDown={(event) => { if (event.key === "Escape") props.onCancel(); }}>
        <h1>DL Tracker - new entry (row {start.row})</h1>
        <div className="dl-grid">
          <div className="dl-col">
            {text("run", "Run #", 50)}
            {combo("page", "DL Page", 50, start.lists.pages)}
            <div className="dl-field">
              <label htmlFor={id("cmd")}>Dnlk Cmd.</label>
              <span>
                <input id={id("cmd")} autoFocus list={id("cmd-list")} style={{ width: 70 }} value={form.cmd} onChange={(event) => change("cmd", event.target.value)} />
                <datalist id={id("cmd-list")}>{commands.map((command) => <option key={command.key} value={command.key} label={command.setting} />)}</datalist>
              </span>
            </div>
          </div>
          <div className="dl-col">
            {text("depth", "Bit Depth (m)", 80, spin("Bit depth", () => bumpDepth(1), () => bumpDepth(-1)))}
            {text("date", "Date", 80, spin("Date", () => bumpDays(1), () => bumpDays(-1)))}
            {text("time", "Time", 80, spin("Time", () => showAt(new Date(at.getTime() + 60_000)), () => showAt(new Date(at.getTime() - 60_000))))}
          </div>
        </div>
        <div className="dl-field dl-wide">
          <label>D/L Setting</label>
          <div className={`dl-setting ${setting.tone}`} role="note" aria-label="D/L Setting">{setting.text}</div>
        </div>
        <div className="dl-grid">
          <div className="dl-col">
            {text("i", "TF / INC", 70)}
            {text("j", "SR% / AZI", 70)}
            {combo("period", "Bit Period", 70, start.lists.periods)}
            {combo("type", "Dnlk Type", 70, start.lists.types)}
            {combo("bottom", "ON/OFF Bottom", 70, start.lists.bottoms)}
            {text("low", "Low Flow / Rpm", 70)}
          </div>
          <div className="dl-col">
            {combo("code", "DLNK_PD", 80, codes)}
            {text("mon", "DL Monitor %", 80)}
            <div className="dl-field">
              <label htmlFor={id("pass")}>Pass / Fail</label>
              <span>
                <select id={id("pass")} style={{ width: 86 }} value={form.pass} onChange={(event) => change("pass", event.target.value)}>
                  <option value="PASS">PASS</option>
                  <option value="FAIL">FAIL</option>
                </select>
              </span>
            </div>
            {text("ref", "Ref. Flow / Rpm", 80)}
            {text("drop", "% Drop", 80)}
            {text("rpm", "TD RPM", 80)}
          </div>
        </div>
        <div className="dl-field dl-wide">
          <label htmlFor={id("note")}>Comments</label>
          <input id={id("note")} value={form.note} onChange={(event) => change("note", event.target.value)} />
        </div>
        <p className={status.error ? "status-line error" : "status-line"} role="status">{status.text}</p>
        <div className="tool-row dl-buttons">
          <button type="submit">Save Entry</button>
          <button type="button" onClick={props.onCancel}>Cancel</button>
        </div>
      </form>
    </div>
  );
}
