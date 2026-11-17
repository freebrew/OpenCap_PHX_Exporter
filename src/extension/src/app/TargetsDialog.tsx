import { useState, type FormEvent } from "react";
import { fixed2, TARGET_NAMES, type NamedTarget } from "../../../packages/workbook-commands/src/index";

export interface TargetEdit {
  planRow: number;
  name: string;
  md: number;
  inc: number;
  azm: number;
  tvd: number;
}

interface TargetsDialogProps {
  targets: NamedTarget[];
  /** Stores every row and rebuilds T2:Y5; returns the first store message, or null when saved. */
  onSave: (edits: TargetEdit[]) => string | null;
  onCancel: () => void;
}

interface Row {
  planRow: number;
  name: string;
  md: string;
  inc: string;
  azm: string;
  tvd: string;
}

const NOTE = "Saved targets replace the plan on the Slidesheet and Cont DI. They stay until the next plan import.";

/** NumOf: commas stripped, then VBA IsNumeric. */
function numOf(text: string): number | null {
  const s = text.replace(/,/g, "").trim();
  if (s === "") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

/** frmTargets: the named plan stations (KOP, BUILD, HEEL ...) as one editable table. */
export function TargetsDialog(props: TargetsDialogProps) {
  const [rows, setRows] = useState<Row[]>(() => props.targets.map((target) => ({
    planRow: target.planRow,
    name: target.name,
    md: fixed2(target.md),
    inc: fixed2(target.inc),
    azm: fixed2(target.azm),
    tvd: fixed2(target.tvd),
  })));
  const [note, setNote] = useState(props.targets.length < 1 ? "No named targets on this plan." : NOTE);

  function update(index: number, field: keyof Omit<Row, "planRow">, value: string) {
    setRows(rows.map((row, i) => (i === index ? { ...row, [field]: value } : row)));
  }

  function save(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (rows.length < 1) return;
    const edits: TargetEdit[] = [];
    for (const row of rows) {
      const md = numOf(row.md);
      const inc = numOf(row.inc);
      let azm = numOf(row.azm);
      const tvd = numOf(row.tvd);
      if (md == null || inc == null || azm == null || tvd == null) {
        setNote(`MD, INC, AZM and TVD have to be numbers (${row.name.trim()}).`);
        return;
      }
      if (!(TARGET_NAMES as readonly string[]).includes(row.name.trim().toUpperCase())) {
        setNote(`Name must be one of: ${TARGET_NAMES.join(", ")}`);
        return;
      }
      if (azm >= 360) azm -= 360;
      edits.push({ planRow: row.planRow, name: row.name, md, inc, azm, tvd });
    }
    const message = props.onSave(edits);
    if (message) setNote(message);
  }

  return (
    <div className="modal-backdrop">
      <form className="modal targets" role="dialog" aria-modal="true" aria-label="Targets" onSubmit={save} onKeyDown={(event) => { if (event.key === "Escape") props.onCancel(); }}>
        <h1>Targets</h1>
        <p className="targets-note" role="status">{note}</p>
        {rows.length > 0 ? (
          <table>
            <thead>
              <tr><th>Name</th><th>MD</th><th>INC</th><th>AZM</th><th>TVD</th></tr>
            </thead>
            <tbody>
              {rows.map((row, index) => (
                <tr key={row.planRow}>
                  <td><input aria-label={`Target ${index + 1} name`} value={row.name} onChange={(event) => update(index, "name", event.target.value)} /></td>
                  <td><input aria-label={`${row.name} MD`} value={row.md} onChange={(event) => update(index, "md", event.target.value)} /></td>
                  <td><input aria-label={`${row.name} INC`} value={row.inc} onChange={(event) => update(index, "inc", event.target.value)} /></td>
                  <td><input aria-label={`${row.name} AZM`} value={row.azm} onChange={(event) => update(index, "azm", event.target.value)} /></td>
                  <td><input aria-label={`${row.name} TVD`} value={row.tvd} onChange={(event) => update(index, "tvd", event.target.value)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : null}
        <div className="tool-row">
          {rows.length > 0 ? <button type="submit">Save</button> : null}
          <button type="button" onClick={props.onCancel}>Cancel</button>
        </div>
      </form>
    </div>
  );
}
