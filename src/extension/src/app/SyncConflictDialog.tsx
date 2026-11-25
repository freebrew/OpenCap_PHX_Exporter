import { useState } from "react";
import { defaultConflictChoice, type SyncConflict, type VisibleMerge } from "@phx/xlsx-bridge";

interface SyncConflictDialogProps {
  fileName: string;
  jobLabel?: string;
  merge: VisibleMerge;
  preferBrowserMachine: boolean;
  onApply: (picks: Record<string, "excel" | "browser">) => void;
  onCancel: () => void;
}

function conflictKey(conflict: SyncConflict): string {
  return `${conflict.sheet}!${conflict.a1}`;
}

function showCell(kind: string, value: string): string {
  if (kind === "blank" || value === "") return "(blank)";
  return value;
}

/** Excel changed on disk. Pick a side for each cell both shifts edited. */
export function SyncConflictDialog(props: SyncConflictDialogProps) {
  const [picks, setPicks] = useState<Record<string, "excel" | "browser">>(() => {
    const initial: Record<string, "excel" | "browser"> = {};
    for (const conflict of props.merge.conflicts) {
      initial[conflictKey(conflict)] = defaultConflictChoice(conflict, props.preferBrowserMachine);
    }
    return initial;
  });

  function setAll(choice: "excel" | "browser") {
    const next: Record<string, "excel" | "browser"> = {};
    for (const conflict of props.merge.conflicts) next[conflictKey(conflict)] = choice;
    setPicks(next);
  }

  const { merge } = props;
  return (
    <div className="modal-backdrop">
      <form
        className="modal sync-conflicts"
        role="dialog"
        aria-modal="true"
        aria-label="Sync with Excel"
        onSubmit={(event) => {
          event.preventDefault();
          props.onApply(picks);
        }}
      >
        <h1>{props.fileName} changed in Excel</h1>
        {props.jobLabel ? <p>{props.jobLabel}</p> : null}
        <p>Setup and Slidesheet stay from the browser (one-way Browser → Excel). Derived Data tables rebuild automatically, and the newest BHA # owns Data H3, C22, and the day rows (H4:Q22). Only conflicting same-BHA Data, Cont DI, and DL Tracker user cells appear below.</p>
        <p>
          {merge.fromExcel.length} from Excel, {merge.keepBrowser.length} kept in the browser, {merge.conflicts.length} need a choice.
          {merge.converged.length > 0 ? ` ${merge.converged.length} already match.` : ""}
        </p>
        <p>Browser values stay here and go out on the next Export handoff. Values from Excel are from the last saved workbook.</p>
        {merge.conflicts.length > 0 ? (
          <table>
            <thead>
              <tr>
                <th>Cell</th>
                <th>Baseline</th>
                <th>Excel</th>
                <th>Browser</th>
                <th>Use</th>
              </tr>
            </thead>
            <tbody>
              {merge.conflicts.map((conflict) => {
                const key = conflictKey(conflict);
                const pick = picks[key] ?? "excel";
                return (
                  <tr key={key}>
                    <td>{key}</td>
                    <td>{showCell(conflict.base.kind, conflict.base.value)}</td>
                    <td>{showCell(conflict.excel.kind, conflict.excel.value)}</td>
                    <td>{showCell(conflict.browser.kind, conflict.browser.value)}</td>
                    <td>
                      <button
                        type="button"
                        onClick={() => setPicks({ ...picks, [key]: pick === "excel" ? "browser" : "excel" })}
                      >
                        {pick === "excel" ? "Excel" : "Browser"}
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        ) : null}
        <div className="tool-row">
          {merge.conflicts.length > 0 ? (
            <>
              <button type="button" onClick={() => setAll("excel")}>All Excel</button>
              <button type="button" onClick={() => setAll("browser")}>All Browser</button>
            </>
          ) : null}
          <button type="submit">Merge into this shift</button>
          <button type="button" onClick={props.onCancel}>Not now</button>
        </div>
      </form>
    </div>
  );
}
