import { useEffect, useState } from "react";
import { fixed2, type TallyView } from "../../../packages/workbook-commands/src/index";

interface PipeTallyDialogProps {
  view: TallyView;
  status: string;
  busy: boolean;
  onImport: () => void;
  onClear: () => void;
  onAdd: (stands: string, split: string) => void;
  onRefresh: () => void;
  /** txtStands / txtSplit lost focus: settings go to C11 / D11 and the preview recalculates. */
  onSettings: (stands: string, split: string) => void;
  onClose: () => void;
}

/** frmPipeTally (modeless): CSV lengths in, stands and split, bit depths out to Slidesheet D. */
export function PipeTallyDialog(props: PipeTallyDialogProps) {
  const { view } = props;
  const [stands, setStands] = useState(view.stands);
  const [split, setSplit] = useState(view.split);
  const rows = Math.max(view.lengths.length, view.depths.length);

  useEffect(() => {
    setStands(view.stands);
    setSplit(view.split);
  }, [view.stands, view.split]);

  const settingsChanged = () => {
    if (stands !== view.stands || split !== view.split) props.onSettings(stands, split);
  };

  return (
    <section className="modal floating pipe-tally" role="dialog" aria-label="Pipe Tally Calculator">
      <h1>Pipe Tally Calculator</h1>
      <dl>
        <dt>Start depth</dt>
        <dd aria-label="Start depth">{view.start}</dd>
      </dl>
      <label>Stands <input aria-label="Stands" value={stands} onChange={(event) => setStands(event.target.value)} onBlur={settingsChanged} /></label>
      <label>Split <input aria-label="Split" value={split} onChange={(event) => setSplit(event.target.value)} onBlur={settingsChanged} /></label>
      <div className="tally-list" role="table" aria-label="Pipe tally preview">
        <div role="row" className="tally-head"><span role="columnheader">Length</span><span role="columnheader">Bit depth</span></div>
        {Array.from({ length: rows }, (_, i) => (
          <div role="row" key={i}>
            <span role="cell">{i < view.lengths.length ? fixed2(view.lengths[i]!) : ""}</span>
            <span role="cell">{i < view.depths.length ? fixed2(view.depths[i]!) : ""}</span>
          </div>
        ))}
      </div>
      <p className="delta" aria-label="Tally count">{view.lengths.length} joints / {view.depths.length} depths</p>
      <p className="status-line" role="status">{props.status}</p>
      <div className="tool-row">
        <button type="button" disabled={props.busy} onClick={props.onImport}>Import CSV</button>
        <button type="button" disabled={props.busy} onClick={props.onClear}>Clear</button>
        <button type="button" disabled={props.busy} onClick={() => props.onAdd(stands, split)}>Add</button>
        <button type="button" disabled={props.busy} onClick={props.onRefresh}>Refresh</button>
        <button type="button" onClick={props.onClose}>Close</button>
      </div>
    </section>
  );
}
