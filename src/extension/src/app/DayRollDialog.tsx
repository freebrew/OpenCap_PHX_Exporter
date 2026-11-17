import { useState, type FormEvent } from "react";
import { checkDayRoll, formatNumber, type DayRollContext, type DayRollInputs } from "../../../packages/workbook-commands/src/index";

interface DayRollDialogProps {
  context: DayRollContext;
  onApply: (inputs: DayRollInputs) => string | null;
  onCancel: () => void;
}

/** Data!H4:H22 period picked: End Depth and Pason running totals, like frmDayHours. */
export function DayRollDialog(props: DayRollDialogProps) {
  const { context } = props;
  const [inputs, setInputs] = useState<DayRollInputs>({ endDepth: "", bitTotal: "", circTotal: "" });
  const [applyError, setApplyError] = useState<string | null>(null);
  const empty = !inputs.endDepth.trim() && !inputs.bitTotal.trim() && !inputs.circTotal.trim();
  const check = checkDayRoll(context, inputs);
  const status = applyError ?? (empty ? "Enter End Depth and Pason running totals." : check.ok ? "Ready to apply End Depth and hour deltas." : check.error);
  const isError = applyError != null || (!empty && !check.ok);
  const delta = (value: number | undefined) => (empty || value == null ? "-" : formatNumber(value, 4));

  function update(field: keyof DayRollInputs, value: string) {
    setApplyError(null);
    setInputs({ ...inputs, [field]: value });
  }

  function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!check.ok) {
      setApplyError(check.error);
      return;
    }
    setApplyError(props.onApply(inputs));
  }

  return (
    <div className="modal-backdrop">
      <form className="modal day-roll" role="dialog" aria-modal="true" aria-label="Day Roll - Depth and Hours" onSubmit={submit}>
        <h1>Day Roll - Depth and Hours</h1>
        <p>BHA # {context.bha} &middot; Period: {context.period} (row {context.row})</p>
        <dl>
          <dt>Start Depth</dt>
          <dd>{context.startDepth == null ? "(none)" : formatNumber(context.startDepth, 2)}</dd>
          <dt>Prior Bit hrs</dt>
          <dd>{formatNumber(context.priorBit, 2)}</dd>
          <dt>Prior Circ hrs</dt>
          <dd>{formatNumber(context.priorCirc, 2)}</dd>
          <dt>Slide time</dt>
          <dd>{formatNumber(context.slideHours, 2)} hrs</dd>
        </dl>
        <label>End Depth <input autoFocus inputMode="decimal" value={inputs.endDepth} onChange={(event) => update("endDepth", event.target.value)} /></label>
        <label>Pason Bit Total <input inputMode="decimal" value={inputs.bitTotal} onChange={(event) => update("bitTotal", event.target.value)} /></label>
        <span className="delta">Bit delta {delta(check.bitDelta)}</span>
        <label>Pason Circ Total <input inputMode="decimal" value={inputs.circTotal} onChange={(event) => update("circTotal", event.target.value)} /></label>
        <span className="delta">Circ delta {delta(check.circDelta)}</span>
        <p className={isError ? "status-line error" : "status-line"} role="status">{status}</p>
        <div className="tool-row">
          <button type="submit">Apply</button>
          <button type="button" onClick={props.onCancel}>Cancel</button>
        </div>
      </form>
    </div>
  );
}
