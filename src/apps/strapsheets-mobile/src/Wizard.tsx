// The entry wizard, as deployed in Entry.html: pick a tool, then one screen per schema step in
// steps[] order. Text is ALL CAPS as typed, every measure opens the number pad and may be left
// blank, Enter moves on, and choices are full-width buttons. Ranges warn; they never block.

import { useEffect, useRef, useState, type KeyboardEvent, type ReactNode } from "react";
import {
  choiceOptions,
  entryLines,
  findFamily,
  liveValue,
  normalizeEntry,
  prefillFromItem,
  setCell,
  stepProblem,
  stepWarning,
  threadParts,
  validateEntry,
  type BoxPinStep,
  type Cells,
  type EntryFamily,
  type EntryPayload,
  type EntryStep,
  type InventoryItem,
  type MeasureRange,
  type ThreadsStep,
} from "../../../packages/strapsheets-core/src/index";
import { Diagram } from "./Diagram";
import { letterForStep } from "./diagramData";
import { TOOL_DIAGRAMS } from "./diagrams.gen";
import { SerialPicker } from "./SerialPicker";

interface WizardProps {
  families: EntryFamily[];
  whichOptions: string[];
  threads: string[];
  items: InventoryItem[];
  /** Normalised serials that already have a strap sheet on this job. */
  measured: Set<string>;
  history: EntryPayload[];
  ranges: MeasureRange[];
  onSave: (entry: EntryPayload) => void;
}

type Phase = "pick" | "steps" | "confirm" | "done";

const AUTO_ADVANCE_MS = 200;

export function Wizard({ families, whichOptions, threads, items, measured, history, ranges, onSave }: WizardProps) {
  const [phase, setPhase] = useState<Phase>("pick");
  const [family, setFamily] = useState<EntryFamily | null>(null);
  const [cells, setCells] = useState<Cells>({});
  const [index, setIndex] = useState(0);
  const [error, setError] = useState("");
  const advanceTimer = useRef<number | null>(null);

  const step = family?.steps[index] ?? null;

  function pick(which: string) {
    const next = findFamily(families, which);
    if (!next) return;
    setFamily(next);
    setCells({});
    setIndex(0);
    setError("");
    setPhase("steps");
  }

  function next(from: Cells = cells) {
    if (advanceTimer.current != null) window.clearTimeout(advanceTimer.current);
    advanceTimer.current = null;
    if (!family || !step) return;
    const problem = stepProblem(step, from);
    if (problem) {
      setError(problem);
      return;
    }
    setError("");
    if (index < family.steps.length - 1) setIndex(index + 1);
    else setPhase("confirm");
  }

  /** A completed button step moves on by itself, after a beat so the selection is seen. */
  function advance(nextCells: Cells) {
    setCells(nextCells);
    if (advanceTimer.current != null) window.clearTimeout(advanceTimer.current);
    advanceTimer.current = window.setTimeout(() => next(nextCells), AUTO_ADVANCE_MS);
  }

  function back() {
    if (advanceTimer.current != null) window.clearTimeout(advanceTimer.current);
    advanceTimer.current = null;
    setError("");
    if (phase === "confirm") setPhase("steps");
    else if (index > 0) setIndex(index - 1);
    else setPhase("pick");
  }

  function save() {
    if (!family) return;
    const entry = normalizeEntry(family, { which: family.which, cells });
    const problems = validateEntry(family, entry);
    if (problems.length > 0) {
      setError(problems.map((problem) => problem.message).join(" "));
      return;
    }
    onSave(entry);
    setPhase("done");
  }

  if (phase === "pick") {
    return (
      <section className="wizard pick">
        <label className="step-label">Tool type</label>
        <div className="choices tool-list">
          {whichOptions.map((which) => (
            <button key={which} type="button" className={family?.which === which ? "selected" : ""} onClick={() => pick(which)}>
              {which}
            </button>
          ))}
        </div>
      </section>
    );
  }

  if (!family) return null;

  if (phase === "done") {
    return (
      <section className="wizard done-screen">
        <p className="done">Saved {family.which} on this phone.</p>
        <button type="button" className="primary" onClick={() => pick(family.which)}>Another {family.which}</button>
        <button type="button" onClick={() => setPhase("pick")}>Other tool</button>
      </section>
    );
  }

  if (phase === "confirm") {
    return (
      <section className="wizard">
        <h2>{family.which}</h2>
        <FamilyDiagram which={family.which} />
        <dl className="confirm">
          {entryLines(family, cells).map((line) => (
            <div key={line.label}>
              <dt>{line.label}</dt>
              <dd>{line.value}</dd>
            </div>
          ))}
        </dl>
        {error && <p className="error">{error}</p>}
        <div className="nav">
          <button type="button" onClick={back}>Back</button>
          <button type="button" className="primary" onClick={save}>Save</button>
        </div>
      </section>
    );
  }

  if (!step) return null;
  const warning = stepWarning(family, step, cells, ranges);
  const diagram = TOOL_DIAGRAMS[family.which];

  return (
    <section className="wizard">
      <div className="progress">{family.which} · {index + 1} of {family.steps.length}</div>
      {diagram && <Diagram diagram={diagram} letter={letterForStep(step, diagram)} />}
      <div className="step-body">
        <StepView
          key={`${family.which}-${step.id}`}
          family={family}
          step={step}
          cells={cells}
          threads={threads}
          items={items}
          measured={measured}
          onCells={setCells}
          onAdvance={advance}
          onPrefill={(item) => setCells(prefillFromItem(family, cells, item, history))}
          onEnter={() => next()}
        />
      </div>
      {warning && <p className="warning">Outside the usual range ({warning === "below" ? "low" : "high"}). Check it, then carry on.</p>}
      {error && <p className="error">{error}</p>}
      <div className="nav">
        <button type="button" onClick={back}>Back</button>
        <button type="button" className="primary" onClick={() => next()}>{index < family.steps.length - 1 ? "Next" : "Review"}</button>
      </div>
    </section>
  );
}

function FamilyDiagram({ which }: { which: string }) {
  const diagram = TOOL_DIAGRAMS[which];
  return diagram ? <Diagram diagram={diagram} letter="" /> : null;
}

interface StepViewProps {
  family: EntryFamily;
  step: EntryStep;
  cells: Cells;
  threads: string[];
  items: InventoryItem[];
  measured: Set<string>;
  onCells: (cells: Cells) => void;
  /** Set the cells and move on: the step is complete. */
  onAdvance: (cells: Cells) => void;
  onPrefill: (item: InventoryItem) => void;
  onEnter: () => void;
}

function StepView({ family, step, cells, threads, items, measured, onCells, onAdvance, onPrefill, onEnter }: StepViewProps) {
  switch (step.kind) {
    case "text":
      return (
        <>
          <TypedField label={step.label} value={cells[step.col] ?? ""} mode="text" onChange={(value) => onCells(setCell(cells, step.col, liveValue(step, value)))} onEnter={onEnter} />
          <div className="hint">Letters and numbers OK</div>
          {step.id === "serial" && items.length > 0 && <SerialPicker items={items} measured={measured} query={cells[step.col] ?? ""} onPick={onPrefill} />}
        </>
      );
    case "decimal":
      return (
        <>
          <TypedField label={step.label} value={cells[step.col] ?? ""} mode="decimal" onChange={(value) => onCells(setCell(cells, step.col, liveValue(step, value)))} onEnter={onEnter} />
          <div className="hint">Number pad · leave blank to skip</div>
        </>
      );
    case "choice":
      return (
        <>
          <label className="step-label">{step.label}</label>
          <Buttons options={choiceOptions(family, step)} value={cells[step.col] ?? ""} onPick={(value) => onAdvance(setCell(cells, step.col, value))} />
        </>
      );
    case "boxPin":
      return (
        <BoxPin
          step={step}
          options={family.boxPinOptions.length > 0 ? family.boxPinOptions : ["PIN", "BOX"]}
          cells={cells}
          onCells={(next) => (next[step.cols.upper] && next[step.cols.lower] ? onAdvance(next) : onCells(next))}
        />
      );;
    case "threads":
      return <Threads step={step} threads={threads} cells={cells} onCells={onCells} onAdvance={onAdvance} onEnter={onEnter} />;
    default: {
      const never: never = step;
      return never;
    }
  }
}

function TypedField({ label, value, mode, onChange, onEnter }: { label: string; value: string; mode: "text" | "decimal"; onChange: (value: string) => void; onEnter: () => void }) {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => ref.current?.focus(), []);
  const keyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key !== "Enter") return;
    event.preventDefault();
    onEnter();
  };
  return (
    <>
      <label className="step-label" htmlFor="field">{label}</label>
      {mode === "text" ? (
        <input id="field" ref={ref} type="text" inputMode="text" autoCapitalize="characters" autoComplete="off" autoCorrect="off" spellCheck={false} enterKeyHint="next" value={value} onChange={(event) => onChange(event.target.value)} onKeyDown={keyDown} />
      ) : (
        <input id="field" ref={ref} type="text" inputMode="decimal" pattern="[0-9]*[.]?[0-9]*" autoComplete="off" enterKeyHint="next" value={value} onChange={(event) => onChange(event.target.value)} onKeyDown={keyDown} />
      )}
    </>
  );
}

function Buttons({ options, value, onPick, grid = false, children }: { options: string[]; value: string; onPick: (value: string) => void; grid?: boolean; children?: ReactNode }) {
  return (
    <div className={grid ? "choices grid" : "choices"}>
      {options.map((option) => (
        <button key={option} type="button" className={option === value ? "selected" : ""} onClick={() => onPick(option)}>
          {grid ? <ThreadLabel option={option} /> : option}
        </button>
      ))}
      {children}
    </div>
  );
}

/** A thread's second name (NC equivalent or maker) goes on its own line. */
function ThreadLabel({ option }: { option: string }) {
  const { name, alias } = threadParts(option);
  return (
    <>
      <span>{name}</span>
      {alias && <span className="alt">[{alias}]</span>}
    </>
  );
}

function BoxPin({ step, options, cells, onCells }: { step: BoxPinStep; options: string[]; cells: Cells; onCells: (cells: Cells) => void }) {
  return (
    <>
      <label className="step-label">{step.label}</label>
      <div className="pair-row">
        <div className="sub">Upper</div>
        <div className="pair"><Buttons options={options} value={cells[step.cols.upper] ?? ""} onPick={(value) => onCells(setCell(cells, step.cols.upper, value))} /></div>
      </div>
      <div className="pair-row">
        <div className="sub">Lower</div>
        <div className="pair"><Buttons options={options} value={cells[step.cols.lower] ?? ""} onPick={(value) => onCells(setCell(cells, step.cols.lower, value))} /></div>
      </div>
    </>
  );
}

function Threads({ step, threads, cells, onCells, onAdvance, onEnter }: { step: ThreadsStep; threads: string[]; cells: Cells; onCells: (cells: Cells) => void; onAdvance: (cells: Cells) => void; onEnter: () => void }) {
  // Both ends picked from the buttons moves on; an end typed in Other waits for Next or Enter.
  const pick = (col: number, value: string) => {
    const next = setCell(cells, col, value);
    const bothListed = [step.cols.upper, step.cols.lower].every((end) => threads.includes(next[end] ?? ""));
    if (bothListed) onAdvance(next);
    else onCells(next);
  };
  const row = (label: string, col: number, last: boolean) => {
    const current = cells[col] ?? "";
    const listed = threads.includes(current);
    return (
      <div className="thread-row">
        <div className="sub">{label}</div>
        <Buttons grid options={threads} value={listed ? current : ""} onPick={(value) => pick(col, value)}>
          <input
          type="text"
          className="thread-other"
          placeholder="Other"
          autoCapitalize="characters"
          autoComplete="off"
          enterKeyHint={last ? "next" : "done"}
          value={listed ? "" : current}
          onChange={(event) => onCells(setCell(cells, col, event.target.value.toUpperCase()))}
          onKeyDown={(event) => {
            if (event.key === "Enter" && last) {
              event.preventDefault();
              onEnter();
            }
          }}
          />
        </Buttons>
      </div>
    );
  };
  return (
    <>
      <label className="step-label">{step.label}</label>
      {row("Upper Thread Type", step.cols.upper, false)}
      {row("Lower Thread Type", step.cols.lower, true)}
    </>
  );
}
