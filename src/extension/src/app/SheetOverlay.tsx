import { useEffect, useRef, useState, type ReactNode } from "react";
import type { MacroButton } from "@phx/xlsx-bridge";
import type { GaugeModel } from "../../../packages/workbook-commands/src/index";
import type { SheetAnchor, SheetPlacement, WorkbookHost } from "../workbook/facade";
import { PlanGaugeSvg } from "./PlanGaugeSvg";

/** Slidesheet!AA1:AC7, where MDL_PlanGauge draws. */
const GAUGE_ANCHOR: SheetAnchor = { col: 26, colOff: 0, row: 0, rowOff: 0, col2: 29, colOff2: 0, row2: 7, rowOff2: 0 };

/** Buttons the 35783 template carries. A workbook that lost one gets it back where the VBA puts it. */
const TEMPLATE_BUTTONS: MacroButton[] = [
  { sheet: "Slidesheet", id: "btnClearSlidesheetRanges", caption: "Clear Ranges", macro: "ClearSlidesheetRanges", anchor: "25, 0, 0, 0, 26, 0, 2, 0" },
  { sheet: "Slidesheet", id: "btnPipeTallyForm", caption: "Pipe Tally Calculator", macro: "ShowPipeTallyForm", anchor: "25, 0, 2, 0, 26, 0, 5, 0" },
  { sheet: "Slidesheet", id: "btnEditTargets", caption: "Targets", macro: "ShowTargetEditor", anchor: "25, 0, 5, 0, 26, 0, 7, 0" },
  { sheet: "DL Tracker", id: "btnDLTAdd", caption: "ADD ENTRY", macro: "DLT_AddEntry", anchor: "0, 6, 0, 3, 2, 3, 0, 38" },
  { sheet: "DL Tracker", id: "btnDLTClear", caption: "CLEAR", macro: "DLT_ClearEntries", anchor: "20, 283, 0, 3, 20, 423, 0, 38" },
];

/** "leftCol, leftOff, topRow, topOff, rightCol, rightOff, bottomRow, bottomOff" (0-based, px). */
export function parseAnchor(text: string): SheetAnchor | null {
  const parts = text.split(",").map((part) => Number(part.trim()));
  if (parts.length !== 8 || parts.some((part) => !Number.isFinite(part))) return null;
  const [col, colOff, row, rowOff, col2, colOff2, row2, rowOff2] = parts as [number, number, number, number, number, number, number, number];
  return { col, colOff, row, rowOff, col2, colOff2, row2, rowOff2 };
}

/** Workbook buttons on this sheet, plus any template button whose macro is missing. */
export function sheetButtons(buttons: MacroButton[], sheet: string): MacroButton[] {
  const own = buttons.filter((button) => button.sheet === sheet);
  const missing = TEMPLATE_BUTTONS.filter((template) => template.sheet === sheet && !own.some((button) => button.macro === template.macro));
  return [...own, ...missing];
}

interface SheetOverlayProps {
  host: WorkbookHost | null;
  sheet: string;
  buttons: MacroButton[];
  gauge: GaugeModel | null;
  busy: boolean;
  onMacro: (button: MacroButton) => void;
}

interface Placed {
  button: MacroButton;
  at: SheetPlacement;
}

export function SheetOverlay(props: SheetOverlayProps) {
  const { host, sheet } = props;
  const [placed, setPlaced] = useState<Placed[]>([]);
  const [gaugeAt, setGaugeAt] = useState<SheetPlacement | null>(null);
  const buttons = sheetButtons(props.buttons, sheet);
  const buttonKey = buttons.map((button) => `${button.macro}@${button.anchor}`).join("|");
  const latest = useRef(buttons);
  latest.current = buttons;

  useEffect(() => {
    if (!host) return undefined;
    const layout = () => {
      const next: Placed[] = [];
      for (const button of latest.current) {
        const anchor = parseAnchor(button.anchor);
        const at = anchor ? host.placeOnSheet(sheet, anchor) : null;
        if (at && at.width > 2 && at.height > 2) next.push({ button, at });
      }
      setPlaced(next);
      setGaugeAt(sheet === "Slidesheet" ? host.placeOnSheet(sheet, GAUGE_ANCHOR) : null);
    };
    layout();
    // The canvas sizes itself a frame after a sheet switch.
    const late = window.setTimeout(layout, 120);
    const stop = host.onViewChange(layout);
    return () => {
      window.clearTimeout(late);
      stop();
    };
  }, [host, sheet, buttonKey]);

  return (
    <div className="sheet-overlay" aria-label={`${sheet} sheet controls`}>
      {gaugeAt && props.gauge ? (
        <Clip at={gaugeAt}>
          <div className="plan-gauge-box" style={box(gaugeAt)}>
            <PlanGaugeSvg model={props.gauge} width={(gaugeAt.width / gaugeAt.zoom) * 0.75} height={(gaugeAt.height / gaugeAt.zoom) * 0.75} />
          </div>
        </Clip>
      ) : null}
      {placed.map(({ button, at }) => (
        <Clip key={`${button.id}-${button.macro}`} at={at}>
          <button
            type="button"
            className="sheet-button"
            title={button.macro}
            disabled={props.busy}
            style={{ ...box(at), fontSize: `${10 * at.zoom}pt` }}
            onClick={() => props.onMacro(button)}
          >
            {button.caption}
          </button>
        </Clip>
      ))}
    </div>
  );
}

function box(at: SheetPlacement) {
  return { left: at.left - at.clipLeft, top: at.top - at.clipTop, width: at.width, height: at.height };
}

/** Hides the part of a shape that has scrolled under the headers or frozen panes. */
function Clip(props: { at: SheetPlacement; children: ReactNode }) {
  return <div className="sheet-clip" style={{ left: props.at.clipLeft, top: props.at.clipTop }}>{props.children}</div>;
}
