// The tool drawing with the active callout letter ringed. The letters are small against the whole
// tool, so the ring sits well outside the glyph and the rest of the drawing is dimmed.

import { useId, type CSSProperties } from "react";
import type { ToolDiagram } from "./diagramData";

const RING = "#1a73e8";

export function Diagram({ diagram, letter }: { diagram: ToolDiagram; letter: string }) {
  const maskId = useId();
  const { w, h } = diagram.size;
  const mark = letter ? diagram.marks.find((candidate) => candidate.letter === letter) : undefined;
  const cx = mark ? mark.x + mark.w / 2 : 0;
  const cy = mark ? mark.y + mark.h / 2 : 0;
  const radius = mark ? Math.max(mark.w, mark.h) * 1.5 + 10 : 0;
  return (
    <div className="diagram" style={{ "--ratio": h / w } as CSSProperties}>
      <div className="diagram-frame">
        <img src={diagram.file} alt="Tool measurement diagram" />
        <svg viewBox={`0 0 ${w} ${h}`} preserveAspectRatio="xMidYMid meet">
          {mark && (
            <>
              <defs>
                <mask id={maskId}>
                  <rect x={0} y={0} width={w} height={h} fill="#fff" />
                  <circle cx={cx} cy={cy} r={radius} fill="#000" />
                </mask>
              </defs>
              <rect x={0} y={0} width={w} height={h} fill="#fff" fillOpacity={0.55} mask={`url(#${maskId})`} />
              <circle cx={cx} cy={cy} r={radius} fill={RING} fillOpacity={0.08} />
              <circle className="mark-ring" cx={cx} cy={cy} r={radius} fill="none" stroke={RING} strokeWidth={5} />
            </>
          )}
        </svg>
      </div>
      {mark && <div className="hint">Highlighted: {letter} on the drawing</div>}
    </div>
  );
}
