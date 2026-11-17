import { axisLabel, fixed2, type GaugeModel } from "../../../packages/workbook-commands/src/index";

const FS_HDR = 12;
const FS_VAL = 11;
const FS_AXIS = 10;
const INK = "#282828";
const GRID = "#B4B4B4";
const AXIS = "#1E1E1E";
const PLAN = "#005A5A";
const ACT = "#C00000";
const PTB = "#005ABE";
const BAND_IN = "#C8EBD2";
const BAND_OUT = "#FFD7D7";
const BAND_EDGE = "#238241";
const BAND_OUT_INK = "#B42828";

interface PlanGaugeSvgProps {
  model: GaugeModel;
  /** Slidesheet!AA1:AC7 in points, as the VBA lays it out. */
  width: number;
  height: number;
}

function clamp(value: number, lo: number, hi: number): number {
  return value < lo ? lo : value > hi ? hi : value;
}

function Label(props: { x: number; y: number; size: number; color: string; bold?: boolean; text: string; center?: number }) {
  return (
    <text
      x={props.center != null ? props.x + props.center / 2 : props.x}
      y={props.y}
      fontFamily="Consolas, monospace"
      fontSize={props.size}
      fontWeight={props.bold ? 700 : 400}
      fill={props.color}
      dominantBaseline="text-before-edge"
      textAnchor={props.center != null ? "middle" : "start"}
      xmlSpace="preserve"
    >
      {props.text}
    </text>
  );
}

/** MDL_PlanGauge.DrawGauge: dial on the left, read-outs on the right. */
export function PlanGaugeSvg(props: PlanGaugeSvgProps) {
  const { model } = props;
  const w = props.width;
  const h = props.height;
  const textW = 72;
  const gap = 3;
  const dialLeft = 1;
  let r = (h - 4) / 2;
  if (dialLeft + 2 * r > w - textW - gap - 1) r = (w - textW - gap - 1 - dialLeft) / 2;
  if (r < 18) r = 18;
  const cx = dialLeft + r;
  const cy = h / 2;
  let textX = w - textW - 1;
  if (textX < dialLeft + 2 * r + gap) textX = dialLeft + 2 * r + gap;
  const box = { l: 1, t: 1, r: Math.max(textX - 1, dialLeft + 2 * r + 1), b: h - 1 };
  const scale = model.scale > 0 ? model.scale : 1;
  const grav = model.mode === "GRAV";
  const aw = FS_AXIS + 4;

  function band() {
    if (!model.hasData || !model.hasBand || model.scale <= 0) return null;
    const hi = Math.max(model.yTop, model.yBot);
    const lo = Math.min(model.yTop, model.yBot);
    const xL = cx - r;
    const stripW = 2 * r;
    if (stripW < 4) return null;
    const topEdge = cy - r;
    const botEdge = cy + r;
    const syHi = cy - r * (hi / scale);
    const syLo = cy - r * (lo / scale);
    const gTop = clamp(syHi, topEdge, botEdge);
    const gBot = clamp(syLo, topEdge, botEdge);
    const rects: Array<{ key: string; y: number; height: number; fill: string }> = [];
    if (gTop > topEdge + 1) rects.push({ key: "out-u", y: topEdge, height: gTop - topEdge, fill: BAND_OUT });
    if (Math.abs(gBot - gTop) > 1) rects.push({ key: "in", y: gTop, height: gBot - gTop, fill: BAND_IN });
    if (botEdge > gBot + 1) rects.push({ key: "out-d", y: gBot, height: botEdge - gBot, fill: BAND_OUT });
    return (
      <g>
        {rects.filter((rect) => rect.height >= 1).map((rect) => <rect key={rect.key} x={xL} y={rect.y} width={stripW} height={rect.height} fill={rect.fill} fillOpacity={0.65} />)}
        {syHi > topEdge + 0.5 && syHi < botEdge - 0.5 ? <line x1={xL} x2={xL + stripW} y1={syHi} y2={syHi} stroke={BAND_EDGE} strokeWidth={1.5} /> : null}
        {syLo > topEdge + 0.5 && syLo < botEdge - 0.5 ? <line x1={xL} x2={xL + stripW} y1={syLo} y2={syLo} stroke={BAND_EDGE} strokeWidth={1.5} /> : null}
      </g>
    );
  }

  function marker(tag: string, vx: number, vy: number, color: string, hollow: boolean) {
    if (Math.hypot(vx, vy) <= 0.005) return <circle key={tag} cx={cx} cy={cy} r={6} fill="none" stroke={color} strokeWidth={2} />;
    const margin = 6;
    const dx = clamp(cx + r * (vx / scale), box.l + margin, box.r - margin);
    const dy = clamp(cy - r * (vy / scale), box.t + margin, box.b - margin);
    return (
      <g key={tag}>
        <line x1={cx} y1={cy} x2={dx} y2={dy} stroke={color} strokeWidth={hollow ? 1.5 : 2} strokeDasharray={hollow ? "4 3" : undefined} markerEnd={`url(#gauge-arrow-${tag})`} />
        <circle cx={dx} cy={dy} r={4.5} fill={hollow ? "#FFFFFF" : color} stroke={hollow ? color : "none"} strokeWidth={hollow ? 2 : 0} />
      </g>
    );
  }

  const rowH = FS_VAL + 1;
  const y0 = 1;
  const magA = Math.hypot(model.ax, model.ay);
  const magB = Math.hypot(model.bx, model.by);
  const aOut = model.hasBand && (model.ay > model.yTop || model.ay < model.yBot);
  const bOut = model.hasBand && model.showPtb && (model.by > model.yTop || model.by < model.yBot);
  const yPtb = y0 + 4 * rowH + 2;
  const total = grav ? "TOT" : "PRP";

  return (
    <svg className="plan-gauge" viewBox={`0 0 ${w} ${h}`} width="100%" height="100%" role="img" aria-label="Plan proximity gauge">
      <defs>
        {[["A", ACT], ["B", PTB]].map(([tag, color]) => (
          <marker key={tag} id={`gauge-arrow-${tag}`} viewBox="0 0 6 6" refX={5} refY={3} markerWidth={4} markerHeight={4} orient="auto-start-reverse">
            <path d="M0,0 L6,3 L0,6 Z" fill={color} />
          </marker>
        ))}
      </defs>
      {band()}
      <circle cx={cx} cy={cy} r={r} fill="none" stroke={AXIS} strokeWidth={1.75} />
      <line x1={cx - r} y1={cy} x2={cx + r} y2={cy} stroke={AXIS} strokeWidth={0.75} strokeDasharray="3 2" />
      <line x1={cx} y1={cy - r} x2={cx} y2={cy + r} stroke={AXIS} strokeWidth={0.75} strokeDasharray="3 2" />
      <Label x={cx - aw / 2} y={cy - r - 1} center={aw} size={FS_AXIS} color={AXIS} bold text={grav ? "U" : "A"} />
      <Label x={cx - aw / 2} y={cy + r - aw + 2} center={aw} size={FS_AXIS} color={AXIS} bold text={grav ? "D" : "B"} />
      <Label x={cx - r + 1} y={cy - aw / 2} center={aw} size={FS_AXIS} color={AXIS} bold text="L" />
      <Label x={cx + r - aw + 1} y={cy - aw / 2} center={aw} size={FS_AXIS} color={AXIS} bold text="R" />
      <circle cx={cx} cy={cy} r={3} fill={PLAN} />
      {model.hasData && model.showPtb ? marker("B", model.bx, model.by, PTB, true) : null}
      {model.hasData ? marker("A", model.ax, model.ay, ACT, false) : null}
      {!model.hasData ? (
        <Label x={textX} y={2} size={FS_HDR} color={INK} bold text={model.caption} />
      ) : (
        <g>
          <Label x={textX} y={y0} size={FS_HDR} color={INK} bold text={grav ? "GRAVITY" : "MAGNETIC"} />
          <Label x={textX} y={y0 + rowH} size={FS_VAL} color={aOut ? BAND_OUT_INK : ACT} text={axisLabel(model.mode, true, model.ay) + (aOut ? " OUT" : "")} />
          <Label x={textX} y={y0 + 2 * rowH} size={FS_VAL} color={ACT} text={axisLabel(model.mode, false, model.ax)} />
          <Label x={textX} y={y0 + 3 * rowH} size={FS_VAL} color={INK} text={`${total} ${fixed2(magA)}m`} />
          {model.showPtb ? (
            <g>
              <Label x={textX} y={yPtb} size={FS_HDR} color={PTB} bold text="PTB" />
              <Label x={textX} y={yPtb + rowH} size={FS_VAL} color={bOut ? BAND_OUT_INK : PTB} text={axisLabel(model.mode, true, model.by) + (bOut ? " OUT" : "")} />
              <Label x={textX} y={yPtb + 2 * rowH} size={FS_VAL} color={PTB} text={axisLabel(model.mode, false, model.bx)} />
              <Label x={textX} y={yPtb + 3 * rowH} size={FS_VAL} color={INK} text={`${total} ${fixed2(magB)}m`} />
            </g>
          ) : (
            <g>
              <Label x={textX} y={yPtb} size={FS_HDR} color={GRID} bold text="PTB" />
              <Label x={textX} y={yPtb + rowH} size={FS_VAL} color={GRID} text="at survey" />
            </g>
          )}
        </g>
      )}
    </svg>
  );
}
