import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from "react";
import type { SnapshotBook } from "../../../packages/workbook-commands/src/cells";
import { SheetBook } from "../../../packages/workbook-commands/src/sheetBook";
import { shownOffsets } from "../../../packages/wells-3d/src/wellFilter";
import { shadowBoxOf } from "../../../packages/wells-3d/src/shadowBox";
import { buildWellScene, type OffsetPath, type WellReport, type WellScene } from "../../../packages/wells-3d/src/wellScene";
import type { HoverPick } from "../../../packages/wells-3d/src/wellPick";
import type { WorkbookHost } from "../workbook/facade";
import type { WellsIn, WellsOut } from "../wells/wellsMessages";

interface WellsViewProps {
  host: WorkbookHost | null;
  open: boolean;
  active: boolean;
}

interface Tip {
  x: number;
  y: number;
  hit: HoverPick;
}

const SF_LIMIT: Record<string, number | null> = { any: null, "1.5": 1.5, "2": 2, "5": 5 };
const C2C_LIMIT: Record<string, number | null> = { any: null, "50": 50, "100": 100, "300": 300 };

function bearingText(bearing: number): string {
  return String(bearing).padStart(3, "0");
}

/** The 3D well picture. It stays mounted, and hidden, so the worker can draw it before it is opened. */
export function WellsView(props: WellsViewProps) {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const frameRef = useRef<HTMLDivElement | null>(null);
  const workerRef = useRef<Worker | null>(null);
  const hashRef = useRef<string | null>(null);
  const dragging = useRef(false);
  const pointerAt = useRef({ x: 0, y: 0 });
  const [drilled, setDrilled] = useState(true);
  const [showAll, setShowAll] = useState(false);
  const [site, setSite] = useState("all");
  const [sfChoice, setSfChoice] = useState("any");
  const [c2cChoice, setC2cChoice] = useState("any");
  const [scene, setScene] = useState<WellScene | null>(null);
  const [badge, setBadge] = useState("");
  const [badgeOk, setBadgeOk] = useState(true);
  const [missing, setMissing] = useState<string[]>([]);
  const [hasPlan, setHasPlan] = useState(false);
  const [gridStep, setGridStep] = useState(0);
  const [tip, setTip] = useState<Tip | null>(null);
  const [fault, setFault] = useState("");
  const filterRef = useRef({ drilled: true, showAll: false, site: "all", sfChoice: "any", c2cChoice: "any" });
  filterRef.current = { drilled, showAll, site, sfChoice, c2cChoice };

  function post(message: WellsIn): void {
    workerRef.current?.postMessage(message);
  }

  function sendOptions(next: WellScene | null): void {
    const filter = filterRef.current;
    const shown = next
      ? shownOffsets(next.offsets, {
        allOffsets: filter.showAll,
        site: next.report.stored ? filter.site : "all",
        maxSf: SF_LIMIT[filter.sfChoice] ?? null,
        maxC2c: C2C_LIMIT[filter.c2cChoice] ?? null,
      })
      : [];
    post({ type: "options", drilled: filter.drilled, shown });
  }

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return undefined;
    const offscreen = canvas.transferControlToOffscreen();
    const worker = new Worker(new URL("../wells/wells.worker.ts", import.meta.url), { type: "module" });
    workerRef.current = worker;
    const rect = canvas.getBoundingClientRect();
    worker.postMessage({
      type: "init",
      canvas: offscreen,
      width: Math.max(1, rect.width),
      height: Math.max(1, rect.height),
      dpr: Math.min(window.devicePixelRatio || 1, 2),
      forceWebGL: false,
    }, [offscreen]);
    worker.onmessage = (event: MessageEvent<WellsOut>) => {
      const message = event.data;
      if (message.type === "hover") setTip(message.hit ? { x: pointerAt.current.x, y: pointerAt.current.y, hit: message.hit } : null);
      if (message.type === "error") setFault(message.message);
    };
    const frame = frameRef.current;
    const resize = () => {
      const box = canvas.getBoundingClientRect();
      worker.postMessage({
        type: "resize",
        width: Math.max(1, box.width),
        height: Math.max(1, box.height),
        dpr: Math.min(window.devicePixelRatio || 1, 2),
      });
    };
    const observer = new ResizeObserver(resize);
    if (frame) observer.observe(frame);
    const onWheel = (event: WheelEvent) => {
      if (frame?.dataset.active !== "true") return;
      event.preventDefault();
      worker.postMessage({ type: "pointer", phase: "wheel", x: 0, y: 0, button: 0, shift: event.shiftKey, deltaY: event.deltaY });
    };
    frame?.addEventListener("wheel", onWheel, { passive: false });
    return () => {
      observer.disconnect();
      frame?.removeEventListener("wheel", onWheel);
      worker.terminate();
      workerRef.current = null;
    };
  }, []);

  useEffect(() => {
    sendOptions(scene);
  }, [scene, drilled, showAll, site, sfChoice, c2cChoice]);

  useEffect(() => {
    if (!props.active) return;
    const box = canvasRef.current?.getBoundingClientRect();
    if (!box) return;
    post({ type: "resize", width: Math.max(1, box.width), height: Math.max(1, box.height), dpr: Math.min(window.devicePixelRatio || 1, 2) });
  }, [props.active]);

  useEffect(() => {
    const host = props.host;
    if (!host || !props.open) {
      if (hashRef.current !== null) {
        hashRef.current = null;
        setScene(null);
        setHasPlan(false);
        setGridStep(0);
        setBadge("");
        setMissing([]);
        post({ type: "scene", scene: null });
        sendOptions(null);
      }
      return undefined;
    }
    let timer = 0;
    const sync = () => {
      const snap = host.snapshot();
      if (!snap) return;
      let scene;
      try {
        scene = buildWellScene(new SheetBook(snap as SnapshotBook, []));
      } catch (error) {
        setFault(error instanceof Error ? error.message : String(error));
        return;
      }
      const hash = scene?.hash ?? "";
      if (hash === hashRef.current) return;
      hashRef.current = hash;
      setScene(scene);
      setHasPlan(scene != null);
      setGridStep(scene ? shadowBoxOf(scene)?.step ?? 0 : 0);
      setBadge(scene?.badge ?? "");
      setBadgeOk(scene?.badgeOk ?? true);
      setMissing(scene?.missing ?? []);
      if (scene?.report.stored) {
        const sites = new Set(scene.offsets.map((offset) => offset.site).filter((name) => name !== ""));
        const chosen = filterRef.current.site;
        if (chosen !== "all" && chosen !== "pad" && !sites.has(chosen)) setSite("all");
      } else if (filterRef.current.site !== "all") setSite("all");
      post({ type: "scene", scene });
      sendOptions(scene);
    };
    sync();
    const stop = host.onDirty(() => {
      window.clearTimeout(timer);
      timer = window.setTimeout(sync, 400);
    });
    return () => {
      window.clearTimeout(timer);
      stop();
    };
  }, [props.host, props.open]);

  function pointer(phase: "down" | "move" | "up", event: ReactPointerEvent<HTMLDivElement>): void {
    const box = frameRef.current?.getBoundingClientRect();
    if (!box) return;
    post({
      type: "pointer",
      phase,
      x: event.clientX - box.left,
      y: event.clientY - box.top,
      button: event.button,
      shift: event.shiftKey,
      deltaY: 0,
    });
  }

  return (
    <div ref={frameRef} className="wells-view" data-active={props.active ? "true" : "false"} aria-hidden={props.active ? undefined : true}>
      <canvas ref={canvasRef} className="wells-canvas" />
      <div
        className="wells-hit"
        onPointerDown={(event) => {
          dragging.current = true;
          setTip(null);
          event.currentTarget.setPointerCapture(event.pointerId);
          pointer("down", event);
        }}
        onPointerMove={(event) => {
          const box = frameRef.current?.getBoundingClientRect();
          if (!box) return;
          const x = event.clientX - box.left;
          const y = event.clientY - box.top;
          pointerAt.current = { x, y };
          if (dragging.current) {
            pointer("move", event);
            return;
          }
          post({ type: "hover", x, y });
        }}
        onPointerUp={(event) => {
          dragging.current = false;
          pointer("up", event);
        }}
        onContextMenu={(event) => event.preventDefault()}
      />
      {props.active ? (
        <>
          <div className="wells-legend">
            <p><i className="swatch plan" /> Planned Path</p>
            <p><i className="swatch drilled" /> Drilled Path</p>
            <p><i className="swatch station" /> Station</p>
            <p><i className="swatch target" /> Target</p>
            <p><i className="swatch red" /> Below 1.5</p>
            <p><i className="swatch yellow" /> Below 2.0</p>
            <p><i className="swatch green" /> 2.0 and Above</p>
            <p><i className="swatch shadow" /> Shadow of the Active Well</p>
            <p><i className="swatch formation" /> Formation</p>
            <p className="wells-axes">{gridStep > 0 ? `East, North, Up. Grid ${gridStep.toLocaleString("en-US")} m` : "East, North, Up"}</p>
          </div>
          <div className="wells-controls">
            <label><input type="checkbox" checked={drilled} onChange={(event) => setDrilled(event.target.checked)} /> Drilled Path</label>
            <label><input type="checkbox" checked={showAll} onChange={(event) => setShowAll(event.target.checked)} /> Show All Offset Wells</label>
            <FilterSelect label="Site" value={site} onChange={setSite} options={siteOptions(scene)} />
            <FilterSelect label="Lowest SF" value={sfChoice} onChange={setSfChoice} options={[
              ["any", "Any"],
              ["1.5", "Below 1.5"],
              ["2", "Below 2.0"],
              ["5", "Below 5.0"],
            ]} />
            <FilterSelect label="Closest C2C" value={c2cChoice} onChange={setC2cChoice} options={[
              ["any", "Any"],
              ["50", "Within 50 m"],
              ["100", "Within 100 m"],
              ["300", "Within 300 m"],
            ]} />
            {scene ? <p className="wells-count">{shownCount(scene, showAll, site, sfChoice, c2cChoice)} of {scene.offsets.length} Wells Shown</p> : null}
            <button type="button" onClick={() => post({ type: "reset" })}>Reset View</button>
          </div>
          {hasPlan ? null : <p className="wells-empty">Import a well plan to draw the paths.</p>}
          <div className="wells-status">
            {badge ? <p className={badgeOk ? "wells-badge" : "wells-badge bad"}>{badge}</p> : null}
            {fault ? <p className="wells-badge bad">{fault}</p> : null}
            <ReportBlock report={scene?.report ?? null} />
          </div>
          {missing.length > 0 ? (
            <div className="wells-missing">
              <p>No Path in Report</p>
              <ul>{missing.map((name) => <li key={name}>{name}</li>)}</ul>
            </div>
          ) : null}
          {tip && tip.hit.name ? <TipView tip={tip} offset={scene?.offsets.find((offset) => offset.name === tip.hit.name) ?? null} /> : null}
        </>
      ) : null}
    </div>
  );
}

function siteOptions(scene: WellScene | null): Array<[string, string]> {
  const options: Array<[string, string]> = [["all", "All Sites"]];
  if (!scene?.report.stored) return options;
  options.push(["pad", "Pad Wells"]);
  const sites = [...new Set(scene.offsets.map((offset) => offset.site).filter((name) => name !== ""))];
  sites.sort((a, b) => a.localeCompare(b));
  for (const name of sites) options.push([name, name]);
  return options;
}

function shownCount(scene: WellScene, showAll: boolean, site: string, sfChoice: string, c2cChoice: string): number {
  return shownOffsets(scene.offsets, {
    allOffsets: showAll,
    site: scene.report.stored ? site : "all",
    maxSf: SF_LIMIT[sfChoice] ?? null,
    maxC2c: C2C_LIMIT[c2cChoice] ?? null,
  }).length;
}

function FilterSelect(props: { label: string; value: string; onChange: (value: string) => void; options: Array<[string, string]> }) {
  return (
    <label className="wells-filter">{props.label}
      <select aria-label={props.label} value={props.value} onChange={(event) => props.onChange(event.target.value)}>
        {props.options.map(([value, label]) => <option key={value} value={value}>{label}</option>)}
      </select>
    </label>
  );
}

function ReportBlock(props: { report: WellReport | null }) {
  const report = props.report;
  if (!report?.stored) return <div className="wells-report"><p>Run Import AC again to store the report header</p></div>;
  return (
    <div className="wells-report">
      <p>Reference Well {report.referenceWell}</p>
      <p>Design {report.design}</p>
      <p>Coordinates Relative To {report.coordinatesRelativeTo}</p>
      <p>North {report.north}</p>
      <p>TVD Reference {report.tvdReference}</p>
      <p>Offset TVD Reference {report.offsetTvdReference}</p>
      <p>Report Date {report.reportDate}</p>
      <p className={report.verdictOk ? undefined : "bad"}>{report.verdict}</p>
    </div>
  );
}

function TipView(props: { tip: Tip; offset: OffsetPath | null }) {
  const hit = props.tip.hit;
  const wellhead = props.offset?.wellhead;
  return (
    <div className="wells-tip" style={{ left: props.tip.x + 14, top: props.tip.y + 14 }}>
      <strong>{hit.name}</strong>
      {props.offset ? <p>{props.offset.site !== "" ? `Site ${props.offset.site}${props.offset.onPad ? " (Pad)" : ""}` : "Site Not in Report"}</p> : null}
      {props.offset ? <p>{wellhead ? `Wellhead ${wellhead.distance.toFixed(2)} m at ${bearingText(wellhead.bearing)} from Active Well` : "Wellhead Not in Report"}</p> : null}
      <p>Ref MD {hit.refMd.toFixed(2)} m</p>
      <p>TVD {hit.tvd.toFixed(2)} m</p>
      {hit.c2c != null ? <p>C2C {hit.c2c.toFixed(2)} m</p> : null}
      {hit.sf != null ? <p>SF {hit.sf.toFixed(3)}</p> : null}
      <p>Measured Here {hit.measuredM.toFixed(2)} m</p>
    </div>
  );
}
