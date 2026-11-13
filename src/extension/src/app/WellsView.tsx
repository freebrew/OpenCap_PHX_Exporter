import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent } from "react";
import type { SnapshotBook } from "../../../packages/workbook-commands/src/cells";
import { SheetBook } from "../../../packages/workbook-commands/src/sheetBook";
import { buildWellScene } from "../../../packages/wells-3d/src/wellScene";
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
  const [badge, setBadge] = useState("");
  const [badgeOk, setBadgeOk] = useState(true);
  const [missing, setMissing] = useState<string[]>([]);
  const [hasPlan, setHasPlan] = useState(false);
  const [tip, setTip] = useState<Tip | null>(null);
  const [fault, setFault] = useState("");

  function post(message: WellsIn): void {
    workerRef.current?.postMessage(message);
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
    post({ type: "options", drilled, allOffsets: showAll });
  }, [drilled, showAll]);

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
        setHasPlan(false);
        setBadge("");
        setMissing([]);
        post({ type: "scene", scene: null });
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
      setHasPlan(scene != null);
      setBadge(scene?.badge ?? "");
      setBadgeOk(scene?.badgeOk ?? true);
      setMissing(scene?.missing ?? []);
      post({ type: "scene", scene });
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
            <p><i className="swatch red" /> Below 1.5</p>
            <p><i className="swatch yellow" /> Below 2.0</p>
            <p><i className="swatch green" /> 2.0 and Above</p>
            <p className="wells-axes">East, North, Up</p>
          </div>
          <div className="wells-controls">
            <label><input type="checkbox" checked={drilled} onChange={(event) => setDrilled(event.target.checked)} /> Drilled Path</label>
            <label><input type="checkbox" checked={showAll} onChange={(event) => setShowAll(event.target.checked)} /> Show All Offset Wells</label>
            <button type="button" onClick={() => post({ type: "reset" })}>Reset View</button>
          </div>
          {hasPlan ? null : <p className="wells-empty">Import a well plan to draw the paths.</p>}
          {badge ? <p className={badgeOk ? "wells-badge" : "wells-badge bad"}>{badge}</p> : null}
          {fault ? <p className="wells-badge bad">{fault}</p> : null}
          {missing.length > 0 ? (
            <div className="wells-missing">
              <p>No Path in Report</p>
              <ul>{missing.map((name) => <li key={name}>{name}</li>)}</ul>
            </div>
          ) : null}
          {tip && tip.hit.name ? <TipView tip={tip} /> : null}
        </>
      ) : null}
    </div>
  );
}

function TipView(props: { tip: Tip }) {
  const hit = props.tip.hit;
  return (
    <div className="wells-tip" style={{ left: props.tip.x + 14, top: props.tip.y + 14 }}>
      <strong>{hit.name}</strong>
      <p>Ref MD {hit.refMd.toFixed(2)} m</p>
      <p>TVD {hit.tvd.toFixed(2)} m</p>
      {hit.c2c != null ? <p>C2C {hit.c2c.toFixed(2)} m</p> : null}
      {hit.sf != null ? <p>SF {hit.sf.toFixed(3)}</p> : null}
      <p>Measured Here {hit.measuredM.toFixed(2)} m</p>
    </div>
  );
}
