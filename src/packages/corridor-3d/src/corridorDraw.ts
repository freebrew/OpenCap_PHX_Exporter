import { planAt } from "../../workbook-commands/src/planGauge";
import { cellText, numeric, type SheetBook } from "../../workbook-commands/src/sheetBook";
import { SS_SHEET } from "../../workbook-commands/src/slidesheet";
import { CANVAS_W, Painter, type Anchor, type Dash, type PaintContext, type Theme } from "./corridorPaint";
import { arcTvdGain, CorridorScene, loadFormations, niceStep, rad, solveArcCurvature } from "./corridorScene";

// Drawing half of MDL_CorridorImage: DrawRoom (tunnel), DrawShaftProfile (1:1 shadow box),
// formations, shadows, bullseye targets, and the live picture's section / day / metrics panels.

const CORRIDOR_H = 445;
const TYPE_SCALE_META = 1.2;
const LAT_VIEW_INC = 55;
const DEG = 180 / Math.PI;

const fmt = (v: number, digits: number, group = true) =>
  v.toLocaleString("en-US", { minimumFractionDigits: digits, maximumFractionDigits: digits, useGrouping: group });
const fmt0 = (v: number) => fmt(v, 0);

function signedM(v: number): string {
  if (Math.abs(v) < 0.001) return "0";
  return v > 0 ? `+${fmt(v, 1, false)}` : `\u2212${fmt(Math.abs(v), 1, false)}`;
}

/** Shadow-box yaw: world (n, e) to N' (into the box) and E' (left to right). */
export interface ShaftYaw {
  c: number;
  s: number;
  face: boolean;
  /** Degrees added to a world azimuth to get its direction in the yawed frame. */
  azOff: number;
}

const PLAIN_YAW: ShaftYaw = { c: 1, s: 0, face: false, azOff: 0 };

/**
 * ComputeShaftView's yaw for a hole heading azS (radians from grid north). Lateral puts the
 * hole on +E'; otherwise the bit faces the viewer. Facing is a 180 degree turn of both axes:
 * negating N' alone (as the VBA did) is a reflection and draws every turn the wrong way.
 */
export function shaftYaw(azS: number, lateral: boolean): ShaftYaw {
  if (lateral) return { c: Math.sin(azS), s: Math.cos(azS), face: false, azOff: 90 - azS * DEG };
  return { c: Math.cos(azS), s: -Math.sin(azS), face: true, azOff: 180 - azS * DEG };
}

export function yawPoint(view: ShaftYaw, n: number, e: number): { n: number; e: number } {
  const sign = view.face ? -1 : 1;
  return { n: sign * (n * view.c - e * view.s), e: sign * (n * view.s + e * view.c) };
}

/**
 * Signed screen turn in degrees. Canvas Y points down, so positive is visually clockwise. The shaft
 * runs down the page, and a path heading down that turns clockwise bends toward the page's LEFT
 * (counter-clockwise bends right). A left turn of the hole (azimuth decreasing) has to read
 * positive here, a right turn negative.
 */
function signedScreenBend(view: ShaftYaw, a: Pt3, b: Pt3): number {
  const onScreen = (d: Pt3) => {
    const y = yawPoint(view, d.n, d.e);
    return Math.atan2(d.v - y.n * 0.4 + y.e * 0.18, y.e * 0.88 + y.n * 0.52);
  };
  const delta = onScreen(b) - onScreen(a);
  return Math.atan2(Math.sin(delta), Math.cos(delta)) * DEG;
}

/** Angle in degrees between two 3D directions as ShX/ShY draw them under this yaw. */
export function screenBend(view: ShaftYaw, a: Pt3, b: Pt3): number {
  return Math.abs(signedScreenBend(view, a, b));
}

const TURN_GAIN_DEG = 10;
/** A world turn smaller than this is noise; it does not decide which way the picture must bend. */
const TURN_MIN_DEG = 3;

/**
 * Facing-the-bit heading for the shaft. Facing the bit can put the eye in the plane of a turn,
 * which draws a curved day as a straight line (the shadows still show it). When another heading
 * shows at least TURN_GAIN_DEG more of the bend, use the one that shows the most, nearest the bit.
 *
 * Only a heading that bends the same way the hole turns is acceptable, the facing one included:
 * a left turn (azimuth decreasing) has to bend toward the page's left, a right turn toward the right.
 */
export function shaftFacing(endAz: number, start: Pt3, end: Pt3): number {
  const startAz = Math.atan2(start.e, start.n);
  const worldTurn = Math.atan2(Math.sin(endAz - startAz), Math.cos(endAz - startAz));
  const turnSign = Math.abs(worldTurn) * DEG < TURN_MIN_DEG ? 0 : Math.sign(worldTurn);
  // On a hole running down the page, clockwise on screen (positive) bends left.
  const wanted = -turnSign;
  const tries = Array.from({ length: 72 }, (_, k) => {
    const az = endAz + k * 5 * Math.PI / 180;
    const signedBend = signedScreenBend(shaftYaw(az, false), start, end);
    return { az, bend: Math.abs(signedBend), handed: Math.sign(signedBend), off: Math.min(k, 72 - k) };
  });
  const handed = turnSign === 0 ? tries : tries.filter((t) => t.handed === wanted);
  const atBit = turnSign === 0 || tries[0]!.handed === wanted ? tries[0]!.bend : 0;
  if (handed.length === 0) return endAz;
  const most = Math.max(...handed.map((t) => t.bend));
  // A facing view that visibly bends the wrong way is replaced even when the right-way view shows little.
  const wrongAtBit = turnSign !== 0 && tries[0]!.handed !== wanted && tries[0]!.bend >= TURN_MIN_DEG;
  if (!wrongAtBit && most < atBit + TURN_GAIN_DEG) return endAz;
  const near = handed.filter((t) => t.bend >= most - 3).sort((a, b) => a.off - b.off);
  return near[0]!.az;
}

/** LatScr: positive data lateral (right of plan) is the near / front side of the room. */
const latScr = (l: number) => -l;

/** Picture size in points: the corridor band, or just the room for a segment picture. */
export function corridorSize(scene: CorridorScene): { width: number; height: number } {
  if (scene.isExample) return { width: scene.vx + scene.vw + 28, height: scene.vy + scene.vh + 40 };
  return { width: CANVAS_W, height: CORRIDOR_H };
}

/** RenderCorridorPng(path, "corridor") onto a context already scaled to points. */
export function paintCorridor(ctx: PaintContext, scene: CorridorScene, theme: Theme = "dark"): void {
  const p = new Painter(ctx, theme);
  const size = corridorSize(scene);
  p.rect(0, 0, size.width, size.height, "FFFFFF", null, 0);
  p.typeScale = TYPE_SCALE_META;
  const room = new Room(scene, p);
  room.draw();
  if (!scene.isExample) {
    room.drawSection();
    room.drawDay();
    room.drawMetrics();
  }
}

/** A cell's printed text: the value through its number format, trailing colon dropped. */
function cellDisplay(book: SheetBook, sheet: string, a1: string): string {
  const match = /^([A-Z]+)(\d+)$/i.exec(a1);
  if (!match) return "";
  const col = match[1]!.toUpperCase().split("").reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0);
  const row = Number(match[2]);
  const value = book.value(sheet, row, col);
  let text: string;
  if (typeof value === "number") text = formatNumber(value, book.style(sheet, row, col)?.n?.pattern ?? "");
  else text = cellText(value);
  text = text.trim();
  return text.endsWith(":") ? text.slice(0, -1) : text;
}

/** The numeric part of an Excel format ("#,##0.00", 0.00"°", 0.0 "m"); dates fall back to the raw value. */
function formatNumber(value: number, pattern: string): string {
  const section = pattern.split(";")[0] ?? "";
  if (section === "" || /general/i.test(section)) return String(value);
  const bare = section.replace(/"[^"]*"/g, "").replace(/\\./g, "");
  if (/[dyhs]|m{1,2}(?![^[]*\])/i.test(bare.replace(/0|#|,|\.|%|\s/g, ""))) return String(value);
  const core = /[#0,]*\.?[#0]*/.exec(section.replace(/"[^"]*"/g, (m) => " ".repeat(m.length)))?.[0] ?? "";
  const start = section.indexOf(core);
  const digits = (core.split(".")[1] ?? "").length;
  const number = fmt(section.includes("%") ? value * 100 : value, digits, core.includes(","));
  const literal = (s: string) => s.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1").replace(/[_*]./g, "");
  return `${literal(section.slice(0, start))}${number}${literal(section.slice(start + core.length))}`;
}

class Room {
  // shadow-box frame (ComputeShaftView / DrawShaftProfile)
  private view: ShaftYaw = PLAIN_YAW;
  private shOx = 0;
  private shOy = 0;
  private shS = 1;
  private shNS = 1;
  private shVS = 1;
  private tgtV = 0;

  constructor(private readonly s: CorridorScene, private readonly p: Painter) {}

  private get book(): SheetBook {
    return this.s.book;
  }

  private ob(md: number, plot: number, lat: number): [number, number] {
    return this.s.ob(md, plot, lat);
  }

  private ln3(a: [number, number], b: [number, number], color: string, weight: number, dash: Dash): void {
    this.p.ln(a[0], a[1], b[0], b[1], color, weight, dash);
  }

  /** Quad: a face of the tunnel given as (md, plot, lat) corners. */
  private quad(corners: Array<[number, number, number]>, fill: string | null, line: string, weight: number): void {
    const pts = corners.map(([md, plot, lat]) => this.ob(md, plot, lat));
    this.p.polyFill(pts.map((q) => q[0]), pts.map((q) => q[1]), pts.length, fill, line, weight);
  }

  private txAt(at: [number, number], dx: number, dy: number, text: string, size: number, color: string, anchor: Anchor, bold = false): void {
    this.p.tx(at[0] + dx, at[1] + dy, text, size, color, anchor, bold);
  }

  draw(): void {
    if (this.s.shaft) {
      this.drawShaftProfile();
      return;
    }
    const s = this.s;
    const p = this.p;
    const { md0, md1, devHi, devLo, latBox, latTol, geoHalf } = s;
    p.tx(s.vx, 20, `WELLBORE CORRIDOR \u2014 ${s.section} to ${s.tgtName}  ${fmt0(md0)} to ${fmt0(md1)} mMD  \u00b7 24 h ${fmt0(md1 - md0)} m`, 12.5, "FFFFFF", "start", true);
    p.rect(s.vx, s.vy, s.vw, s.vh, "222222", "3A3A3A", 0.75);

    this.quad([[md0, devHi, latBox], [md1, devHi, latBox], [md1, devLo, latBox], [md0, devLo, latBox]], null, "4A4A4A", 0.75);
    this.quad([[md0, devHi, -latBox], [md1, devHi, -latBox], [md1, devLo, -latBox], [md0, devLo, -latBox]], null, "4A4A4A", 0.75);

    // Two-metre ticks on a sail day; a room that still holds the curve would stack 60 labels.
    const tick = devHi - devLo <= 40 ? 2 : niceStep(devHi - devLo, 12);
    for (let dV = Math.floor(devHi / tick) * tick; dV >= devLo; dV -= tick) {
      const isZero = Math.abs(dV) < 0.001;
      this.ln3(this.ob(md0, dV, latBox), this.ob(md1, dV, latBox), isZero ? "555555" : "3A3A3A", 0.7, isZero ? "dash" : "dot");
      this.ln3(this.ob(md0, dV, -latBox), this.ob(md1, dV, -latBox), isZero ? "444444" : "333333", 0.6, isZero ? "dash" : "dot");
      this.ln3(this.ob(md0, dV, -latBox), this.ob(md0, dV, latBox), "3A3A3A", 0.6, "solid");
      this.txAt(this.ob(md0, dV, -latBox), -5, 3, signedM(dV), 8.5, "A8A8A8", "end");
    }
    this.drawFormationsInTunnel();

    if (s.useSailBands) {
      this.drawGeoCorridorRibbon(latBox, "4CAF7A", 1.15);
      this.drawGeoCorridorRibbon(-latBox, "3D6B54", 1);
      const endGeo = s.waypointTvd(md1) ?? s.wpTvd;
      this.txAt(this.ob(md1, s.plotY(endGeo - geoHalf), latBox), -4, -4, `geo \u00b1${fmt(geoHalf, 2)} m`, 8, "4CAF7A", "end");
    }

    this.quad([[md0, devLo, -latBox], [md1, devLo, -latBox], [md1, devLo, latBox], [md0, devLo, latBox]], null, "4A4A4A", 0.75);
    if (s.useSailBands) this.quad([[md0, devLo, -latTol], [md1, devLo, -latTol], [md1, devLo, latTol], [md0, devLo, latTol]], "1E3A2F", "4CAF7A", 1);
    this.ln3(this.ob(md0, devLo, 0), this.ob(md1, devLo, 0), "4CAF7A", 0.7, "dash");
    if (s.useSailBands) {
      const latLbl = fmt(latTol, 2);
      this.txAt(this.ob(md1, devLo, -latTol), 4, 3, `${latLbl} R`, 8, "4CAF7A", "start");
      this.txAt(this.ob(md1, devLo, 0), 4, 3, "plan", 8, "4CAF7A", "start");
      this.txAt(this.ob(md1, devLo, latTol), 4, 3, `${latLbl} L`, 8, "4CAF7A", "start");
    } else {
      this.txAt(this.ob(md1, devLo, 0), 4, 3, "plan", 8, "A8A8A8", "start");
    }

    this.quad([[md0, devHi, -latBox], [md0, devHi, latBox], [md0, devLo, latBox], [md0, devLo, -latBox]], null, "4A4A4A", 0.75);
    if (s.useSailBands) this.quad([[md0, geoHalf, -latTol], [md0, geoHalf, latTol], [md0, -geoHalf, latTol], [md0, -geoHalf, -latTol]], "1E3A2F", "4CAF7A", 1.2);

    this.ln3(this.ob(md0, devHi, -latBox), this.ob(md1, devHi, -latBox), "4A4A4A", 0.75, "solid");
    this.ln3(this.ob(md1, devHi, -latBox), this.ob(md1, devHi, latBox), "4A4A4A", 0.75, "solid");
    this.ln3(this.ob(md1, devHi, -latBox), this.ob(md1, devLo, -latBox), "4A4A4A", 0.75, "solid");
    this.ln3(this.ob(md0, devLo, -latBox), this.ob(md1, devLo, -latBox), "4A4A4A", 0.75, "solid");

    // Back wall carries the only TVD shadow; the floor carries left / right from plan.
    const n = s.sCount;
    const back = s.sMd.map((md, i) => this.ob(md, s.plotY(s.sTvd[i]!), latBox));
    p.polyLine(back.map((q) => q[0]), back.map((q) => q[1]), n, "A8CBB8", 2.4, "dash");
    const floor = s.sMd.map((md, i) => this.ob(md, devLo, latScr(s.sLat[i]!)));
    p.polyLine(floor.map((q) => q[0]), floor.map((q) => q[1]), n, "B6DBC8", 3, "solid");
    const iLast = n - 1;
    const lastFloor = floor[iLast]!;
    p.dot(lastFloor[0], lastFloor[1], 3.4, "B6DBC8", "1A1A1A", 1);
    this.txAt(lastFloor, 0, 12, `${fmt(Math.abs(s.sLat[iLast]!), 1)} m ${s.sLat[iLast]! < 0 ? "L" : "R"}`, 8.5, "B6DBC8", "middle");

    const hole = s.sMd.map((md, i) => this.ob(md, s.plotY(s.sTvd[i]!), latScr(s.sLat[i]!)));
    const xs = hole.map((q) => q[0]);
    const ys = hole.map((q) => q[1]);
    p.polyLine(xs, ys, n, "FFFFFF", 2.8, "solid");
    this.drawBitProjectionInTunnel();
    this.drawPlanPathInRoom();
    this.drawWaypointMarks();

    let iWorstDev = 0;
    let iWorstLat = 0;
    for (let i = 0; i < n; i += 1) {
      const bust = s.useSailBands && (Math.abs(s.sDev[i]!) > geoHalf || Math.abs(s.sLat[i]!) > latTol);
      p.dot(xs[i]!, ys[i]!, 2.2, bust ? "FF5555" : "111111", null, 0);
      if (Math.abs(s.sDev[i]!) > Math.abs(s.sDev[iWorstDev]!)) iWorstDev = i;
      if (Math.abs(s.sLat[i]!) > Math.abs(s.sLat[iWorstLat]!)) iWorstLat = i;
    }
    if (s.useSailBands) {
      const [wx, wy] = this.ob(s.sMd[iWorstDev]!, s.plotY(s.sTvd[iWorstDev]!), latBox);
      p.ln(wx, wy - 5, wx + 14, wy - 18, "FF5555", 0.8, "solid");
      p.tx(wx + 17, wy - 16, `${fmt(s.sDev[iWorstDev]!, 2, false)} m high at ${fmt0(s.sMd[iWorstDev]!)}`, 8.5, "FF5555", "start");
      if (Math.abs(this.ob(s.sMd[iWorstLat]!, 0, 0)[0] - this.ob(s.sMd[iLast]!, 0, 0)[0]) > 40) {
        this.txAt(this.ob(s.sMd[iWorstLat]!, devLo, latScr(s.sLat[iWorstLat]!)), 0, 11, `${fmt(Math.abs(s.sLat[iWorstLat]!), 1)} m ${s.sLat[iWorstLat]! < 0 ? "L" : "R"}`, 8, "8AAB9A", "middle");
      }
    }

    // Forward projections, only when the last survey sits in this room.
    if (this.surveyInRoom()) {
      let drawLen = s.toGo;
      if (s.svyMd + drawLen > md1) drawLen = md1 - s.svyMd;
      if (drawLen < 1) drawLen = 1;
      const kap = solveArcCurvature(s.svyInc, s.toGo, s.wpTvd - s.svyTvd);
      const count = Math.floor(drawLen / 4) + 2;
      const rx: number[] = [];
      const ry: number[] = [];
      const hx: number[] = [];
      const hy: number[] = [];
      for (let k = 0; k < count; k += 1) {
        const a = Math.min(k * 4, drawLen);
        const md = s.svyMd + a;
        const latR = latScr(s.svyLat * (1 - a / s.toGo));
        const req = this.ob(md, s.plotY(s.svyTvd + arcTvdGain(rad(s.svyInc), kap, a)), latR);
        rx.push(req[0]);
        ry.push(req[1]);
        const latH = latScr(s.svyLat + a * Math.sin(rad(s.svyAzi - s.svyPlanAzi)));
        const held = this.ob(md, s.plotY(s.svyTvd + a * Math.cos(rad(s.svyInc))), latH);
        hx.push(held[0]);
        hy.push(held[1]);
      }
      p.polyLine(rx, ry, count, "FF5555", 2.4, "solid");
      p.polyLine(hx, hy, count, "E09A3D", 2.2, "dash");
      const [bx, by] = this.ob(s.svyMd, s.plotY(s.svyTvd), latScr(s.svyLat));
      p.dot(bx, by, 5, "FF5555", "1A1A1A", 1.4);
      p.tx(bx - 10, by - 11, "SVY", 10, "FF5555", "end", true);
    }

    if (s.wpMd <= md1 + 0.1) {
      const [ax, ay] = this.ob(s.wpMd, s.plotY(s.svyTvd + s.toGo * Math.cos(rad(s.svyInc))), latScr(s.holdLatEnd));
      p.dot(ax, ay, 4, "1A1A1A", "E09A3D", 2);
      if (s.useSailBands) p.tx(ax + 10, ay + 12, `${fmt(s.holdDevEnd, 2, false)} m high`, 8.5, "E09A3D", "start");
      else if (Math.abs(s.holdDevEnd) > 1) p.tx(ax + 10, ay + 12, `${fmt(Math.abs(s.holdDevEnd), 1)} m TVD if held`, 8.5, "E09A3D", "start");
    }

    // A segment picture stands in for a report day that covered its window.
    const day = s.isExample ? { d0: s.exA, d1: s.exB } : s.reportDayMd();
    const d0 = s.clampMd(day.d0);
    const d1 = s.clampMd(day.d1);
    const dy = s.vy + s.vh - 16;
    const x0 = this.ob(d0, 0, -latBox)[0];
    const x1 = this.ob(d1, 0, -latBox)[0];
    p.rect(x0, dy - 8, x1 - x0, 11, "1A3348", "5BA3D9", 0.75);
    p.tx((x0 + x1) / 2, dy, `24 h \u00b7 ${fmt0(d0)} \u2192 ${fmt0(d1)} m \u00b7 ${fmt0(d1 - d0)} m`, 8.5, "5BA3D9", "middle");

    p.tx(s.vx + 6, s.vy + 16, "back wall \u2014 TVD plot (0 = geo at window start), m", 8.5, "A8A8A8", "start");
    p.tx(s.vx + s.vw - 6, s.vy + s.vh - 44, "floor \u2014 left / right from plan, m", 8.5, "A8A8A8", "end");

    const ly = s.vy + s.vh + 15;
    const key = (x: number, color: string, weight: number, dash: Dash, label: string, size: number, labelColor = "FFFFFF") => {
      p.ln(s.vx + x, ly, s.vx + x + 16, ly, color, weight, dash);
      p.tx(s.vx + x + 20, ly + 4, label, size, labelColor, "start");
    };
    key(2, "FFFFFF", 2.6, "solid", "hole", 10);
    key(52, "5BA3D9", 2.2, "solid", "plan", 10);
    key(108, "9EC9B0", 2, "dash", "shadows", 10);
    key(178, "FF5555", 2.4, "solid", s.showBurr ? `BURR ${fmt(s.burr, 2)}\u00b0/30m` : `required ${fmt(s.reqInc, 2)}\u00b0`, 10);
    key(292, "E09A3D", 2.2, "dash", `if ${fmt(s.svyInc, 2)}\u00b0 held`, 10);
    key(392, "008C95", 2.8, "solid", "to PTB", 9, "008C95");
    p.tx(s.vx + s.vw, ly + 17, `MD \u00f7${fmt(1 / s.sxd, 1, false)} \u00b7 TVD \u00d7${fmt(s.sv, 0, false)} \u00b7 L/R \u00d7${fmt(s.kzx, 0, false)}`, 8.5, "A8A8A8", "end");
  }

  /** A live day keeps a survey up to 40 m past midnight; a segment holds only its own window. */
  private surveyInRoom(): boolean {
    const s = this.s;
    if (s.isExample) return s.svyMd >= s.exA - 0.5 && s.svyMd <= s.exB + 0.5;
    return s.svyMd >= s.md0 - 0.5 && s.svyMd <= s.md1 + 40;
  }

  private drawBitProjectionInTunnel(): void {
    const s = this.s;
    const count = s.ptbCount;
    if (count < 2 || (s.isExample && !this.surveyInRoom())) return;
    if (s.ptbMd[count - 1]! < s.md0 - 1 || s.ptbMd[0]! > s.md1 + 1) return;
    const pts = s.ptbMd.map((md, i) => {
      const lat = s.stationOffsets(md, s.ptbInc[i]!, s.ptbNs[i]!, s.ptbEw[i]!, s.ptbTvd[i]!).lat;
      return this.ob(md, s.plotY(s.ptbTvd[i]!), latScr(lat));
    });
    this.p.polyLine(pts.map((q) => q[0]), pts.map((q) => q[1]), count, "008C95", 3, "solid");
    const last = pts[count - 1]!;
    this.p.dot(last[0], last[1], 4.2, "FFFFFF", "008C95", 2);
    this.p.tx(last[0] + 8, last[1] - 7, "PTB", 9, "008C95", "start", true);
  }

  /** DrawPlanPathInRoom: the plan through the room, clipped where it leaves the TVD window. */
  private drawPlanPathInRoom(): void {
    const s = this.s;
    const step = 8;
    const count = Math.floor((s.md1 - s.md0) / step) + 2;
    let xs: number[] = [];
    let ys: number[] = [];
    let last: [number, number] | null = null;
    const flush = () => {
      if (xs.length >= 2) {
        this.p.polyLine(xs, ys, xs.length, "5BA3D9", 2.3, "solid");
        last = [xs[xs.length - 1]!, ys[ys.length - 1]!];
      }
      xs = [];
      ys = [];
    };
    for (let i = 0; i < count; i += 1) {
      const md = Math.min(s.md0 + i * step, s.md1);
      const plot = s.plotY(planAt(s.plan, md).v);
      if (s.plotInRoom(plot)) {
        const q = this.ob(md, plot, 0);
        xs.push(q[0]);
        ys.push(q[1]);
      } else {
        flush();
      }
      if (md >= s.md1) break;
    }
    flush();
    const end = last as [number, number] | null;
    if (!end) return;
    const name = cellDisplay(this.book, "Data", "C2") || "Plan";
    this.p.tx(end[0] + 4, end[1] - 6, name, 9, "5BA3D9", "start", true);
  }

  /** DrawGeoCorridorRibbon: geo +/- outline between waypoint TVDs on one wall. */
  private drawGeoCorridorRibbon(lat: number, line: string, weight: number): void {
    const s = this.s;
    const mdK: number[] = [];
    const tvdK: number[] = [];
    const g0 = s.waypointTvd(s.md0);
    if (g0 != null) {
      mdK.push(s.md0);
      tvdK.push(g0);
    }
    for (let i = 0; i < s.wMd.length; i += 1) {
      if (s.wMd[i]! > s.md0 + 0.05 && s.wMd[i]! < s.md1 - 0.05) {
        mdK.push(s.wMd[i]!);
        tvdK.push(s.wTvd[i]!);
      }
    }
    const g1 = s.waypointTvd(s.md1);
    if (g1 != null) {
      mdK.push(s.md1);
      tvdK.push(g1);
    }
    if (mdK.length < 2) return;
    const run = (offset: number, w: number, dash: Dash) => {
      const pts = mdK.map((md, i) => this.ob(md, s.plotY(tvdK[i]! + offset), lat));
      this.p.polyLine(pts.map((q) => q[0]), pts.map((q) => q[1]), pts.length, line, w, dash);
    };
    run(-s.geoHalf, weight + 0.55, "solid");
    run(s.geoHalf, weight + 0.55, "solid");
    run(0, 1, "dash");
  }

  /** DrawWaypointMarks: MD and Inc-to-next on the left wall, with an attitude tick. */
  private drawWaypointMarks(): void {
    const s = this.s;
    for (let i = 0; i < s.wMd.length; i += 1) {
      const md = s.wMd[i]!;
      if (md < s.md0 || md > s.md1) continue;
      const at = s.holeAtMdPlot(md);
      const plot = at ? at.plot : s.plotY(s.wTvd[i]!);
      const isNext = Math.abs(md - s.wpMd) < 0.001;
      const [wx, wy] = this.ob(md, plot, -s.latBox);
      this.p.dot(wx, wy, isNext ? 3.2 : 2.4, "4CAF7A", "1A1A1A", 1);
      if (s.wInc[i]! > 0.1) {
        const ang = rad(s.wInc[i]! - 90);
        this.p.ln(wx, wy, wx + 14 * Math.cos(ang), wy + 14 * Math.sin(ang), "4CAF7A", isNext ? 1.8 : 1.2, "solid");
        this.p.tx(wx, wy - 10, `${fmt(s.wInc[i]!, 2)}\u00b0`, 8, isNext ? "4CAF7A" : "8AAB9A", "middle", isNext);
      }
      const isHeel = s.userWpCount >= 1 && Math.abs(md - s.heelMd) < 0.01;
      this.p.tx(wx, s.vy + s.vh - 30, `${isHeel ? "HEEL " : ""}${fmt0(md)}`, 8.5, isNext || isHeel ? "4CAF7A" : "A8A8A8", "middle", isNext || isHeel);
    }
  }

  private drawFormationsInTunnel(): void {
    const s = this.s;
    const ink = "C4A574";
    let lastPlot = 1e30;
    for (const fm of loadFormations(this.book)) {
      const plot = s.plotY(fm.tvd);
      if (!s.plotInRoom(plot)) continue;
      this.ln3(this.ob(s.md0, plot, s.latBox), this.ob(s.md1, plot, s.latBox), ink, 1.15, "dash");
      this.ln3(this.ob(s.md0, plot, -s.latBox), this.ob(s.md1, plot, -s.latBox), ink, 1, "dash");
      const mdL = s.md0 + (s.md1 - s.md0) * 0.08;
      const mdR = s.md0 + (s.md1 - s.md0) * 0.92;
      const dy = Math.abs(plot - lastPlot) < 1.2 ? 11 : -4;
      const a = this.ob(mdL, plot, s.latBox);
      const b = this.ob(mdR, plot, s.latBox);
      this.p.wallLabel(a[0], a[1] + dy, b[0], b[1] + dy, fm.name, `TVD ${fmt0(fm.tvd)} m`, fm.lith, 9, ink);
      lastPlot = plot;
    }
  }

  // ---- the 1:1 shadow box ----------------------------------------------------

  private vwN(n: number, e: number): number {
    return yawPoint(this.view, n, e).n;
  }

  private vwE(n: number, e: number): number {
    return yawPoint(this.view, n, e).e;
  }

  private shX(n: number, e: number): number {
    return this.shOx + e * this.shS * 0.88 + n * this.shNS * 0.52;
  }

  private shY(n: number, e: number, v: number): number {
    return this.shOy + (v - this.tgtV) * this.shVS - n * this.shNS * 0.4 + e * this.shS * 0.18;
  }

  private sh(n: number, e: number, v: number): [number, number] {
    return [this.shX(n, e), this.shY(n, e, v)];
  }

  /** ComputeShaftView: yaw the frame from the drawn line at the end of the window. */
  private computeShaftView(plan: Pt3[], hole: Pt3[]): void {
    this.view = PLAIN_YAW;
    const line = plan.length >= 2 ? plan : hole.length >= 2 ? hole : null;
    if (!line) return;
    const end = line[line.length - 1]!;
    let dN = 0;
    let dE = 0;
    let dV = 0;
    for (let i = line.length - 2; i >= 0; i -= 1) {
      dN = end.n - line[i]!.n;
      dE = end.e - line[i]!.e;
      dV = end.v - line[i]!.v;
      const horiz = Math.hypot(dN, dE);
      if (horiz >= 8 || Math.hypot(horiz, dV) >= 30) break;
    }
    const horiz = Math.hypot(dN, dE);
    if (horiz < 1.5) return;
    const incEst = dV <= 0.01 ? 90 : Math.atan2(horiz, dV) * DEG;
    const endAz = Math.atan2(dE, dN);
    if (incEst >= LAT_VIEW_INC) {
      this.view = shaftYaw(endAz, true);
      return;
    }
    const first = line[0]!;
    const reach = line.find((pt) => Math.hypot(pt.n - first.n, pt.e - first.e, pt.v - first.v) >= 30) ?? end;
    const start = { n: reach.n - first.n, e: reach.e - first.e, v: reach.v - first.v };
    this.view = shaftYaw(shaftFacing(endAz, start, { n: dN, e: dE, v: dV }), false);
  }

  private yaw(pt: Pt3): Pt3 {
    return { n: this.vwN(pt.n, pt.e), e: this.vwE(pt.n, pt.e), v: pt.v };
  }

  /** CollectDayHole: as-drilled XYZ from day0 (interpolated) through day1. */
  private collectDayHole(day0: number, day1: number): Pt3[] {
    const s = this.s;
    const count = s.aMd.length;
    if (count < 2) return [];
    const out: Pt3[] = [];
    if (day0 < s.aMd[count - 1]! - 0.01) {
      const start = s.holeXyzAtMd(day0);
      if (start) out.push(start);
    }
    for (let i = 0; i < count; i += 1) {
      if (s.aMd[i]! > day0 + 0.01 && s.aMd[i]! <= day1 + 0.05) out.push({ n: s.aNs[i]!, e: s.aEw[i]!, v: s.aTvd[i]! });
    }
    if (out.length < 2 && s.svyMd >= day0 - 0.5 && s.svyMd <= day1 + 0.5) {
      for (let i = count - 1; i >= 1; i -= 1) {
        if (s.aMd[i]! <= day1) {
          return [{ n: s.aNs[i - 1]!, e: s.aEw[i - 1]!, v: s.aTvd[i - 1]! }, { n: s.aNs[i]!, e: s.aEw[i]!, v: s.aTvd[i]! }];
        }
      }
    }
    return out.length < 2 ? [] : out;
  }

  private polyline3(pts: Pt3[], color: string, weight: number, dash: Dash = "solid"): void {
    const q = pts.map((pt) => this.sh(pt.n, pt.e, pt.v));
    this.p.polyLine(q.map((a) => a[0]), q.map((a) => a[1]), q.length, color, weight, dash);
  }

  /** ShadowOnWall: orthographic shadow onto the back (n = nMax), left (e = eMin), or floor (v = vMax). */
  private shadow(pts: Pt3[], wallN: number, wallE: number, floorV: number, color: string): void {
    if (pts.length < 2) return;
    this.polyline3(pts.map((pt) => ({ ...pt, n: wallN })), color, 1.05);
    this.polyline3(pts.map((pt) => ({ ...pt, e: wallE })), color, 1.05);
    this.polyline3(pts.map((pt) => ({ ...pt, v: floorV })), color, 1.05);
  }

  private quad3(corners: Pt3[], fill: string, edge: string): void {
    const q = corners.map((pt) => this.sh(pt.n, pt.e, pt.v));
    this.p.polyFill(q.map((a) => a[0]), q.map((a) => a[1]), 4, fill, edge, 0.7);
  }

  private drawShaftProfile(): void {
    const s = this.s;
    const p = this.p;
    const ink = "2C3338";
    const shd = "7A848C";
    const planClr = "1A3A5C";
    const holeClr = "6B2020";
    const gridClr = "B4BCC2";

    const day = s.reportDayMd();
    let visEnd = day.d1;
    if (s.svyMd > visEnd && s.svyMd - visEnd <= 40) visEnd = s.svyMd;
    let mdA = day.d0;
    if (s.isExample) {
      mdA = s.exA;
      visEnd = s.exB;
    }
    const ahead = Math.min(90, Math.max(40, (visEnd - mdA) * 0.55));
    let mdB = visEnd + ahead;
    if (s.wpMd > visEnd && s.wpMd <= visEnd + ahead + 0.5) mdB = s.wpMd + 8;
    if (s.bitMd > mdB && s.bitMd - visEnd <= 40) mdB = s.bitMd + 3;
    if (mdB <= mdA) mdB = mdA + 40;

    p.tx(s.vx, 18, `${s.isExample ? "view  " : "24 h  "}${fmt0(mdA)} \u2192 ${fmt0(visEnd)} mMD   ${fmt(visEnd - mdA, 1)} m`, 12, ink, "start", true);
    p.tx(s.vx, s.vy + 12, "plan  \u00b7  last 24 h hole  \u00b7  shadows on back / left / floor", 8, "6B7A85", "start");
    p.rect(s.vx, s.vy, s.vw, s.vh, "FFFFFF", "8A949C", 0.6);

    const planMd = s.plan.md;
    if (mdA < planMd[0]!) mdA = planMd[0]!;
    if (mdB > planMd[planMd.length - 1]!) mdB = planMd[planMd.length - 1]!;
    const stepMd = Math.min(8, Math.max(2, (mdB - mdA) / 80));
    let plan: Pt3[] = [];
    for (let md = mdA; plan.length <= 160;) {
      const pt = planAt(s.plan, md);
      plan.push({ n: pt.n, e: pt.e, v: pt.v });
      if (md >= mdB) break;
      md = Math.min(md + stepMd, mdB);
    }
    let hole = this.collectDayHole(mdA, visEnd);

    this.computeShaftView(plan, hole);
    plan = plan.map((pt) => this.yaw(pt));
    hole = hole.map((pt) => this.yaw(pt));
    let proj: Pt3[] = [];
    for (let i = 0; i < s.ptbCount; i += 1) {
      if (s.ptbMd[i]! < mdA - 1 || s.ptbMd[i]! > mdB + 1) continue;
      proj.push(this.yaw({ n: s.ptbNs[i]!, e: s.ptbEw[i]!, v: s.ptbTvd[i]! }));
    }

    const seed = plan[0] ?? hole[0] ?? { n: this.vwN(s.svyNs, s.svyEw), e: this.vwE(s.svyNs, s.svyEw), v: s.svyTvd };
    const ext = { nMin: seed.n, nMax: seed.n, eMin: seed.e, eMax: seed.e, vMin: seed.v, vMax: seed.v };
    if (!plan[0] && !hole[0]) {
      ext.nMin = 0; ext.nMax = 0; ext.eMin = 0; ext.eMax = 0;
    }
    const grow = (pt: Pt3) => {
      ext.nMin = Math.min(ext.nMin, pt.n); ext.nMax = Math.max(ext.nMax, pt.n);
      ext.eMin = Math.min(ext.eMin, pt.e); ext.eMax = Math.max(ext.eMax, pt.e);
      ext.vMin = Math.min(ext.vMin, pt.v); ext.vMax = Math.max(ext.vMax, pt.v);
    };
    plan.forEach(grow);
    hole.forEach(grow);
    proj.forEach(grow);
    const targets = s.windowTargets(mdA, mdB);
    for (const target of targets) {
      const pt = planAt(s.plan, target.md);
      grow(this.yaw({ n: pt.n, e: pt.e, v: pt.v }));
    }

    const frmN = (ext.nMin + ext.nMax) / 2;
    const frmE = (ext.eMin + ext.eMax) / 2;
    const centre = (pt: Pt3): Pt3 => ({ n: pt.n - frmN, e: pt.e - frmE, v: pt.v });
    plan = plan.map(centre);
    hole = hole.map(centre);
    proj = proj.map(centre);
    let nMin = ext.nMin - frmN;
    let nMax = ext.nMax - frmN;
    let eMin = ext.eMin - frmE;
    let eMax = ext.eMax - frmE;
    let vMin = ext.vMin;
    let vMax = ext.vMax;

    const padN = Math.max(20, (nMax - nMin) * 0.12);
    const padE = Math.max(20, (eMax - eMin) * 0.12);
    const padV = Math.max(4, (vMax - vMin) * 0.08);
    nMin -= padN; nMax += padN;
    eMin -= padE; eMax += padE;
    vMin -= padV; vMax += padV;
    if (vMax - vMin < 12) {
      const mid = (vMin + vMax) / 2;
      vMin = mid - 6;
      vMax = mid + 6;
    }
    const needLat = Math.max(80, vMax - vMin);
    if (nMax - nMin < needLat) {
      const mid = (nMin + nMax) / 2;
      nMin = mid - needLat / 2;
      nMax = mid + needLat / 2;
    }
    if (eMax - eMin < needLat) {
      const mid = (eMin + eMax) / 2;
      eMin = mid - needLat / 2;
      eMax = mid + needLat / 2;
    }
    this.tgtV = (vMin + vMax) / 2;

    let uxMin = Infinity;
    let uxMax = -Infinity;
    let uyMin = Infinity;
    let uyMax = -Infinity;
    for (let i = 0; i < 8; i += 1) {
      const cN = i & 1 ? nMax : nMin;
      const cE = i & 2 ? eMax : eMin;
      const cV = i & 4 ? vMax : vMin;
      const ux = cE * 0.88 + cN * 0.52;
      const uy = cV - this.tgtV - cN * 0.4 + cE * 0.18;
      uxMin = Math.min(uxMin, ux); uxMax = Math.max(uxMax, ux);
      uyMin = Math.min(uyMin, uy); uyMax = Math.max(uyMax, uy);
    }
    if (uxMax - uxMin < 1) uxMax = uxMin + 1;
    if (uyMax - uyMin < 1) uyMax = uyMin + 1;
    const usableW = Math.max(80, s.vw - 70);
    const usableH = Math.max(80, s.vh - 40);
    this.shS = Math.min(12, Math.max(0.2, Math.min(usableW / (uxMax - uxMin), usableH / (uyMax - uyMin))));
    this.shNS = this.shS;
    this.shVS = this.shS;
    this.shOx = s.vx + 56 + (usableW - this.shS * (uxMax - uxMin)) / 2 - uxMin * this.shS;
    this.shOy = s.vy + 22 + (usableH - this.shS * (uyMax - uyMin)) / 2 - uyMin * this.shS;

    const wallEdge = "5C656C";
    this.quad3([{ n: nMax, e: eMin, v: vMin }, { n: nMax, e: eMax, v: vMin }, { n: nMax, e: eMax, v: vMax }, { n: nMax, e: eMin, v: vMax }], "FFFFFF", wallEdge);
    this.quad3([{ n: nMin, e: eMin, v: vMin }, { n: nMax, e: eMin, v: vMin }, { n: nMax, e: eMin, v: vMax }, { n: nMin, e: eMin, v: vMax }], "FFFFFF", wallEdge);
    this.quad3([{ n: nMin, e: eMin, v: vMax }, { n: nMax, e: eMin, v: vMax }, { n: nMax, e: eMax, v: vMax }, { n: nMin, e: eMax, v: vMax }], "E8EBEE", wallEdge);

    const stepV = Math.max(1, niceStep(vMax - vMin, 5));
    let v = stepV * Math.floor(vMin / stepV);
    if (v < vMin) v += stepV;
    for (let tick = 0; v <= vMax + 0.01 && tick < 16; tick += 1, v += stepV) {
      this.ln3(this.sh(nMin, eMin, v), this.sh(nMax, eMin, v), gridClr, 0.5, "solid");
      this.ln3(this.sh(nMax, eMin, v), this.sh(nMax, eMax, v), gridClr, 0.5, "solid");
      this.txAt(this.sh(nMin, eMin, v), -5, 3, fmt0(v), 7, "5C6770", "end");
    }
    this.drawFormationsOnShaft(nMin, nMax, eMin, eMax, vMin, vMax);

    const stepN = Math.max(1, niceStep(nMax - nMin, 4));
    const stepE = Math.max(1, niceStep(eMax - eMin, 4));
    let g = stepN * Math.floor(nMin / stepN);
    if (g < nMin) g += stepN;
    for (let tick = 0; g <= nMax + 0.01 && tick < 16; tick += 1, g += stepN) {
      this.ln3(this.sh(g, eMin, vMax), this.sh(g, eMax, vMax), gridClr, 0.5, "solid");
      this.txAt(this.sh(g, eMax, vMax), 2, 11, fmt(g, 0, false), 6.5, "5C6770", "start");
    }
    g = stepE * Math.floor(eMin / stepE);
    if (g < eMin) g += stepE;
    for (let tick = 0; g <= eMax + 0.01 && tick < 16; tick += 1, g += stepE) {
      this.ln3(this.sh(nMin, g, vMax), this.sh(nMax, g, vMax), gridClr, 0.5, "solid");
      this.txAt(this.sh(nMin, g, vMax), -3, 5, fmt(g, 0, false), 6.5, "5C6770", "end");
    }

    this.shadow(plan, nMax, eMin, vMax, shd);
    this.shadow(hole, nMax, eMin, vMax, shd);
    this.shadow(proj, nMax, eMin, vMax, "8FC7C9");
    this.polyline3(plan, planClr, 2);
    this.polyline3(hole, holeClr, 2.5);
    if (proj.length >= 2) {
      this.polyline3(proj, "008C95", 3);
      const end = proj[proj.length - 1]!;
      const [x, y] = this.sh(end.n, end.e, end.v);
      p.dot(x, y, 4.2, "FFFFFF", "008C95", 2);
      p.tx(x + 8, y - 7, "PTB", 9, "008C95", "start", true);
    }

    this.drawNamedTargets(targets, frmN, frmE);

    const bN = this.vwN(s.svyNs, s.svyEw) - frmN;
    const bE = this.vwE(s.svyNs, s.svyEw) - frmE;
    if (!s.isExample || (s.svyMd >= mdA - 0.5 && s.svyMd <= visEnd + 0.5)) {
      const [x, y] = this.sh(bN, bE, s.svyTvd);
      p.dot(x, y, 3.6, holeClr, "1A1A1A", 0.8);
      p.tx(x - 7, y + 2, "SVY", 8.5, holeClr, "end", true);
    }

    const gx = s.vx + s.vw - 40;
    const gy = s.vy + s.vh - 12;
    const gnN = this.vwN(1, 0);
    const gnE = this.vwE(1, 0);
    let gdx = gnE * 0.88 + gnN * 0.52;
    let gdy = -gnN * 0.4 + gnE * 0.18;
    const gm = Math.hypot(gdx, gdy) || 1;
    gdx = gdx / gm * 13;
    gdy = gdy / gm * 13;
    p.ln(gx, gy, gx + gdx, gy + gdy, planClr, 1.1, "solid");
    p.tx(gx + gdx + 3, gy + gdy + 4, "GN", 8, planClr, "start", true);

    const ly = s.vy + s.vh + 14;
    p.ln(s.vx + 2, ly, s.vx + 18, ly, holeClr, 2.4, "solid");
    p.tx(s.vx + 22, ly + 4, "24 h hole", 9, ink, "start");
    p.ln(s.vx + 78, ly, s.vx + 94, ly, planClr, 2, "solid");
    p.tx(s.vx + 98, ly + 4, "plan", 9, ink, "start");
    p.ln(s.vx + 128, ly, s.vx + 144, ly, shd, 1.1, "solid");
    p.tx(s.vx + 148, ly + 4, "wall / floor shadows", 9, ink, "start");
    p.ln(s.vx + 270, ly, s.vx + 286, ly, "008C95", 3, "solid");
    p.tx(s.vx + 290, ly + 4, "survey to PTB", 9, "008C95", "start");
  }

  /** DrawFormationsOnShaft: dashed tops on the left and back walls at TVD, labelled on the back wall. */
  private drawFormationsOnShaft(nMin: number, nMax: number, eMin: number, eMax: number, vMin: number, vMax: number): void {
    const ink = "6B4F2A";
    let lastY = -1e30;
    for (const fm of loadFormations(this.book)) {
      if (fm.tvd < vMin - 0.5 || fm.tvd > vMax + 0.5) continue;
      this.ln3(this.sh(nMin, eMin, fm.tvd), this.sh(nMax, eMin, fm.tvd), ink, 1.35, "dash");
      this.ln3(this.sh(nMax, eMin, fm.tvd), this.sh(nMax, eMax, fm.tvd), ink, 1.35, "dash");
      const e0 = eMin + (eMax - eMin) * 0.08;
      const e1 = eMin + (eMax - eMin) * 0.92;
      const [x0, wallY0] = this.sh(nMax, e0, fm.tvd);
      const [x1, wallY1] = this.sh(nMax, e1, fm.tvd);
      const nudge = Math.abs(wallY0 - lastY) < 11 ? 12 : 0;
      this.p.wallLabel(x0, wallY0 + nudge, x1, wallY1 + nudge, fm.name, `TVD ${fmt0(fm.tvd)} m`, fm.lith, 9, ink);
      lastY = wallY0 + nudge;
    }
  }

  /** DrawNamedTargetsOnPlan: traveling-cylinder bullseye (outer + inner ring) at each named target. */
  private drawNamedTargets(targets: Array<{ md: number; name: string }>, frmN: number, frmE: number): void {
    const s = this.s;
    const rt = Math.max(1.2, 16 / Math.max(0.05, this.shS));
    const local = (md: number): Pt3 => {
      const pt = planAt(s.plan, md);
      const y = this.yaw({ n: pt.n, e: pt.e, v: pt.v });
      return { n: y.n - frmN, e: y.e - frmE, v: y.v };
    };
    for (const target of targets) {
      const raw = planAt(s.plan, target.md);
      const c = local(target.md);
      const a = local(target.md - 6);
      const b = local(target.md + 6);
      const dN = b.n - a.n;
      const dE = b.e - a.e;
      const dV = b.v - a.v;
      const horiz = Math.hypot(dN, dE);
      let inc: number;
      let azi: number;
      if (horiz < 0.05 && Math.abs(dV) < 0.05) {
        inc = raw.inc;
        azi = raw.azi + this.view.azOff;
      } else {
        inc = dV <= 0.01 ? 90 : Math.atan2(horiz, dV) * DEG;
        azi = Math.atan2(dE, dN) * DEG;
      }
      this.ring(rt, c, inc, azi, "1B4D2E", 1.15);
      this.ring(rt * 0.45, c, inc, azi, "1B4D2E", 0.8);
      const [x, y] = this.sh(c.n, c.e, c.v);
      this.p.dot(x, y, 2.6, "3D2458", "1A1A1A", 0.7);
      this.p.tx(x + 7, y + 12, `${target.name}  ${fmt0(target.md)}`, 8, "1B4D2E", "start", true);
    }
  }

  /** SampleTargetRingAt: circle square to the wellbore tangent at the target, 29 points. */
  private ring(radius: number, c: Pt3, incDeg: number, aziDeg: number, color: string, weight: number): void {
    const sE = Math.max(0.001, this.shS);
    const sN = this.shNS < 0.001 ? sE : this.shNS;
    const sV = Math.max(0.001, this.shVS);
    let tN = Math.cos(rad(aziDeg)) * Math.sin(rad(incDeg)) * sN;
    let tE = Math.sin(rad(aziDeg)) * Math.sin(rad(incDeg)) * sE;
    let tV = Math.cos(rad(incDeg)) * sV;
    let mag = Math.hypot(tN, tE, tV);
    if (mag < 0.0001) {
      tN = 0; tE = 0; tV = 1;
    } else {
      tN /= mag; tE /= mag; tV /= mag;
    }
    const [hN, hE, hV] = Math.abs(tV) < 0.95 ? [0, 0, 1] : [1, 0, 0];
    let uN = tE * hV - tV * hE;
    let uE = tV * hN - tN * hV;
    let uV = tN * hE - tE * hN;
    mag = Math.hypot(uN, uE, uV);
    if (mag < 0.0001) {
      uN = 0; uE = 1; uV = 0;
    } else {
      uN /= mag; uE /= mag; uV /= mag;
    }
    const wN = tE * uV - tV * uE;
    const wE = tV * uN - tN * uV;
    const wV = tN * uE - tE * uN;
    const r = radius * sE;
    const pts: Pt3[] = [];
    const count = 29;
    for (let i = 0; i < count; i += 1) {
      const a = 2 * Math.PI * i / (count - 1);
      const cA = Math.cos(a);
      const sa = Math.sin(a);
      pts.push({ n: c.n + r * (cA * uN + sa * wN) / sN, e: c.e + r * (cA * uE + sa * wE) / sE, v: c.v + r * (cA * uV + sa * wV) / sV });
    }
    this.polyline3(pts, color, weight);
  }

  // ---- the live picture's side panels ------------------------------------------

  /** DrawSection: true 1:1 vertical section of the whole plan; drilled / remaining / today. */
  drawSection(): void {
    const s = this.s;
    const p = this.p;
    const plan = s.plan;
    const np = plan.md.length;
    const vsAzi = plan.azi[np - 1]!;
    const vsOf = (n: number, e: number) => n * Math.cos(rad(vsAzi)) + e * Math.sin(rad(vsAzi));
    let maxVs = 0;
    let maxTvd = 0;
    for (let i = 0; i < np; i += 1) {
      maxVs = Math.max(maxVs, vsOf(plan.ns[i]!, plan.ew[i]!));
      maxTvd = Math.max(maxTvd, plan.tvd[i]!);
    }
    if (maxVs <= 0 || maxTvd <= 0) return;
    const pad = 20;
    const sc = Math.max(maxVs / (s.pw - 2 * pad), maxTvd / (s.ph - 2 * pad));
    const gx = s.px + pad;
    const gy = s.py + pad;
    const at = (md: number): [number, number] => {
      const pt = planAt(plan, md);
      return [gx + vsOf(pt.n, pt.e) / sc, gy + pt.v / sc];
    };

    p.tx(s.px, 20, "VERTICAL SECTION \u00b7 1:1", 12.5, "FFFFFF", "start", true);
    p.rect(s.px, s.py, s.pw, s.ph, "222222", "3A3A3A", 0.75);
    p.ln(gx - 8, gy, s.px + s.pw - 8, gy, "4A4A4A", 1, "solid");
    for (let tv = 1000; tv <= maxTvd; tv += 1000) {
      p.ln(gx - 4, gy + tv / sc, s.px + s.pw - 8, gy + tv / sc, "3A3A3A", 0.6, "dot");
      p.tx(gx - 6, gy + tv / sc + 3, fmt0(tv), 7.5, "A8A8A8", "end");
    }

    const seg = (mdA: number, mdB: number, color: string, weight: number, dash: Dash) => {
      if (mdB <= mdA) return;
      const pts: Array<[number, number]> = [at(mdA)];
      for (let i = 0; i < np; i += 1) if (plan.md[i]! > mdA && plan.md[i]! < mdB) pts.push([gx + vsOf(plan.ns[i]!, plan.ew[i]!) / sc, gy + plan.tvd[i]! / sc]);
      pts.push(at(mdB));
      p.polyLine(pts.map((q) => q[0]), pts.map((q) => q[1]), pts.length, color, weight, dash);
    };
    const day = s.reportDayMd();
    seg(0, s.bitMd, "FFFFFF", 2.4, "solid");
    seg(s.bitMd, plan.md[np - 1]!, "6A6A6A", 1.6, "dash");
    seg(day.d0, day.d1, "5BA3D9", 4.2, "solid");

    for (const md of s.wMd) {
      const [x, y] = at(md);
      p.dot(x, y, 2, "4CAF7A", null, 0);
    }
    let prevName = "";
    for (let r = 2; r <= 5; r += 1) {
      const md = numeric(this.book.value(SS_SHEET, r, 21)) ?? 0;
      const name = cellText(this.book.value(SS_SHEET, r, 25));
      if (md <= 0 || name === "") continue;
      const [x, y] = at(md);
      p.dot(x, y, 3, "1A1A1A", "A8A8A8", 1);
      if (name !== prevName && planAt(plan, md).v > 500) p.tx(x - 7, y + 3, name, 8, "A8A8A8", "end");
      prevName = name;
    }
    const [bx, by] = at(s.bitMd);
    p.dot(bx, by, 4, "FF5555", "1A1A1A", 1.2);
    const bitLabel = `BIT ${fmt0(s.bitMd)}`;
    p.tx(this.keepInX(bx + 8, bitLabel, 8.5, true, "start", s.px + 8, s.px + s.pw - 8), by + 4, bitLabel, 8.5, "FF5555", "start", true);
    const [tx, ty] = at(plan.md[np - 1]!);
    const tdLabel = `TD ${fmt0(plan.md[np - 1]!)}`;
    p.tx(this.keepInX(tx, tdLabel, 8, false, "middle", s.px + 8, s.px + s.pw - 8), ty - 8, tdLabel, 8, "A8A8A8", "middle");
    p.tx(gx - 14, gy - 7, "TVD m", 8, "A8A8A8", "start");
    p.tx(s.px + s.pw - 8, s.py + s.ph - 7, "vertical section m \u2192", 8, "A8A8A8", "end");
  }

  private keepInX(x: number, text: string, size: number, bold: boolean, anchor: Anchor, lo: number, hi: number): number {
    const w = this.p.measure(text, size, bold);
    let l = anchor === "end" ? x - w : anchor === "middle" ? x - w / 2 : x;
    if (l + w > hi) {
      const over = l + w - hi;
      x -= over;
      l -= over;
    }
    if (l < lo) x += lo - l;
    return x;
  }

  private leaderAfter(x0: number, y: number, x1: number, text: string, size: number): void {
    const tw = this.p.measure(text, size, true) + 14;
    if (x0 + tw < x1 - 12) this.p.ln(x0 + tw, y - size * this.p.typeScale * 0.35, x1, y - size * this.p.typeScale * 0.35, "4A4A4A", 0.55, "solid");
  }

  private kv(px: number, pw: number, y: number, key: string, value: string, size: number, color: string, bold: boolean): void {
    this.p.tx(px + 11, y, key, size, "A8A8A8", "start");
    this.p.tx(px + pw - 11, y, value, size, color, "end", bold);
  }

  private data(a1: string): string {
    return cellDisplay(this.book, "Data", a1);
  }

  /** DrawDay: the reporting-day block, the Data sheet's own printed text. */
  drawDay(): void {
    const s = this.s;
    this.p.rect(s.px, s.qy, s.pw, s.qh, "222222", "3A3A3A", 0.6);
    this.p.tx(s.px + 11, s.qy + 16, "REPORTING DAY", 9.5, "FFFFFF", "start", true);
    this.leaderAfter(s.px + 11, s.qy + 16, s.px + s.pw - 11, "REPORTING DAY", 9.5);
    let y = s.qy + 34;
    const step = Math.min(16, (s.qy + s.qh - 9 - y) / 5);
    const rows: Array<[string, string, number, string, boolean]> = [
      ["From / to", `${this.data("C4")} \u2192 ${this.data("C5")}`, 9, "FFFFFF", false],
      ["Drilled", this.data("C6"), 9, "5BA3D9", true],
      ["Slid / rot", `${this.data("C7")} / ${this.data("C8")}`, 9, "FFFFFF", false],
      ["Sliding", this.data("C9"), 8.5, "FFFFFF", false],
      ["Rotating", this.data("C10"), 8.5, "FFFFFF", false],
    ];
    for (const [key, value, size, color, bold] of rows) {
      this.kv(s.px, s.pw, y, key, value, size, color, bold);
      y += step;
    }
  }

  /** DrawMetrics: last survey, position of wellbore, requirement to the next target. */
  drawMetrics(): void {
    const s = this.s;
    const p = this.p;
    p.tx(s.mx, 20, "LAST SURVEY \u00b7 POSITION \u00b7 REQUIREMENT", 12.5, "FFFFFF", "start", true);
    p.rect(s.mx, s.vy, s.mw, s.mh, "222222", "3A3A3A", 0.75);
    let y = s.vy + 20;
    const head = (text: string) => {
      y += 8;
      p.tx(s.mx + 11, y, text, 9.5, "FFFFFF", "start", true);
      this.leaderAfter(s.mx + 11, y, s.mx + s.mw - 11, text, 9.5);
      y += 18;
    };
    const row = (key: string, value: string, color = "FFFFFF", bold = false) => {
      this.kv(s.mx, s.mw, y, key, value, 10.5, color, bold);
      y += 16;
    };
    head("LAST SURVEY");
    row("Depth", `${fmt(s.svyMd, 2)} m`);
    row("INC", this.data("E8"));
    row("AZM", this.data("E9"));
    row("TVD", `${fmt(s.svyTvd, 2)} m`);
    row("DLS", this.data("E11"));
    row("Bit depth", `${fmt(s.bitMd, 2)} m`);
    head("POSITION OF WELLBORE");
    row("R/L from plan", `${this.data("E13")} ${this.data("F13")}`);
    row("Above/Below plan", `${this.data("E14")} ${this.data("F14")}`, "4CAF7A", true);
    if (s.useSailBands) {
      row("Dist. from geo", `${this.data("E15")} ${this.data("F15")}`);
      row("Corridor used", `${fmt(100 * Math.abs(s.svyDev) / s.geoHalf, 0)}% of \u00b1${fmt(s.geoHalf, 2)} m`);
      row("Lateral used", `${fmt(100 * Math.abs(s.svyLat) / s.latTol, 0)}% of \u00b1${fmt(s.latTol, 2)} m`);
    } else {
      row("Offset from plan", `${fmt(Math.abs(s.svyLat), 2)} m ${s.svyLat < 0 ? "L" : "R"}`);
      row("TVD vs plan", `${fmt(Math.abs(s.svyDev), 2)} m`);
    }
    head(`REQUIREMENT TO ${s.tgtName} ${fmt0(s.wpMd)}`);
    row("Distance to go", `${fmt(s.toGo, 1)} m`);
    row("Target TVD", `${fmt(s.wpTvd, 2)} m`);
    if (s.showBurr) {
      row("BURR", `${fmt(s.burr, 2)}\u00b0/30 m`, "FF5555", true);
      row("Holding", `${fmt(s.svyInc, 2)}\u00b0`);
      row("Landing inclination", `${fmt(s.aimInclination(), 2)}\u00b0`);
    } else {
      row("Required inclination", `${fmt(s.reqInc, 2)}\u00b0`, "FF5555", true);
      row("Holding", `${fmt(s.svyInc, 2)}\u00b0`);
      row("Correction", `${fmt(Math.abs(s.reqInc - s.svyInc), 2)}\u00b0 ${s.reqInc < s.svyInc ? "drop" : "build"}`, "FF5555", true);
    }
    row("Arrival if held", `${fmt(Math.abs(s.holdDevEnd), 2)} m ${s.holdDevEnd >= 0 ? "above" : "below"}`, "E09A3D");
  }
}

interface Pt3 {
  n: number;
  e: number;
  v: number;
}
