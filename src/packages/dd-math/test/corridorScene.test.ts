import { readFileSync } from "node:fs";
import path from "node:path";
import { beforeAll, describe, expect, it } from "vitest";
import { corridorSize, paintCorridor, screenBend, shaftFacing, shaftYaw, yawPoint } from "../../corridor-3d/src/corridorDraw";
import type { PaintContext } from "../../corridor-3d/src/corridorPaint";
import { CorridorScene, TUNNEL_PITCH } from "../../corridor-3d/src/corridorScene";
import type { SnapshotBook } from "../../workbook-commands/src/cells";
import { SheetBook } from "../../workbook-commands/src/sheetBook";
import { readWorkbook } from "../../xlsx-bridge/src/readWorkbook";

const demoPath = path.resolve(import.meta.dirname, "../../../../Slide Sheet - Demo.xlsm");

function recorder(): PaintContext & { texts: string[] } {
  const texts: string[] = [];
  const noop = () => undefined;
  return {
    texts,
    fillStyle: "", strokeStyle: "", lineWidth: 1, lineCap: "butt", lineJoin: "miter", font: "", textAlign: "start", textBaseline: "alphabetic",
    beginPath: noop, closePath: noop, moveTo: noop, lineTo: noop, arc: noop, rect: noop, fill: noop, stroke: noop, setLineDash: noop,
    fillText: (text: string) => { texts.push(text); },
    measureText: (text: string) => ({ width: text.length * 5 }),
  };
}

describe("corridor picture (MDL_CorridorImage port) on the Demo well", () => {
  let book: SheetBook;

  beforeAll(async () => {
    const mapped = await readWorkbook(new Uint8Array(readFileSync(demoPath)));
    book = new SheetBook(mapped.workbook as unknown as SnapshotBook);
  }, 60_000);

  it("classifies a segment from the survey at the end of its window", () => {
    expect(CorridorScene.build(book, { from: 0, to: 500 })?.section).toBe("VERTICAL");
    expect(CorridorScene.build(book, { from: 2500, to: 3000 })?.section).toBe("BUILD");
  });

  it("keeps a segment to its own window instead of reopening it to surface or TD", () => {
    const scene = CorridorScene.build(book, { from: 1000, to: 1500 })!;
    expect([scene.md0, scene.md1]).toEqual([1000, 1500]);
    const stations = scene.sMd.slice(0, scene.sCount).filter((md) => md > 0);
    expect(Math.min(...stations)).toBeGreaterThanOrEqual(1000 - 40);
    expect(Math.max(...stations)).toBeLessThanOrEqual(1500 + 40);
  });

  it("draws today's survey only on a segment that contains it", () => {
    const live = CorridorScene.build(book)!;
    const early = recorder();
    paintCorridor(early, CorridorScene.build(book, { from: 0, to: 500 })!);
    expect(early.texts).not.toContain("SVY");
    const containing = recorder();
    const from = Math.floor(live.svyMd / 500) * 500;
    paintCorridor(containing, CorridorScene.build(book, { from, to: from + 500 })!);
    expect(containing.texts).toContain("SVY");
  });

  it("paints the live report day at the mail width", () => {
    const live = CorridorScene.build(book)!;
    expect(live.shaft).toBe(live.section !== "LATERAL");
    expect(corridorSize(live).width).toBe(1023);
    const context = recorder();
    paintCorridor(context, live);
    expect(context.texts.length).toBeGreaterThan(20);
  });

  it("yaws the shadow box by rotation only, so turns are never mirrored", () => {
    for (const lateral of [false, true]) {
      for (let az = 0; az < 360; az += 15) {
        const view = shaftYaw(az * Math.PI / 180, lateral);
        const n = yawPoint(view, 1, 0);
        const e = yawPoint(view, 0, 1);
        expect(n.n * e.e - n.e * e.n).toBeCloseTo(1, 12);
      }
    }
  });

  it("puts the start of a left-turning build day on the viewer's right when facing the bit", () => {
    // 35783 2,500 to 3,000 m: azimuth 261 to 165; the window start is 145 m N, 92 m E of its end.
    const view = shaftYaw(165 * Math.PI / 180, false);
    const start = yawPoint(view, 145, 92);
    const end = yawPoint(view, 0, 0);
    expect(start.e).toBeGreaterThan(end.e);
    expect(start.n).toBeGreaterThan(end.n);
  });

  it("turns the shaft camera off the bit when facing it would hide the day's turn", () => {
    const dir = (inc: number, az: number) => {
      const [i, a] = [inc * Math.PI / 180, az * Math.PI / 180];
      return { n: Math.sin(i) * Math.cos(a), e: Math.sin(i) * Math.sin(a), v: Math.cos(i) };
    };
    const signedScreenTurn = (az: number, start: ReturnType<typeof dir>, end: ReturnType<typeof dir>) => {
      const view = shaftYaw(az, false);
      const angle = (d: ReturnType<typeof dir>) => {
        const y = yawPoint(view, d.n, d.e);
        return Math.atan2(d.v - y.n * 0.4 + y.e * 0.18, y.e * 0.88 + y.n * 0.52);
      };
      return Math.atan2(Math.sin(angle(end) - angle(start)), Math.cos(angle(end) - angle(start)));
    };
    // 35783 2,517 to 3,000 m: inc 26 at az 261 bending to inc 37 at az 168 (46 degrees in 3D).
    const start = dir(26.3, 261.2);
    const end = dir(37.2, 168.1);
    const endAz = 168.1 * Math.PI / 180;
    expect(screenBend(shaftYaw(endAz, false), start, end)).toBeLessThan(10);
    const az = shaftFacing(endAz, start, end);
    expect(screenBend(shaftYaw(az, false), start, end)).toBeGreaterThan(40);
    // Azimuth decreasing is a left turn. The shaft runs down the page, so it has to bend toward the page's left
    // (clockwise on screen, positive).
    expect(signedScreenTurn(az, start, end)).toBeGreaterThan(0);

    const straight = dir(30, 120);
    expect(shaftFacing(120 * Math.PI / 180, straight, straight)).toBe(120 * Math.PI / 180);
  });

  it("bends the shaft the way the hole turns, left for a falling azimuth and right for a rising one", () => {
    const dir = (inc: number, az: number) => {
      const [i, a] = [inc * Math.PI / 180, az * Math.PI / 180];
      return { n: Math.sin(i) * Math.cos(a), e: Math.sin(i) * Math.sin(a), v: Math.cos(i) };
    };
    const signedScreenTurn = (az: number, start: ReturnType<typeof dir>, end: ReturnType<typeof dir>) => {
      const view = shaftYaw(az, false);
      const angle = (d: ReturnType<typeof dir>) => {
        const y = yawPoint(view, d.n, d.e);
        return Math.atan2(d.v - y.n * 0.4 + y.e * 0.18, y.e * 0.88 + y.n * 0.52);
      };
      return Math.atan2(Math.sin(angle(end) - angle(start)), Math.cos(angle(end) - angle(start)));
    };
    // Builds past 42 degrees that turn each way, from several headings.
    for (let endAzDeg = 0; endAzDeg < 360; endAzDeg += 30) {
      for (const turn of [-60, -25, 25, 60]) {
        const start = dir(38, endAzDeg - turn);
        const end = dir(46, endAzDeg);
        const az = shaftFacing(endAzDeg * Math.PI / 180, start, end);
        const bend = signedScreenTurn(az, start, end);
        if (turn < 0) expect(bend).toBeGreaterThan(0);
        else expect(bend).toBeLessThan(0);
      }
    }
  });

  it("tilts the tunnel 0.18 for every 0.88 along hole", () => {
    expect(TUNNEL_PITCH).toBeCloseTo(0.18 / 0.88, 12);
  });
});
