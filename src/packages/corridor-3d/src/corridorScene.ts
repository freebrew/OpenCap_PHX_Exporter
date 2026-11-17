import { minCurveStep } from "../../dd-math/src/angles";
import { walkUnseenToBit } from "../../dd-math/src/project";
import { frameComponents, isSummaryRow, loadPlan, planAt, slideIntervals, waypointHalfWidth, waypointTvdAtMd, type Plan } from "../../workbook-commands/src/planGauge";
import { cellText, numeric, type SheetBook } from "../../workbook-commands/src/sheetBook";
import { isGoodSurveyRow, PLANSEC_SHEET, seedNameFromText, SS_SHEET } from "../../workbook-commands/src/slidesheet";

// Scene half of MDL_CorridorImage: BuildScene and everything it reads. The drawing half
// (corridorDraw.ts) only reads this object, the same way the VBA draw subs read module state.

export type Section = "VERTICAL" | "BUILD" | "LATERAL";

const V_X = 12;
const V_Y = 30;
const V_W = 490;
const V_H = 376;
const KZX = 3;
const KZY = 2.2;
const DEV_SPAN_PT = 220;
const P_X = 530;
const P_Y = 30;
const P_W = 232;
const P_H = 250;
const Q_Y = 290;
const Q_H = 116;
const M_X = 758;
const M_W = 253;
const M_H = 376;

const SURV_FIRST = 13;
const SURV_LAST = 520;
const WP_FIRST = 14;
const WP_LAST = 33;
const RUN_AHEAD = 20;
const BIT_COL = 4;
const CROSSOVER_INC = 5;
export const TUNNEL_PITCH = 0.18 / 0.88;

const PLOT_TARGET_NAMES = ["KOP", "TANGENT", "HEEL", "SOT", "EOT", "TD", "NUDGE", "VERTICAL", "ICP", "BUILD"];

export const rad = (d: number) => d * Math.PI / 180;

/** ClassifySection: vertical shaft under 5 degrees, build shaft under 80, MD tunnel from 80. */
export function sectionFor(inc: number): Section {
  if (inc < 5) return "VERTICAL";
  if (inc < 80) return "BUILD";
  return "LATERAL";
}

/** VBA IsNumeric: an empty cell counts and reads as 0. */
function vbaNumeric(value: unknown): number | null {
  if (value == null || value === "") return 0;
  return numeric(value);
}

function numCell(value: unknown): number {
  return numeric(value) ?? 0;
}

function filled(value: unknown): boolean {
  return value != null && String(value).trim() !== "";
}

/** AA14 as a number, or the first number in text like "+/-10.00 m". */
function latTolFromCell(value: unknown): number {
  const direct = numeric(value);
  if (direct != null) return Math.abs(direct);
  const match = /[\d.]+/.exec(cellText(value));
  const parsed = match ? Number(match[0]) : Number.NaN;
  return Number.isFinite(parsed) ? Math.abs(parsed) : 0;
}

export interface Formation {
  md: number;
  tvd: number;
  name: string;
  lith: string;
}

/** PI_LoadFormations: _OC_Formations A=MD, B=TVD, C=name, D=lithology from row 3. */
export function loadFormations(book: SheetBook): Formation[] {
  const sheet = "_OC_Formations";
  if (!book.hasSheet(sheet)) return [];
  const out: Formation[] = [];
  let lastR = 0;
  for (let r = 3; r <= 2000; r += 1) if (filled(book.value(sheet, r, 1))) lastR = r;
  for (let r = 3; r <= lastR; r += 1) {
    const md = vbaNumeric(book.value(sheet, r, 1));
    const tvd = vbaNumeric(book.value(sheet, r, 2));
    if (md == null || tvd == null) continue;
    out.push({ md, tvd, name: cellText(book.value(sheet, r, 3)), lith: cellText(book.value(sheet, r, 4)) });
  }
  return out;
}

export interface WindowTarget {
  md: number;
  name: string;
}

export class CorridorScene {
  // layout (ApplySectionLayout)
  vx = V_X; vy = V_Y; vw = V_W; vh = V_H;
  px = P_X; py = P_Y; pw = P_W; ph = P_H;
  qy = Q_Y; qh = Q_H;
  mx = M_X; mw = M_W; mh = M_H;
  sxd = 0.297; kzx = KZX; kzy = KZY; devSpan = DEV_SPAN_PT;
  leftPad = 70;
  section: Section = "BUILD";
  tgtName = "WP";
  shaft = true;
  useSailBands = false;
  tvdTop = 0;

  md0 = 0; md1 = 0;
  devHi = 0; devLo = 0;
  latTol = 10; latBox = 12.5; geoHalf = 2;
  x0 = 0; y0 = 0; sv = 1;
  tvdDatum = 0;

  // stations inside the window
  sMd: number[] = []; sDev: number[] = []; sLat: number[] = []; sTvd: number[] = []; sNs: number[] = []; sEw: number[] = [];
  // every station, surface to bit
  aMd: number[] = []; aNs: number[] = []; aEw: number[] = []; aTvd: number[] = [];

  svyMd = 0; svyInc = 0; svyAzi = 0; svyTvd = 0; svyNs = 0; svyEw = 0;
  svyDev = 0; svyLat = 0; svyPlanAzi = 0; svyRow = 0; bitMd = 0;
  ptbMd: number[] = []; ptbInc: number[] = []; ptbAzi: number[] = []; ptbNs: number[] = []; ptbEw: number[] = []; ptbTvd: number[] = [];
  wpMd = 0; wpTvd = 0; wpInc = 0;
  toGo = 1; reqInc = 0; holdDevEnd = 0; holdLatEnd = 0;
  burr = 0; showBurr = false;

  wMd: number[] = []; wTvd: number[] = []; wInc: number[] = [];
  userWpCount = 0;
  heelMd = 0; heelTvd = 0;

  /** RenderCorridorExample window; 0/0 = the live report day. */
  exA = 0; exB = 0;

  plan: Plan;

  private constructor(readonly book: SheetBook, plan: Plan) {
    this.plan = plan;
  }

  get isExample(): boolean {
    return this.exB > this.exA && this.exB > 0;
  }

  get sCount(): number {
    return this.sMd.length;
  }

  get ptbCount(): number {
    return this.ptbMd.length;
  }

  /** BuildScene. Null when the workbook has no plan or fewer than two stations in the window. */
  static build(book: SheetBook, example?: { from: number; to: number }): CorridorScene | null {
    const plan = loadPlan(book);
    if (!plan || plan.md.length < 2) return null;
    const s = new CorridorScene(book, plan);
    if (example) {
      s.exA = example.from;
      s.exB = example.to;
    }
    return s.buildScene() ? s : null;
  }

  private value(r: number, c: number): unknown {
    return this.book.value(SS_SHEET, r, c);
  }

  waypointTvd(md: number): number | null {
    return waypointTvdAtMd(this.book, md);
  }

  private buildScene(): boolean {
    this.geoHalf = waypointHalfWidth(this.book);
    if (this.geoHalf <= 0) this.geoHalf = 2;
    this.latTol = latTolFromCell(this.value(14, 27));
    if (this.latTol <= 0) this.latTol = 10;
    this.latBox = this.latTol * 1.25;

    this.loadWaypoints();
    this.mergePlanTargetsAsWaypoints();
    if (!this.walkSurveys(true)) return false;
    this.buildBitProjection();

    this.tgtName = "WP";
    this.wpMd = 0;
    this.wpInc = 0;
    for (let i = 0; i < this.wMd.length; i += 1) {
      if (this.wMd[i]! > this.svyMd + 0.005) {
        this.wpMd = this.wMd[i]!;
        this.wpTvd = this.wTvd[i]!;
        this.wpInc = this.wInc[i]!;
        break;
      }
    }
    const next = this.nextPlanTarget(this.svyMd);
    if (this.wpMd > 0) {
      if (next && Math.abs(next.md - this.wpMd) < 0.75) {
        this.tgtName = next.name;
        if (next.inc > 0.01) this.wpInc = next.inc;
      }
      if (this.userWpCount >= 1 && Math.abs(this.wpMd - this.heelMd) < 0.01) this.tgtName = "HEEL";
    } else if (next) {
      this.wpMd = next.md;
      this.wpTvd = next.tvd;
      this.wpInc = next.inc;
      this.tgtName = next.name;
    } else {
      this.wpMd = this.svyMd + 80;
      this.wpTvd = this.svyTvd;
      this.wpInc = this.svyInc;
    }
    if (this.wpTvd === 0) this.wpTvd = this.waypointTvd(this.wpMd) ?? this.svyTvd;

    this.classifySection(this.svyInc);
    this.frameWindowToTarget();
    if (this.isExample) this.classifySection(this.incAtMd(this.md1));
    this.applySectionLayout();

    if (!this.walkSurveys(false)) return false;
    if (this.sCount < 2 && !this.isExample) {
      if (this.section === "LATERAL") {
        this.section = "BUILD";
        this.frameWindowToTarget();
        this.applySectionLayout();
      } else {
        this.md0 = 0;
        if (this.md1 < this.svyMd + RUN_AHEAD) this.md1 = this.svyMd + RUN_AHEAD;
        this.applySectionLayout();
      }
      if (!this.walkSurveys(false)) return false;
    }
    if (this.sCount < 2) return false;

    this.shaft = this.section !== "LATERAL";
    this.useSailBands = this.section === "LATERAL";

    const datum = this.waypointTvd(this.md0);
    if (datum != null) this.tvdDatum = datum;
    else if (this.userWpCount >= 1 && this.md0 < this.heelMd) this.tvdDatum = this.heelTvd;
    else this.tvdDatum = this.sTvd[0]!;

    if (this.shaft) {
      this.devHi = this.sTvd[0]!;
      this.devLo = this.sTvd[0]!;
      for (const tvd of this.sTvd) {
        if (tvd < this.devHi) this.devHi = tvd;
        if (tvd > this.devLo) this.devLo = tvd;
      }
      if (this.wpTvd > 0) {
        if (this.wpTvd < this.devHi) this.devHi = this.wpTvd;
        if (this.wpTvd > this.devLo) this.devLo = this.wpTvd;
      }
      this.devHi -= 8;
      this.devLo += 8;
      if (this.devLo - this.devHi < 40) {
        const mid = (this.devHi + this.devLo) / 2;
        this.devHi = mid - 20;
        this.devLo = mid + 20;
      }
      this.tvdTop = this.devHi;
      const tvdSpan = this.devLo - this.devHi;
      this.latBox = Math.min(140, Math.max(50, tvdSpan * 0.4));
      this.sv = (this.vh - 80) / tvdSpan;
      this.kzx = this.sv;
      this.sxd = this.sv * 0.08;
      this.kzy = this.sv * 0.12;
      this.x0 = this.vx + this.vw * 0.52;
      this.y0 = this.vy + 48;
    } else {
      this.devHi = -1e99;
      this.devLo = 1e99;
      const grow = (hi: number, lo: number) => {
        if (hi > this.devHi) this.devHi = hi;
        if (lo < this.devLo) this.devLo = lo;
      };
      for (const tvd of this.sTvd) grow(this.plotY(tvd), this.plotY(tvd));
      for (let i = 0; i < this.wMd.length; i += 1) {
        if (this.wMd[i]! >= this.md0 - 0.1 && this.wMd[i]! <= this.md1 + 0.1) {
          grow(this.plotY(this.wTvd[i]! - this.geoHalf), this.plotY(this.wTvd[i]! + this.geoHalf));
        }
      }
      for (const md of [this.md0, this.md1]) {
        const gT = this.waypointTvd(md);
        if (gT != null) grow(this.plotY(gT - this.geoHalf), this.plotY(gT + this.geoHalf));
      }
      this.devHi += 1;
      this.devLo -= 1;
      if (this.devHi - this.devLo < 8) {
        const mid = (this.devHi + this.devLo) / 2;
        this.devHi = mid + 4;
        this.devLo = mid - 4;
      }
      this.sv = this.devSpan / (this.devHi - this.devLo);
      this.x0 = this.vx + this.leftPad;
      this.y0 = this.vy + 42.5 + this.devHi * this.sv + this.latBox * this.kzy - (this.md1 - this.md0) * this.sxd * TUNNEL_PITCH * 0.5;
    }

    this.toGo = this.wpMd - this.svyMd;
    if (this.toGo <= 0) this.toGo = 1;
    const c = Math.max(-1, Math.min(1, (this.wpTvd - this.svyTvd) / this.toGo));
    this.reqInc = Math.acos(c) * 180 / Math.PI;
    this.holdDevEnd = this.wpTvd - (this.svyTvd + this.toGo * Math.cos(rad(this.svyInc)));
    this.holdLatEnd = this.svyLat + this.toGo * Math.sin(rad(this.svyAzi - this.svyPlanAzi));
    this.computeBurrMetrics();
    return true;
  }

  private loadWaypoints(): void {
    this.wMd = [];
    this.wTvd = [];
    this.wInc = [];
    for (let r = WP_FIRST; r <= WP_LAST; r += 1) {
      const vM = this.value(r, 29);
      const vT = this.value(r, 30);
      const vI = this.value(r, 31);
      if (!filled(vM) || !filled(vT)) continue;
      const md = numeric(vM);
      const tvd = numeric(vT);
      if (md == null || tvd == null) continue;
      if (this.wMd.length > 0 && !(md > this.wMd[this.wMd.length - 1]!)) continue;
      this.wMd.push(md);
      this.wTvd.push(tvd);
      this.wInc.push(filled(vI) ? numeric(vI) ?? 0 : 0);
    }
    this.userWpCount = this.wMd.length;
    this.heelMd = this.wMd[0] ?? 0;
    this.heelTvd = this.wTvd[0] ?? 0;
  }

  /** Named stations in T2:Y5 (U=MD, V=INC, X=TVD, Y=name) after bitMd. */
  private nextPlanTarget(bitMd: number): { md: number; tvd: number; inc: number; name: string } | null {
    for (let r = 2; r <= 5; r += 1) {
      const name = cellText(this.value(r, 25));
      const md = vbaNumeric(this.value(r, 21));
      if (name === "" || md == null || md <= bitMd + 0.005) continue;
      return { md, inc: numCell(this.value(r, 22)), tvd: numCell(this.value(r, 24)), name };
    }
    return null;
  }

  private mergePlanTargetsAsWaypoints(): void {
    for (let r = 2; r <= 5; r += 1) {
      const vMd = this.value(r, 21);
      const md = vbaNumeric(vMd);
      const tvd = vbaNumeric(this.value(r, 24));
      if (md == null || tvd == null || !filled(vMd)) continue;
      if (this.userWpCount >= 1 && cellText(this.value(r, 25)).toUpperCase() === "HEEL") continue;
      if (this.wMd.some((existing) => Math.abs(existing - md) < 0.5)) continue;
      let i = this.wMd.length;
      while (i > 0 && this.wMd[i - 1]! > md) i -= 1;
      this.wMd.splice(i, 0, md);
      this.wTvd.splice(i, 0, tvd);
      this.wInc.splice(i, 0, numCell(this.value(r, 22)));
    }
  }

  /** AimInclination: a HEEL aim under 80 degrees lands at 90. */
  aimInclination(): number {
    if (this.tgtName.trim().toUpperCase() === "HEEL" && this.wpInc < 80) return 90;
    return this.wpInc;
  }

  private computeBurrMetrics(): void {
    this.burr = 0;
    this.showBurr = false;
    if (this.section === "VERTICAL" || this.section === "LATERAL") return;
    const aim = this.aimInclination();
    if (Math.abs(aim - this.svyInc) <= 2.5 || this.toGo <= 0.01) return;
    this.showBurr = true;
    if (aim >= 80) {
      const dTvd = this.wpTvd - this.svyTvd;
      const dSin = Math.sin(rad(aim)) - Math.sin(rad(this.svyInc));
      if (dTvd > 0.01 && dSin > 0.01) {
        this.burr = (30 * 180 / Math.PI) / (dTvd / dSin);
        return;
      }
    }
    this.burr = (aim - this.svyInc) * 30 / this.toGo;
  }

  private surveyAt(r: number): { md: number; inc: number; azi: number } | null {
    if (isSummaryRow(this.book, r) || !isGoodSurveyRow(this.book, r)) return null;
    const vMd = this.value(r, 5);
    const vInc = this.value(r, 6);
    const vAzi = this.value(r, 7);
    if (!filled(vMd) || !filled(vInc) || !filled(vAzi)) return null;
    const md = numeric(vMd);
    const inc = numeric(vInc);
    const azi = numeric(vAzi);
    if (md == null || inc == null || azi == null) return null;
    return { md, inc, azi };
  }

  /** Inclination of the last good survey at or before md; the live survey otherwise. */
  private incAtMd(md: number): number {
    let bestMd = -1;
    let bestInc = this.svyInc;
    for (let r = SURV_FIRST; r <= SURV_LAST; r += 1) {
      const st = this.surveyAt(r);
      if (st && st.md <= md + 0.5 && st.md >= bestMd) {
        bestMd = st.md;
        bestInc = st.inc;
      }
    }
    return bestInc;
  }

  private classifySection(inc: number): void {
    this.section = sectionFor(inc);
  }

  /** ReportDayMd: Data C4/C5 (00:00 / Midnight depth), else I23/J23. */
  reportDayMd(): { d0: number; d1: number } {
    const data = (a1: string) => numCell(this.book.at("Data", a1));
    let d0 = data("C4");
    let d1 = data("C5");
    if (d1 <= 0) {
      d0 = data("I23");
      d1 = data("J23");
    }
    if (d1 <= 0) d1 = this.svyMd;
    if (d0 < 0) d0 = 0;
    if (d0 >= d1) d0 = Math.max(0, d1 - 30);
    return { d0, d1 };
  }

  private frameWindowToTarget(): void {
    if (this.isExample) {
      this.md0 = this.exA;
      this.md1 = this.exB;
      return;
    }
    if (this.section === "LATERAL") {
      const day = this.reportDayMd();
      let day1 = day.d1;
      if (day1 < this.svyMd && this.svyMd - day1 <= 40) day1 = this.svyMd;
      this.md0 = Math.max(0, day.d0);
      this.md1 = day1;
      if (this.md1 - this.md0 < 15) this.md0 = Math.max(0, this.md1 - 15);
      if (this.svyMd <= this.md1 + 40 && this.md1 < this.svyMd + 8) this.md1 = this.svyMd + 8;
      if (this.bitMd > this.svyMd + 0.01 && this.bitMd <= this.md1 + 40 && this.md1 < this.bitMd + 3) this.md1 = this.bitMd + 3;
      return;
    }
    const vertical = this.section === "VERTICAL";
    const lookBack = vertical ? 80 : 250;
    const maxWin = vertical ? 400 : 800;
    const minWin = vertical ? 60 : 120;
    const maxAhead = vertical ? 200 : 400;
    this.md1 = this.wpMd + RUN_AHEAD;
    if (this.md1 < this.svyMd + 15) this.md1 = this.svyMd + 15;
    if (this.md1 > this.svyMd + maxAhead) this.md1 = this.svyMd + maxAhead;
    this.md0 = Math.max(0, this.svyMd - lookBack);
    if (this.md1 - this.md0 > maxWin) this.md0 = this.md1 - maxWin;
    if (this.md1 - this.md0 < minWin) this.md0 = Math.max(0, this.md1 - minWin);
  }

  private applySectionLayout(): void {
    this.vy = V_Y;
    switch (this.section) {
      case "VERTICAL":
        Object.assign(this, { vx: 12, vw: 560, vh: 368, px: 584, py: 30, pw: 176, ph: 196, qy: 234, qh: 164, mx: 772, mw: 239, mh: 376, kzx: 2.2, kzy: 2, devSpan: 280 });
        break;
      case "BUILD":
        Object.assign(this, { vx: 12, vw: 520, vh: 368, px: 544, py: 30, pw: 196, ph: 214, qy: 252, qh: 146, mx: 752, mw: 259, mh: 376, kzx: 2.6, kzy: 2.1, devSpan: 240 });
        break;
      case "LATERAL":
        Object.assign(this, { vx: V_X, vw: V_W, vh: V_H, px: P_X, py: P_Y, pw: P_W, ph: P_H, qy: Q_Y, qh: Q_H, mx: M_X, mw: M_W, mh: M_H, kzx: KZX, kzy: KZY, devSpan: DEV_SPAN_PT });
        break;
      default: {
        const never: never = this.section;
        throw new Error(`section ${String(never)}`);
      }
    }
    const span = Math.max(40, this.md1 - this.md0);
    this.leftPad = 70;
    if (this.latBox * this.kzx + 25 > this.leftPad) this.leftPad = this.latBox * this.kzx + 25;
    let rightPad = this.latTol * this.kzx + 48;
    if (this.latBox * this.kzx + 8 > rightPad) rightPad = this.latBox * this.kzx + 8;
    this.sxd = Math.min(4, Math.max(0.02, (this.vw - this.leftPad - rightPad) / span));
  }

  /** WalkSurveys: integrate the survey block; pass 1 records the last survey, pass 2 the window. */
  private walkSurveys(findLastOnly: boolean): boolean {
    let prev = { md: 0, inc: 0, azi: 0 };
    let curN = 0;
    let curE = 0;
    let curV = 0;
    let lastRow = 0;
    if (!findLastOnly) {
      this.sMd = []; this.sDev = []; this.sLat = []; this.sTvd = []; this.sNs = []; this.sEw = [];
      this.aMd = [0]; this.aNs = [0]; this.aEw = [0]; this.aTvd = [0];
    }
    for (let r = SURV_FIRST; r <= SURV_LAST; r += 1) {
      const st = this.surveyAt(r);
      if (!st || !(st.md > prev.md)) continue;
      const step = minCurveStep(prev.md, prev.inc, prev.azi, st.md, st.inc, st.azi);
      curN += step.n;
      curE += step.e;
      curV += step.tvd;
      prev = st;
      lastRow = r;
      if (findLastOnly) continue;
      if (this.aMd.length <= 1023) {
        this.aMd.push(st.md); this.aNs.push(curN); this.aEw.push(curE); this.aTvd.push(curV);
      }
      if (st.md >= this.md0 && st.md <= this.md1 && this.sMd.length <= 511) {
        const off = this.stationOffsets(st.md, st.inc, curN, curE, curV);
        this.sMd.push(st.md); this.sDev.push(off.dev); this.sLat.push(off.lat);
        this.sTvd.push(curV); this.sNs.push(curN); this.sEw.push(curE);
      }
    }
    if (prev.md <= 0) return false;
    if (findLastOnly) {
      this.svyMd = prev.md;
      this.svyInc = prev.inc;
      this.svyAzi = prev.azi;
      this.svyTvd = curV;
      this.svyNs = curN;
      this.svyEw = curE;
      this.svyRow = lastRow;
      const pl = planAt(this.plan, prev.md);
      this.svyPlanAzi = pl.azi;
      this.svyLat = frameComponents(prev.inc >= CROSSOVER_INC, pl.azi, curN - pl.n, curE - pl.e, pl.v - curV).x;
      const wt = this.waypointTvd(prev.md);
      if (wt != null) this.svyDev = wt - curV;
      this.bitMd = numCell(this.value(lastRow, BIT_COL));
      if (this.bitMd < prev.md) this.bitMd = prev.md;
    }
    return true;
  }

  /** StationOffsets: up/down against the geo line (plan TVD above the sail table), L/R of plan. */
  stationOffsets(md: number, inc: number, n: number, e: number, v: number): { dev: number; lat: number } {
    const pl = planAt(this.plan, md);
    const wt = this.waypointTvd(md) ?? pl.v;
    const lat = frameComponents(inc >= CROSSOVER_INC, pl.azi, n - pl.n, e - pl.e, pl.v - v).x;
    return { dev: wt - v, lat };
  }

  private buildBitProjection(): void {
    this.ptbMd = []; this.ptbInc = []; this.ptbAzi = []; this.ptbNs = []; this.ptbEw = []; this.ptbTvd = [];
    if (this.svyRow < 1 || this.bitMd <= this.svyMd + 0.01) return;
    const span = this.bitMd - this.svyMd;
    const nSeg = Math.min(120, Math.max(2, Math.floor(span) + 1));
    const motorOut = numeric(this.value(this.svyRow, 35)) ?? 0;
    const slides = slideIntervals(this.book);
    const push = (md: number, inc: number, azi: number, n: number, e: number, tvd: number) => {
      this.ptbMd.push(md); this.ptbInc.push(inc); this.ptbAzi.push(azi);
      this.ptbNs.push(n); this.ptbEw.push(e); this.ptbTvd.push(tvd);
    };
    push(this.svyMd, this.svyInc, this.svyAzi, this.svyNs, this.svyEw, this.svyTvd);
    for (let i = 1; i <= nSeg; i += 1) {
      const md = this.svyMd + span * i / nSeg;
      const walked = walkUnseenToBit(this.svyMd, this.svyInc, this.svyAzi, md, motorOut, slides);
      const out = walked ?? { inc: this.svyInc, azm: this.svyAzi, ...minCurveStep(this.svyMd, this.svyInc, this.svyAzi, md, this.svyInc, this.svyAzi) };
      push(md, out.inc, out.azm, this.svyNs + out.n, this.svyEw + out.e, this.svyTvd + out.tvd);
    }
  }

  /** LoadWindowTargets: named stations in the MD window from _OC_PlanSec, then T2:Y5. */
  windowTargets(mdA: number, mdB: number): WindowTarget[] {
    const out: WindowTarget[] = [];
    const take = (md: number, name: string) => {
      if (md < mdA - 0.5 || md > mdB + 0.5) return;
      if (out.some((target) => Math.abs(target.md - md) < 0.5)) return;
      out.push({ md, name });
    };
    if (this.book.hasSheet(PLANSEC_SHEET)) {
      let lastR = 0;
      for (let r = 3; r <= 2000; r += 1) if (filled(this.book.value(PLANSEC_SHEET, r, 1))) lastR = r;
      for (let r = 3; r <= lastR && out.length <= 15; r += 1) {
        const md = vbaNumeric(this.book.value(PLANSEC_SHEET, r, 1));
        if (md == null) continue;
        let name = cellText(this.book.value(PLANSEC_SHEET, r, 12)).toUpperCase();
        if (name === "") name = cellText(this.book.value(PLANSEC_SHEET, r, 11)).toUpperCase();
        if (name === "") name = seedNameFromText(cellText(this.book.value(PLANSEC_SHEET, r, 10)));
        if (!PLOT_TARGET_NAMES.includes(name)) continue;
        take(md, name);
      }
    }
    for (let r = 2; r <= 5; r += 1) {
      const name = cellText(this.value(r, 25));
      const md = vbaNumeric(this.value(r, 21));
      if (name === "" || md == null) continue;
      take(md, name);
    }
    return out;
  }

  // ---- projection ------------------------------------------------------------

  /** Tunnel: screen-up metres against geo at the window start. Shaft: true TVD. */
  plotY(tvd: number): number {
    return this.shaft ? tvd : this.tvdDatum - tvd;
  }

  plotInRoom(plot: number): boolean {
    return this.shaft ? plot >= this.devHi && plot <= this.devLo : plot >= this.devLo && plot <= this.devHi;
  }

  /** ObX / ObY in one call (the VBA passes MD to ObY through mObMd). */
  ob(md: number, plot: number, lat: number): [number, number] {
    if (this.shaft) {
      return [this.x0 + lat * this.kzx + (md - this.md0) * this.sxd, this.y0 + (plot - this.tvdTop) * this.sv + lat * this.kzy];
    }
    return [
      this.x0 + (md - this.md0) * this.sxd + lat * this.kzx,
      this.y0 - plot * this.sv - lat * this.kzy + (md - this.md0) * this.sxd * TUNNEL_PITCH,
    ];
  }

  clampMd(md: number): number {
    return Math.min(this.md1, Math.max(this.md0, md));
  }

  /** HoleAtMdPlot: as-drilled (plot, lat) interpolated at an MD inside the window. */
  holeAtMdPlot(md: number): { plot: number; lat: number } | null {
    const n = this.sCount;
    if (n < 1 || md < this.sMd[0]! - 0.01 || md > this.sMd[n - 1]! + 0.01) return null;
    if (n === 1 || md <= this.sMd[0]!) return { plot: this.plotY(this.sTvd[0]!), lat: this.sLat[0]! };
    if (md >= this.sMd[n - 1]!) return { plot: this.plotY(this.sTvd[n - 1]!), lat: this.sLat[n - 1]! };
    for (let i = 0; i < n - 1; i += 1) {
      if (md >= this.sMd[i]! && md <= this.sMd[i + 1]!) {
        const f = this.sMd[i + 1]! > this.sMd[i]! ? (md - this.sMd[i]!) / (this.sMd[i + 1]! - this.sMd[i]!) : 0;
        return { plot: this.plotY(this.sTvd[i]! + f * (this.sTvd[i + 1]! - this.sTvd[i]!)), lat: this.sLat[i]! + f * (this.sLat[i + 1]! - this.sLat[i]!) };
      }
    }
    return null;
  }

  /** HoleXyzAtMd over every station, clamped at both ends. */
  holeXyzAtMd(md: number): { n: number; e: number; v: number } | null {
    const count = this.aMd.length;
    if (count < 2) return null;
    if (md <= this.aMd[0]!) return { n: this.aNs[0]!, e: this.aEw[0]!, v: this.aTvd[0]! };
    if (md >= this.aMd[count - 1]!) return { n: this.aNs[count - 1]!, e: this.aEw[count - 1]!, v: this.aTvd[count - 1]! };
    for (let i = 0; i < count - 1; i += 1) {
      if (md >= this.aMd[i]! && md <= this.aMd[i + 1]!) {
        const span = this.aMd[i + 1]! - this.aMd[i]!;
        const f = span <= 0.0001 ? 0 : (md - this.aMd[i]!) / span;
        return {
          n: this.aNs[i]! + f * (this.aNs[i + 1]! - this.aNs[i]!),
          e: this.aEw[i]! + f * (this.aEw[i + 1]! - this.aEw[i]!),
          v: this.aTvd[i]! + f * (this.aTvd[i + 1]! - this.aTvd[i]!),
        };
      }
    }
    return null;
  }
}

/** ArcTvdGain: TVD gained over a along a constant-curvature arc from inc0 (rad), kap rad/m. */
export function arcTvdGain(inc0: number, kap: number, a: number): number {
  if (Math.abs(kap) < 0.0000001) return a * Math.cos(inc0);
  return (Math.sin(inc0 + kap * a) - Math.sin(inc0)) / kap;
}

/** SolveArcCurvature: signed curvature that gains dTvd over toGo, clamped to +-40 deg/30 m. */
export function solveArcCurvature(inc0Deg: number, toGo: number, dTvd: number): number {
  if (toGo <= 0) return 0;
  const i1 = rad(inc0Deg);
  const kMax = rad(40 / 30);
  let lo = -kMax;
  let hi = kMax;
  if (arcTvdGain(i1, lo, toGo) <= dTvd) return lo;
  if (arcTvdGain(i1, hi, toGo) >= dTvd) return hi;
  for (let k = 0; k < 48; k += 1) {
    const mid = (lo + hi) / 2;
    if (arcTvdGain(i1, mid, toGo) > dTvd) lo = mid;
    else hi = mid;
  }
  return (lo + hi) / 2;
}

/** NiceStep: a 1/2/5 grid step giving about the requested tick count. */
export function niceStep(span: number, ticks: number): number {
  if (span <= 0 || ticks < 1) return 1;
  const raw = span / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const m = raw / mag;
  if (m < 1.5) return mag;
  if (m < 3.5) return 2 * mag;
  if (m < 7.5) return 5 * mag;
  return 10 * mag;
}
