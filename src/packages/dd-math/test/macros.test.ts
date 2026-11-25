import { describe, expect, it } from "vitest";
import { minCurveStep } from "../src/angles";
import {
  appendTallyBitDepths,
  axisLabel,
  checkEntry,
  clearLog,
  clearSlidesheetRanges,
  bindSurveyDepthMetres,
  computePlanGauge,
  DLC,
  highlightNewestSurvey,
  lastSurveyRow,
  lengthsFromCsv,
  onTrackerEdit,
  previousIJ,
  refreshCommand,
  repairSurveyFormulas,
  SheetBook,
  slideComment,
  syncLastSurveyStrip,
  startEntry,
  storeTarget,
  surveyBarState,
  updateWellborePosition,
  settingFormula,
  writeEntry,
  type CellWrite,
  type DlEntryForm,
  type ProtectionInfo,
} from "../../workbook-commands/src/index";

type CellInput = string | number | { v?: unknown; f?: string; s?: unknown };

function rc(a1: string): [number, number] {
  const match = /^([A-Z]+)(\d+)$/.exec(a1)!;
  let col = 0;
  for (const ch of match[1]!) col = col * 26 + ch.charCodeAt(0) - 64;
  return [Number(match[2]) - 1, col - 1];
}

/** A saved workbook from { sheet: { A1: value } }, with a "good" style for the survey green. */
function makeBook(sheets: Record<string, Record<string, CellInput>>, protection: ProtectionInfo[] = []): SheetBook {
  const out: Record<string, { id: string; name: string; cellData: Record<number, Record<number, unknown>> }> = {};
  Object.entries(sheets).forEach(([name, cells], index) => {
    const cellData: Record<number, Record<number, unknown>> = {};
    for (const [a1, input] of Object.entries(cells)) {
      const [row, col] = rc(a1);
      cellData[row] ??= {};
      cellData[row]![col] = typeof input === "object" ? input : { v: input };
    }
    out[`s${index}`] = { id: `s${index}`, name, cellData };
  });
  return new SheetBook({ sheets: out, styles: { good: { bg: { rgb: "#CCFFCC" } }, blue: { bg: { rgb: "#A4CAFE" } }, nearGreen: { bg: { rgb: "#CCFFCB" } } } } as never, protection);
}

function form(overrides: Partial<DlEntryForm>): DlEntryForm {
  return {
    run: "1", page: "0", cmd: "1-5", period: "", type: "", bottom: "", i: "", j: "", depth: "1000",
    date: "2026-11-03", time: "06:30", code: "", mon: "", pass: "PASS", ref: "", drop: "", low: "", rpm: "", note: "",
    ...overrides,
  };
}

function wrote(writes: CellWrite[], sheet: string, row: number, col: number): CellWrite | undefined {
  return [...writes].reverse().find((write) => write.kind !== "sheet" && write.sheet === sheet && write.row === row && write.col === col);
}

const TRACKER_UNLOCKED: ProtectionInfo[] = [{ sheet: "DL Tracker", unlockedRanges: ["A3:A674", "C3:G674", "I3:U674"] }];

describe("DL Tracker ADD ENTRY (frmDLEntry)", () => {
  it("starts the first entry on row 3 and refuses a missing tab", () => {
    expect(startEntry(makeBook({ Data: {} }))).toEqual({ ok: false, message: "The 'DL Tracker' tab is not in this workbook." });
    const started = startEntry(makeBook({ "DL Tracker": {} }));
    expect(started.ok && started.start.row).toBe(3);
    expect(started.ok && started.start.status).toBe("First entry - fill in the command and settings.");
  });

  it("copies the last row, follows its setting, and keeps the code family it used", () => {
    const book = makeBook({
      "DL Tracker": {
        AA3: "1-5", AB3: "TF = 45 SR = 60",
        AA4: "2-1", AB4: "Go to Page 3: HIA Mode",
        A3: 7, C3: 0, D3: "1-5", I3: 10, J3: 20, N3: 261, P3: "PASS", U3: "first",
      },
    });
    const started = startEntry(book, new Date(2026, 10, 3, 6, 29, 40));
    expect(started.ok).toBe(true);
    if (!started.ok) return;
    expect(started.start).toMatchObject({ row: 4, source: 3, high: false, status: "Copied from row 3. Change what differs, then Save." });
    expect(started.start.form).toMatchObject({ run: "7", page: "0", cmd: "1-5", i: "45", j: "60", code: "261", date: "2026-11-03", time: "06:30", note: "first" });

    const goTo = makeBook({ "DL Tracker": { AA3: "2-1", AB3: "Go to Page 3: HIA Mode", C3: 0, D3: "2-1", P3: "PASS" } });
    const moved = startEntry(goTo);
    expect(moved.ok && moved.start.form.page).toBe("3");
    expect(moved.ok && moved.start.form.cmd).toBe("");
  });

  it("checks the form with the VBA messages", () => {
    expect(checkEntry(form({ cmd: " " }))).toEqual({ ok: false, error: "Pick a Dnlk command." });
    expect(checkEntry(form({ date: "2026-13-40" }))).toEqual({ ok: false, error: "Date or time is not valid (yyyy-mm-dd, hh:mm)." });
    expect(checkEntry(form({ i: "abc" }))).toEqual({ ok: false, error: "Numbers only in TF/INC, SR/AZI, depth, DLNK_PD, %, flow and RPM." });
    const odd = checkEntry(form({ code: "999" }));
    expect(odd.ok && odd.confirm).toBe("DLNK_PD 999 is neither 325 nor 261 for command 1-5.\nSave anyway?");
    const good = checkEntry(form({ code: "325", mon: "50", pass: "pass" }));
    expect(good.ok).toBe(true);
    if (!good.ok) return;
    expect(good.confirm).toBeNull();
    expect(good.vals[DLC.CODE]).toBe(325);
    expect(good.vals[DLC.MON]).toBe(0.5);
    expect(good.vals[DLC.PASS]).toBe("PASS");
    expect(Number.isInteger(good.vals[DLC.DATE])).toBe(true);
    expect(good.vals[DLC.TIME]).toBeCloseTo(6.5 / 24, 9);
    const failed = checkEntry(form({ code: "325", pass: "FAIL" }));
    expect(failed.ok && failed.vals[DLC.CODE]).toBeNull();
  });

  it("writes only the typed columns and leaves B, H and formulas alone", () => {
    const book = makeBook({ "DL Tracker": { B3: { f: "=IF(D3=\"\",\"\",1)" }, H3: { f: "=X" }, D3: "1-5", E4: "old" } }, TRACKER_UNLOCKED);
    writeEntry(book, "DL Tracker", 4, { [DLC.CMD]: "1-6", [DLC.SET]: "nope", [DLC.NUM]: 99, [DLC.PERIOD]: null });
    expect(wrote(book.writes, "DL Tracker", 4, DLC.NUM)).toMatchObject({ kind: "formula", formula: "=IF(D4=\"\",\"\",B3+1)" });
    expect(wrote(book.writes, "DL Tracker", 4, DLC.CMD)).toMatchObject({ kind: "value", value: "1-6" });
    expect(wrote(book.writes, "DL Tracker", 4, DLC.PERIOD)).toMatchObject({ kind: "clear" });
    expect(wrote(book.writes, "DL Tracker", 4, DLC.SET)).toMatchObject({ kind: "formula", formula: settingFormula(4) });
    expect(wrote(book.writes, "DL Tracker", 3, DLC.SET)).toBeUndefined();
  });

  it("puts the D/L Setting lookup back when H has no formula", () => {
    const book = makeBook({ "DL Tracker": { AA7: "1-4", AB7: "Set TF = 0 degrees, SR = 100%", C3: 0, D3: "1-4" } }, TRACKER_UNLOCKED);
    writeEntry(book, "DL Tracker", 3, { [DLC.CMD]: "1-4", [DLC.PAGE]: 0 });
    expect(wrote(book.writes, "DL Tracker", 3, DLC.SET)).toMatchObject({ kind: "formula", formula: settingFormula(3) });
    onTrackerEdit(book, "DL Tracker", { row: 3, col: 4, row2: 3, col2: 4 });
    expect(wrote(book.writes, "DL Tracker", 3, DLC.SET)?.kind).toBe("formula");
  });

  it("clears typed log cells only and counts them like DLT_ClearLog", () => {
    const book = makeBook({ "DL Tracker": { A1: "Client", B3: { f: "=1", v: 1 }, D3: "1-5", E3: 1, H3: { f: "=X", v: "s" }, D4: "1-6" } }, TRACKER_UNLOCKED);
    expect(clearLog(book, "DL Tracker")).toBe(3);
    expect(book.writes.filter((write) => write.kind === "clear").every((write) => write.kind === "clear" && write.col !== 2 && write.col !== 8)).toBe(true);
    expect(wrote(book.writes, "DL Tracker", 3, DLC.SET)).toMatchObject({ kind: "formula", formula: settingFormula(3) });
  });

  it("recomputes I/J when the command is typed on the log", () => {
    const book = makeBook({ "DL Tracker": { AA3: "1-5", AB3: "TF = 45 SR = 60", C3: 0, D3: "1-5" } });
    onTrackerEdit(book, "DL Tracker", { row: 3, col: 4, row2: 3, col2: 4 });
    expect(wrote(book.writes, "DL Tracker", 3, DLC.I)).toMatchObject({ value: 45 });
    expect(wrote(book.writes, "DL Tracker", 3, DLC.J)).toMatchObject({ value: 60 });
    const other = makeBook({ "DL Tracker": { AA3: "1-5", AB3: "TF = 45 SR = 60", C3: 0, D3: "1-5" } });
    onTrackerEdit(other, "DL Tracker", { row: 3, col: 5, row2: 3, col2: 5 });
    expect(other.writes).toEqual([]);
  });

  const hiaTables = {
    AA3: "2-28", AB3: "Go to Page 3: HIA Mode",
    AE3: "2-28", AF3: "Go to Page 3: HIA Mode",
    AG3: "2-4", AH3: "Decrease Target AZI by 2.5 deg instantaneous",
  };

  it("steps the first HIA AZI command from INC/AZM typed on the inbound go-to, not an older page-3 row", () => {
    const book = makeBook({
      "DL Tracker": {
        ...hiaTables,
        C3: 3, D3: "2-4", I3: 90, J3: 152.06, P3: "PASS",
        C4: 2, D4: "2-28", I4: 90.45, J4: 133.06, P4: "PASS",
        C5: 3, D5: "2-4", P5: "PASS",
      },
    });
    expect(previousIJ(book, "DL Tracker", 5, "3")).toEqual({ i: 90.45, j: 133.06 });
    onTrackerEdit(book, "DL Tracker", { row: 5, col: 4, row2: 5, col2: 4 });
    expect(wrote(book.writes, "DL Tracker", 5, DLC.I)).toMatchObject({ value: 90.45 });
    expect(wrote(book.writes, "DL Tracker", 5, DLC.J)).toMatchObject({ value: 130.56 });
  });

  it("steps AZI from INC/AZM typed on the first HIA command when the go-to row is blank", () => {
    const book = makeBook({
      "DL Tracker": {
        ...hiaTables,
        C3: 0, D3: "2-28", P3: "PASS",
        C4: 3, D4: "2-4", I4: 81, J4: 154.56, P4: "PASS",
      },
    });
    onTrackerEdit(book, "DL Tracker", { row: 4, col: 4, row2: 4, col2: 4 });
    expect(wrote(book.writes, "DL Tracker", 4, DLC.I)).toMatchObject({ value: 81 });
    expect(wrote(book.writes, "DL Tracker", 4, DLC.J)).toMatchObject({ value: 152.06 });
  });

  it("keeps later HIA commands on the same visit stepping from the previous page-3 row", () => {
    const book = makeBook({
      "DL Tracker": {
        ...hiaTables,
        C3: 2, D3: "2-28", I3: 90.45, J3: 133.06, P3: "PASS",
        C4: 3, D4: "2-4", I4: 90.45, J4: 130.56, P4: "PASS",
        C5: 3, D5: "2-4", P5: "PASS",
      },
    });
    onTrackerEdit(book, "DL Tracker", { row: 5, col: 4, row2: 5, col2: 4 });
    expect(wrote(book.writes, "DL Tracker", 5, DLC.I)).toMatchObject({ value: 90.45 });
    expect(wrote(book.writes, "DL Tracker", 5, DLC.J)).toMatchObject({ value: 128.06 });
  });

  it("ADD ENTRY after a filled HIA go-to applies 2-4 to those INC/AZM", () => {
    const book = makeBook({
      "DL Tracker": {
        ...hiaTables,
        C3: 3, D3: "2-4", I3: 90, J3: 152.06, P3: "PASS",
        C4: 2, D4: "2-28", I4: 90.45, J4: 133.06, P4: "PASS",
      },
    });
    const started = startEntry(book);
    expect(started.ok).toBe(true);
    if (!started.ok) return;
    const next = refreshCommand(book, started.start, { ...started.start.form, cmd: "2-4" });
    expect(next.i).toBe("90.45");
    expect(next.j).toBe("130.56");
  });
});

describe("Slidesheet macros", () => {
  it("Clear Ranges clears the entry ranges and restores RKB", () => {
    const book = makeBook({ Slidesheet: { D13: 100, Y12: "x", AC14: 1 } });
    clearSlidesheetRanges(book);
    const cleared = book.writes.filter((write) => write.kind === "clear").map((write) => `${write.col}:${write.row}-${write.row2}`);
    expect(cleared).toEqual(expect.arrayContaining(["4:13-505", "6:13-505", "7:13-505", "20:13-505", "21:13-505", "29:14-33", "21:2-5", "25:12-505", "26:13-505"]));
    expect(wrote(book.writes, "Slidesheet", 12, 25)).toMatchObject({ kind: "value", value: "RKB" });
  });

  it("formats the auto slide comment like ProjSlideComment", () => {
    expect(slideComment(3.456, "45L", 2.5, 9)).toBe("Sliding 3.46m @ 45L\u0001DLR 2.50");
    expect(slideComment(12, "", 1, 9)).toBe("Sliding 9.00m @ -\u0001DLR 1.00");
    expect(slideComment(0.5, "45L", 1, 9)).toBe("");
    expect(slideComment(3, "45L", "", 9)).toBe("");
  });

  it("stores a target with the frmTargets range checks", () => {
    const book = makeBook({ _OC_PlanSec: { A5: 1500, B5: 30, C5: 0, D5: 1400 } });
    expect(storeTarget(book, 5, "kop?", 1, 1, 1, 1)).toBe("Name must be one of: NUDGE,VERTICAL,KOP,TANGENT,SOT,EOT,ICP,BUILD,HEEL,TD");
    expect(storeTarget(book, 5, "kop", 1500, 200, 0, 1400)).toBe("KOP: MD, INC (0-180), AZM (0-360) and TVD have to be in range.");
    expect(storeTarget(book, 5, "kop", 1510, 30, 360, 1400)).toBe("");
    expect(wrote(book.writes, "_OC_PlanSec", 5, 12)).toMatchObject({ value: "KOP" });
    expect(wrote(book.writes, "_OC_PlanSec", 5, 19)).toMatchObject({ value: 1510 });
    expect(wrote(book.writes, "_OC_PlanSec", 5, 20)).toMatchObject({ kind: "clear" });
  });
});

describe("Pipe Tally Calculator", () => {
  it("takes the lengths below the column G total that matches the bit", () => {
    const csv = ["Joint,a,b,c,d,Length,Total", "1,,,,,9.5,980.00", "2,,,,,9.6,1000.02", "3,,,,,9.70,", "4,,,,,55,", "5,,,,,9.80,"].join("\r\n");
    expect(lengthsFromCsv(csv, 1000)).toEqual({ ok: true, lengths: [9.7, 9.8], note: "" });
    expect(lengthsFromCsv(csv, null)).toEqual({ ok: false, error: "No bit depth in Slidesheet column D." });
    const far = lengthsFromCsv(csv, 1100);
    expect(far).toEqual({ ok: false, error: "This file only reaches 1000.02 m in column G, and the bit is at 1100.00 m.\nPick a newer pipe tally export." });
  });

  it("appends the tally bit depths under the last Slidesheet D and clears F", () => {
    const book = makeBook({
      Slidesheet: { D12: 950, D13: 960 },
      "Pipe Tally Calculator": { H7: 969.7, H8: 979.5, H9: "" },
    });
    const result = appendTallyBitDepths(book);
    expect(result).toEqual({
      ok: true,
      message: "Added 2 bit depth(s) to Slidesheet starting at D14. Pipe tally F7:F315 cleared.",
      warning: null,
      added: { row: 14, row2: 15 },
    });
    expect(wrote(book.writes, "Slidesheet", 15, 4)).toMatchObject({ value: 979.5 });
    expect(book.writes.some((write) => write.sheet === "Pipe Tally Calculator" && write.kind === "clear" && write.col === 6)).toBe(true);
  });
});

describe("Slidesheet plan gauge (MDL_PlanGauge)", () => {
  it("shows NO DATA until a plan and surveys exist", () => {
    expect(computePlanGauge(makeBook({ Slidesheet: {} })).model).toMatchObject({ hasData: false, caption: "NO DATA - Import Plan / enter surveys" });
  });

  it("reads MAGNETIC off a vertical plan at the actual TVD using pure min-curvature", () => {
    const book = makeBook({
      _OC_Survey: { A2: "MD", B2: "INC", C2: "AZI", D2: "TVD", E2: "NS", F2: "EW", A3: 0, B3: 0, C3: 0, D3: 0, E3: 0, F3: 0, A4: 2000, B4: 0, C4: 0, D4: 2000, E4: 0, F4: 0 },
      Slidesheet: {
        E13: 0, F13: { v: 0, s: "good" }, G13: { v: 0, s: "good" },
        E14: 1000, F14: { v: 2, s: "good" }, G14: { v: 90, s: "good" },
      },
    });
    const expected = minCurveStep(0, 0, 0, 1000, 2, 90);
    const { model, wellbore } = computePlanGauge(book);
    expect(model).toMatchObject({ hasData: true, mode: "MAG", showPtb: false });
    expect(model.ax).toBeCloseTo(expected.e, 4);
    expect(model.ay).toBeCloseTo(expected.n, 4);
    expect(wellbore).toMatchObject({ row: 14, md: 1000 });
  });

  it("reads GRAVITY in the closest plan point's High-Side / Right normal plane", () => {
    // Plan runs past the last survey so the closest-point search is not clamped at its end.
    const stations = [[0, 0, 0], [1000, 0, 0], [1500, 30, 0], [1600, 30, 0]] as const;
    const plan: Record<string, CellInput> = { A2: "MD", B2: "INC", C2: "AZI", D2: "TVD", E2: "NS", F2: "EW" };
    let n = 0;
    let v = 0;
    stations.forEach(([md, inc, azi], i) => {
      if (i > 0) {
        const step = minCurveStep(stations[i - 1]![0], stations[i - 1]![1], stations[i - 1]![2], md, inc, azi);
        n += step.n;
        v += step.tvd;
      }
      const r = 3 + i;
      Object.assign(plan, { [`A${r}`]: md, [`B${r}`]: inc, [`C${r}`]: azi, [`D${r}`]: v, [`E${r}`]: n, [`F${r}`]: 0 });
    });
    // Actual wellbore builds 2 deg faster and turns 3 deg right of the North-bearing plan.
    const slides: Record<string, CellInput> = {
      E12: 5.55, F12: 0, G12: 0, H12: 5.55,
      E13: 1000, F13: { v: 0, s: "good" }, G13: { v: 0, s: "good" },
      E14: 1500, F14: { v: 32, s: "good" }, G14: { v: 3, s: "good" },
    };
    const { model } = computePlanGauge(makeBook({ _OC_Survey: plan, Slidesheet: slides }));
    expect(model.mode).toBe("GRAV");
    // Steeper build than plan puts wellbore UP (high-side > 0) and RIGHT (side > 0).
    expect(model.ax).toBeGreaterThan(5);
    expect(model.ay).toBeGreaterThan(8);
    expect(axisLabel("GRAV", true, -6.055)).toBe("DN 6.06m");
    expect(axisLabel("GRAV", false, 11.96)).toBe("RT 11.96m");
    expect(axisLabel("MAG", true, 0.4)).toBe("AH 0.40m");
  });

  it("auto-rebuilds plan N/E when stored _OC_Survey has glued Crystal EW, but keeps accurate Compass N/E", () => {
    const s1 = minCurveStep(100, 0, 0, 200, 30, 90);
    const s2 = minCurveStep(200, 30, 90, 300, 30, 90);
    // Cumulative 0.12 m North offset representing full-precision Compass plan coords vs 2-decimal angle rebuild
    const cleanPlan = {
      A2: "MD", B2: "INC", C2: "AZI", D2: "TVD", E2: "NS", F2: "EW",
      A3: 0, B3: 0, C3: 0, D3: 0, E3: 0, F3: 0,
      A4: 100, B4: 0, C4: 0, D4: 100, E4: 0.12, F4: 0,
      A5: 200, B5: 30, C5: 90, D5: 100 + s1.tvd, E5: 0.12 + s1.n, F5: s1.e,
      A6: 300, B6: 30, C6: 90, D6: 100 + s1.tvd + s2.tvd, E6: 0.12 + s1.n + s2.n, F6: s1.e + s2.e,
    };
    const corruptPlan = { ...cleanPlan, F3: 891.82, F4: 791.82, F5: 696.33, F6: 609.73 };
    const slides: Record<string, CellInput> = {
      E13: 100, F13: { v: 0, s: "good" }, G13: { v: 0, s: "good" },
      E14: 200, F14: { v: 30, s: "good" }, G14: { v: 90, s: "good" },
    };
    // Exact Compass N/E (within 2 m of MC) is preserved, giving a 0.12 m side offset.
    const clean = computePlanGauge(makeBook({ _OC_Survey: cleanPlan, Slidesheet: slides })).model;
    expect(Math.abs(clean.ax)).toBeCloseTo(0.12, 2);
    // Glued Crystal EW (891.82 at MD 0) is rejected and rebuilt via MC so actual matches plan.
    const repaired = computePlanGauge(makeBook({ _OC_Survey: corruptPlan, Slidesheet: slides })).model;
    expect(Math.abs(repaired.ax)).toBeLessThan(0.01);
    expect(Math.abs(repaired.ay)).toBeLessThan(0.01);
  });

  it("skips unparseable toolface notes in Slidesheet column U without throwing VALUE", () => {
    const plan: Record<string, CellInput> = {
      A2: "MD", B2: "INC", C2: "AZI", D2: "TVD", E2: "NS", F2: "EW",
      A3: 0, B3: 0, C3: 0, D3: 0, E3: 0, F3: 0,
      A4: 300, B4: 0, C4: 0, D4: 300, E4: 0, F4: 0,
    };
    const slides: Record<string, CellInput> = {
      D13: 220, E13: 200, F13: { v: 10, s: "good" }, G13: { v: 90, s: "good" }, AI13: 10,
      R13: 200, S13: 210, U13: "45L/60L",
      R14: 210, S14: 220, U14: "30R",
    };
    const { model } = computePlanGauge(makeBook({ _OC_Survey: plan, Slidesheet: slides }));
    expect(model.hasData).toBe(true);
    expect(model.showPtb).toBe(true);
  });
});

describe("Slidesheet survey formula repair", () => {
  const cum = (name: string, fn: string, r: number) => `=IF(E${r}="","",${fn}(${name}${r - 1},E${r - 1},F${r - 1},G${r - 1},E${r},F${r},G${r},FALSE))`;

  /** Rows 14:60 with the cumulative N/E pattern, plus whatever the test damages. */
  function surveyBook(extra: Record<string, CellInput> = {}): SheetBook {
    const slides: Record<string, CellInput> = {};
    for (let r = 14; r <= 60; r += 1) {
      slides[`AL${r}`] = { f: cum("AL", "ProjCumN", r) };
      slides[`AM${r}`] = { f: cum("AM", "ProjCumE", r) };
      slides[`D${r}`] = r * 10;
    }
    return makeBook({ Slidesheet: { ...slides, ...extra } });
  }

  /** What a deleted F cell leaves behind in Excel: #REF! in row n, and the next two rows pointing one row up. */
  function damaged(): Record<string, CellInput> {
    return {
      AL30: { f: `=IF(E30="","",ProjCumN(AL29,E29,F29,G29,E30,#REF!,G30,FALSE))` },
      AM30: { f: `=IF(E30="","",ProjCumE(AM29,E29,F29,G29,E30,#REF!,G30,FALSE))` },
      AL31: { f: `=IF(E31="","",ProjCumN(AL30,E30,#REF!,G30,E31,F30,G31,FALSE))` },
      AM31: { f: `=IF(E31="","",ProjCumE(AM30,E30,#REF!,G30,E31,F30,G31,FALSE))` },
      AL32: { f: `=IF(E32="","",ProjCumN(AL31,E31,F30,G31,E32,F32,G32,FALSE))` },
      AM32: { f: `=IF(E32="","",ProjCumE(AM31,E31,F30,G31,E32,F32,G32,FALSE))` },
    };
  }

  it("rewrites the cells a deleted survey cell broke, from the pattern the rest of the column follows", () => {
    const book = surveyBook(damaged());
    const fixed = repairSurveyFormulas(book);
    expect(fixed.map((cell) => `${cell.col === 38 ? "AL" : "AM"}${cell.row}`).sort()).toEqual(["AL30", "AL31", "AL32", "AM30", "AM31", "AM32"]);
    for (const r of [30, 31, 32]) {
      expect(book.formula("Slidesheet", r, 38)).toBe(cum("AL", "ProjCumN", r));
      expect(book.formula("Slidesheet", r, 39)).toBe(cum("AM", "ProjCumE", r));
    }
    expect(book.writes).toHaveLength(6);
  });

  it("changes nothing on a healthy sheet, so repeating it is silent", () => {
    const book = surveyBook();
    expect(repairSurveyFormulas(book)).toEqual([]);
    expect(book.writes).toHaveLength(0);
    const again = surveyBook(damaged());
    repairSurveyFormulas(again);
    const writesAfterFirst = again.writes.length;
    expect(repairSurveyFormulas(again)).toEqual([]);
    expect(again.writes).toHaveLength(writesAfterFirst);
  });

  it("leaves a typed value in the column alone, and the rows around it", () => {
    const book = surveyBook({ AL40: 12.5 });
    expect(repairSurveyFormulas(book)).toEqual([]);
    expect(book.writes).toHaveLength(0);
  });

  it("does not touch a column whose rows do not follow one pattern", () => {
    const mixed: Record<string, CellInput> = {};
    for (let r = 14; r <= 60; r += 1) mixed[`AX${r}`] = { f: r % 2 === 0 ? `=D${r}*2` : `=D${r}+1` };
    const book = surveyBook(mixed);
    expect(repairSurveyFormulas(book)).toEqual([]);
  });

  it("does not touch a column with only a few formulas", () => {
    const few: Record<string, CellInput> = {};
    for (let r = 14; r <= 19; r += 1) few[`AX${r}`] = { f: r === 16 ? `=D15*3` : `=D${r}*2` };
    const book = surveyBook(few);
    expect(repairSurveyFormulas(book)).toEqual([]);
  });

  it("reads the survey rows only: row 13 (the tie-in formula) and rows past 505 are left as they are", () => {
    const book = surveyBook({ AL13: { f: `=IF(E13="","",ProjCumN(0,E12,F12,G12,E13,F13,G13,TRUE))` }, AL506: { f: "=AL505" } });
    expect(repairSurveyFormulas(book)).toEqual([]);
  });
});

describe("Last Survey strip", () => {
  it("follows the deepest typed survey and shows that depth in metres", () => {
    const book = makeBook({
      Slidesheet: {
        E20: 4591.76, F20: 89.99, G20: 148.7,
        E21: 4706.9, F21: { v: 89.61, f: "=F20" }, G21: 147.74,
        E22: 4800, F22: 90, G22: 150,
        E506: { f: "=E20" },
      },
      Data: {
        E7: { v: 4591.76, f: "=Slidesheet!E506", s: { n: { pattern: "####.00°" } } },
      },
    });
    expect(syncLastSurveyStrip(book)).toBe(true);
    expect(book.formula("Slidesheet", 506, 5)).toBe("=E22");
    expect(book.formula("Slidesheet", 506, 6)).toBe("=F22");
    expect(book.formula("Slidesheet", 506, 7)).toBe("=G22");
    expect(syncLastSurveyStrip(book)).toBe(false);
    expect(bindSurveyDepthMetres(book)).toBe(true);
    expect(book.writes.some((write) => write.kind === "format" && write.row === 7 && write.col === 5 && write.pattern === '#####.00"m"')).toBe(true);
    expect(bindSurveyDepthMetres(book)).toBe(false);
  });

  it("names the geo-target offset Distance", () => {
    const book = makeBook({ Data: {}, Slidesheet: {} });
    updateWellborePosition(book, { lat: 0.08, planUp: -6.38, row: 20, md: 100, inc: 10 });
    expect(book.value("Data", 13, 6)).toBe("Right");
    expect(book.value("Data", 14, 6)).toBe("Below");
    expect(book.value("Data", 15, 6)).toBe("Distance");
    expect(book.value("Data", 15, 5)).toBe("6.38m");
  });
});

describe("Slidesheet green bar (last survey row painted the survey green)", () => {
  /** Rows 13..last: E, F, G numbers; `fills` names each row's Inc/Azm style; `barRow` carries the bar on its J cell. */
  function surveyRows(fills: Record<number, string>, barRow: number): SheetBook {
    const cells: Record<string, CellInput> = {};
    for (let r = 13; r <= 18; r += 1) {
      const s = fills[r];
      cells[`E${r}`] = 100 * r;
      cells[`F${r}`] = s ? { v: 20, s } : 20;
      cells[`G${r}`] = s ? { v: 90, s } : 90;
      if (r === barRow) cells[`J${r}`] = { v: 1, s: "good" };
    }
    return makeBook({ Slidesheet: cells });
  }

  it("is on the last green row, past rows painted another colour", () => {
    const book = surveyRows({ 13: "good", 14: "good", 15: "good", 16: "blue", 17: "blue" }, 15);
    expect(lastSurveyRow(book)).toBe(15);
    expect(surveyBarState(book)).toEqual({ newest: 15, stale: false });
  });

  it("is stale when the bar sits above a later green row, and a redraw moves it down", () => {
    const book = surveyRows({ 13: "good", 14: "good", 15: "good" }, 13);
    expect(surveyBarState(book)).toEqual({ newest: 15, stale: true });
    highlightNewestSurvey(book);
    expect(book.fill("Slidesheet", 15, 10)).toBe("#CCFFCC");
    expect(book.fill("Slidesheet", 13, 10)).not.toBe("#CCFFCC");
    expect(surveyBarState(makeBook({ Slidesheet: {} }))).toEqual({ newest: 0, stale: false });
  });

  it("steps back when the newest rows are recoloured, and returns when they are painted green again", () => {
    const recoloured = surveyRows({ 13: "good", 14: "good", 15: "blue" }, 15);
    expect(surveyBarState(recoloured)).toEqual({ newest: 14, stale: true });
    const painted = surveyRows({ 13: "good", 14: "good", 15: "good" }, 14);
    expect(surveyBarState(painted)).toEqual({ newest: 15, stale: true });
  });

  it("does not count a near green: only the exact survey green is a survey", () => {
    const book = surveyRows({ 13: "good", 14: "good", 15: "nearGreen" }, 14);
    expect(lastSurveyRow(book)).toBe(14);
    expect(surveyBarState(book).stale).toBe(false);
  });

  it("leaves the caller's book untouched while it looks", () => {
    const book = surveyRows({ 13: "good", 14: "good" }, 13);
    surveyBarState(book);
    expect(book.writes).toHaveLength(0);
  });

  it("is the reference of the gauge: the station is the last green row, not a projection below it", () => {
    const book = makeBook({
      _OC_Survey: { A2: "MD", B2: "INC", C2: "AZI", D2: "TVD", E2: "NS", F2: "EW", A3: 0, B3: 0, C3: 0, D3: 0, E3: 0, F3: 0, A4: 2000, B4: 0, C4: 0, D4: 2000, E4: 0, F4: 0 },
      Slidesheet: {
        E13: 0, F13: { v: 0, s: "good" }, G13: { v: 0, s: "good" },
        E14: 1000, F14: { v: 2, s: "good" }, G14: { v: 90, s: "good" }, AL14: 0, AM14: 3,
        E15: 1100, F15: { v: 3, s: "blue" }, G15: { v: 90, s: "blue" }, AL15: 0, AM15: 9,
      },
    });
    expect(computePlanGauge(book).wellbore).toMatchObject({ row: 14, md: 1000 });
  });
});
