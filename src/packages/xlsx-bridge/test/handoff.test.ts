import { describe, expect, it } from "vitest";
import {
  auditBundle,
  buildBundleFiles,
  canonicalize,
  classifyChange,
  commandBlocksHandoff,
  commitPlan,
  diffHandoff,
  parseCsv,
  sha256Text,
  type CanonBook,
  type CanonCell,
  type CanonSheet,
} from "../src/handoff.js";
import {
  defaultConflictChoice,
  diffVisibleUserData,
  machineWrittenCell,
  mergeVisibleUserData,
  scalarEqual,
} from "../src/syncMerge.js";

function book(cells: Record<string, Record<string, unknown>>, extra: Record<string, unknown> = {}) {
  return {
    styles: {
      fill: { bg: { rgb: "#112233" }, bl: 1 },
      odd: { it: 1, zz: "nope" },
    },
    sheetOrder: ["s1"],
    sheets: {
      s1: {
        name: "Calc",
        hidden: 0,
        rowCount: 20,
        columnCount: 10,
        freeze: { xSplit: 0, ySplit: 1 },
        mergeData: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }],
        cellData: cells,
        rowData: { 1: { h: 30 } },
        columnData: { 0: { w: 70 } },
        ...extra,
      },
    },
    definedNames: [{ name: "Answer", formulaOrRefString: "Calc!$C$1", hidden: false }],
    resources: [
      { name: "SHEET_DATA_VALIDATION_PLUGIN", data: JSON.stringify({ rules: [1] }) },
      { name: "SHEETS_FORMULA_PLUGIN", data: "{\"cache\":1}" },
      { name: "SHEET_VIEW_PLUGIN", data: "{\"scroll\":9}" },
    ],
  };
}

describe("handoff contract", () => {
  it("ignores formula cached values, view resources and formula-plugin cache", () => {
    const base = canonicalize(book({
      0: { 0: { v: "Title" }, 2: { f: "=SUM(B1:B3)", v: 0 } },
    }));
    const live = canonicalize({
      ...book({
        0: { 0: { v: "Title" }, 2: { f: "=SUM(B1:B3)", v: 99 } },
      }),
      resources: [
        { name: "SHEET_DATA_VALIDATION_PLUGIN", data: JSON.stringify({ rules: [1] }) },
        { name: "SHEETS_FORMULA_PLUGIN", data: "{\"cache\":999}" },
        { name: "SHEET_VIEW_PLUGIN", data: "{\"scroll\":1}" },
      ],
    });
    const diff = diffHandoff(base, live);
    expect(diff.blocked).toEqual([]);
    expect(diff.ops).toEqual([]);
    expect(base.resources.SHEET_VIEW_PLUGIN).toBeUndefined();
  });

  it("emits a cell edit that was not a CellWrite and blocks an unrepresentable command", () => {
    const base = canonicalize(book({ 0: { 0: { v: "Title" } } }));
    const live = canonicalize(book({
      0: { 0: { v: "Title", s: "fill" }, 3: { v: "typed in the grid" } },
    }));
    const diff = diffHandoff(base, live, ["sheet.command.insert-row-before"]);
    expect(diff.blocked).toContain("command blocked: sheet.command.insert-row-before");
    expect(diff.ops.some((op) => op.kind === "cell" && op.a1 === "D1" && op.nextValue === "typed in the grid")).toBe(true);
    expect(diff.ops.some((op) => op.kind === "style" && op.field === "bg" && op.next === "112233")).toBe(true);
    expect(commandBlocksHandoff("sheet.command.set-range-values")).toBe(false);
    expect(commandBlocksHandoff("sheet.command.insert-sheet")).toBe(false);
    expect(commandBlocksHandoff("sheet.mutation.insert-row")).toBe(false);
    expect(commandBlocksHandoff("sheet.command.remove-sheet")).toBe(true);
  });

  it("refuses a validation resource change, a deleted sheet and an unknown style", () => {
    const base = canonicalize(book({ 0: { 0: { v: 1 } } }));
    const live = canonicalize(book({ 0: { 0: { v: 1, s: "odd" } } }));
    live.resources.SHEET_DATA_VALIDATION_PLUGIN = "{\"rules\":[2]}";
    const styled = diffHandoff(base, live);
    expect(styled.blocked.some((item) => item.startsWith("unsupported resource change"))).toBe(true);
    expect(styled.blocked.some((item) => item.startsWith("unsupported style change"))).toBe(true);
    const removed = diffHandoff(base, { ...live, sheets: [] });
    expect(removed.blocked.some((item) => item.startsWith("sheet removed"))).toBe(true);
  });

  it("classifies an overlapping Excel edit as a conflict and an already-applied cell as idempotent", () => {
    expect(classifyChange("1", "1", "2")).toBe("apply");
    expect(classifyChange("2", "1", "2")).toBe("idempotent");
    expect(classifyChange("9", "1", "2")).toBe("conflict");
  });

  it("round-trips quoted csv, hashes every payload, and commits READY last", async () => {
    const base = canonicalize(book({ 0: { 0: { v: "a,b" } } }));
    const live = canonicalize(book({ 0: { 0: { v: "a,\"b\"" } } }));
    const diff = diffHandoff(base, live);
    expect(diff.blocked).toEqual([]);
    const files = await buildBundleFiles({
      jobNumber: "35783",
      sourceWorkbookHash: "abc",
      createdAt: "2026-11-04T00:00:00.000Z",
      exporter: "browser shift",
      ops: diff.ops,
    });
    expect(files.READY).toBeUndefined();
    const parsed = parseCsv(files["cells.csv"]);
    expect(parsed[1][3]).toBe("a,b");
    expect(parsed[1][5]).toBe("a,\"b\"");
    expect(await auditBundle(files)).toEqual([]);
    const plan = commitPlan(files);
    expect(plan.payloads.some((file) => file.name === "READY")).toBe(false);
    expect(plan.ready.name).toBe("READY");
    expect(await sha256Text("abc")).toBe("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
  });

  it("rejects a bundle whose payload does not match the manifest", async () => {
    const files = await buildBundleFiles({
      jobNumber: "Demo",
      sourceWorkbookHash: "abc",
      createdAt: "2026-11-04T00:00:00.000Z",
      exporter: "browser shift",
      ops: [],
    });
    files["cells.csv"] = "tampered\n";
    const problems = await auditBundle(files);
    expect(problems.some((item) => item.includes("cells.csv"))).toBe(true);
  });
});

function cell(kind: CanonCell["k"], value = ""): CanonCell {
  return { k: kind, v: value };
}

function visible(name: string, cells: Record<string, CanonCell>, hidden: 0 | 1 | 2 = 0): CanonSheet {
  return { name, hidden, rowCount: 40, columnCount: 20, freeze: "0,0", merges: [], rows: {}, cols: {}, cells };
}

function canon(sheets: CanonSheet[]): CanonBook {
  return { sheets, names: [], resources: {} };
}

const TABS = ["Setup", "Data", "Slidesheet", "Cont DI", "DL Tracker"] as const;

describe("visible user-data sync", () => {
  it("treats the same number, and numeric text, as one value, and keeps command text", () => {
    expect(scalarEqual(cell("number", "1.75"), cell("number", "1.750"))).toBe(true);
    expect(scalarEqual(cell("number", "0"), cell("text", "-0"))).toBe(true);
    expect(scalarEqual(cell("number", "0.2708333333333333"), cell("number", "0.27083333333333331"))).toBe(true);
    expect(scalarEqual(cell("number", "1"), cell("number", "1.0001"))).toBe(false);
    expect(scalarEqual(cell("text", "00"), cell("number", "0"))).toBe(false);
    expect(scalarEqual(cell("text", "2-4"), cell("text", "2-4"))).toBe(true);
    expect(scalarEqual(cell("text", "2-4"), cell("number", "2"))).toBe(false);
    expect(scalarEqual(cell("formula", "=A1"), cell("number", "1"))).toBe(false);
    expect(scalarEqual(cell("blank", ""), cell("text", ""))).toBe(true);
  });

  it("emits visible scalar edits and ignores hidden sheets, formulas and identical numbers", () => {
    const base = canon([
      visible("Data", { A1: cell("text", "keep"), C4: cell("number", "1.75"), K4: cell("formula", "=J4-I4") }),
      visible("Setup", { B2: cell("text", "old") }),
      visible("_OC_Survey", { A1: cell("text", "cache") }, 1),
    ]);
    const live = canon([
      visible("Data", { A1: cell("text", "keep"), C4: cell("number", "1.750"), K4: cell("number", "9"), D4: cell("text", "00") }),
      visible("Setup", { B2: cell("text", "new") }, 2),
      visible("_OC_Survey", { A1: cell("text", "changed") }, 1),
    ]);
    const ops = diffVisibleUserData(base, live);
    expect(ops.map((op) => `${op.sheet}!${op.a1}=${op.nextKind}:${op.nextValue}`)).toEqual(["Data!D4=text:00"]);
    expect(ops.every((op) => op.kind === "cell")).toBe(true);
  });

  it("classifies a 3-way merge across the five visible tabs and drops a hidden cache", () => {
    const sheets = TABS.map((name) => visible(name, { A1: cell("text", "base"), B1: cell("number", "1") }));
    const base = canon([...sheets, visible("_OC_Plan", { A1: cell("text", "hidden") }, 2)]);
    const browser = canon([
      visible("Setup", { A1: cell("text", "browser-only"), B1: cell("number", "1") }),
      visible("Data", { A1: cell("text", "base"), B1: cell("number", "1"), L4: cell("number", "12") }),
      visible("Slidesheet", { A1: cell("text", "same"), B1: cell("number", "1") }),
      visible("Cont DI", { A1: cell("text", "from-browser"), B1: cell("number", "1") }),
      visible("DL Tracker", { A1: cell("text", "base"), B1: cell("formula", "=H3"), D3: cell("text", "2-4") }),
      visible("_OC_Plan", { A1: cell("text", "browser-cache") }, 2),
    ]);
    const excel = canon([
      visible("Setup", { A1: cell("text", "base"), B1: cell("number", "1") }),
      visible("Data", { A1: cell("text", "excel-only"), B1: cell("number", "1.000") }),
      visible("Slidesheet", { A1: cell("text", "same"), B1: cell("number", "1") }),
      visible("Cont DI", { A1: cell("text", "from-excel"), B1: cell("number", "1") }),
      visible("DL Tracker", { A1: cell("text", "base"), B1: cell("number", "5"), D3: cell("text", "2-4") }),
      visible("_OC_Plan", { A1: cell("text", "excel-cache") }, 2),
    ]);
    const merge = mergeVisibleUserData(base, browser, excel);
    expect(merge.keepBrowser.map((op) => `${op.sheet}!${op.a1}`)).toEqual(["Setup!A1", "Data!L4"]);
    expect(merge.fromExcel.map((op) => `${op.sheet}!${op.a1}`)).toEqual(["Data!A1"]);
    expect(merge.converged.map((op) => `${op.sheet}!${op.a1}`)).toEqual(["Slidesheet!A1", "DL Tracker!D3"]);
    expect(merge.conflicts.map((item) => `${item.sheet}!${item.a1}:${item.browser.value}/${item.excel.value}`)).toEqual([
      "Cont DI!A1:from-browser/from-excel",
    ]);
    expect(merge.conflicts.some((item) => item.sheet === "_OC_Plan")).toBe(false);
    expect(merge.fromExcel.some((op) => op.a1 === "B1")).toBe(false);
  });

  it("resolves a conflict by advancing the baseline and is idempotent once the books match", () => {
    const base = canon([visible("DL Tracker", { I3: cell("number", "90.45"), J3: cell("number", "133.06") })]);
    const browser = canon([visible("DL Tracker", { I3: cell("number", "90.45"), J3: cell("number", "130.56") })]);
    const excel = canon([visible("DL Tracker", { I3: cell("number", "81"), J3: cell("number", "154.56") })]);
    const first = mergeVisibleUserData(base, browser, excel);
    expect(first.fromExcel.map((op) => `${op.a1}=${op.nextValue}`)).toEqual(["I3=81"]);
    expect(first.conflicts.map((item) => item.a1)).toEqual(["J3"]);
    expect(defaultConflictChoice(first.conflicts[0]!, false)).toBe("excel");
    const dataConflict = {
      sheet: "Data",
      a1: "L4",
      base: { kind: "number" as const, value: "1" },
      excel: { kind: "number" as const, value: "2" },
      browser: { kind: "number" as const, value: "3" },
    };
    expect(machineWrittenCell("Data", "L4")).toBe(true);
    expect(machineWrittenCell("Setup", "A1")).toBe(false);
    expect(defaultConflictChoice(dataConflict, true)).toBe("browser");
    expect(defaultConflictChoice(dataConflict, false)).toBe("excel");

    const taken = canon([visible("DL Tracker", {
      I3: cell("number", "81"),
      J3: cell("number", "130.56"),
    })]);
    const exported = diffVisibleUserData(excel, taken);
    expect(exported.map((op) => `${op.a1}=${op.nextValue}`)).toEqual(["J3=130.56"]);

    const again = mergeVisibleUserData(excel, taken, excel);
    expect(again.fromExcel).toEqual([]);
    expect(again.conflicts).toEqual([]);
    expect(again.keepBrowser.map((op) => op.a1)).toEqual(["J3"]);
    const settled = mergeVisibleUserData(excel, excel, excel);
    expect(settled.fromExcel).toEqual([]);
    expect(settled.keepBrowser).toEqual([]);
    expect(settled.converged).toEqual([]);
    expect(settled.conflicts).toEqual([]);
  });

  it("keeps Setup and Slidesheet strictly Browser-authoritative and excludes derived Data tables", () => {
    const base = canon([
      visible("Setup", { D4: cell("text", "35784"), D5: cell("text", "old-well") }),
      visible("Slidesheet", { T150: cell("blank", ""), Z150: cell("blank", "") }),
      visible("Data", {
        X26: cell("text", "HMJ 625041"),
        O31: cell("text", "HMJ 625041"),
        Q31: cell("number", "10"),
        R31: cell("number", "41.81"),
        O45: cell("text", "PHX-725031-PTS"),
        Q45: cell("number", "5"),
        R45: cell("number", "41.81"),
        B47: cell("text", "PHX-650076-PTS"),
      }),
    ]);
    const browser = canon([
      visible("Setup", { D4: cell("text", "35784"), D5: cell("text", "browser-well") }),
      visible("Slidesheet", { T150: cell("number", "6"), Z150: cell("number", "13.18") }),
      visible("Data", {
        X26: cell("text", "Excalibur 14 X-Over"),
        O31: cell("text", "HMJ 625-50"),
        Q31: cell("number", "10"),
        R31: cell("number", "81.92"),
        O45: cell("text", "24X-19722-PTS"),
        Q45: cell("number", "5"),
        R45: cell("number", "13"),
        B47: cell("text", "PHX-576154-PTS"),
      }),
    ]);
    const excel = canon([
      visible("Setup", { D4: cell("text", "35784"), D5: cell("text", "excel-well") }),
      visible("Slidesheet", { T150: cell("number", "10"), Z150: cell("number", "9.18") }),
      visible("Data", {
        X26: cell("text", "VC650789"),
        O31: cell("text", "1142"),
        Q31: cell("number", "12.5"),
        R31: cell("number", "25.75"),
        O45: cell("text", "PHX-576039-PTS"),
        Q45: cell("number", "5"),
        R45: cell("number", "25.75"),
        B47: cell("blank", ""),
      }),
    ]);

    const merge = mergeVisibleUserData(base, browser, excel);
    // Setup & Slidesheet never pull from Excel and never conflict — browser stays authoritative
    expect(merge.conflicts).toEqual([]);
    expect(merge.fromExcel.map((op) => `${op.sheet}!${op.a1}=${op.nextValue}`)).toEqual(["Data!Q31=12.5"]);
    expect(merge.keepBrowser.map((op) => `${op.sheet}!${op.a1}=${op.nextValue}`)).toEqual([
      "Setup!D5=browser-well",
      "Slidesheet!T150=6",
      "Slidesheet!Z150=13.18",
    ]);
    // Derived Data tables are also excluded from outbound export
    const outbound = diffVisibleUserData(excel, browser);
    expect(outbound.map((op) => `${op.sheet}!${op.a1}`)).toEqual([
      "Setup!D5",
      "Slidesheet!T150",
      "Slidesheet!Z150",
      "Data!Q31",
    ]);
  });

  it("skips all derived Data tables and gives the newer BHA # atomic priority on H3, C22, and H4:Q22", () => {
    const base = canon([
      visible("Data", {
        H3: cell("number", "1"),
        C22: cell("text", "PHX-725031-PTS"),
        E13: cell("text", "2.09m"),
        E14: cell("text", "4.24m"),
        E15: cell("text", "4.73m"),
        L54: cell("number", "-4319.5"),
        M54: cell("number", "2063.75"),
        B35: cell("text", "Offset-A"),
        Z58: cell("number", "5.23"),
        H4: cell("text", "Midnight"),
        I4: cell("blank", ""),
        J4: cell("blank", ""),
        L4: cell("blank", ""),
        N4: cell("blank", ""),
        P4: cell("blank", ""),
        Q4: cell("blank", ""),
        I5: cell("number", "629"),
      }),
    ]);
    const browser = canon([
      visible("Data", {
        H3: cell("number", "1"),
        C22: cell("text", "24X-19722-PTS"),
        E13: cell("text", "1.11m"),
        E14: cell("text", "5.79m"),
        E15: cell("text", "5.89m"),
        L54: cell("number", "-3878.27"),
        M54: cell("number", "743.86"),
        B35: cell("text", "Offset-Browser"),
        Z58: cell("number", "9.99"),
        H4: cell("text", "Midnight"),
        I4: cell("number", "29"),
        J4: cell("number", "629"),
        L4: cell("number", "31"),
        N4: cell("number", "41"),
        P4: cell("number", "5"),
        Q4: cell("number", "13"),
        I5: cell("number", "629"),
      }),
    ]);
    const excel = canon([
      visible("Data", {
        H3: cell("number", "3"),
        C22: cell("text", "PHX-576039-PTS"),
        E13: cell("text", "0.85m"),
        E14: cell("text", "5.66m"),
        E15: cell("text", "5.73m"),
        L54: cell("number", "0"),
        M54: cell("number", "-2020"),
        B35: cell("text", "Offset-Excel"),
        Z58: cell("number", "12.5"),
        H4: cell("text", "Midnight"),
        I4: cell("number", "3161"),
        J4: cell("number", "3161"),
        L4: cell("number", "0"),
        N4: cell("text", "00"),
        P4: cell("number", "0"),
        Q4: cell("number", "0"),
        // I5 is blank in Excel BHA #3 (old BHA #1 row 5 must clear cleanly)
      }),
    ]);

    // Case 1: Excel BHA #3 > Browser BHA #1 -> Excel wins H3, C22, H4:Q22 atomically, zero conflicts, clears stale I5
    const mergeExcelNewer = mergeVisibleUserData(base, browser, excel);
    expect(mergeExcelNewer.conflicts).toEqual([]);
    expect(mergeExcelNewer.keepBrowser).toEqual([]);
    expect(mergeExcelNewer.fromExcel.map((op) => `${op.a1}=${op.nextKind === "blank" ? "" : op.nextValue}`)).toEqual([
      "H3=3",
      "I4=3161",
      "J4=3161",
      "L4=0",
      "N4=00",
      "P4=0",
      "Q4=0",
      "I5=",
      "C22=PHX-576039-PTS",
    ]);

    // Case 2: Browser BHA #4 > Excel BHA #3 -> Browser keeps H3, C22, H4:Q22 atomically, zero conflicts
    const browserBha4 = canon([
      visible("Data", {
        H3: cell("number", "4"),
        C22: cell("text", "PHX-999999-PTS"),
        H4: cell("text", "Midnight"),
        I4: cell("number", "4200"),
        J4: cell("number", "4500"),
        E13: cell("text", "0.10m"),
        L54: cell("number", "-100"),
      }),
    ]);
    const mergeBrowserNewer = mergeVisibleUserData(base, browserBha4, excel);
    expect(mergeBrowserNewer.conflicts).toEqual([]);
    expect(mergeBrowserNewer.fromExcel).toEqual([]);
    expect(mergeBrowserNewer.keepBrowser.map((op) => `${op.a1}=${op.nextKind === "blank" ? "" : op.nextValue}`)).toEqual([
      "H3=4",
      "I4=4200",
      "J4=4500",
      "L4=",
      "N4=",
      "P4=",
      "Q4=",
      "C22=PHX-999999-PTS",
    ]);
  });

  it("stamps a visible-data bundle the workbook can tell from a full one", async () => {
    const files = await buildBundleFiles({
      jobNumber: "Demo",
      sourceWorkbookHash: "abc",
      createdAt: "2026-11-21T00:00:00.000Z",
      exporter: "browser shift",
      mode: "visible-data",
      ops: diffVisibleUserData(
        canon([visible("Data", { A1: cell("text", "a") })]),
        canon([visible("Data", { A1: cell("text", "b") })]),
      ),
    });
    expect(files["manifest.csv"]).toContain("mode,visible-data");
    expect(parseCsv(files["cells.csv"])[1]?.[5]).toBe("b");
    expect(await auditBundle(files)).toEqual([]);
  });
});
