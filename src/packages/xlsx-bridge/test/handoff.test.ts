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
} from "../src/handoff.js";

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
