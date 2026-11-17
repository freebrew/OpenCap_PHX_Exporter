import "fake-indexeddb/auto";
import { describe, expect, it } from "vitest";
import { MemoryDirectory } from "../src/directory.js";
import { openRegistry } from "../src/registry.js";
import { openJob } from "../src/session.js";
import { createEnvelope, encodeEnvelope, parseEnvelope } from "../src/snapshot.js";
import { StoreError } from "../src/types.js";

function envelope(savedAt: string) {
  return createEnvelope({
    univerAppVersion: "1.3.1",
    sourceWorkbookHash: "abc",
    savedAt,
    workbook: { id: "wb", sheetOrder: ["s"], resources: [{ name: "custom-plugin", data: "{}" }] },
    sidecar: { veryHiddenSheets: ["Cache"] },
  });
}

describe("snapshot envelope", () => {
  it("round-trips known fields and preserves unknown workbook resources", () => {
    const encoded = encodeEnvelope(envelope("2026-11-03T00:00:00.000Z"));
    const parsed = parseEnvelope(encoded);
    expect(parsed.schemaVersion).toBe(1);
    expect(parsed.sidecar).toEqual({ veryHiddenSheets: ["Cache"] });
    expect((parsed.workbook as { resources: unknown[] }).resources).toHaveLength(1);
  });

  it("keeps the import baseline on a schema 1 snapshot", () => {
    const baseline = { sheets: [], names: [], resources: {} };
    const encoded = encodeEnvelope(createEnvelope({
      univerAppVersion: "1.3.1",
      sourceWorkbookHash: "abc",
      workbook: { id: "wb" },
      sidecar: {},
      importBaseline: baseline,
      handoffBlocks: ["sheet.command.insert-row-before"],
    }));
    const parsed = parseEnvelope(encoded);
    expect(parsed.importBaseline).toEqual(baseline);
    expect(parsed.handoffBlocks).toEqual(["sheet.command.insert-row-before"]);
    const legacy = parseEnvelope(encodeEnvelope(envelope("2026-11-03T00:00:00.000Z")));
    expect(legacy.importBaseline).toBeUndefined();
    expect(legacy.handoffBlocks).toBeUndefined();
  });

  it("rejects a future schema and a corrupt payload", () => {
    const future = new TextEncoder().encode(JSON.stringify({ schemaVersion: 9, workbook: {} }));
    expect(() => parseEnvelope(future)).toThrow(StoreError);
    expect(() => parseEnvelope(new TextEncoder().encode("{"))).toThrow(StoreError);
  });
});

describe("job session", () => {
  it("saves exclusively, reloads, and restores a backup after a bad verification", async () => {
    const directory = new MemoryDirectory();
    const registry = await openRegistry();
    const first = envelope("2026-11-03T00:00:00.000Z");
    const opened = await openJob({ jobId: "job", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    await opened.session.replace(first, { backup: true });
    await opened.session.replace(envelope("2026-11-03T00:00:02.000Z"), { backup: true });
    const again = await openJob({ jobId: "job", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    expect(again.result.envelope?.savedAt).toBe("2026-11-03T00:00:02.000Z");
    expect(again.result.recovery.kind).toBe("none");
    const names = await directory.list();
    expect(names.some((name) => name.includes("-recent-"))).toBe(true);
  });

  it("does not create a backup for the automatic folder write", async () => {
    const directory = new MemoryDirectory();
    const registry = await openRegistry();
    const opened = await openJob({ jobId: "job2", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    opened.session.noteEdit(envelope("2026-11-03T00:00:01.000Z"));
    await opened.session.writeMirror();
    await opened.session.writeFolder(false);
    const names = await directory.list();
    expect(names.some((name) => name.includes("backup"))).toBe(false);
    expect(opened.session.state).toBe("folder-safe");
  });

  it("keeps a corrupt canonical and offers backups", async () => {
    const directory = new MemoryDirectory();
    const registry = await openRegistry();
    const seeded = await openJob({ jobId: "job3", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    await seeded.session.replace(envelope("2026-11-03T00:00:00.000Z"), { backup: true });
    await seeded.session.replace(envelope("2026-11-03T00:00:02.000Z"), { backup: true });
    const file = await directory.getFile("job.phx.json", true);
    const writer = await file?.writeExclusive();
    await writer?.write(new TextEncoder().encode("{"));
    await writer?.close();
    const opened = await openJob({ jobId: "job3", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    expect(opened.result.recovery.kind).toBe("corrupt");
    if (opened.result.recovery.kind === "corrupt") expect(opened.result.recovery.backups.length).toBeGreaterThan(0);
    const still = await file?.read();
    expect(new TextDecoder().decode(still ?? new Uint8Array())).toBe("{");
  });

  it("reports a newer mirror without overwriting the folder", async () => {
    const directory = new MemoryDirectory();
    const registry = await openRegistry();
    const opened = await openJob({ jobId: "job4", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    await opened.session.replace(envelope("2026-11-03T00:00:00.000Z"), { backup: false });
    await registry.putMirror({ jobId: "job4", savedAt: "2026-11-03T01:00:00.000Z", bytes: "{}" });
    const again = await openJob({ jobId: "job4", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    expect(again.result.recovery.kind).toBe("mirror-newer");
    expect(again.result.envelope?.savedAt).toBe("2026-11-03T00:00:00.000Z");
  });

  it("blocks a second exclusive writer", async () => {
    const directory = new MemoryDirectory();
    const file = await directory.getFile("job.phx.json", true);
    const first = await file?.writeExclusive();
    await expect(file?.writeExclusive()).rejects.toBeInstanceOf(StoreError);
    await first?.abort();
  });
});
