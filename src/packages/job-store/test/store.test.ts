import "fake-indexeddb/auto";
import { describe, expect, it, vi } from "vitest";
import { BACKUP_DIR, BACKUP_FILE, listBackups } from "../src/backups.js";
import { directoryFromHandle, MemoryDirectory } from "../src/directory.js";
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
  it("keeps only a single backup file inside OpenCap/ and never clutters the job root", async () => {
    const directory = new MemoryDirectory();
    const registry = await openRegistry();
    const first = envelope("2026-11-03T00:00:00.000Z");
    const second = envelope("2026-11-03T00:00:02.000Z");
    const third = envelope("2026-11-03T00:00:05.000Z");
    const opened = await openJob({ jobId: "job", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    await opened.session.replace(first, { backup: true });
    await opened.session.replace(second, { backup: true });
    await opened.session.replace(third, { backup: true });
    const again = await openJob({ jobId: "job", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    expect(again.result.envelope?.savedAt).toBe("2026-11-03T00:00:05.000Z");
    expect(again.result.recovery.kind).toBe("none");

    const rootNames = await directory.list();
    expect(rootNames).toEqual(["job.phx.json"]);
    const openCap = await directory.getDirectory(BACKUP_DIR, false);
    expect(openCap).not.toBeNull();
    expect(await openCap!.list()).toEqual([BACKUP_FILE]);
    const backupBytes = await (await openCap!.getFile(BACKUP_FILE, false))!.read();
    expect(parseEnvelope(backupBytes!).savedAt).toBe("2026-11-03T00:00:02.000Z");
    expect(await listBackups(directory)).toEqual([`${BACKUP_DIR}/${BACKUP_FILE}`]);
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
    expect(await directory.getDirectory(BACKUP_DIR, false)).toBeNull();
    expect(opened.session.state).toBe("folder-safe");
  });

  it("keeps a corrupt canonical and offers the single OpenCap backup", async () => {
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
    if (opened.result.recovery.kind === "corrupt") {
      expect(opened.result.recovery.backups).toEqual([`${BACKUP_DIR}/${BACKUP_FILE}`]);
    }
    const still = await file?.read();
    expect(new TextDecoder().decode(still ?? new Uint8Array())).toBe("{");
  });

  it("migrates legacy root multi-backups into a single OpenCap/.phx-backup.json on open", async () => {
    const directory = new MemoryDirectory();
    const registry = await openRegistry();
    const writeRaw = async (dir: MemoryDirectory, name: string, bytes: Uint8Array) => {
      const f = await dir.getFile(name, true);
      const w = await f!.writeExclusive();
      await w.write(bytes);
      await w.close();
    };
    await writeRaw(directory, "job.phx.json", encodeEnvelope(envelope("2026-11-05T00:00:00.000Z")));
    await writeRaw(directory, ".phx-backup-weekly-2026-W44.json", encodeEnvelope(envelope("2026-11-01T00:00:00.000Z")));
    await writeRaw(directory, ".phx-backup-daily-2026-11-03.json", encodeEnvelope(envelope("2026-11-03T00:00:00.000Z")));
    await writeRaw(directory, ".phx-backup-recent-2026-11-04T01-00-00-000Z.json", encodeEnvelope(envelope("2026-11-04T01:00:00.000Z")));
    await writeRaw(directory, ".phx-backup-recent-2026-11-04T05-00-00-000Z.json", encodeEnvelope(envelope("2026-11-04T05:00:00.000Z")));

    const opened = await openJob({ jobId: "job-legacy", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    expect(opened.result.recovery.kind).toBe("none");
    expect(await directory.list()).toEqual(["job.phx.json"]);
    const openCap = await directory.getDirectory(BACKUP_DIR, false);
    expect(openCap).not.toBeNull();
    expect(await openCap!.list()).toEqual([BACKUP_FILE]);
    const migrated = await (await openCap!.getFile(BACKUP_FILE, false))!.read();
    expect(parseEnvelope(migrated!).savedAt).toBe("2026-11-04T05:00:00.000Z");
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

  it("rethrows NotFoundError when create=true so a disconnected folder is reported instead of null", async () => {
    const notFound = new DOMException("The path supplied exists, but was not an entry of requested type.", "NotFoundError");
    const directory = directoryFromHandle({
      getFileHandle: async () => { throw notFound; },
      getDirectoryHandle: async () => { throw notFound; },
      entries: async function* () {},
      queryPermission: async () => "granted",
      requestPermission: async () => "granted",
    });
    expect(await directory.getFile("job.phx.json", false)).toBeNull();
    expect(await directory.getDirectory("OpenCap", false)).toBeNull();
    await expect(directory.getFile("job.phx.json", true)).rejects.toBe(notFound);
    await expect(directory.getDirectory("OpenCap", true)).rejects.toBe(notFound);
  });

  it("transitions to save-error and notifies onStateChange without unhandled rejection when background autosave fails", async () => {
    const notFound = new DOMException("Folder gone", "NotFoundError");
    const directory = directoryFromHandle({
      getFileHandle: async () => { throw notFound; },
      entries: async function* () {},
      queryPermission: async () => "granted",
      requestPermission: async () => "granted",
    });
    const registry = await openRegistry();
    const opened = await openJob({ jobId: "job-gone", label: "Gone", fileName: "job.phx.json", directory, registry });
    const seen: Array<{ state: string; error: unknown }> = [];
    opened.session.onStateChange = (state, error) => seen.push({ state, error });
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    try {
      opened.session.noteEdit(envelope("2026-11-03T00:00:01.000Z"));
      await vi.advanceTimersByTimeAsync(30_000);
      expect(opened.session.state).toBe("save-error");
      expect(seen.some((entry) => entry.state === "save-error" && entry.error === notFound)).toBe(true);
      opened.session.dispose();
    } finally {
      vi.useRealTimers();
    }
  });
});
