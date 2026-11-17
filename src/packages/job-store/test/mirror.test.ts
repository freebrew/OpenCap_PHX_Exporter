import "fake-indexeddb/auto";
import { describe, expect, it } from "vitest";
import { MemoryDirectory } from "../src/directory.js";
import { mirrorSnapshot, packMirror } from "../src/mirror.js";
import { openRegistry, requestToPromise, type JobRegistry } from "../src/registry.js";
import { openJob } from "../src/session.js";
import { createEnvelope, encodeEnvelope, parseEnvelope } from "../src/snapshot.js";

function envelope(savedAt: string, rows = 1) {
  return createEnvelope({
    univerAppVersion: "1.3.1",
    sourceWorkbookHash: "abc",
    savedAt,
    // A character outside Latin-1, as real workbooks have: IndexedDB then stores text at two bytes a character.
    workbook: { id: "wb", sheetOrder: ["s"], cells: Array.from({ length: rows }, (_, row) => ({ row, v: "12.5 m — slide", s: "style-7" })) },
    sidecar: {},
  });
}

describe("browser safety copy", () => {
  it("is stored gzipped and reads back byte for byte", async () => {
    const snapshot = encodeEnvelope(envelope("2026-11-03T00:00:00.000Z", 5_000));
    const record = await packMirror("job", "2026-11-03T00:00:00.000Z", snapshot);
    expect(record.bytes).toBeUndefined();
    expect(record.gzip!.byteLength).toBeLessThan(snapshot.byteLength / 10);
    expect(await mirrorSnapshot(record)).toEqual(snapshot);
  });

  it("still reads a copy written as text by an older preview", async () => {
    const text = new TextDecoder().decode(encodeEnvelope(envelope("2026-11-03T00:00:00.000Z")));
    const restored = parseEnvelope(await mirrorSnapshot({ jobId: "job", savedAt: "2026-11-03T00:00:00.000Z", bytes: text }));
    expect(restored.savedAt).toBe("2026-11-03T00:00:00.000Z");
  });

  it("goes through IndexedDB and restores the workbook", async () => {
    const registry = await openRegistry();
    const opened = await openJob({ jobId: "mirror-job", label: "Willowbend", fileName: "job.phx.json", directory: new MemoryDirectory(), registry });
    opened.session.noteEdit(envelope("2026-11-03T00:00:05.000Z", 200));
    await opened.session.writeMirror();
    opened.session.dispose();
    expect(opened.session.state).toBe("mirror-safe");
    expect(opened.session.mirrorError).toBe("");
    const stored = await registry.getMirror("mirror-job");
    expect(stored?.gzip).toBeInstanceOf(Uint8Array);
    const restored = parseEnvelope(await mirrorSnapshot(stored!));
    expect((restored.workbook as { cells: unknown[] }).cells).toHaveLength(200);
  });
});

describe("a safety copy that cannot be written", () => {
  function refusing(registry: JobRegistry): JobRegistry {
    return { ...registry, putMirror: async () => { throw new Error("The serialized keys and/or value are too large"); } };
  }

  it("does not stop the import or Save now from reaching the folder", async () => {
    const directory = new MemoryDirectory();
    const registry = refusing(await openRegistry());
    const opened = await openJob({ jobId: "refused", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    await opened.session.replace(envelope("2026-11-03T00:00:00.000Z"), { backup: true });
    expect(opened.session.state).toBe("folder-safe");
    expect(opened.session.mirrorError).toContain("too large");
    opened.session.noteEdit(envelope("2026-11-03T00:00:09.000Z"));
    await opened.session.flushFolder();
    expect(opened.session.state).toBe("folder-safe");
    const again = await openJob({ jobId: "refused", label: "Willowbend", fileName: "job.phx.json", directory, registry });
    expect(again.result.envelope?.savedAt).toBe("2026-11-03T00:00:09.000Z");
  });

  it("leaves the edit marked unsaved until the folder has it", async () => {
    const registry = refusing(await openRegistry());
    const opened = await openJob({ jobId: "refused-2", label: "Willowbend", fileName: "job.phx.json", directory: new MemoryDirectory(), registry });
    opened.session.noteEdit(envelope("2026-11-03T00:00:01.000Z"));
    await opened.session.writeMirror();
    opened.session.dispose();
    expect(opened.session.state).toBe("dirty");
    expect(opened.session.dirty()).toBe(true);
  });
});

describe("requestToPromise", () => {
  function fakeRequest() {
    const transaction = Object.assign(new EventTarget(), { error: new DOMException("The serialized keys and/or value are too large", "UnknownError") });
    const request = { transaction, result: undefined, error: null, onsuccess: null, onerror: null } as unknown as IDBRequest<undefined>;
    return { request, transaction };
  }

  it("rejects when Chrome aborts the transaction without failing the request", async () => {
    const { request, transaction } = fakeRequest();
    const pending = requestToPromise(request);
    transaction.dispatchEvent(new Event("abort"));
    await expect(pending).rejects.toThrow("too large");
  });

  it("resolves with the request result", async () => {
    const { request } = fakeRequest();
    const pending = requestToPromise(request);
    request.onsuccess?.call(request, new Event("success"));
    await expect(pending).resolves.toBeUndefined();
  });
});
