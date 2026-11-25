import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import "fake-indexeddb/auto";
import { BACKUP_DIR, BACKUP_FILE } from "../src/backups.ts";
import { nodeDirectory } from "../src/nodeDirectory.ts";
import { openRegistry } from "../src/registry.ts";
import { openJob } from "../src/session.ts";
import { createEnvelope } from "../src/snapshot.ts";

function envelope(savedAt: string, marker: string) {
  return createEnvelope({
    univerAppVersion: "1.3.1",
    sourceWorkbookHash: marker,
    savedAt,
    workbook: { id: "wb", sheetOrder: ["s"], marker },
    sidecar: { veryHiddenSheets: [] },
  });
}

async function main(): Promise<void> {
  const root = await mkdtemp(path.join(tmpdir(), "phx-folder-smoke-"));
  try {
    const registry = await openRegistry();
    const first = await nodeDirectory(root);
    const opened = await openJob({ jobId: "smoke", label: "Smoke", fileName: "job.phx.json", directory: first, registry });
    if ((await first.queryPermission()) !== "granted") throw new Error("initial grant missing");
    await opened.session.replace(envelope("2026-11-01T00:00:00.000Z", "one"), { backup: true });
    opened.session.dispose();

    const restarted = await nodeDirectory(root);
    const reopened = await openJob({ jobId: "smoke", label: "Smoke", fileName: "job.phx.json", directory: restarted, registry });
    if (reopened.result.envelope?.sourceWorkbookHash !== "one") throw new Error("reconnect did not reload the canonical snapshot");

    await reopened.session.replace(envelope("2026-11-02T00:00:00.000Z", "two"), { backup: true });
    const locked = await restarted.getFile("job.phx.json", false);
    const writer = await locked?.writeExclusive();
    let concurrent = false;
    try {
      await reopened.session.writeFolder(false);
    } catch {
      concurrent = true;
    }
    await writer?.abort();
    if (!concurrent) throw new Error("locked file did not fail the second writer");

    const canonical = await restarted.getFile("job.phx.json", true);
    const corrupt = await canonical?.writeExclusive();
    await corrupt?.write(new TextEncoder().encode("not-json"));
    await corrupt?.close();
    const corruptOpen = await openJob({ jobId: "smoke", label: "Smoke", fileName: "job.phx.json", directory: restarted, registry });
    if (corruptOpen.result.recovery.kind !== "corrupt") throw new Error("corrupt file was not preserved for recovery");

    await registry.putMirror({ jobId: "mirror", savedAt: "2026-11-03T00:00:00.000Z", bytes: "{}" });
    const mirrorDir = await nodeDirectory(root);
    const mirrorFile = await mirrorDir.getFile("mirror.phx.json", true);
    const seed = await mirrorFile?.writeExclusive();
    await seed?.write(new TextEncoder().encode(JSON.stringify(envelope("2026-11-02T00:00:00.000Z", "folder"))));
    await seed?.close();
    const mirrorOpen = await openJob({ jobId: "mirror", label: "Mirror", fileName: "mirror.phx.json", directory: mirrorDir, registry });
    if (mirrorOpen.result.recovery.kind !== "mirror-newer") throw new Error("newer mirror was not offered");

    const rootNames = await restarted.list();
    if (rootNames.some((name) => name.includes("backup"))) throw new Error(`unexpected root backup clutter: ${rootNames.join(", ")}`);
    const openCap = await restarted.getDirectory(BACKUP_DIR, false);
    if (!openCap) throw new Error(`${BACKUP_DIR} folder missing`);
    const openCapNames = await openCap.list();
    if (openCapNames.length !== 1 || openCapNames[0] !== BACKUP_FILE) {
      throw new Error(`expected single ${BACKUP_DIR}/${BACKUP_FILE}, got: ${openCapNames.join(", ")}`);
    }
    process.stdout.write(`folder smoke ok ${root}\n`);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
}

main().catch((error: unknown) => {
  process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
  process.exitCode = 1;
});
