// The browser safety copy ("mirror") of a job's snapshot, as it is kept in IndexedDB.
//
// It is gzipped. Kept as text it cost two bytes a character, and IndexedDB refuses a value over
// 127 MiB: a 66 MB snapshot sat at that limit, and one over it left the write hanging.
import type { MirrorRecord } from "./types.js";

async function throughStream(bytes: Uint8Array, stream: CompressionStream | DecompressionStream): Promise<Uint8Array> {
  const piped = new Blob([bytes as Uint8Array<ArrayBuffer>]).stream().pipeThrough(stream);
  return new Uint8Array(await new Response(piped).arrayBuffer());
}

/** The record to store for a snapshot file's bytes. */
export async function packMirror(jobId: string, savedAt: string, snapshot: Uint8Array): Promise<MirrorRecord> {
  return { jobId, savedAt, gzip: await throughStream(snapshot, new CompressionStream("gzip")) };
}

/** The snapshot file's bytes held by a stored record, whichever preview wrote it. */
export async function mirrorSnapshot(record: MirrorRecord): Promise<Uint8Array> {
  if (record.gzip) return throughStream(record.gzip, new DecompressionStream("gzip"));
  return new TextEncoder().encode(record.bytes ?? "");
}
