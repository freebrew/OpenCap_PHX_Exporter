import { encodeBase64, outlookList, type MailDraft, type TransportResult } from "@phx/mail";
import { NATIVE_HOST_NAME } from "../../extensionIdentity";

const CHUNK_BYTES = 48_000;

export interface HostReply {
  ok: boolean;
  command?: string;
  error?: string;
  code?: string;
  displayed?: boolean;
  created?: boolean;
  skipped?: boolean;
  protocol?: number;
  hostVersion?: string;
  outlookRegistered?: boolean;
  outlookRunning?: boolean;
  mailtoCommand?: string;
  defaultClient?: string;
}

export interface HostPort {
  postMessage(message: unknown): void;
  onMessage: {
    addListener(listener: (message: HostReply) => void): void;
    removeListener(listener: (message: HostReply) => void): void;
  };
  onDisconnect: {
    addListener(listener: () => void): void;
    removeListener(listener: () => void): void;
  };
  disconnect(): void;
}

export interface HostFile {
  fileId: string;
  name: string;
  role: "attachment" | "visible-png" | "inline-png";
  bytes: Uint8Array;
}

function hex(bytes: number): string {
  const data = crypto.getRandomValues(new Uint8Array(bytes));
  return [...data].map((value) => value.toString(16).padStart(2, "0")).join("");
}

function replyResult(reply: HostReply): TransportResult | null {
  if (reply.ok && reply.displayed) return { status: "created-and-opened", location: "Classic Outlook" };
  if (reply.ok && reply.skipped) return { status: "no-side-effect-failure", error: "The helper did not open Outlook." };
  if (!reply.ok && reply.code === "indeterminate") return { status: "indeterminate", error: reply.error ?? "Outlook did not confirm the compose." };
  if (!reply.ok && reply.created) return { status: "created-not-opened", id: "classic", location: "Classic Outlook", error: reply.error ?? "The compose did not stay open." };
  if (!reply.ok) return { status: "no-side-effect-failure", error: reply.error ?? "The Outlook helper refused the report." };
  return null;
}

function roundTrip(port: HostPort, message: unknown): Promise<HostReply> {
  return new Promise((resolve, reject) => {
    const finish = () => {
      clearTimeout(timer);
      port.onMessage.removeListener(onMessage);
      port.onDisconnect.removeListener(onDisconnect);
    };
    const timer = setTimeout(() => {
      finish();
      reject(new Error("The Outlook helper did not answer."));
    }, 20_000);
    const onMessage = (reply: HostReply) => {
      finish();
      resolve(reply);
    };
    const onDisconnect = () => {
      const runtimeError = chrome.runtime.lastError?.message;
      finish();
      reject(new Error(runtimeError || "The Outlook helper closed."));
    };
    port.onMessage.addListener(onMessage);
    port.onDisconnect.addListener(onDisconnect);
    port.postMessage(message);
  });
}

export function hostFiles(draft: MailDraft): HostFile[] {
  const files: HostFile[] = [];
  for (const part of draft.inline) files.push({ fileId: hex(4), name: part.name, role: "inline-png", bytes: part.bytes });
  for (const part of draft.attachments) {
    const role = part.name === "daily_report.png" && part.contentType === "image/png" ? "visible-png" : "attachment";
    files.push({ fileId: hex(4), name: part.name, role, bytes: part.bytes });
  }
  return files;
}

/** Chunked compose. A failure before commit has no Outlook item. Commit is sent once. */
export async function composeOnPort(port: HostPort, draft: MailDraft, origin: string, onOpening: () => void): Promise<TransportResult> {
  const runId = hex(16);
  let commitSent = false;
  let disconnected: TransportResult | null = null;
  port.onDisconnect.addListener(() => {
    disconnected = commitSent
      ? { status: "indeterminate", error: "The Outlook helper closed after the compose was requested. Check Outlook before trying again." }
      : { status: "no-side-effect-failure", error: "The Outlook helper closed before opening a compose." };
  });
  const send = async (message: unknown, committing = false): Promise<TransportResult | null> => {
    if (disconnected) return disconnected;
    if (committing) commitSent = true;
    const reply = await roundTrip(port, { ...message as object, origin, runId });
    return replyResult(reply);
  };
  try {
    const begun = await send({ command: "compose.begin", subject: draft.subject, to: outlookList(draft.to), cc: outlookList(draft.cc), html: draft.html });
    if (begun) return begun;
    for (const file of hostFiles(draft)) {
      const count = Math.max(1, Math.ceil(file.bytes.byteLength / CHUNK_BYTES));
      for (let index = 0; index < count; index += 1) {
        const slice = file.bytes.subarray(index * CHUNK_BYTES, (index + 1) * CHUNK_BYTES);
        const chunked = await send({ command: "file.chunk", fileId: file.fileId, name: file.name, role: file.role, index, count, data: encodeBase64(slice) });
        if (chunked) {
          await send({ command: "cancel" }).catch(() => undefined);
          return chunked;
        }
      }
    }
    onOpening();
    const committed = await send({ command: "compose.commit" }, true);
    return committed ?? { status: "indeterminate", error: "The Outlook helper did not confirm the compose." };
  } catch (error) {
    if (commitSent) return { status: "indeterminate", error: error instanceof Error ? error.message : "The Outlook helper did not answer." };
    await send({ command: "cancel" }).catch(() => undefined);
    return { status: "no-side-effect-failure", error: error instanceof Error ? error.message : "The Outlook helper did not answer." };
  } finally {
    port.disconnect();
  }
}

export function connectClassic(): HostPort {
  return chrome.runtime.connectNative(NATIVE_HOST_NAME);
}

export async function probeClassic(port: HostPort, origin: string): Promise<HostReply> {
  try {
    return await roundTrip(port, { command: "probe", origin });
  } finally {
    port.disconnect();
  }
}
