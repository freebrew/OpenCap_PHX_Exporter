import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const script = fileURLToPath(new URL("../phx-outlook-host.ps1", import.meta.url));
const origin = "chrome-extension://glkcmpmiomoppdkmdcjaepneimgbipdl/";

function encode(message) {
  const body = Buffer.from(JSON.stringify(message));
  const header = Buffer.alloc(4);
  header.writeUInt32LE(body.length, 0);
  return Buffer.concat([header, body]);
}

function start(extra = {}) {
  const child = spawn("powershell.exe", ["-NoProfile", "-Sta", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script], {
    env: { ...process.env, PHX_MAIL_ORIGIN: origin, PHX_MAIL_SKIP_OUTLOOK: "1", ...extra },
    stdio: ["pipe", "pipe", "pipe"],
  });
  let stderr = "";
  child.stderr.on("data", (chunk) => { stderr += chunk.toString(); });
  const pending = [];
  let buffer = Buffer.alloc(0);
  child.stdout.on("data", (chunk) => {
    buffer = Buffer.concat([buffer, chunk]);
    while (buffer.length >= 4) {
      const length = buffer.readUInt32LE(0);
      if (buffer.length < 4 + length) return;
      const json = buffer.subarray(4, 4 + length).toString("utf8");
      buffer = buffer.subarray(4 + length);
      const waiter = pending.shift();
      if (!waiter) throw new Error(`unexpected host message ${json}`);
      waiter(JSON.parse(json));
    }
  });
  return {
    stderr: () => stderr,
    send(message) {
      child.stdin.write(encode({ origin, ...message }));
      return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error(`no reply for ${message.command}. ${stderr}`)), 15000);
        pending.push((reply) => {
          clearTimeout(timer);
          resolve(reply);
        });
      });
    },
    close() {
      child.stdin.end();
      child.kill();
    },
  };
}

const host = start();
try {
  const probe = await host.send({ command: "probe" });
  if (probe.ok !== true || probe.protocol !== 1 || typeof probe.outlookRegistered !== "boolean") {
    throw new Error(`probe failed ${JSON.stringify(probe)}`);
  }
  const rejected = await host.send({ command: "probe", origin: "chrome-extension://aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa/" });
  if (rejected.ok !== false || rejected.code !== "origin") throw new Error(`origin was accepted ${JSON.stringify(rejected)}`);

  const runId = "a".repeat(32);
  const early = await host.send({ command: "file.chunk", runId, fileId: "bbbbbbbb", name: "note.txt", role: "attachment", index: 1, count: 1, data: "YQ==" });
  if (early.ok !== false) throw new Error(`chunk before begin was accepted ${JSON.stringify(early)}`);

  const begun = await host.send({ command: "compose.begin", runId, subject: "Day", to: "a@example.com", cc: "", html: "<p>Hi</p>" });
  if (begun.ok !== true) throw new Error(`begin failed ${JSON.stringify(begun)}`);
  const unordered = await host.send({ command: "file.chunk", runId, fileId: "cccccccc", name: "note.txt", role: "attachment", index: 1, count: 2, data: "YQ==" });
  if (unordered.ok !== false) throw new Error(`out-of-order chunk was accepted ${JSON.stringify(unordered)}`);

  const again = await host.send({ command: "compose.begin", runId: "b".repeat(32), subject: "Day", to: "", cc: "", html: "<p>Hi</p>" });
  if (again.ok !== true) throw new Error(`second begin failed ${JSON.stringify(again)}`);
  const folder = join(process.env.TEMP ?? tmpdir(), "phx-mail", "b".repeat(32));
  const chunk = await host.send({ command: "file.chunk", runId: "b".repeat(32), fileId: "dddddddd", name: "note.txt", role: "attachment", index: 0, count: 1, data: Buffer.from("hello").toString("base64") });
  if (chunk.ok !== true) throw new Error(`chunk failed ${JSON.stringify(chunk)}`);
  const { existsSync } = await import("node:fs");
  if (!existsSync(join(folder, "note.txt"))) throw new Error("chunk was not written");
  const canceled = await host.send({ command: "cancel", runId: "b".repeat(32) });
  if (canceled.ok !== true || existsSync(folder)) throw new Error("cancel did not remove the temp folder");

  const small = start({ PHX_MAIL_MAX_BYTES: "4" });
  try {
    const began = await small.send({ command: "compose.begin", runId: "c".repeat(32), subject: "Day", to: "", cc: "", html: "<p>Hi</p>" });
    if (began.ok !== true) throw new Error(`limited begin failed ${JSON.stringify(began)}`);
    const oversized = await small.send({ command: "file.chunk", runId: "c".repeat(32), fileId: "eeeeeeee", name: "note.txt", role: "attachment", index: 0, count: 1, data: Buffer.from("hello").toString("base64") });
    if (oversized.ok !== false) throw new Error(`oversize chunk was accepted ${JSON.stringify(oversized)}`);
  } finally {
    small.close();
  }

  const commitId = "d".repeat(32);
  const ready = await host.send({ command: "compose.begin", runId: commitId, subject: "Day", to: "a@example.com", cc: "b@example.com", html: "<p>Hi</p>" });
  if (ready.ok !== true) throw new Error(`commit begin failed ${JSON.stringify(ready)}`);
  const committed = await host.send({ command: "compose.commit", runId: commitId });
  if (committed.skipped !== true || existsSync(join(process.env.TEMP ?? tmpdir(), "phx-mail", commitId))) {
    throw new Error(`commit did not skip Outlook and clean up ${JSON.stringify(committed)}`);
  }
  console.log("native host protocol ok");
} finally {
  host.close();
}
