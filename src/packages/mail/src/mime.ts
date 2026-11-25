import { encodeBase64, wrapBase64 } from "./bytes";
import type { MailBytes, MailDraft } from "./draft";
import { stripHeaderControls } from "./validate";

export interface MimeBoundaries {
  mixed: string;
  related: string;
  alternative: string;
}

function crlf(value: string): string {
  return value.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n/g, "\r\n");
}

function fnv(text: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function fileToken(name: string): string {
  const leaf = stripHeaderControls(name).split(/[\\/]/).pop() ?? "file";
  const clean = leaf.replace(/["\r\n]/g, "").replace(/[^A-Za-z0-9._ -]/g, "_").slice(0, 80);
  return clean === "" || clean === "." || clean === ".." ? "file" : clean;
}

function encodeHeader(value: string): string {
  const clean = stripHeaderControls(value);
  if (/^[\u0020-\u007E]*$/.test(clean)) return clean;
  const words: string[] = [];
  let piece = "";
  for (const char of clean) {
    const next = piece + char;
    if (encodeBase64(new TextEncoder().encode(next)).length > 45 && piece !== "") {
      words.push(`=?UTF-8?B?${encodeBase64(new TextEncoder().encode(piece))}?=`);
      piece = char;
    } else {
      piece = next;
    }
  }
  if (piece !== "") words.push(`=?UTF-8?B?${encodeBase64(new TextEncoder().encode(piece))}?=`);
  return words.join("\r\n ");
}

function defaultBoundaries(draft: MailDraft): MimeBoundaries {
  const seed = `${draft.subject}\n${draft.html}\n${draft.text}`;
  const pick = (label: string) => {
    let value = `phx_${label}_${fnv(seed + label)}`;
    while (seed.includes(value)) value += "x";
    return value;
  };
  return { mixed: pick("mixed"), related: pick("related"), alternative: pick("alt") };
}

function partHeaders(part: MailBytes, disposition: "inline" | "attachment"): string[] {
  const name = fileToken(part.name);
  const headers = [
    `Content-Type: ${part.contentType || "application/octet-stream"}`,
    "Content-Transfer-Encoding: base64",
    `Content-Disposition: ${disposition}; filename="${name}"`,
  ];
  if (disposition === "inline" && part.contentId) headers.push(`Content-ID: <${stripHeaderControls(part.contentId)}>`);
  return headers;
}

function pushFile(lines: string[], boundary: string, part: MailBytes, disposition: "inline" | "attachment"): void {
  lines.push(`--${boundary}`, ...partHeaders(part, disposition), "", wrapBase64(encodeBase64(part.bytes)));
}

/** Canonical archive and Gmail body. CRLF, 76-column base64, visible files and a separate inline CID part. */
export function toMime(draft: MailDraft, boundaries = defaultBoundaries(draft)): string {
  const lines = [
    `To: ${draft.to.map(stripHeaderControls).join(", ")}`,
    `Cc: ${draft.cc.map(stripHeaderControls).join(", ")}`,
    `Subject: ${encodeHeader(draft.subject)}`,
    "MIME-Version: 1.0",
    "X-PHX-Preview: evaluation",
    `Content-Type: multipart/mixed; boundary="${boundaries.mixed}"`,
    "",
    `--${boundaries.mixed}`,
    `Content-Type: multipart/related; boundary="${boundaries.related}"`,
    "",
    `--${boundaries.related}`,
    `Content-Type: multipart/alternative; boundary="${boundaries.alternative}"`,
    "",
    `--${boundaries.alternative}`,
    "Content-Type: text/plain; charset=utf-8",
    "Content-Transfer-Encoding: 8bit",
    "",
    crlf(draft.text),
    `--${boundaries.alternative}`,
    "Content-Type: text/html; charset=utf-8",
    "Content-Transfer-Encoding: 8bit",
    "",
    crlf(draft.html),
    `--${boundaries.alternative}--`,
  ];
  for (const part of draft.inline) pushFile(lines, boundaries.related, part, "inline");
  lines.push(`--${boundaries.related}--`);
  for (const part of draft.attachments) pushFile(lines, boundaries.mixed, part, "attachment");
  lines.push(`--${boundaries.mixed}--`, "");
  return lines.join("\r\n");
}

/** Editable desktop-client draft (.eml). X-Unsent: 1 opens as an unsent compose in Classic Outlook and New Outlook. */
export function toDraftMime(draft: MailDraft, boundaries = defaultBoundaries(draft)): string {
  return `X-Unsent: 1\r\n${toMime(draft, boundaries)}`;
}
