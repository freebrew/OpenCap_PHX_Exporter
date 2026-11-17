import { createDraft, decodeBase64, splitAddresses, toMime, type MailDraft } from "../../mail/src/index";

export interface MailFile {
  name: string;
  /** Base64 of the file bytes. Empty when the file was named but not found. */
  base64: string;
  missing?: boolean;
}

export interface MailMessage {
  to: string;
  cc?: string;
  subject: string;
  html: string;
  pngBase64?: string;
  /** Attachment 1-5, the files named on Data. Missing ones are listed in the body, not attached. */
  files?: MailFile[];
}

export function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

export function dailyText(fields: Record<string, string>): string {
  return Object.entries(fields).map(([label, value]) => `${label}: ${value}`).join("\n");
}

function plainText(html: string): string {
  return html.replace(/<br\s*\/?>/gi, "\n").replace(/<[^>]+>/g, " ").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

function fromLegacy(message: MailMessage): MailDraft {
  const files = (message.files ?? []).filter((file) => file.base64 !== "");
  const missing = (message.files ?? []).filter((file) => file.missing).map((file) => file.name);
  const note = missing.length > 0 ? `<p>Not in the job folder: ${missing.map(escapeHtml).join(", ")}</p>` : "";
  const picture = message.pngBase64 ? `<p><img src="cid:daily_report.png" alt="Corridor"></p>` : "";
  const html = `<!doctype html><html><body>${message.html}${picture}${note}</body></html>`;
  const png = message.pngBase64 ? decodeBase64(message.pngBase64) : null;
  return createDraft({
    to: splitAddresses(message.to),
    cc: splitAddresses(message.cc ?? ""),
    subject: message.subject,
    html,
    text: plainText(html),
    inline: png ? [{ name: "daily_report_inline.png", contentType: "image/png", bytes: png, contentId: "daily_report.png" }] : [],
    attachments: [
      ...(png ? [{ name: "daily_report.png", contentType: "image/png", bytes: png }] : []),
      ...files.map((file) => ({ name: file.name, contentType: "application/octet-stream", bytes: decodeBase64(file.base64) })),
    ],
    missing: missing.map((name) => ({ name })),
  });
}

/** Archival message. The canonical serializer owns the bytes; this keeps the old call shape. */
export function buildEml(message: MailMessage): string {
  return toMime(fromLegacy(message));
}

export function emlBytes(message: MailMessage): Uint8Array {
  return new TextEncoder().encode(buildEml(message));
}
