import { encodeBase64 } from "./bytes";

export interface PdfMailFile {
  name: string;
  pdf: Uint8Array;
}

export interface PdfMailMessage {
  to: string[];
  subject: string;
  text: string;
  files: PdfMailFile[];
}

function encodeHeader(text: string): string {
  return /^[\u0020-\u007E]*$/.test(text) ? text : `=?UTF-8?B?${encodeBase64(new TextEncoder().encode(text))}?=`;
}

/** StrapSheets direct-send body. The boundary is injectable so the existing golden stays exact. */
export function pdfTextMime(message: PdfMailMessage, boundary = `strap-${crypto.randomUUID()}`): string {
  const lines = [
    `To: ${message.to.join(", ")}`,
    `Subject: ${encodeHeader(message.subject)}`,
    "MIME-Version: 1.0",
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    "",
    `--${boundary}`,
    "Content-Type: text/plain; charset=UTF-8",
    "",
    message.text,
  ];
  for (const file of message.files) {
    lines.push(
      `--${boundary}`,
      `Content-Type: application/pdf; name="${file.name}"`,
      `Content-Disposition: attachment; filename="${file.name}"`,
      "Content-Transfer-Encoding: base64",
      "",
      ...(encodeBase64(file.pdf).match(/.{1,76}/g) ?? []),
    );
  }
  lines.push(`--${boundary}--`, "");
  return lines.join("\r\n");
}
