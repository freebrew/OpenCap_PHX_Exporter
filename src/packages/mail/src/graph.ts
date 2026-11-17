import { encodeBase64 } from "./bytes";
import type { MailBytes, MailDraft } from "./draft";

export const GRAPH_SMALL_LIMIT = 3 * 1024 * 1024;
export const GRAPH_FILE_LIMIT = 150 * 1024 * 1024;

export interface GraphRequest {
  method: "POST" | "PUT";
  url: string;
  body: unknown;
  headers?: Record<string, string>;
}

export type AttachmentSize = "small" | "upload" | "too-big";

export function attachmentSize(part: MailBytes): AttachmentSize {
  if (part.bytes.byteLength > GRAPH_FILE_LIMIT) return "too-big";
  if (part.bytes.byteLength < GRAPH_SMALL_LIMIT) return "small";
  return "upload";
}

export function graphDraftRequest(draft: MailDraft): GraphRequest {
  return {
    method: "POST",
    url: "https://graph.microsoft.com/v1.0/me/messages",
    body: {
      subject: draft.subject,
      body: { contentType: "HTML", content: draft.html },
      toRecipients: draft.to.map((address) => ({ emailAddress: { address } })),
      ccRecipients: draft.cc.map((address) => ({ emailAddress: { address } })),
    },
  };
}

export function graphSmallAttachmentRequest(messageId: string, part: MailBytes, inline: boolean): GraphRequest {
  const body: Record<string, unknown> = {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: part.name,
    contentType: part.contentType,
    contentBytes: encodeBase64(part.bytes),
    isInline: inline,
  };
  if (inline && part.contentId) body.contentId = part.contentId;
  return {
    method: "POST",
    url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments`,
    body,
  };
}

export function graphUploadSessionRequest(messageId: string, part: MailBytes, inline: boolean): GraphRequest {
  return {
    method: "POST",
    url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`,
    body: {
      AttachmentItem: {
        attachmentType: "file",
        name: part.name,
        size: part.bytes.byteLength,
        contentType: part.contentType,
        isInline: inline,
        contentId: inline ? part.contentId ?? null : null,
      },
    },
  };
}

export function graphUploadPut(uploadUrl: string, bytes: Uint8Array): GraphRequest {
  const last = Math.max(0, bytes.byteLength - 1);
  return {
    method: "PUT",
    url: uploadUrl,
    headers: {
      "content-type": "application/octet-stream",
      "content-range": `bytes 0-${last}/${bytes.byteLength}`,
    },
    body: bytes,
  };
}

export function outlookDraftsUrl(email: string): string {
  return /@(outlook|hotmail|live)\.[a-z.]+$/i.test(email)
    ? "https://outlook.live.com/mail/0/drafts"
    : "https://outlook.office.com/mail/drafts";
}

/** StrapSheets mobile Mail.Send body. saveToSentItems stays true. */
export function graphSendMailBody(message: {
  subject: string;
  text: string;
  to: string[];
  files: Array<{ name: string; contentType: string; contentBytes: string }>;
}): { message: unknown; saveToSentItems: true } {
  return {
    message: {
      subject: message.subject,
      body: { contentType: "Text", content: message.text },
      toRecipients: message.to.map((address) => ({ emailAddress: { address } })),
      attachments: message.files.map((file) => ({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: file.name,
        contentType: file.contentType,
        contentBytes: file.contentBytes,
      })),
    },
    saveToSentItems: true,
  };
}
