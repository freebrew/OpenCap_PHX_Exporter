import { textBase64Url } from "./bytes";

/** Encoded-message ceiling used before users.drafts.create. Gmail's published limit is 25 MiB. */
export const GMAIL_ENCODED_LIMIT = 25 * 1024 * 1024;

export const GMAIL_DRAFTS_URL = "https://mail.google.com/mail/u/0/#drafts";

export function gmailRaw(mime: string): string {
  return textBase64Url(mime);
}

export function gmailEncodedSize(mime: string): number {
  return new TextEncoder().encode(gmailRaw(mime)).byteLength;
}

export function gmailWithinLimit(mime: string, limit = GMAIL_ENCODED_LIMIT): boolean {
  return gmailEncodedSize(mime) <= limit;
}

export function gmailDraftRequest(mime: string): { method: "POST"; url: string; body: { message: { raw: string } } } {
  return {
    method: "POST",
    url: "https://gmail.googleapis.com/gmail/v1/users/me/drafts",
    body: { message: { raw: gmailRaw(mime) } },
  };
}

export function gmailSendBody(mime: string): { raw: string } {
  return { raw: gmailRaw(mime) };
}
