import { createDraft, type MailDraft } from "./draft";

export class MailValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "MailValidationError";
  }
}

const EMAIL = /^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;

export function stripHeaderControls(value: string): string {
  return value.replace(/[\u0000-\u001F\u007F]/g, "");
}

export function splitAddresses(value: string): string[] {
  return stripHeaderControls(value).split(/[;,]/).map((part) => part.trim()).filter((part) => part !== "");
}

export function sanitizeFileName(name: string): string {
  const leaf = stripHeaderControls(name).split(/[\\/]/).pop() ?? "";
  const clean = leaf.replace(/["]/g, "").replace(/[^A-Za-z0-9._ -]/g, "_").trim().slice(0, 80);
  if (clean === "" || clean === "." || clean === "..") throw new MailValidationError("An attachment name is not usable.");
  return clean;
}

export interface SizeLimits {
  fileBytes?: number;
  totalBytes?: number;
}

export function draftByteSize(draft: MailDraft): number {
  const files = [...draft.inline, ...draft.attachments].reduce((sum, part) => sum + part.bytes.byteLength, 0);
  return files + new TextEncoder().encode(draft.html).byteLength + new TextEncoder().encode(draft.text).byteLength;
}

function cleanAddresses(label: string, addresses: string[]): string[] {
  return addresses.map((address) => {
    const clean = stripHeaderControls(address).trim();
    if (!EMAIL.test(clean)) throw new MailValidationError(`${label} has an invalid address.`);
    return clean;
  });
}

export function validateDraft(draft: MailDraft, limits: SizeLimits = {}): MailDraft {
  const next = createDraft(draft);
  next.to = cleanAddresses("To", next.to);
  next.cc = cleanAddresses("Cc", next.cc);
  next.subject = stripHeaderControls(next.subject).trim();
  if (next.subject.length === 0) throw new MailValidationError("The subject is empty.");
  if (next.subject.length > 200) throw new MailValidationError("The subject is too long.");
  for (const part of [...next.inline, ...next.attachments]) {
    part.name = sanitizeFileName(part.name);
    if (part.contentId) part.contentId = stripHeaderControls(part.contentId);
    if (limits.fileBytes != null && part.bytes.byteLength > limits.fileBytes) {
      throw new MailValidationError(`${part.name} is over the file size limit.`);
    }
  }
  for (const file of next.missing) file.name = sanitizeFileName(file.name);
  if (limits.totalBytes != null && draftByteSize(next) > limits.totalBytes) {
    throw new MailValidationError("The message is over the size limit.");
  }
  return next;
}

export function outlookList(addresses: string[]): string {
  return addresses.join("; ");
}
