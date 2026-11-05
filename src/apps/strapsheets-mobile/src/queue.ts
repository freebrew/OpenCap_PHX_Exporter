// Every entry the phone has taken. It stays on the phone until the user emails the PDF.

import type { EntryPayload } from "../../../packages/strapsheets-core/src/index";

export interface StrapSession {
  id: string;
  jobNumber: string;
  entry: EntryPayload;
  createdAt: string;
}

export function newSession(jobNumber: string, entry: EntryPayload, now = new Date()): StrapSession {
  return { id: crypto.randomUUID(), jobNumber, entry, createdAt: now.toISOString() };
}
