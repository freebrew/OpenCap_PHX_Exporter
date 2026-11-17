import { GMAIL_DRAFTS_URL, gmailDraftRequest, gmailWithinLimit, toMime, type MailDraft, type TransportResult } from "@phx/mail";

type FetchLike = (input: string, init: { method: string; headers: Record<string, string>; body?: string }) => Promise<{ ok: boolean; status: number; json(): Promise<unknown> }>;

export async function createGmailDraft(draft: MailDraft, token: string, deps: {
  fetchImpl?: FetchLike;
  open: (url: string) => Promise<void>;
  onOpening: () => void;
}): Promise<TransportResult> {
  const mime = toMime(draft);
  if (!gmailWithinLimit(mime)) return { status: "no-side-effect-failure", error: "The report is over Gmail’s 25 MiB encoded limit. Nothing was created." };
  const request = gmailDraftRequest(mime);
  const fetchImpl = deps.fetchImpl ?? (fetch as unknown as FetchLike);
  let response: { ok: boolean; status: number; json(): Promise<unknown> };
  try {
    response = await fetchImpl(request.url, {
      method: request.method,
      headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
      body: JSON.stringify(request.body),
    });
  } catch {
    return { status: "indeterminate", error: "Gmail did not confirm whether a draft was created. Check Drafts before trying again." };
  }
  if (!response.ok && response.status >= 500) return { status: "indeterminate", error: "Gmail did not confirm whether a draft was created. Check Drafts before trying again." };
  if (!response.ok) return { status: "no-side-effect-failure", error: `Gmail refused the draft (${response.status}). Nothing was created.` };
  const json = await response.json().catch(() => null) as { id?: string } | null;
  const id = json?.id ?? "";
  if (id === "") return { status: "indeterminate", error: "Gmail did not return a draft id. Check Drafts before trying again." };
  deps.onOpening();
  try {
    await deps.open(GMAIL_DRAFTS_URL);
  } catch (error) {
    return { status: "created-not-opened", id, location: GMAIL_DRAFTS_URL, error: error instanceof Error ? error.message : "Drafts did not open." };
  }
  return { status: "created-and-opened", id, location: GMAIL_DRAFTS_URL };
}
