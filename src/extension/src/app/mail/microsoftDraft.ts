import {
  attachmentSize,
  graphDraftRequest,
  graphSmallAttachmentRequest,
  graphUploadPut,
  graphUploadSessionRequest,
  outlookDraftsUrl,
  type MailDraft,
  type TransportResult,
} from "@phx/mail";

interface HttpResult {
  kind: "ok";
  json: unknown;
}

type FetchLike = (input: string, init: { method: string; headers: Record<string, string>; body?: string | Uint8Array }) => Promise<{ ok: boolean; status: number; json(): Promise<unknown>; text(): Promise<string> }>;

async function call(fetchImpl: FetchLike, token: string, request: { method: string; url: string; body: unknown; headers?: Record<string, string> }, raw?: Uint8Array): Promise<HttpResult | { kind: "http"; status: number; text: string } | { kind: "network" }> {
  try {
    const headers = { authorization: `Bearer ${token}`, ...(request.headers ?? {}) };
    const response = await fetchImpl(request.url, {
      method: request.method,
      headers: raw ? headers : { ...headers, "content-type": "application/json" },
      body: raw ?? JSON.stringify(request.body),
    });
    if (!response.ok) return { kind: "http", status: response.status, text: await response.text() };
    const json = response.status === 204 ? {} : await response.json().catch(() => ({}));
    return { kind: "ok", json };
  } catch {
    return { kind: "network" };
  }
}

function messageId(json: unknown): string {
  return json && typeof json === "object" && typeof (json as { id?: unknown }).id === "string" ? (json as { id: string }).id : "";
}

function uploadUrl(json: unknown): string {
  return json && typeof json === "object" && typeof (json as { uploadUrl?: unknown }).uploadUrl === "string" ? (json as { uploadUrl: string }).uploadUrl : "";
}

function allowedUpload(url: string): boolean {
  try {
    const host = new URL(url).host;
    return host === "outlook.office.com" || host === "graph.microsoft.com";
  } catch {
    return false;
  }
}

export async function createMicrosoftDraft(draft: MailDraft, token: string, email: string, deps: {
  fetchImpl?: FetchLike;
  open: (url: string) => Promise<void>;
  onOpening: () => void;
}): Promise<TransportResult> {
  const fetchImpl = deps.fetchImpl ?? (fetch as unknown as FetchLike);
  const location = outlookDraftsUrl(email);
  const parts = [...draft.inline.map((part) => ({ part, inline: true })), ...draft.attachments.map((part) => ({ part, inline: false }))];
  if (parts.some(({ part }) => attachmentSize(part) === "too-big")) {
    return { status: "no-side-effect-failure", error: "An attachment is over Graph’s 150 MiB limit. Nothing was created." };
  }
  const created = await call(fetchImpl, token, graphDraftRequest(draft));
  if (created.kind === "network" || (created.kind === "http" && created.status >= 500)) {
    return { status: "indeterminate", error: "Microsoft mail did not confirm whether a draft was created. Check Drafts before trying again." };
  }
  if (created.kind === "http") return { status: "no-side-effect-failure", error: `Microsoft mail refused the draft (${created.status}). Nothing was created.` };
  const id = messageId(created.json);
  if (id === "") return { status: "indeterminate", error: "Microsoft mail did not return a draft id. Check Drafts before trying again." };

  for (const { part, inline } of parts) {
    const size = attachmentSize(part);
    const posted = size === "small"
      ? await call(fetchImpl, token, graphSmallAttachmentRequest(id, part, inline))
      : await call(fetchImpl, token, graphUploadSessionRequest(id, part, inline));
    if (posted.kind !== "ok") {
      return { status: "created-not-opened", id, location, error: "A draft exists, but an attachment did not upload." };
    }
    if (size === "upload") {
      const target = uploadUrl(posted.json);
      if (!allowedUpload(target)) return { status: "created-not-opened", id, location, error: "A draft exists, but the upload URL was not on the expected host." };
      const put = graphUploadPut(target, part.bytes);
      const uploaded = await call(fetchImpl, token, { ...put, body: undefined }, part.bytes);
      if (uploaded.kind !== "ok") return { status: "created-not-opened", id, location, error: "A draft exists, but an attachment did not finish uploading." };
    }
  }

  deps.onOpening();
  try {
    await deps.open(location);
  } catch (error) {
    return { status: "created-not-opened", id, location, error: error instanceof Error ? error.message : "Drafts did not open." };
  }
  return { status: "created-and-opened", id, location };
}
