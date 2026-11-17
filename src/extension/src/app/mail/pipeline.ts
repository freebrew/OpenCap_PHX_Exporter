import {
  GMAIL_ENCODED_LIMIT,
  GRAPH_FILE_LIMIT,
  MailValidationError,
  gmailWithinLimit,
  toMime,
  validateDraft,
  type MailDraft,
  type SizeLimits,
  type TransportResult,
} from "@phx/mail";
import type { MailClient } from "./clients";
import type { MailPhase } from "./phases";

const CLASSIC_LIMIT = 64 * 1024 * 1024;

export function limitsFor(client: MailClient): SizeLimits {
  switch (client) {
    case "classic":
      return { fileBytes: CLASSIC_LIMIT, totalBytes: CLASSIC_LIMIT };
    case "microsoft":
      return { fileBytes: GRAPH_FILE_LIMIT, totalBytes: GRAPH_FILE_LIMIT };
    case "gmail":
      return { fileBytes: GMAIL_ENCODED_LIMIT, totalBytes: GMAIL_ENCODED_LIMIT };
    case "generic":
    case "archive":
      return {};
    default: {
      const never: never = client;
      return never;
    }
  }
}

export interface ReportOutcome {
  status: TransportResult["status"] | "archive-only" | "manual-handoff";
  message: string;
  fileName?: string;
  archived: boolean;
  canRetryOpen: boolean;
}

export interface ReportTransport {
  run(draft: MailDraft, onPhase: (phase: MailPhase) => void): Promise<TransportResult>;
  retryOpen?(): Promise<TransportResult>;
}

function describe(client: MailClient, fileName: string, draft: MailDraft, result: TransportResult): string {
  const archive = `Archive: Midnight Reports/${fileName}.`;
  const names = [...draft.attachments.map((part) => part.name), ...draft.missing.map((file) => file.name)];
  switch (result.status) {
    case "created-and-opened":
      if (client === "generic") {
        return `The mail window has the recipients and subject. Attach these files yourself: ${names.join(", ") || "none"}. ${archive}`;
      }
      if (client === "classic") return `Classic Outlook is open with the report. Nothing was sent. ${archive}`;
      return `Draft “${draft.subject}” is in Drafts. Nothing was sent. ${archive}`;
    case "created-not-opened":
      return `The draft was created, but it did not open (${result.error}). Use Open Drafts. Emailing again would create a second draft. ${archive}`;
    case "indeterminate":
      return `${result.error} Check Drafts before trying again. ${archive}`;
    case "no-side-effect-failure":
      return `${result.error} ${archive}`;
    default: {
      const never: never = result;
      return never;
    }
  }
}

function outcomeOf(client: MailClient, fileName: string, draft: MailDraft, result: TransportResult): ReportOutcome {
  const manual = client === "generic" && result.status === "created-and-opened";
  return {
    status: manual ? "manual-handoff" : result.status,
    message: describe(client, fileName, draft, result),
    fileName,
    archived: true,
    canRetryOpen: result.status === "created-not-opened",
  };
}

export async function deliverReport(input: {
  draft: MailDraft;
  client: MailClient;
  archive: { write(eml: string): Promise<string> };
  transport: ReportTransport;
  onPhase?: (phase: MailPhase) => void;
  gmailEnabled?: boolean;
}): Promise<ReportOutcome> {
  const onPhase = input.onPhase ?? (() => undefined);
  onPhase("validating");
  let draft: MailDraft;
  try {
    draft = validateDraft(input.draft, limitsFor(input.client));
    if (input.client === "gmail") {
      if (input.gmailEnabled === false) {
        return { status: "no-side-effect-failure", archived: false, canRetryOpen: false, message: "Gmail drafts are off until Google has verified this app. Turn the switch on only for a listed test user." };
      }
      if (!gmailWithinLimit(toMime(draft))) {
        return { status: "no-side-effect-failure", archived: false, canRetryOpen: false, message: "The report is over Gmail’s 25 MiB encoded limit. Nothing was created." };
      }
    }
  } catch (error) {
    const message = error instanceof MailValidationError ? error.message : "The report could not be checked.";
    return { status: "no-side-effect-failure", archived: false, canRetryOpen: false, message };
  }

  onPhase("writing-archive");
  let fileName: string;
  try {
    fileName = await input.archive.write(toMime(draft));
  } catch {
    return {
      status: "no-side-effect-failure",
      archived: false,
      canRetryOpen: false,
      message: "Could not write Midnight Reports. Grant the job folder again, then retry Email report.",
    };
  }

  if (input.client === "archive") {
    onPhase("complete");
    return {
      status: "archive-only",
      archived: true,
      canRetryOpen: false,
      fileName,
      message: `Saved Midnight Reports/${fileName}. This did not open a mail compose.`,
    };
  }

  onPhase("creating");
  let result: TransportResult;
  try {
    result = await input.transport.run(draft, onPhase);
  } catch (error) {
    result = { status: "indeterminate", error: error instanceof Error ? error.message : "The mail client did not answer." };
  }
  const outcome = outcomeOf(input.client, fileName, draft, result);
  onPhase(outcome.canRetryOpen || result.status === "indeterminate" ? "partial" : result.status === "no-side-effect-failure" ? "failed" : "complete");
  return outcome;
}

let retry: (() => Promise<TransportResult>) | null = null;

export function rememberRetry(transport: ReportTransport, outcome: ReportOutcome): void {
  retry = outcome.canRetryOpen && transport.retryOpen ? transport.retryOpen.bind(transport) : null;
}

export async function retryOpen(onPhase: (phase: MailPhase) => void = () => undefined): Promise<ReportOutcome> {
  const open = retry;
  if (!open) {
    return {
      status: "no-side-effect-failure",
      archived: true,
      canRetryOpen: false,
      message: "Check Drafts before trying again. A new draft was not created.",
    };
  }
  onPhase("opening");
  const result = await open();
  const outcome: ReportOutcome = {
    status: result.status,
    archived: true,
    canRetryOpen: result.status === "created-not-opened",
    message: result.status === "created-and-opened"
      ? `Drafts is open. Nothing was sent.`
      : result.status === "created-not-opened"
        ? `Drafts still did not open (${result.error}). A new draft was not created.`
        : `${result.status === "indeterminate" || result.status === "no-side-effect-failure" ? result.error : "Drafts did not open."} A new draft was not created.`,
  };
  if (!outcome.canRetryOpen) retry = null;
  onPhase(outcome.status === "created-and-opened" ? "complete" : "partial");
  return outcome;
}
