import { decodeBase64 } from "@phx/mail";
import { jobField, readJobFields, type SheetBook, type SnapshotBook } from "../../../../packages/workbook-commands/src/index";
import { nativeHostOrigin } from "../../extensionIdentity";
import type { FolderEntry } from "../pickFile";
import { writeArchive } from "./archive";
import { composeOnPort, connectClassic, probeClassic } from "./classicHost";
import { DEFAULT_MAIL_CLIENT } from "./clients";
import { openDesktopDraft, openDownloadedDraft } from "./desktopDraft";
import { looksLikeClassic } from "./detect";
import { createGmailDraft } from "./gmailDraft";
import { openMailto } from "./genericMail";
import { readNamedFiles } from "./jobFiles";
import { createMicrosoftDraft } from "./microsoftDraft";
import { getMailUi, phaseActive, setMailUi } from "./phases";
import { deliverReport, rememberRetry, retryOpen, type ReportTransport } from "./pipeline";
import { readClientConfig, readMailClient, usableToken } from "./platform";
import { attachmentNames, buildReportDraft } from "./reportDraft";

export interface DailySource {
  snapshot: SnapshotBook | null;
  book: SheetBook | null;
  pngBase64?: string;
  jobFolder: () => Promise<FolderEntry | null>;
  notify: (message: string) => void;
  openChooser: () => void;
}

let running = false;

async function openUrl(url: string): Promise<void> {
  await chrome.tabs.create({ url });
}

function transportFor(client: "outlook" | "classic" | "microsoft" | "gmail" | "generic"): ReportTransport {
  let draftsUrl = "";
  let draftDownloadId: number | null = null;
  switch (client) {
    case "outlook":
      return {
        async run(draft, onPhase) {
          // If the optional Classic Outlook native helper is already installed and Classic Outlook is active,
          // use it directly. Otherwise save an editable X-Unsent: 1 .eml draft and launch Outlook (Classic or New).
          try {
            const origin = nativeHostOrigin(chrome.runtime.id);
            const probe = await probeClassic(connectClassic(), origin, 2000);
            if (
              probe.ok === true &&
              looksLikeClassic({
                outlookRegistered: probe.outlookRegistered === true,
                mailtoCommand: probe.mailtoCommand ?? "",
                defaultClient: probe.defaultClient ?? "",
              })
            ) {
              const classicResult = await composeOnPort(connectClassic(), draft, origin, () => onPhase("opening"));
              if (classicResult.status !== "no-side-effect-failure") return classicResult;
            }
          } catch {
            // Helper not installed or unavailable: continue straight to the default Outlook .eml draft.
          }
          const opened = await openDesktopDraft(draft, { onOpening: () => onPhase("opening") });
          draftDownloadId = opened.downloadId;
          return opened.result;
        },
        async retryOpen() {
          if (draftDownloadId != null) return openDownloadedDraft(draftDownloadId);
          return { status: "created-not-opened", id: "outlook", location: "Outlook", error: "Open Outlook and look for the compose window. A new draft was not created." };
        },
      };
    case "classic":
      return {
        async run(draft, onPhase) {
          let port;
          try {
            port = connectClassic();
          } catch {
            return { status: "no-side-effect-failure", error: "The Classic Outlook helper is not installed. Choose “Outlook (Classic or New) — default”, or install the optional helper in Mail settings." };
          }
          return composeOnPort(port, draft, nativeHostOrigin(chrome.runtime.id), () => onPhase("opening"));
        },
        async retryOpen() {
          return { status: "created-not-opened", id: "classic", location: "Classic Outlook", error: "Open Outlook and look for the compose window. A new message was not created." };
        },
      };
    case "microsoft":
    case "gmail":
      return {
        async run(draft, onPhase) {
          onPhase("signing-in");
          const config = await readClientConfig();
          let token;
          try {
            token = await usableToken(client, config);
          } catch (error) {
            return { status: "no-side-effect-failure", error: error instanceof Error ? error.message : "Mailbox sign-in failed." };
          }
          if (token === "canceled") return { status: "no-side-effect-failure", error: "Sign-in was canceled." };
          onPhase("creating");
          const created = client === "microsoft"
            ? await createMicrosoftDraft(draft, token.accessToken, token.email, { open: openUrl, onOpening: () => onPhase("opening") })
            : await createGmailDraft(draft, token.accessToken, { open: openUrl, onOpening: () => onPhase("opening") });
          if ((created.status === "created-and-opened" || created.status === "created-not-opened") && created.location.startsWith("http")) {
            draftsUrl = created.location;
          }
          return created;
        },
        async retryOpen() {
          if (draftsUrl === "") return { status: "no-side-effect-failure", error: "There is no Drafts link to open. A new draft was not created." };
          await openUrl(draftsUrl);
          return { status: "created-and-opened", location: draftsUrl };
        },
      };
    case "generic":
      return {
        async run(draft, onPhase) {
          onPhase("opening");
          await openMailto(draft);
          return { status: "created-and-opened", location: "mailto" };
        },
      };
    default: {
      const never: never = client;
      return never;
    }
  }
}

export async function sendDailyReport(source: DailySource): Promise<void> {
  if (running || phaseActive(getMailUi().phase)) {
    source.notify("An email report is already running.");
    return;
  }
  const client = (await readMailClient()) ?? DEFAULT_MAIL_CLIENT;
  if (!source.snapshot) {
    const message = "The workbook is not open.";
    setMailUi({ phase: "failed", message, canRetryOpen: false });
    source.notify(message);
    return;
  }
  running = true;
  setMailUi({ phase: "preparing", message: "Preparing the report.", canRetryOpen: false });
  try {
    const fields = source.book ? readJobFields(source.book) : null;
    const root = await source.jobFolder();
    const files = await readNamedFiles(root, attachmentNames(source.snapshot));
    const draft = buildReportDraft({
      snapshot: source.snapshot,
      jobCode: fields ? jobField(fields, "Job Code", "Job ID") : "",
      jobName: fields ? jobField(fields, "Job Name") : "",
      png: source.pngBase64 ? decodeBase64(source.pngBase64) : null,
      files,
    });
    const config = await readClientConfig();
    const transport = client === "archive" ? { async run() { return { status: "no-side-effect-failure" as const, error: "" }; } } : transportFor(client);
    const outcome = await deliverReport({
      draft,
      client,
      gmailEnabled: config.gmailEnabled,
      archive: {
        async write(eml) {
          if (!root) throw new Error("folder");
          return writeArchive(root, eml);
        },
      },
      transport,
      onPhase: (phase) => setMailUi({ phase, message: getMailUi().message, canRetryOpen: false }),
    });
    rememberRetry(transport, outcome);
    if (outcome.canRetryOpen) source.openChooser();
    const missing = draft.missing.map((file) => file.name);
    const message = missing.length > 0 ? `Missing from the job folder: ${missing.join(", ")}. ${outcome.message}` : outcome.message;
    setMailUi({ phase: getMailUi().phase, message, canRetryOpen: outcome.canRetryOpen });
    source.notify(message);
  } catch (error) {
    const message = error instanceof Error ? error.message : "The email report failed.";
    setMailUi({ phase: "failed", message, canRetryOpen: false });
    source.notify(message);
  } finally {
    running = false;
  }
}

export async function retryDailyOpen(notify: (message: string) => void): Promise<void> {
  if (running) return;
  running = true;
  try {
    const outcome = await retryOpen((phase) => setMailUi({ phase, message: getMailUi().message, canRetryOpen: true }));
    setMailUi({ phase: getMailUi().phase, message: outcome.message, canRetryOpen: outcome.canRetryOpen });
    notify(outcome.message);
  } finally {
    running = false;
  }
}
