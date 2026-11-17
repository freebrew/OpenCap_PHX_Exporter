import { createDraft, splitAddresses, type MailDraft } from "@phx/mail";
import { escapeHtml } from "../../../../packages/reports/src/index";
import { readCell, type SnapshotBook } from "../../../../packages/workbook-commands/src/index";
import { rangeToHtml } from "./rangeHtml";

export interface RoleRow {
  role: string;
  address: string;
}

function cellText(snapshot: SnapshotBook, a1: string): string {
  const value = readCell(snapshot, "Data", a1);
  return value == null ? "" : String(value);
}

/** Module11 MailToFromSheet / MailCcFromSheet. The first TO or blank role is To; other non-TO rows are Cc. */
export function module11Recipients(rows: RoleRow[]): { to: string[]; cc: string[] } {
  let to = "";
  for (const row of rows) {
    const role = row.role.trim().toUpperCase();
    const address = row.address.trim();
    if (address === "") continue;
    if (role === "TO" || role === "") {
      to = address;
      break;
    }
  }
  if (to === "") to = rows[0]?.address.trim() ?? "";
  const cc: string[] = [];
  rows.forEach((row, index) => {
    const role = row.role.trim().toUpperCase();
    const address = row.address.trim();
    if (address === "") return;
    if (index === 0 && (role === "TO" || role === "")) return;
    if (role === "TO") return;
    cc.push(address);
  });
  return { to: splitAddresses(to), cc: cc.flatMap(splitAddresses) };
}

/** Module11 DailyMailSubject. A blank job code becomes Job. A blank job name omits its segment. */
export function dailyMailSubject(jobCode: string, jobName: string): string {
  const code = jobCode.trim() || "Job";
  const name = jobName.trim();
  return name === "" ? `Daily Drilling Summary - ${code}` : `Daily Drilling Summary - ${code} - ${name}`;
}

export function recipientRows(snapshot: SnapshotBook): RoleRow[] {
  const rows: RoleRow[] = [];
  for (let row = 29; row <= 33; row += 1) rows.push({ role: cellText(snapshot, `K${row}`), address: cellText(snapshot, `L${row}`) });
  return rows;
}

export function attachmentNames(snapshot: SnapshotBook): string[] {
  const names: string[] = [];
  for (let row = 29; row <= 33; row += 1) {
    const name = cellText(snapshot, `I${row}`).trim();
    if (name !== "") names.push(name);
  }
  return names;
}

export function buildReportDraft(input: {
  snapshot: SnapshotBook;
  jobCode: string;
  jobName: string;
  png: Uint8Array | null;
  files: Array<{ name: string; bytes: Uint8Array | null }>;
}): MailDraft {
  const recipients = module11Recipients(recipientRows(input.snapshot));
  const missing = input.files.filter((file) => file.bytes == null).map((file) => file.name);
  const note = missing.length > 0 ? `<p>Not in the job folder: ${missing.map((name) => escapeHtml(name)).join(", ")}</p>` : "";
  const subject = dailyMailSubject(input.jobCode, input.jobName);
  const html = `${rangeToHtml(input.snapshot, input.png != null)}${note}`;
  return createDraft({
    to: recipients.to,
    cc: recipients.cc,
    subject,
    html,
    text: [subject, missing.length > 0 ? `Not in the job folder: ${missing.join(", ")}` : ""].filter((line) => line !== "").join("\n"),
    inline: input.png ? [{ name: "daily_report_inline.png", contentType: "image/png", bytes: input.png, contentId: "daily_report.png" }] : [],
    attachments: [
      ...(input.png ? [{ name: "daily_report.png", contentType: "image/png", bytes: input.png }] : []),
      ...input.files.flatMap((file) => file.bytes ? [{ name: file.name, contentType: "application/octet-stream", bytes: file.bytes }] : []),
    ],
    missing: missing.map((name) => ({ name })),
  });
}
