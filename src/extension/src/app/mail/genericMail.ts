import type { MailDraft } from "@phx/mail";

export function mailtoUrl(draft: MailDraft): string {
  const names = [...draft.attachments.map((part) => part.name), ...draft.missing.map((file) => file.name)];
  const lines = [draft.text, "", "Attach these files manually:", ...names];
  let body = lines.join("\n");
  if (body.length > 1500) body = `${body.slice(0, 1400)}\n…`;
  const query: string[] = [];
  if (draft.cc.length > 0) query.push(`cc=${encodeURIComponent(draft.cc.join(","))}`);
  query.push(`subject=${encodeURIComponent(draft.subject)}`);
  query.push(`body=${encodeURIComponent(body)}`);
  return `mailto:${encodeURIComponent(draft.to.join(","))}?${query.join("&")}`;
}

export async function openMailto(draft: MailDraft): Promise<void> {
  const names = [...draft.attachments.map((part) => part.name), ...draft.missing.map((file) => file.name)];
  try {
    if (typeof ClipboardItem !== "undefined" && navigator.clipboard?.write) {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/html": new Blob([draft.html], { type: "text/html" }),
          "text/plain": new Blob([`${draft.text}\n\nAttachments:\n${names.join("\n")}`], { type: "text/plain" }),
        }),
      ]);
    } else {
      await navigator.clipboard.writeText(names.join("\n"));
    }
  } catch {
    // The names are also in the message body and the status line.
  }
  window.open(mailtoUrl(draft));
}
