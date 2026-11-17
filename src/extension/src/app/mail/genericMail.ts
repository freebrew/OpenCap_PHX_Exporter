import type { MailDraft } from "@phx/mail";

export function mailtoUrl(draft: MailDraft): string {
  const names = [...draft.attachments.map((part) => part.name), ...draft.missing.map((file) => file.name)];
  const lines = [draft.text, "", "Attach these files manually:", ...names];
  let body = lines.join("\n");
  if (body.length > 1500) body = `${body.slice(0, 1400)}\n…`;
  const params = new URLSearchParams();
  if (draft.cc.length > 0) params.set("cc", draft.cc.join(","));
  params.set("subject", draft.subject);
  params.set("body", body);
  return `mailto:${encodeURIComponent(draft.to.join(","))}?${params.toString()}`;
}

export async function openMailto(draft: MailDraft): Promise<void> {
  const names = [...draft.attachments.map((part) => part.name), ...draft.missing.map((file) => file.name)];
  try {
    await navigator.clipboard.writeText(names.join("\n"));
  } catch {
    // The names are also in the message body and the status line.
  }
  window.open(mailtoUrl(draft));
}
