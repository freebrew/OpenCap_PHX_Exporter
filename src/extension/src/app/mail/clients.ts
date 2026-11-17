export const MAIL_CLIENTS = ["classic", "microsoft", "gmail", "generic", "archive"] as const;

export type MailClient = (typeof MAIL_CLIENTS)[number];

export function parseMailClient(value: unknown): MailClient | null {
  return MAIL_CLIENTS.find((client) => client === value) ?? null;
}

export const MAIL_CLIENT_LABELS: Record<MailClient, string> = {
  classic: "Classic Outlook",
  microsoft: "New Outlook / Outlook web",
  gmail: "Gmail",
  generic: "Other / default webmail",
  archive: "Save archive only",
};
