export const MAIL_CLIENTS = ["outlook", "classic", "microsoft", "gmail", "generic", "archive"] as const;

export type MailClient = (typeof MAIL_CLIENTS)[number];

export const DEFAULT_MAIL_CLIENT: MailClient = "outlook";

export function parseMailClient(value: unknown): MailClient | null {
  return MAIL_CLIENTS.find((client) => client === value) ?? null;
}

export const MAIL_CLIENT_LABELS: Record<MailClient, string> = {
  outlook: "Outlook (Classic or New) — default",
  classic: "Classic Outlook (installed helper)",
  microsoft: "Outlook on the web (Microsoft sign-in)",
  gmail: "Gmail (Google sign-in)",
  generic: "Other / default mail app (mailto)",
  archive: "Save archive only (.eml)",
};
