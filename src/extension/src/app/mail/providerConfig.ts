import { MAIL_REDIRECT_PATH, mailRedirectUri, NATIVE_HOST_NAME, PREVIEW_EXTENSION_ID } from "../../extensionIdentity";

/**
 * Mailbox OAuth clients are not the StrapSheets API login. Paste each client id into Mail
 * settings on this browser. Nothing here is a secret: a public client uses PKCE and no secret.
 */
export const MAIL_STORAGE = {
  client: "phx-mail-client",
  gmailEnabled: "phx-mail-gmail-enabled",
  microsoftClientId: "phx-mail-microsoft-client-id",
  googleClientId: "phx-mail-google-client-id",
  tenant: "phx-mail-tenant",
  sessionToken: "phx-mail-token",
  pendingAuth: "phx-mail-pending",
  lastDraft: "phx-mail-last-draft",
} as const;

export const MICROSOFT_DRAFT_SCOPE = "openid email https://graph.microsoft.com/Mail.ReadWrite";
export const GOOGLE_DRAFT_SCOPE = "openid email https://www.googleapis.com/auth/gmail.compose";

export const PROVIDER_SETUP = [
  `Chrome Web Store id, if it differs from ${PREVIEW_EXTENSION_ID}, is registered on its own.`,
  `Redirect for this unpacked preview: ${mailRedirectUri()}.`,
  `Entra app: public client, redirect ${mailRedirectUri()}, delegated Mail.ReadWrite, no offline_access. Tenant defaults to common.`,
  "Google web client: add the same redirect. gmail.compose stays off until you enable it here. Unverified testing needs the Google account listed as a test user.",
  `Classic Outlook helper name: ${NATIVE_HOST_NAME}. Allowed origin: chrome-extension://${PREVIEW_EXTENSION_ID}/.`,
  `The redirect path is ${MAIL_REDIRECT_PATH}.`,
] as const;
