/** Unpacked preview identity. The private key is not stored; this is the public SPKI only. */
export const PREVIEW_EXTENSION_KEY =
  "MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEArkfIt7phTUEvn0rAzmSFEot2ihb19FxYLpJN0Gwj6JxGNGE52otUD0uspNCBnxvrTtQ4eBTzx5R2CyGbJgC5sjunCcfF0zccjc4UWi62aUxscRWjKXd9pJDYiu8ov2sN2JkyaSejCNr+IdirSpiAWTjuCQG3deJqPCcm4artk6wa/+LqoMyFbBCvXOqBjU8ByG1pSLBZshw0a9L0b7OkAfUEXD2BLfP9+MacWd9qBNmnWZN7aQODdivyV434OxCz2j3aO5pU1sJR1t4iOSjWQvjEP0RM1GHRa66sG3CGLGWyyL1eCVBlfOVUC09wggCJCCgn7NykeWv7bdZQi0NTJQIDAQAB";

export const PREVIEW_EXTENSION_ID = "glkcmpmiomoppdkmdcjaepneimgbipdl";

export const MAIL_REDIRECT_PATH = "/mail";

export const NATIVE_HOST_NAME = "com.phx.outlook";

/** Exact vendor hosts the mailbox calls need. Webmail page detection is optional and separate. */
export const MAIL_HOST_PERMISSIONS = [
  "https://login.microsoftonline.com/*",
  "https://graph.microsoft.com/*",
  "https://outlook.office.com/*",
  "https://oauth2.googleapis.com/*",
  "https://gmail.googleapis.com/*",
] as const;

/** Webmail pages checked when the user turns on open-tab detection. */
export const WEBMAIL_TAB_ORIGINS = [
  "https://outlook.office.com/*",
  "https://outlook.live.com/*",
  "https://mail.google.com/*",
] as const;

/** Tab-detection hosts not already granted by MAIL_HOST_PERMISSIONS. */
export const OPTIONAL_WEBMAIL_ORIGINS = [
  "https://outlook.live.com/*",
  "https://mail.google.com/*",
] as const;

export function mailRedirectUri(extensionId = PREVIEW_EXTENSION_ID): string {
  return `https://${extensionId}.chromiumapp.org${MAIL_REDIRECT_PATH}`;
}

export function nativeHostOrigin(extensionId = PREVIEW_EXTENSION_ID): string {
  return `chrome-extension://${extensionId}/`;
}
