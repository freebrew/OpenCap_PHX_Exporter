# Browser Workbook email — 2026-11-04

Supersedes the email row in `docs/superpowers/specs/2026-11-03-browser-workbook-port-design.md`. The extension at `src/extension` is **PHX Slide Sheet** in the manifest and **Browser Workbook** here. Excel parity is `bas/cache_demo/Module11.bas` (`RTP_printW_Email`). StrapSheets mobile keeps its existing direct-send behavior and shares the message model.

## Goal

The Email report action stops downloading `morning-report.eml` for a mail client to open as a received message. It builds one `MailDraft` from the Data sheet and hands that draft to a mail client the user has confirmed. The Browser Workbook never sends mail.

## Accepted outcomes

| Client | Success | What the product does not promise |
|---|---|---|
| Classic Outlook | The native host opens an editable COM compose with `.Display`. | New Outlook is not COM-capable and is never treated as classic. |
| New Outlook, Outlook on the web, Gmail | The provider API creates the complete draft. The extension opens that account’s Drafts view and names the draft. | No stable exact-draft compose URL is part of the contract. Graph `webLink` and Gmail draft ids are not used as compose links. |
| Other / default webmail | `mailto:` prefills To, Cc, subject, and a plain-text body. The UI lists the attachment names and copies them. | Browsers cannot attach local files to an arbitrary webmail compose. This path is a manual-attachment handoff, not full automation. |
| Save archive only | `Midnight Reports/<yyyy-mm-dd> Daily Report.eml` is written in the job folder. | The archive is not opened and is not a successful compose. |

Every path writes that archive before any compose or draft call. An archive failure stops the pipeline and asks for the job folder grant again. Nothing is auto-sent.

## Unpacked extension identity

The preview manifest commits a public SPKI key so the unpacked extension id stays stable. The private key is not stored.

| | |
|---|---|
| Extension id | `glkcmpmiomoppdkmdcjaepneimgbipdl` |
| OAuth redirect | `https://glkcmpmiomoppdkmdcjaepneimgbipdl.chromiumapp.org/mail` |
| Native host name | `com.phx.outlook` |
| Allowed origin | `chrome-extension://glkcmpmiomoppdkmdcjaepneimgbipdl/` |

A Chrome Web Store listing can have a different id. That production id, when it exists, is recorded beside the preview id and registered on its own. Mailbox client ids are not the StrapSheets API login client.

## Provider limits

Microsoft Graph, delegated `Mail.ReadWrite` (drafts need this scope; it can also read mail):

- The extension does not request `offline_access`. The access token lives in `chrome.storage.session` and is requested again after restart or expiry.
- Create the draft first, then attach files. Under 3 MiB: `POST /me/messages/{id}/attachments`. From 3 MiB through 150 MiB: an upload session. Larger than 150 MiB fails before any request.
- Upload session URLs are on `https://outlook.office.com/`.
- Work and school Drafts: `https://outlook.office.com/mail/drafts`. Personal Drafts: `https://outlook.live.com/mail/0/drafts`, chosen from the id token email.
- If Drafts fails to open after a draft id exists, the result is “draft created, not opened.” Retry opens Drafts and does not create another draft.

Gmail, restricted scope `https://www.googleapis.com/auth/gmail.compose` (this scope can also send):

- `users.drafts.create` with one base64url MIME body. The encoded message is checked against a 25 MiB ceiling before the create call.
- The extension token is session-only. No refresh token is stored.
- Drafts view: `https://mail.google.com/mail/u/0/#drafts`.
- Production use stays behind an explicit local switch until Google verifies the scope. Test users can turn the switch on for unverified-app mode.

Classic Outlook native host:

- Chrome’s extension-to-host message ceiling is 64 MiB and the host-to-extension ceiling is 1 MiB. File bytes move in small chunks on a persistent port. Host replies stay small.
- Files are written under `%TEMP%\phx-mail\<run-id>\` with generated names, attached, then removed in `finally`.
- The host accepts only the origin written at install time.

Generic `mailto:` bodies are truncated to a length browsers accept. Attachment names are listed because they cannot be added by the page.

## Message

`Data!K29:L33` matches Module11. The first row whose role is `TO` or blank supplies To. Other non-`TO` addresses are Cc. Addresses may be separated by `;` or `,`.

Subject matches Module11: `Daily Drilling Summary - <Job Code> - <Job Name>`. A blank job code becomes `Job`. A blank job name omits its segment. The depth suffix is not used.

The HTML body is the displayed `Data!B2:F55` table, with the corridor picture as the last row (`cid:daily_report.png`). Cell text is escaped. Text color in the mail body is black. The picture is also a visible `daily_report.png` attachment with no Content-ID. Module21 `report.html` is not the mail body.

Attachments 1–5 are the file names on `Data!I29:I33`, read from the granted job folder. Missing names are shown before transport and listed in the message. They are not attached.

## Shared package

`src/packages/mail` (`@phx/mail`) owns `MailDraft`, validation, canonical MIME, Graph and Gmail request bodies, PKCE, transport results, and the StrapSheets send order.

Transport results are `created-and-opened`, `created-not-opened`, `no-side-effect-failure`, and `indeterminate`. A later transport may run only after a proven no-side-effect failure. A draft id, or an indeterminate result, ends the attempt.

The Browser Workbook uses the confirmed client only. It does not walk the StrapSheets send order. StrapSheets mobile still sends directly: account (`Mail.Send` or `gmail.send`), then the StrapSheets API, then the share sheet. `/v1/email` stays as it is. Mobile refresh tokens stay in Capacitor secure storage.

`src/packages/reports/src/eml.ts` stays a wrapper over the canonical serializer for the archive. Tests compare parsed recipients, subject, HTML, and image bytes. The canonical serializer has its own exact golden.

## Browser Workbook flow

1. Preparing the report from the current Data snapshot.
2. Validating addresses, file names, and sizes. Failure here has no mail side effect.
3. Writing the Midnight Reports archive.
4. Signing in, when the confirmed client is Microsoft or Gmail.
5. Creating the compose or the draft.
6. Opening the compose or the Drafts view.
7. Complete, partial success, or failed. The button is disabled while a run is active.

Cancel during sign-in returns to the mail-client choices and does not start a second provider. The panel shows each failure and the final outcome. Archive-only and generic webmail are not described as full automation.

Detection may preselect Classic Outlook (native probe and the mailto ProgId), an open Outlook or Gmail tab after the user grants the narrow optional host permission, or a connected session. It never writes the choice. The user confirms it. The confirmed choice is stored. Tokens are not.

## Native host install

The extension cannot start PowerShell. It downloads a zip. The user extracts it and runs `install-host.cmd`. The installer copies the host to `%LOCALAPPDATA%\PHX\OutlookHost\<version>\`, writes the absolute-path host manifest, and registers `com.phx.outlook` under HKCU for Chrome and Edge. The previous version folder is left in place so repair can roll back. Uninstall removes the registration and the host folder.

`probe` reports the host protocol version, whether Classic Outlook’s COM ProgId is registered, whether it is running, and the mailto command. Recheck distinguishes missing, incompatible, ready, and Outlook-unavailable.

## Out of scope

Browser Workbook auto-send. Attachment automation for webmail that has no provider API. New Outlook process automation and undocumented compose deep links. FieldCap write-back. Apps Script `MailApp`. The legacy `src/chrome-extension` exporter.
