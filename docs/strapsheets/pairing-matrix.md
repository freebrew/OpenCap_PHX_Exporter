# StrapSheets pairing and send test matrix

Fill in one row per run. A pairing row passes when the 6-digit code mail arrives, both the code
and the magic link pair the address, and a second code within the hour is refused once five have
been sent (HTTP 429). A send row passes when the PDF arrives and looks like the Sheets printout: the tool drawing, one card per entry in the lettered grid with the yellow length/diameter block, and the family title underneath. A PDF that only arrives does not pass.

Build under test: `src/apps/strapsheets-mobile` debug APK (`android/app/build/outputs/apk/debug/app-debug.apk`)
against the `strapsheets-api` Worker at the URL set in Settings → Server.

## Pairing code mail

| Mail client | Account type | Code arrives | Code pairs | Magic link pairs | 6th code refused | Date | Tester |
|---|---|---|---|---|---|---|---|
| Gmail web | Google | | | | | | |
| Gmail Android app | Google | | | | | | |
| Outlook web (outlook.com) | Microsoft personal | | | | | | |
| Outlook web (Microsoft 365) | Entra work account | | | | | | |
| New Outlook (Windows) | Entra work account | | | | | | |
| Outlook classic (desktop) | Entra work account | | | | | | |
| Outlook Android app | Microsoft | | | | | | |
| Samsung Mail | IMAP | | | | | | |

Signed-in email equal to the FieldCap crew email pairs with no code (`paired: true` straight away):

| Vendor | Pairs without code | Date | Tester |
|---|---|---|---|
| Google | | | |
| Microsoft | | | |

## Send paths

| Path | Setup | One session, one PDF | Two sessions, one PDF | Two sessions, two PDFs | In Sent folder | Reply-To | Date | Tester |
|---|---|---|---|---|---|---|---|---|
| From my account (Outlook, Graph Mail.Send) | Microsoft test account | | | | yes expected | n/a | | |
| From my account (Gmail, gmail.send) | Google test user (unverified-app mode) | | | | yes expected | n/a | | |
| StrapSheets mail (Cloudflare Email) | mail scope declined, recipient on FormOptions!D2:D50 | | | | n/a | user's email expected | | |
| StrapSheets mail, recipient not on list | | refused (403) expected | | | | | | |
| Phone mail app (share sheet) | airplane mode | | | | | | | |
| Fallback: account fails, service sends | revoke Graph consent first | | | | | | | |

## Standalone (no FieldCap, no account, no network)

| Check | Result | Date | Tester |
|---|---|---|---|
| Motor entry completes end to end in airplane mode | | | |
| Serial typed lower-case shows and saves as ALL CAPS | | | |
| Every measure opens the number pad | | | |
| Blank measure skips; Enter moves to the next step | | | |
| Step order equals `entry_schema.js` | | | |
| Queued entry syncs when the network returns | | | |
| With a snapshot: picked serial prefills, edit sticks | | | |
| Out-of-range measure warns and still saves | | | |

## Browser Workbook email

The Browser Workbook never sends. Classic Outlook opens an editable compose through the installed helper. New Outlook, Outlook on the web, and Gmail create a complete draft and open that account’s Drafts folder. Other webmail only prefills `mailto:` and lists the attachments to add by hand. `Midnight Reports/<date> Daily Report.eml` is the archive, and it is not opened.

Unpacked preview extension id: `glkcmpmiomoppdkmdcjaepneimgbipdl`. Redirect: `https://glkcmpmiomoppdkmdcjaepneimgbipdl.chromiumapp.org/mail`. Helper: `com.phx.outlook`.

Automated with this build: Module11 recipients and subject, `Data!B2:F55` table, both PNG parts, Gmail size preflight, Graph small and upload requests, no second draft after a partial open, session token kept out of `chrome.storage.local`, native-host framing, and install / upgrade / rollback / uninstall against a temporary HKCU key.

The mailbox rows below stay blank until a tester signs in with a configured client id. Do not mark them from the automated tests.

| Path | What to confirm | Result | Date | Tester |
|---|---|---|---|---|
| Classic Outlook | Editable compose, not sent. To/Cc, subject, B2:F55, inline picture, visible PNG, attachments 1–5. Archive written. | Host probe on this PC: Classic Outlook is registered, default mail client is Microsoft Outlook. A test compose titled “PHX preview mail check” opened with Display and was not sent. The full workbook report (B2:F55, both PNG parts, attachments 1–5, Midnight Reports archive) still needs a tester in the reloaded preview. | 2026-11-04 | build |
| New Outlook / Outlook web | Complete draft in Drafts. Drafts view opens. The status names the subject. Not sent. | | | |
| Gmail | Same, after the Gmail switch is on for a test user or a verified client. | | | |
| Other / default webmail | To/Cc/subject/body prefilled. Attachment names listed as manual. Not described as fully automatic. | | | |
| Helper install, upgrade, repair, uninstall | HKCU Chrome and Edge registration, previous version kept, uninstall removes it. | Protocol and installer tests passed on the build machine. Live compose not marked here. | 2026-11-04 | build |
