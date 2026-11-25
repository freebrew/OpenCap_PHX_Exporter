# Chrome Web Store Listing — OpenCap Data Exporter

> Last Updated: 2026-11-25
> Store item: `egafclljokidcoiipaiojlaamgkooffb` (update to an existing listing)
> Upload: `dist/OpenCap_DataExporter_v3.2.19.zip` (built by `pack-store.ps1`; this file is not in the zip)

## Store Listing

**Extension Name**
OpenCap Data Exporter

**Short Description** (matches manifest, 121 chars)
Export FieldCap job, crew, BHA, inventory, and daily slide/rotate metres and slide hours to CSV using your login.

**Detailed Description** (paste as plain text; the store strips markdown)

```
Turns the FieldCap job you are already logged into into Excel-ready CSV files, one click per job.

What it exports
Job details and well information
Crew and personnel
BHA equipment, serials and hours
Equipment inventory for every job tool, with hours
Slide and rotate metres and slide/rotate hours by day, with a row for every day a BHA was in the hole
Ticket costs by day
An activity-log detail file that shows every activity block and how its hours were counted, so a day that looks wrong can be traced to the entry behind it

How to use it
1. Log into FieldCap in the same browser.
2. Open the extension. The Job ID fills in from the open FieldCap tab, or type it.
3. Pick the folder your Slide Sheet workbook reads from, and untick any export you do not want.
4. Click Fetch & Save CSVs. Each file is written to that folder (or to Downloads if the folder write is blocked).

Other features
Hides FieldCap's Bottom Line Verification reminder while you browse (optional).
Figures round the same way Excel does, so the CSV matches the workbook to the centimetre.
Shows when a store update is waiting and installs it from the popup.

Privacy
The extension uses the FieldCap session you already have. It does not ask for or store your password. Exported data goes only to files on your computer and the extension's local storage. Nothing is sent to the publisher, to analytics, or to any third party. Full policy: https://freebrew.github.io/OpenCap_PHX_Exporter/privacy.html

Independent tool
OpenCap Data Exporter is an unofficial tool. It is not affiliated with PHX Technology, FieldCap, or Pason Systems.

Support: https://github.com/freebrew/OpenCap_PHX_Exporter/issues
```

**Category**
Productivity → Workflow & Planning

**Single Purpose**
Exports the FieldCap job a logged-in user is viewing to CSV files on their computer.

**Primary Language**
English

## Graphics & Assets

| Asset | Dimensions | Status | Filename |
|-------|-----------|--------|----------|
| Store Icon | 128×128 PNG | ✅ Ready | `icons/icon128.png` |
| Screenshot 1 | 1280×800 | 🟡 Needs update | `remotion/out/screenshot-1280-export.jpg` |
| Screenshot 2 | 1280×800 | 🟡 Needs update | `remotion/out/screenshot-1280-operations.jpg` |
| Screenshot 3 | 1280×800 | 🟡 Needs update | `remotion/out/screenshot-1280-diagnostics.jpg` |
| Small Promo Tile | 440×280 | 🟡 Needs update | `remotion/out/small-promo-export.jpg` |
| Marquee Promo Tile | 1400×560 | 🟡 Needs update | `remotion/out/marquee-export.jpg` |

### Screenshot Notes
All `remotion/out/` renders are from 2026-05-28 and show the v3.1.4 dark popup and "5 structured CSV outputs". The popup is now the light Slide Sheet theme (v3.2.15) with seven outputs (ticket costs and activity-log detail added). Re-render before or with this submission; outdated screenshots can trigger a rejection.

## Permissions Justification

| Permission | Type | Justification |
|------------|------|---------------|
| `storage` | permissions | Keeps the last fetched CSV text, the chosen save folder, the export checkboxes and the Bottom Line Verification setting on this device so the popup shows the last result and remembers choices between sessions. |
| `downloads` | permissions | Saves a CSV to the Downloads folder when writing to the user's chosen folder is blocked, so an export is never lost. |
| `activeTab` | permissions | Reads the Job ID from the FieldCap page the user has open when they click the toolbar icon, so they do not have to type it. |
| `scripting` | permissions | Re-attaches the page reader to a FieldCap tab that was already open before the extension was installed or updated, so the ticket and activity grids on screen can be read without reloading the page. On Microsoft Edge, also runs the FieldCap data request inside the open FieldCap tab, where the login cookies only travel with requests made from that tab. |
| `webRequest` | permissions | When the user turns on the optional "sniffing" diagnostic, records the addresses of FieldCap requests on `*.phxtech.com` so a new FieldCap data endpoint can be identified. Observes only; never blocks or changes requests, and is not used on any other site. |
| `alarms` | permissions | While a store update is waiting, restarts the one-minute toolbar "UP" badge pulse after the background worker sleeps. Cleared once the update is installed. It does not fetch data. |
| `https://fieldcap-cdn.phxtech.com/*` | host_permissions | The FieldCap web app and its data service. The exporter reads the job the user asked for with the session they are already logged into. |
| `*://*.phxtech.com/*` | host_permissions | Other FieldCap hosts under the same company domain that serve job, crew, BHA, inventory and ticket data, and where the content script reads the Job ID and hides the reminder dialog. No access to any other domain. |

**Remote code:** No. All JavaScript ships in the package; nothing is loaded from a server at runtime.

## Privacy & Data Use

### Data Collection

**Does the extension collect user data?** Yes (handled locally; nothing leaves the device except requests to FieldCap the user is already logged into).

| Data Type | Collected? | Transmitted Off-Device? | Purpose | Shared with Third Parties? |
|-----------|-----------|------------------------|---------|---------------------------|
| Personally identifiable info | Yes — crew names, roles, phone numbers, emails from FieldCap | No | Written to the crew CSV the user requested | No |
| Health info | No | | | |
| Financial info | Yes — daily ticket cost totals | No | Written to the ticket-cost CSV | No |
| Authentication info | No — uses the existing browser session; no password read or stored | | | |
| Personal communications | No | | | |
| Location | No (well location fields are job data, not the user's location) | | | |
| Web history | Only if the user turns on sniffing: FieldCap request URLs on `*.phxtech.com` | No | Identify FieldCap endpoints | No |
| User activity | No | | | |
| Website content | Yes — job, BHA, inventory and activity data shown in FieldCap | No | Build the CSV files | No |

### Data Use Certification
- [x] Data is NOT sold to third parties
- [x] Data is NOT used for purposes unrelated to the extension's core functionality
- [x] Data is NOT used for creditworthiness or lending purposes

## Privacy Policy

**Privacy Policy URL**
https://freebrew.github.io/OpenCap_PHX_Exporter/privacy.html
(Live, HTTP 200, H1 "Privacy Policy", effective October 25, 2026. Do not use the homepage or `privacy-disclaimer.html`.)

## Distribution

**Visibility**: unchanged from the current listing
**Regions**: All regions (free; no payments, so the regulated-country rule does not apply)

## Developer Info

**Publisher Name**: Bruno Brottes
**Contact Email**: the address already verified on the developer account
**Support URL**: https://github.com/freebrew/OpenCap_PHX_Exporter/issues
**Homepage URL**: https://freebrew.github.io/OpenCap_PHX_Exporter/

## Version History

| Version | Date | Changes | Status |
|---------|------|---------|--------|
| 3.2.19 | 2026-11-25 | Excel-style rounding; zero rows for in-hole days in the by-day file. Includes 3.2.16–3.2.18: slide hours capped at bit hours, activity-log detail CSV, by-day merge keeps earlier days, `alarms` permission. | Draft |
| 3.2.15 | 2026-10-25 | Spreadsheet popup theme | Last packed |
| 3.2.14 | 2026-10-25 | Standalone privacy policy URL | Resubmitted after rejection |
| 3.2.9 | 2026-10-06 | Update from the popup | Rejected (Purple Nickel) |

## Review Notes

### Known Issues / Limitations
- Works only for users with a FieldCap login; reviewers cannot exercise the fetch without one. Say so in the "notes to reviewer" field if asked for test credentials.
- The name and description mention FieldCap, a PHX product. The listing and policy state the tool is unofficial and unaffiliated.

### Rejection History
| Date | Reason | Fix Applied | Resubmitted |
|------|--------|-------------|-------------|
| 2026-10-07 | Purple Nickel — privacy policy link was not a valid privacy policy | Published `privacy.html` as a standalone policy and pointed the listing at it (v3.2.14) | Yes |
