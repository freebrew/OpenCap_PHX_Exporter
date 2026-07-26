# PHX OpenCap — Data Export & Field Reporting Ecosystem

**Chrome Extension + Excel VBA Dashboard**
*Bridging wellsite data capture, office reporting, and equipment lifecycle tracking.*

**Current Chrome extension revision:** `v3.1.5`

---

## Overview

PHX OpenCap is an independent toolset that extracts operational drilling data from the [FieldCap](https://fieldcap.com/) web application and delivers it into structured formats for downstream use in Excel-based field reports, custom dashboards, and (planned) web-based office/field ecosystems.

The system consists of two main components:

| Component | Location | Purpose |
|-----------|----------|---------|
| Chrome Extension | `src/chrome-extension/` | Extracts job, crew, BHA equipment, inventory, and daily slide/rotate metre data from FieldCap via OData API |
| Excel VBA Module | `src/excel/MDL_DDTools.bas` | Imports exported CSVs and renders an interactive DD Tools dashboard with BHA selectors, hour/meter tracking, and fatigue warnings |

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    FieldCap Web Application                       │
│              (PHX Technology — phxtech.com)                       │
└────────────────────────┬────────────────────────────────────────┘
                         │  OData v4 API + XHR/Fetch Interception
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│              Chrome Extension (Manifest V3)                       │
│  ┌──────────┐  ┌──────────────┐  ┌───────────────────────────┐  │
│  │content.js│  │injected-spy.js│  │    background.js          │  │
│  │DOM scrape│  │XHR/Fetch hook │  │  OData fetch + normalize  │  │
│  └────┬─────┘  └──────┬───────┘  │  CSV generation           │  │
│       │               │          │  ActivityLog metre parsing │  │
│       └───────────────┴──────────┤                            │  │
│                                  └────────────┬───────────────┘  │
│                                               │                  │
│  ┌──────────────────────────────────────────┐ │                  │
│  │           popup.js + popup.html          │ │                  │
│  │  • Job ID input                          │◄┘                  │
│  │  • 5-checkbox export selector            │                    │
│  │  • Bottom Line Verification suppressor   │                    │
│  │  • Fetch & Build CSVs                    │                    │
│  │  • File System Access API → OpenCap/     │                    │
│  │  • Per-export status cards               │                    │
│  └──────────────────────┬───────────────────┘                    │
└─────────────────────────┼────────────────────────────────────────┘
                          │  5 CSV files under OpenCap/
                          ▼
┌─────────────────────────────────────────────────────────────────┐
│              Local File System (workbook folder)                  │
│  OpenCap/                                                         │
│    fieldcap-job-{id}-job-details.csv                             │
│    fieldcap-job-{id}-crew.csv                                    │
│    fieldcap-job-{id}-bha-equipment.csv                           │
│    fieldcap-job-{id}-inventory.csv                               │
│    fieldcap-job-{id}-slide-rotate-metres-by-day.csv              │
└─────────────────────────┬────────────────────────────────────────┘
                          │  VBA Refresh import
                          ▼
┌─────────────────────────────────────────────────────────────────┐
│              Excel DD Tools Dashboard                             │
│  • Interactive BHA selector buttons                              │
│  • Per-BHA stats: meters, hours (slide/rot/circ), fatigue        │
│  • Cumulative serial-hour tracking with 300hr limit warnings     │
│  • Raw BHA Summary table (for VLOOKUP formulas)                  │
│  • Slide Meters / Rotate Meters breakdown                        │
└─────────────────────────────────────────────────────────────────┘
```

---

## Chrome Extension

### What It Does

1. **Fetches** job details, crew schedules, BHA equipment data, full job inventory, and daily slide/rotate metre totals directly from FieldCap's OData API using your active browser session — no separate login needed.
2. **Intercepts** FieldCap's own internal API calls (XHR/Fetch) to capture hour statistics and metric fields not exposed on the public OData endpoints.
3. **Scrapes** the visible DOM (BHA grid) to capture real-time component values.
4. **Suppresses** the FieldCap "Bottom Line Verification" popup when enabled, auto-dismissing it with the same reminder action available in the site UI.
5. **Generates** 5 clean, properly-typed CSV files ready for Excel or any downstream system.

### CSV Outputs

| File | Contents |
|------|----------|
| `fieldcap-job-{id}-job-details.csv` | Core job metadata + all custom field key-value pairs |
| `fieldcap-job-{id}-crew.csv` | One row per crew member with role, contact, dates |
| `fieldcap-job-{id}-bha-equipment.csv` | One row per BHA component with serial, hours, meters, fatigue data |
| `fieldcap-job-{id}-inventory.csv` | One row per job tool / inventory item with serial, item, status, dispatch/return, and hour fields where available |
| `fieldcap-job-{id}-slide-rotate-metres-by-day.csv` | Daily slide and rotate metres per BHA, sourced directly from ActivityLogs |

### FieldCap UI Helper

Version `3.1.5` includes a popup toggle under **FieldCap UI**:

- **Hide Bottom Line Verification popup** — enabled by default.
- Detects only dialogs containing the "Bottom Line Verification" and bottom-hole measurement text.
- Clicks the site's own **Remind me in 10 minutes** control when present, then hides the dialog/backdrop as a fallback.
- The setting is stored locally in `chrome.storage.local` and can be turned off from the extension popup.

### Slide / Rotate Metres by Day

The `slide-rotate-metres-by-day` CSV is built directly from the FieldCap `ActivityLogs` OData endpoint — the same source that powers the Slide Sheet tab inside FieldCap. Each row represents one calendar day of drilling for a specific BHA.

| Column | Description |
|--------|-------------|
| `Job ID` | FieldCap job number |
| `Date` | Calendar date (`YYYY-MM-DD`) from `StartDateTime` |
| `BHA #` | BHA assembly number |
| `Slide Metres` | Sum of `End MD − Start MD` where `ActivityType = 'Sliding'` |
| `Rotate Metres` | Sum of `End MD − Start MD` where `ActivityType = 'Drilling'` |
| `Survey Course Sum` | Total of slide + rotate for the day |

This replaces the previous `SurveySheetEntries`-based calculation which was producing incorrect (often doubled) totals. The ActivityLogs source exactly matches what FieldCap displays on its Slide Sheet tab.

### BHA Equipment Columns

```
Job ID, BHA #, Section, Status, Motor, Guidance,
Metres Drilled, BHA Mtrs Slid, BHA Mtrs Rot,
BHA Total Hrs, BHA Hrs Slid, BHA Hrs Rot, BHA Hrs Circ, BHA Below Rot,
Activated On, Completed On,
Serial #, Item Code, Description, Sub Description,
Length, Accum Length, Top, Bottom, Max OD, Min ID,
Job Hours, HSLS, Strapped, Shipping Status, Dispatched On, Returned On
```

### Installation

1. Open `chrome://extensions` (or `edge://extensions`).
2. Enable **Developer mode**.
3. Click **Load unpacked** → select `src/chrome-extension/`.
4. Navigate to FieldCap, ensure you're logged in.
5. Click the extension icon, enter Job ID, and click **Fetch & Build CSVs**.

### Usage Tips

- All 5 export types are selected by default. Uncheck any you don't need before fetching.
- Choose your Excel workbook folder once — CSVs are written to an `OpenCap` subfolder so the root stays clean.
- The popup shows independent status cards — each updates as its data finishes fetching.
- Slide/Rotate Metres by Day requires no special tab to be open; it is fetched entirely via OData.
- `live bhaRows` and `live activityRows` counts in the status bar reflect DOM-captured data for the BHA equipment CSV.
- The Bottom Line Verification suppressor is enabled by default and can be disabled under **FieldCap UI**.

---

## Excel DD Tools Dashboard

### What It Does

Imports the exported CSVs and renders a fully interactive directional drilling dashboard:

- **BHA selector buttons** — embedded Form Control buttons in the header, one per BHA
- **Per-BHA statistics** — total meters, slide/rotate meters, hours breakdown, below-rotate hours
- **Component detail table** — serial numbers, descriptions, individual hours per tool
- **Cumulative serial-hour table** — aggregates hours by serial number across all BHAs
- **300-hour fatigue warning** — highlights components exceeding the hour limit with serial identification
- **Raw BHA Summary table** — positioned for easy VLOOKUP formulas from other sheets

### Installation

1. Open your Excel workbook.
2. `Alt+F11` → File → Import File → select `src/excel/MDL_DDTools.bas`.
3. Run macro: `RebuildDashboard` (creates the DD Tools sheet structure).
4. Click **Refresh** button to import CSVs from the workbook's `OpenCap` folder.

### How Refresh Works

- Scans `<workbook>\OpenCap\` first (then workbook root) for CSVs matching `bha-equipment`, `crew`, `job-details`.
- Always picks the **newest file** by timestamp when multiple versions exist.
- Shows a confirmation dialog with filenames and timestamps before importing.
- Imports to hidden data sheets (`_FC_BHA`, `_FC_CREW`, `_FC_JOB`).
- Rebuilds the dashboard without changing your active sheet/cell focus.

See [`src/excel/DD_TOOLS_SETUP.md`](src/excel/DD_TOOLS_SETUP.md) for detailed setup instructions.

---

## Project Structure

```
PHX_FieldCap/
├── README.md
├── .gitignore
├── fieldcap-job-20786-*.csv          (sample exports — not committed)
│
└── src/
    ├── chrome-extension/
    │   ├── manifest.json             (Manifest V3, v3.1.5)
    │   ├── background.js            (OData fetch, CSV generation, ActivityLog parsing)
    │   ├── content.js               (DOM scraping, table detection, auto-scrape, popup suppression)
    │   ├── injected-spy.js          (page-context XHR/Fetch interception)
    │   ├── popup.html               (extension UI — PHX dark/teal theme, 5-export layout + UI helper)
    │   ├── popup.js                 (popup logic, File System Access API downloads)
    │   └── icons/
    │       ├── icon16.png
    │       ├── icon48.png
    │       ├── icon128.png
    │       └── make_icons.ps1
    │
    ├── excel/
    │   ├── MDL_DDTools.bas           (DD Tools dashboard VBA module)
    │   └── DD_TOOLS_SETUP.md        (setup instructions)
    │
    ├── js/                           (legacy/diagnostic scripts)
    │   ├── fieldcap-bha-equipment-exporter.js
    │   ├── fieldcap-console-exporter.js
    │   ├── fieldcap-hour-field-diagnostic.js
    │   └── fieldcap-visible-table-exporter.js
    │
    └── vba/
        └── FieldCapInventoryMacro.bas (standalone inventory macro)
```

---

## Changelog

### v3.1.5 — OpenCap CSV Subfolder
- **CSV output folder** — choosing the Excel workbook root now creates/uses an `OpenCap` subfolder; all scraped FieldCap CSVs write there instead of cluttering the root.
- **Excel Refresh aligned** — Setup and DD Tools scan `<workbook>\OpenCap\` first, with a fallback to the workbook root for older exports.
- **UI clarity** — folder picker shows `Root / OpenCap` and success messages report the full save path.

### v3.1.4 — Compact Popup Layout
- **Compact popup UI** — widened the extension popup and reorganized controls/status cards into grids to remove the vertical scrollbar in normal use.
- **Tighter spacing** — reduced header, folder, checkbox, button, card, and footer padding while keeping the same controls visible.

### v3.1.3 — Popup Suppression Safety Fix
- **Fixed**: Bottom Line Verification suppression could hide a large FieldCap page container, causing a blank screen when the extension was enabled.
- **Tightened dialog targeting** — suppression now chooses the smallest safe dialog-like element and rejects page-sized containers.
- **Version bump policy captured** — Chrome extension runtime fixes must update `manifest.json`, README revision text, and the upload ZIP version.

### v3.1.2 — FieldCap UI Helper + Store Upload Package
- **New toggle: Hide Bottom Line Verification popup** — auto-dismisses the FieldCap bottom-hole measurement reminder while browsing, with a user-controlled popup setting.
- **Chrome Web Store-ready package** — extension manifest revision bumped to `3.1.2` for update submission.
- **Docs refreshed** — project page now reflects the 5-export layout, inventory CSV, and local-only request metadata behavior.

### v3.1.1 — Inventory Export + Diagnostics
- **New export: Equipment Inventory** — 5th CSV output for all job tools / inventory records with serial, item, status, dispatch/return, and hour fields where available.
- **Added schema, inventory, sniff, and debug probes** to help discover tenant-specific FieldCap OData fields.

### v3.0.0 — Major Release
- **New export: Slide / Rotate Metres by Day** — 4th CSV output sourced directly from the FieldCap `ActivityLogs` OData endpoint, exactly matching the FieldCap Slide Sheet tab.
  - `ActivityType = 'Sliding'` → slide metres (`End MD − Start MD`)
  - `ActivityType = 'Drilling'` → rotate metres (`End MD − Start MD`)
  - Grouped by calendar date and BHA number
- **UI overhaul** — popup now has 4 independent checkboxes and 4 status cards (one per export type)
- **Fixed**: slide/rotate totals were previously calculated from `SurveySheetEntries` which produced incorrect (often doubled) values; now uses `ActivityLogs` as the authoritative source
- **Fixed**: date parsing of FieldCap's compact `YYYYMMDDHHMI` integer format (eliminated phantom 1976 dates)

### v2.5.x
- OData-driven export of job details, crew, and BHA equipment
- File System Access API for direct-to-folder downloads
- XHR/Fetch interception for hour statistics
- DOM scraping fallback for activity metre data

---

## Roadmap

### Planned: Web Spreadsheet & Office Integration
- **Web-based spreadsheet interface** — browser-native viewer/editor for field reports without Excel dependency
- **Real-time sync** — push FieldCap exports directly to a shared web workspace
- **Template engine** — populate custom report templates (daily drilling reports, BHA run reports, tool failure reports) from the same CSV data
- **Office 365 / Google Sheets bridge** — auto-import into cloud spreadsheet platforms

### Planned: Seamless Client Workflow
- **Field → Office pipeline** — wellsite crew exports data; office receives formatted reports automatically
- **Equipment lifecycle tracking** — serial-hour fatigue integrated with dispatch/return workflows
- **Service company integration** — standardized data handoff for MWD/LWD tool rental tracking

---

## Data Privacy

- The extension operates using your existing FieldCap browser session — no credentials are stored or transmitted externally.
- All data remains local (browser storage + local CSV files).
- The `webRequest` permission is used only on `phxtech.com` to observe FieldCap OData/API request metadata for the optional probe/sniff diagnostics: URL, method, timestamp, and tab ID.
- No request bodies are read, blocked, redirected, or modified.
- No telemetry, analytics, or external API calls beyond FieldCap's own OData/API endpoints.

---

## Disclaimer

This project (OpenCap) is an independent, unofficial tool. It is **not affiliated with, endorsed by, or supported by [PHX Technology](https://www.phxtech.com/), [FieldCap](https://fieldcap.com/), or any other third party.** All trademarks and product names belong to their respective owners.

The name [FieldCap](https://fieldcap.com/) is used solely to describe the third-party web application this tool interacts with. Use of that name does not imply any association or endorsement.

---

## License

MIT — free to use, modify, and distribute.
