# PHX FieldCap — Data Export & Field Reporting Ecosystem

**Chrome Extension + Excel VBA Dashboard**
*Bridging wellsite data capture, office reporting, and equipment lifecycle tracking.*

---

## Overview

PHX FieldCap is an integrated toolset that extracts operational drilling data from the [FieldCap](https://fieldcap-cdn.phxtech.com) web application and delivers it into structured formats for downstream use in Excel-based field reports, custom dashboards, and (planned) web-based office/field ecosystems.

The system consists of two main components:

| Component | Location | Purpose |
|-----------|----------|---------|
| Chrome Extension | `src/chrome-extension/` | Extracts job, crew, and BHA equipment data from FieldCap via OData API + DOM scraping |
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
│       │               │          │  Activity meter parsing    │  │
│       └───────────────┴──────────┤                            │  │
│                                  └────────────┬───────────────┘  │
│                                               │                  │
│  ┌──────────────────────────────────────────┐ │                  │
│  │           popup.js + popup.html          │ │                  │
│  │  • Job ID input                          │◄┘                  │
│  │  • Fetch & Build CSVs                    │                    │
│  │  • File System Access API download       │                    │
│  │  • Debug state / live metrics            │                    │
│  └──────────────────────┬───────────────────┘                    │
└─────────────────────────┼────────────────────────────────────────┘
                          │  3 CSV files
                          ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Local File System                              │
│  fieldcap-job-{id}-job-details.csv                               │
│  fieldcap-job-{id}-crew.csv                                      │
│  fieldcap-job-{id}-bha-equipment.csv                             │
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

1. **Fetches** job details, crew schedules, and BHA equipment data directly from FieldCap's OData API using your active browser session (no separate login needed).
2. **Intercepts** FieldCap's own internal API calls (XHR/Fetch) to capture hour statistics and metric fields that are not exposed on the public OData endpoints.
3. **Scrapes** the visible DOM (BHA grid, Activities tab) to capture real-time values like slide/rotate meters per daily operation.
4. **Generates** 3 clean, properly-typed CSV files ready for Excel or any downstream system.

### CSV Outputs

| File | Contents |
|------|----------|
| `fieldcap-job-{id}-job-details.csv` | Core job metadata + all custom field key-value pairs |
| `fieldcap-job-{id}-crew.csv` | One row per crew member with role, contact, dates |
| `fieldcap-job-{id}-bha-equipment.csv` | One row per BHA component with serial, hours, meters, fatigue data |

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

### Slide/Rotate Meter Extraction

The extension captures per-BHA slide and rotate meters through a multi-tier approach:

1. **Activities tab scraping** — detects rows with `Sliding`/`Code 2A` and `Rotating`/`Code 2` activities, sums the `Course` (meters) column per BHA.
2. **Depth range delta** — if `Course` is blank, computes meters from `Start-End Depth` values.
3. **Hour-proportional fallback** — if Activities data is unavailable, derives meter split from `Total Metres * (slide_hours / (slide_hours + rotate_hours))`.

### Installation

1. Open `chrome://extensions` (or `edge://extensions`).
2. Enable **Developer mode**.
3. Click **Load unpacked** → select `src/chrome-extension/`.
4. Navigate to FieldCap, ensure you're logged in.
5. Click the extension icon, enter Job ID, and click **Fetch & Build CSVs**.

### Usage Tips

- Keep the **Activities tab** open/loaded before fetching to get slide/rotate meter data.
- Keep the **BHAs tab** visited to capture hour data from FieldCap's internal API calls.
- The popup shows `live bhaRows` and `live activityRows` counts after each fetch — if `activityRows=0`, Activities data was not captured for that run.

---

## Excel DD Tools Dashboard

### What It Does

Imports the 3 exported CSVs and renders a fully interactive directional drilling dashboard:

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
4. Click **Refresh** button to import CSVs from the workbook's directory.

### How Refresh Works

- Scans the workbook's directory for CSVs matching `bha-equipment`, `crew`, `job-details`.
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
├── FieldCap-Exporter-Extension*.png  (screenshots)
│
└── src/
    ├── chrome-extension/
    │   ├── manifest.json             (Manifest V3, v2.5.0)
    │   ├── background.js            (OData fetch, CSV generation, activity parsing)
    │   ├── content.js               (DOM scraping, table detection, auto-scrape)
    │   ├── injected-spy.js          (page-context XHR/Fetch interception)
    │   ├── popup.html               (extension UI — PHX dark/teal theme)
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

## Roadmap: Field-Office-Equipment Ecosystem

### Current State (v2.5.0)
- Chrome extension exports structured CSVs from FieldCap
- Excel VBA dashboard provides interactive BHA/fatigue tracking

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
- No telemetry, analytics, or external API calls beyond FieldCap's own OData endpoint.

---

## Support

For PHX Technology FieldCap platform issues: [phxtech.com](https://www.phxtech.com)

---

## License

Proprietary — PHX Technology internal tooling. Not for public distribution.
