# PHX OpenCap — Data Export & Field Reporting Ecosystem

**Chrome Extension + Excel Slide Sheet Workbook**  
*Browser capture. Workbook intelligence. Same job folder.*

**Project site:** [freebrew.github.io/OpenCap_PHX_Exporter](https://freebrew.github.io/OpenCap_PHX_Exporter/)  
**Field workbook overview:** [docs/PHX_OpenCap_Field_Workbook_Overview.md](docs/PHX_OpenCap_Field_Workbook_Overview.md)  
**Data Hub prospect notes:** [docs/FieldCap_Data_Hub_Prospect_Report.md](docs/FieldCap_Data_Hub_Prospect_Report.md)

**Current Chrome extension revision:** `v3.1.8`

---

## Overview

PHX OpenCap is an independent toolset that extracts operational drilling data from the [FieldCap](https://fieldcap.com/) web application and delivers it into structured formats for downstream use in Excel-based field reports, custom dashboards, and (planned) web-based office/field ecosystems.

| Component | Location | Purpose |
|-----------|----------|---------|
| Chrome Extension | `src/chrome-extension/` | Extracts job, crew, BHA, inventory, slide/rotate metres, and ticket costs from FieldCap |
| Slide Sheet workbook | `Slide Sheet - Demo.xlsm` | GitHub-safe Excel macro workbook: surveys, slides, Setup/OpenCap import, proximity gauges, wellbore corridor |
| Dual-workbook tooling | `bas/_dual_*.ps1` | Keeps Demo and local Field VBA identical without touching Field job data |

---

## Dual workbook model (critical)

Two workbooks share **one VBA codebase**:

| Role | File | GitHub | Data |
|------|------|--------|------|
| **Primary (Demo)** | `Slide Sheet - Demo.xlsm` | Yes (mock / sanitized) | Fictional personnel & well data |
| **Field mirror** | `Slide Sheet - 34801.xlsm` | **Never** | Real job data for field test |

**Rules**

1. Edit / prove VBA on **Demo** first, then mirror modules to Field.
2. Field **user-entered data is sacred** — surveys, slides, Setup values, costs, plans, etc. must not be altered by sync tooling (VBA `CodeModule` replace only).
3. Never commit or push the Field workbook (or Field backups / dual caches that may contain real exports).
4. After VBA changes: backup both → mirror → `bas/_dual_diff_vba.ps1` must **PASS**.

```powershell
powershell -NoProfile -File bas/_dual_mirror_demo_to_field.ps1
powershell -NoProfile -File bas/_dual_export_caches.ps1
powershell -NoProfile -File bas/_dual_diff_vba.ps1
```

`Module11` mail headers (`.To` / `.CC` / `.Subject`) may differ (Demo mock vs Field real) until they are sheet-driven.

**Setup sheet warning:** `RebuildSetup` / `InitSetup` clear and redraw the entire Setup tab from `MDL_Setup.bas`. Hand-tuned layout must be encoded in that module or it will be wiped.

---

## Slide Sheet workbook (Demo)

GitHub ships the sanitized Demo workbook. Staged VBA lives under `bas/` (pull from the live `.xlsm` after manual Excel edits — do not treat stale `.bas` as authoritative over the workbook).

### Highlights

- **OpenCap Setup** — import FieldCap CSVs from `OpenCap/`, job/crew/contacts dashboard, plan & anti-collision import.
- **Plan proximity gauge** (Slidesheet `AA1:AC7`) — dial in AA→AB, GRAVITY/PTB metrics ending in AC; Clear Ranges / Pipe Tally parked in merged `Z1:Z3` / `Z4:Z6`.
- **Wellbore corridor** image — 3D-style MD window with geo ribbon, plan path, floor L/R shadow (AA14 lateral tolerance), and reporting-day band for daily email/PNG export.
- **Position of Wellbore** — R/L from plan, above/below plan (gauge), distance from current geo target.
- **Pipe tally, day roll, costs form, TD calc, sheet protect** — supporting field ops macros.

Rendered sample: `docs/corridor_live.png`.

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
│  content.js · injected-spy.js · background.js · popup UI         │
│  → CSV files under workbook OpenCap/ folder                      │
└─────────────────────────┬────────────────────────────────────────┘
                          │
                          ▼
┌─────────────────────────────────────────────────────────────────┐
│  Slide Sheet - Demo.xlsm  (GitHub)  │  Slide Sheet - 34801.xlsm │
│  mock data · same VBA               │  real data · local only   │
│  Setup · Slidesheet · gauges · corridor · daily report PNG      │
└─────────────────────────────────────────────────────────────────┘
```

---

## Chrome Extension

### What It Does

1. **Fetches** job details, crew, BHA equipment, inventory, daily slide/rotate metres, and ticket costs from FieldCap OData using your browser session.
2. **Intercepts** FieldCap XHR/Fetch for hour/metric fields not on public endpoints.
3. **Scrapes** the visible DOM (BHA grid) when needed.
4. **Suppresses** the FieldCap "Bottom Line Verification" popup when enabled.
5. **Writes** CSVs into an `OpenCap/` subfolder under the folder you choose.

### CSV Outputs

| File | Contents |
|------|----------|
| `fieldcap-job-{id}-job-details.csv` | Core job metadata + custom fields |
| `fieldcap-job-{id}-crew.csv` | Crew members with role / contact / dates |
| `fieldcap-job-{id}-bha-equipment.csv` | BHA components with serial, hours, metres |
| `fieldcap-job-{id}-inventory.csv` | Job tools / inventory |
| `fieldcap-job-{id}-slide-rotate-metres-by-day.csv` | Daily slide & rotate metres per BHA (`ActivityLogs`) |
| `fieldcap-job-{id}-ticket-costs-by-day.csv` | Daily ticket cost totals |

### Installation

1. Open `chrome://extensions` (or `edge://extensions`).
2. Enable **Developer mode**.
3. **Load unpacked** → `src/chrome-extension/`.
4. Log into FieldCap, open the extension popup, enter Job ID, **Fetch & Build CSVs**.

---

## Project Structure

```
PHX_FieldCap/
├── README.md
├── Slide Sheet - Demo.xlsm          ← GitHub workbook (sanitized)
├── Slide Sheet - 34801.xlsm         ← local Field only (gitignored)
├── index.html                       ← GitHub Pages showcase
├── bas/                             ← staged VBA + dual-sync scripts
│   ├── MDL_*.bas / Module*.bas
│   ├── _dual_export_caches.ps1
│   ├── _dual_diff_vba.ps1
│   ├── _dual_mirror_demo_to_field.ps1
│   └── _sanitize_setup_for_github.ps1
├── backups/                         ← local only (gitignored)
├── docs/                            ← overviews, corridor/gauge PNGs
└── src/
    ├── chrome-extension/            ← Manifest V3, v3.1.8
    ├── excel/MDL_DDTools.bas
    └── ...
```

---

## Changelog

### Slide Sheet — 2026-08-10

- **Dual-workbook sync** — Demo primary / Field mirror; `bas/_dual_*.ps1` export, mirror, and verify; Field workbook + dual caches gitignored.
- **Wellbore corridor** — sloping geo ribbon, plan path, corrected floor L/R orientation, back-wall TVD shadow only, Position of Wellbore numbers from gauge + geo window.
- **Proximity gauge layout** — dial starts in AA into AB; GRAVITY/PTB metrics end in AC; Clear Ranges / Pipe Tally in `Z1:Z3` / `Z4:Z6`; caption clutter removed.
- **Setup layout persistence** — section gap heights and notes box size encoded in `MDL_Setup` so Rebuild does not silently undo hand-tuned UI.

### v3.1.8 — Ticket Costs by Day Export

- **New export: Ticket Costs by Day** — 6th CSV (`Date`, `Daily Cost`, `Ticket Count`).
- **Costs tab upsert** — Setup **REFRESH CSVs** upserts daily / running totals on the Costs sheet by calendar date.

### v3.1.7 — Rig Name Capture Fix

- Rig Name cached from Well Parameters / Rig pages; empty custom-field values no longer wipe a good scraped name.

### v3.1.6 — Extension List Sort Name

- Display name `.OpenCap Data Exporter` for sorting near the top of Chrome’s extension list.

### v3.1.5 — OpenCap CSV Subfolder

- CSVs write to `OpenCap/`; Excel Refresh scans that folder first.

### Earlier

See git history for v3.1.4–v2.5 (popup layout, BLV suppression, inventory export, ActivityLogs slide/rotate metres).

---

## Roadmap

- Web spreadsheet / office integration and real-time sync
- Template engine for daily / BHA / tool reports
- Field → office pipeline and equipment lifecycle tracking

---

## Data Privacy

- Extension uses your FieldCap browser session — no credentials stored or sent elsewhere.
- CSVs stay local (`OpenCap/` on disk).
- **GitHub contains only the sanitized Demo workbook** — never the Field workbook or real job exports.
- Run `bas/_sanitize_setup_for_github.ps1` before publishing Demo if live data may have been loaded.
- No telemetry beyond FieldCap’s own APIs.

---

## Disclaimer

This project (OpenCap) is an independent, unofficial tool. It is **not affiliated with, endorsed by, or supported by [PHX Technology](https://www.phxtech.com/), [FieldCap](https://fieldcap.com/), or any other third party.** All trademarks and product names belong to their respective owners.

The name [FieldCap](https://fieldcap.com/) is used solely to describe the third-party web application this tool interacts with.

---

## License

MIT — free to use, modify, and distribute.
