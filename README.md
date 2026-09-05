# PHX OpenCap — Data Export & Field Reporting Ecosystem

**Chrome Extension + Excel Slide Sheet Workbook**  
*Browser capture. Workbook intelligence. Same job folder.*

**Project site:** [freebrew.github.io/OpenCap_PHX_Exporter](https://freebrew.github.io/OpenCap_PHX_Exporter/)  
**Field workbook overview:** [docs/PHX_OpenCap_Field_Workbook_Overview.md](docs/PHX_OpenCap_Field_Workbook_Overview.md)  
**Data Hub prospect notes:** [docs/FieldCap_Data_Hub_Prospect_Report.md](docs/FieldCap_Data_Hub_Prospect_Report.md)

**Current Chrome extension revision:** `v3.2.3`

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
| **Field mirror** | `Slide Sheet - 35780.xlsm` | **Never** | Real job data for field test |

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

When a **live Field workbook** is the newest code source of truth, run `bas/_adopt_field_source_of_truth.ps1` (skill `adopt-field-workbook`): VBA only into Demo + `bas/`. Bind by full path if the open book is not the repo copy (e.g. `-FieldWorkbook "D:\Slide Sheet - 35780.xlsm"`). Field job data is not written. Do not run `_dual_export_caches.ps1` while Field is open (it kills Excel).

**Setup sheet warning:** `RebuildSetup` / `InitSetup` clear and redraw the entire Setup tab from `MDL_Setup.bas`. Hand-tuned layout must be encoded in that module or it will be wiped.

---

## Slide Sheet workbook (Demo)

GitHub ships the sanitized Demo workbook. Staged VBA lives under `bas/` (pull from the live `.xlsm` after manual Excel edits — do not treat stale `.bas` as authoritative over the workbook).

### Highlights

- **OpenCap Setup** — import FieldCap CSVs from `OpenCap/`, job/crew/contacts dashboard, plan & anti-collision import.
- **Crew manifest** — six tight rows in the same Setup block (J3:R10); title shows `6+N` when OpenCap crew is longer. Role dropdown remains DD/MWD.
- **Import Plan** — COMPASS well-plan **PDF** (Plan Sections / SECTION DETAILS). Names targets **KOP / TANGENT / SOT / EOT / HEEL / TD** (Y2:Y5 dropdown to override). All sections stay on hidden `_OC_PlanSec`; Slidesheet `T2:Y5` shows four at a time and slides as the bit passes the highlighted next target. Sibling `.csv` next to the PDF still loads the dense plan surveys into `_OC_Survey` (gauge / Planned TD); without a CSV, the sparse section stations are used.
- **Plan proximity gauge** (Slidesheet `AA1:AC7`) — dial in AA→AB, GRAVITY/PTB metrics ending in AC; Clear Ranges / Pipe Tally parked in merged `Z1:Z3` / `Z4:Z6`. GRAVITY frame uses the **Well Seeker convention** (perpendicular offset from the plan line, high-side/right decomposition) so UP/DN·LT/RT match office software; when the sail waypoint corridor is active, UP/DN switches to true TVD so the dial agrees with the geo window.
- **Daily report PNG** (`RenderCorridorPng`) — EMAIL attaches `daily_report.png` as a visible file and inlines a CID copy under the Data table. No daily report PDF is attached. Paper/print theme (white ground, thin strokes). **VERTICAL / BUILD:** three-wall shadow box (back, left, floor) — **plan** in the window, **as-drilled last 24 h** (tight camera on the report day; a long/stale 00:00 day zooms to the last ~150 m + look-ahead), named targets on the plan with a ring **perpendicular to the wellbore**, shadows on all three walls, GN arrow. No min-curve and no surface-to-TD. **LATERAL:** 3D room with geo ± (AB14) and AA14 L/R bands.
- **Position of Wellbore** — R/L from plan, above/below plan (gauge), distance from current geo target.
- **TD calculator** (Data `I53:M55`) — Actual is last Slidesheet survey projected to bit. Planned TD is the last imported plan station. Actual TD is the original remaining-NS/EW MD estimate (`SQRT` × ratio factor from heading vs remaining-vector azimuth), not a copy of Planned TD.
- **Tool hours** — 3rd-party table seeds the active BHA’s (`H3`) rental / Other Inventory serials (works without a BHA `Source` column). MWD kit, bits, and tubulars stay out. **Orbit RSS** and **iCruise** seed into the Motors table. Hours follow BHA Total Hrs / live `Q24`.
- **Pipe tally, day roll, costs form, sheet protect** — supporting field ops macros.

Rendered sample: `docs/corridor_24h.png` (24 h shadow box). Older full-well still: `docs/corridor_live.png`.

---

## Slide advisor math (BURR, slide, toolface, rotate)

Columns (data rows 13:305):

| Col | Meaning |
|-----|---------|
| **C** | Course length of **this stand** (m). Hard cap for any slide instruction. |
| **D** | Bit MD |
| **E** | Survey MD |
| **Q** | Seen motor yield (`C×J/M`) |
| **T** | Metres **already slid** (history). Not used for Y or Z. |
| **U** | User toolface (what they ran): `40R` / `R40` / `-40` / `40` highside, `190M` magnetic. Parsed in hidden `AK`; unparseable text is `#VALUE!` (never silently TF 0). |
| **V / W / X** | TVD / INC / AZM @ bit. INC and AZM use the same seen-yield dogleg (`DoglegBelow`). |
| **AO** | Active target MD |
| **AR** | BURR (°/30 m) |
| **AS** | Metres to slide (along-motor, before TF cosine) |
| **AT** | Required toolface to the aim |
| **Y** | Comment: `Sliding <m> @ <TF>` … `BURR n.nn` |
| **Z** | Leftover rotate: `C − instructed slide` |

`T2:Y5` is a **display window** only. BURR / AS / AT all read the **full** named list on `_OC_PlanSec` (`ProjTargets_MD/INC/AZM/TVD`) and all start from the **projected bit** (`W` / `X` = Inc / Azm @ bit, never survey `F` / `G`), so the metres, BURR and toolface in Y describe one manoeuvre to one target. `RefreshSlideComments` rewires any legacy `ProjBurr(D,F,G,…)` / `ProjMetersToSlide(D,F,G,…)` helper formula to `W`/`X` on the fly. First KOP / TANGENT MD is `ProjBuildStartMd` — that is the gate for AR/AS/AT. **Do not use `$U$2` as a build-start gate.** U2 is whatever target the four-row window currently shows (often SOT after the window has slid).

### Aim

Stay on a build station (e.g. TANGENT) while `inc_bit + 0.5° < target INC`. Do not jump to SOT because leftover MD to TANGENT is small. Tiny leftover MD is not a reason to skip — use **this stand’s C** as the BURR distance instead of a made-up floor (never invent 10 m).

### BURR (AR)

Same rate as a working build stand:

```
BURR = (I_target − I_bit) × 30 / dMD
```

`dMD` = target MD − bit MD. If that is shorter than this stand’s **C** and inclination is still short of the target, **dMD = C**.

When bit TVD is present and remaining TVD ≥ 5 m, use the TVD-arc rate instead (avoids a tiny-dTVD spike):

```
BURR = (sin I_target − sin I_bit) / dTVD × 30     (°/30 m)
```

BURR is what the **next aim demands**, even if the motor cannot deliver it this stand. No BURR above `ProjBuildStartMd` (vertical / nudge).

### Metres to slide (AS) and the comment (Y)

```
AS = min( |BURR| / Q × C , C )
```

AS is along-motor. A toolface off highside delivers less build per metre of hole. The **comment** metres are:

```
instructed = min( AS / |cos(TF)| , C )
```

`TF` is **AT** (required toolface to the target), not user column U. Display is **this row’s C at 2 decimals** — never snap to 0.25 m (that turned 19.20 / 19.16 into a fake 19.25).

**You cannot slide more than this stand.** `/|cos(TF)|` can ask for more hole than **C** (e.g. R42 ≈ 1.35×). That is why a full-stand instruction appears. The instruction is then **hard-capped at C**. If the cap binds: Y = `Sliding <C>m @ TF` and Z = `0.00m ROT`.

### Remaining rotate (Z)

```
Z = C − instructed     (0 if we slide the whole course)
```

Y slide + Z rotate **always equals C**. Z does **not** use column T (what they already slid). T is history; Y/Z is the instruction for this stand. No BURR → instructed = 0 → Z = C (all rotate).

### INC / AZM at bit (W / X)

Same dogleg-below as each other: seen DLS scaled by this stand’s course, metres already seen, and metres below the survey (`DoglegBelow`). Constant-toolface fill-in is ISCWSA §8.2:

```
ΔInc = dB × cos(TF)
ΔAzm = dB × sin(TF) / sin(Inc)     (Inc ≥ 5°)
```

INC uses column **U** (what they ran this stand). The AZM walk is **calibrated by demonstrated results**: N/O (effective TF back-computed from the last surveyed course) caps the dial TF when both agree in direction — the hole delivers less turn than the dial says (reactive torque / rotary dilution; a stated 30R has measured 5–27R effective on this well), and projecting the dial literally overshot the walk up to 10°/stand at low inc, flipping AT to the wrong side of the target. Dial blank/0 → N/O alone (rotary stand); N/O opposite sign or blank → dial (demonstrated TF is noise). Near-vertical (Inc < 5°) skips `/sin(I)` so walk does not blow up.

### Toolface (AT)

Required TF aims at the **nearest full-plan station** at least 15 m past the bit (`_OC_Survey`, 30 m grid → aim lands 15–45 m ahead) — tracking the plan means matching its attitude station by station, correcting azimuth **now** where turning is cheap. Never the great-circle to a distant named target: that let azimuth converge lazily at the end and said R8 in a build the driller had to slide at 30R. Named targets are only the fallback when no plan is imported.

From bit attitude to aim attitude, the TF is the industry-standard (Well Seeker) spherical toolface:

```
TF = Atan2( sinI2·cosI1·cosΔA − sinI1·cosI2 ,  sinI2·sinΔA )
```

A degree of azimuth walk only moves the bit `sin(Inc)` as far as a degree of build, so a raw `Atan2(ΔInc, ΔAzm)` over-weights the turn (~2.4× at Inc 25°) — verified against the Well Seeker plan TF column to ±0.1°. Comments show that TF (`R10`, `L9`, …; ±90 = pure turn). Magnetic vs gravity mode follows bit inc (default 5°); magnetic shows the **bearing of the push** (`azm_bit + TF`, e.g. `20M`).

### What we do not do

- Gate V/W/X/Y on `$U$2` (window MD), or compute BURR on the whole vertical aimed at SOT.
- Invent a 10 m BURR floor, next-row C, or quarter-metre slide lengths.
- Put leftover `m ROT` in Y — rotate lives in **Z**.
- Paint comment cells neon yellow (`65535`). Y comments and Z use faded `RGB(255,255,204)`; blank Y stays the sheet’s pale green.

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
│  Slide Sheet - Demo.xlsm  (GitHub)  │  Slide Sheet - 35780.xlsm │
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
├── Slide Sheet - 35780.xlsm         ← local Field only (gitignored)
├── index.html                       ← GitHub Pages showcase
├── bas/                             ← staged VBA + dual-sync scripts
│   ├── MDL_*.bas / Module*.bas
│   ├── _dual_export_caches.ps1
│   ├── _dual_diff_vba.ps1
│   ├── _dual_mirror_demo_to_field.ps1
│   ├── _adopt_field_source_of_truth.ps1
│   └── _sanitize_setup_for_github.ps1
├── backups/                         ← local only (gitignored)
├── docs/                            ← overviews, corridor/gauge PNGs
└── src/
    ├── chrome-extension/            ← Manifest V3, v3.2.3
    ├── excel/MDL_DDTools.bas
    └── ...
```

---

## Changelog

### Slide Sheet — 2026-09-04

- **Actual TD formula restored** — `TdEstFinal` had copied Planned TD (INC/AZM/NS/EW, and usually MD). `I55:M55` is again the merged original: current MD + remaining NS/EW distance × RF from heading vs remaining-vector azimuth. Planned TD values are unchanged.
- **3rd-party tools seed without `Source`** — official FieldCap `bha-equipment.csv` has no `Source` column, so the old `Other Inventory` gate left the table empty. `MDL_ToolHours` now also matches inventory `Category = Other Inventory` or a **Rental** description/name on the selected BHA (`H3`). MWD is skipped by inventory **Category** (SubCategory is often blank). Slick/pony collars stay out as tubulars.
- **Orbit RSS / iCruise → Motors** — those names seed into *Enter Motors & hours below* (and Motors On Location) with mud motors, not the 3rd-party table.

### Slide Sheet — 2026-08-30

- **AZM-at-bit walk calibrated by results** — X projects the walk with the smaller of dial TF (U) vs demonstrated effective TF (N/O, same sign) instead of trusting the dial literally. Measured on this well: mean walk error 2.48° → 2.05°/stand, worst-case 9.9° → 5.4°, and the AT recommendation at MD 1670.49 lands on the correct side of the target (R13, was L26 from a bit-azm projection 10° hot).
- **Required toolface (AT) matches Well Seeker** — spherical toolface (`Atan2(sinI2·cosI1·cosΔA − sinI1·cosI2, sinI2·sinΔA)`) instead of the flat inc-rate/azm-rate `Atan2`, aimed at the nearest full-plan station 15–45 m ahead of the bit instead of the distant named target. In the 1600–1806 build this turns useless single-digit calls (R8) into the R27–R28 the driller actually had to slide. Magnetic mode now shows the push bearing (`azm_bit + TF`).
- **Proximity gauge frame** — GRAVITY UP/DN·LT/RT now use Well Seeker's perpendicular high-side/right decomposition (matches office "distance from plan" exactly; verified UP 2.25 / LT 1.48 vs WS 2.25 / 1.51). With the sail waypoint corridor active, UP/DN stays true TVD so dial and green band agree with the geo window.
- **Corridor E-axis scale** — floor numbers moved from the back-wall joint to the open front edge.
- **Mirror guard** — `_dual_mirror_demo_to_field.ps1` fails loudly when Field opens read-only (open workbook) instead of silently not saving.
- **BURR / slide / ROT** — Aim stays on TANGENT while inc is still short (do not skip to SOT on leftover MD). BURR uses real dMD or **this stand’s C**, never a 10 m floor. Comment slide = `AS/|cos(TF)|` hard-capped at C (2 dp, no 0.25 snap). Z = C − instructed slide (0 if full-course). Gate AR/AS/AT on `ProjBuildStartMd`, not `$U$2`.
- **24 h shadow box** — Corridor VERTICAL/BUILD shows plan + last 24 h hole, three-wall shadows, targets on the plan with a ring perpendicular to the wellbore, camera on the report footage (not surface-to-TD). Sample: `docs/corridor_24h.png`.
- **Comment yellow** — Y/Z use one faded yellow (`RGB(255,255,204)`).

### Slide Sheet — 2026-08-29

- **Daily EMAIL PNG** — EMAIL attaches a visible `daily_report.png` (plain attachment) and inlines a second CID copy under the Data table. Outlook hides CID-referenced images from the paperclip list, so the visible file is required. No daily report PDF is generated or attached. User-pinned slots (Data `I29:I33`) still attach as before.
- **Vertical / build 3D shadow-box** — whole-well 3D plot in the style of the classic wall/floor-shadow well-path chart: full plan, all surveys from surface, dashed minimum-curvature projection to the next T2:Y5 target, target ring marker, TVD/N/E grids, wall/floor shadows, GN arrow. Replaces the traveling-cylinder bullseye, which only showed a slice around the target. Lateral keeps the geo/AA14 room. Paper theme for print.
- **Adopt live Field** — `_adopt_field_source_of_truth.ps1` accepts a rooted path so `D:\Slide Sheet - 35780.xlsm` can be code SOT without touching job cells.

### Slide Sheet — 2026-08-28

- **Import Plan PDF** — Setup **Import Plan** prefers a COMPASS planning-report PDF. Parses Plan Sections / SECTION DETAILS into `_OC_PlanSec`, auto-names KOP / TANGENT / SOT / EOT / HEEL / TD (Y dropdown override). Slidesheet `T2:Y5` is a 4-row window that advances with bit MD. Dense plan surveys still come from a same-stem `.csv` beside the PDF when present; CSV-only import remains.
- **Crew density** — Setup crew panel holds 6 people in J3:R10 (tighter rows); AC table moves to J11:R16. Overflow count in the section title. `RedrawCrewAndAc` redraws that block without a full Rebuild.

### Slide Sheet — 2026-08-10

- **Dual-workbook sync** — Demo primary / Field mirror; `bas/_dual_*.ps1` export, mirror, and verify; Field workbook + dual caches gitignored.
- **Wellbore corridor** — sloping geo ribbon, plan path, corrected floor L/R orientation, back-wall TVD shadow only, Position of Wellbore numbers from gauge + geo window.
- **Proximity gauge layout** — dial starts in AA into AB; GRAVITY/PTB metrics end in AC; Clear Ranges / Pipe Tally in `Z1:Z3` / `Z4:Z6`; caption clutter removed.
- **Setup layout persistence** — section gap heights and notes box size encoded in `MDL_Setup` so Rebuild does not silently undo hand-tuned UI.

### v3.2.3 — Other Tools are item-less JobTools

- Probe on job 21782 showed the blank BHA rows each carry a `JobToolId` — the third-party tools are ordinary JobTools with no `Item` / `ItemSerial` link, their serial and description typed straight onto the JobTool. v3.2.0 only looked for a separate table.
- `normalizeBhaRow` now reads serial / code / description from the JobTool's own fields when it has no Item link (`Source = Other Inventory`); `buildInventoryCsv` keeps those JobTools instead of dropping them as placeholders and tags them `Category = Other Inventory`. Key picking skips `*Id` / GUID / timestamp fields.
- Probe reports the item-less JobTools it recognised plus the raw JobTool behind any still-unresolved item; `$metadata` parser tolerates namespace-prefixed tags. `Sub Description` / `SubCategory` carry FieldCap's Category + SubCategory (`DD other`, `MWD other`, `DD drill bit`, `DD tubular`).
- **Slide Sheet seeding is BHA-scoped** — Refresh fills *Enter any 3rd Party Tools/hours below* (`O31:O41`) from the selected BHA (`H3`): `Source` / `Category = Other Inventory`, or a Rental description. Skips MWD kit (by Category), drill bits, and tubulars. Serial matching ignores spaces / dashes / case. Orbit RSS and iCruise go to the Motors table (`O45:O55`).

### v3.2.0 — Other Tools (third-party) in the existing CSVs

- **BHA components without a JobTool now resolve** — third-party / rental tools live in FieldCap's *Other Tools* table, so those rows exported with no serial or description. The exporter discovers the `ToolAssemblyItem` navigation property from OData `$metadata` (falls back to sniffing populated `*Id` keys on bare items), verifies it with a live `$expand`, and fills `Serial #` / `Item Code` / `Description` / `Sub Description`.
- **No new CSV** — tagged in place: `bha-equipment.csv` gains a trailing `Source` column (`Other Inventory` or blank); `inventory.csv` appends the job's Other Tools with `Category = Other Inventory` (SubCategory = vendor / type when the table has one).
- **Other Tools Probe** (footer button) — shows the discovered navs / entity sets, the bare BHA items, a sample resolved tool, and the rows that would be appended. Share its output if nothing resolves for a tenant.
- **Slide Sheet Data** — Setup / Data **Refresh** seeds *Enter any 3rd Party Tools/hours below* (`O31:O41`) from the selected BHA. Official FieldCap CSVs without `Source` still seed Rental / Other Inventory rows; Chrome exports that tag `Source` keep working. `MDL_ToolHours` matches them for `Current Hours`. `C31` *Previous Motor Hours* self-heals to VLOOKUP column 3 (Q, previous) — it pointed at column 4 (R, current) and echoed `C32`.

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
