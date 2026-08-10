# PHX OpenCap Field Workbook

### Browser-connected wellsite intelligence — inside Excel

The PHX OpenCap Field Workbook turns a single `.xlsm` into a **field operations cockpit**: job identity, crew, BHA and inventory from FieldCap; daily slide/rotate and cost totals; a live Slidesheet for directional guidance; and one-click morning reports for the office — without rebuilding your workflow in another platform.

It is designed as a **two-layer system**:

1. **Chrome extension** — captures structured data from FieldCap while you work in the browser.
2. **Standalone Excel workbook** — imports those CSVs, then does everything FieldCap does not: plan vs. wellbore geometry, proximity guidance, bit projections, anti-collision context, and outbound HTML/PNG daily reports.

Together they close the gap between “what FieldCap already knows” and “what the directional team needs on the pad and in the morning email.”

---

## Why it exists

Directional drilling teams already live in two places: a web job system for equipment and tickets, and Excel for surveys, plans, and the language the rig actually drills by. Copy-paste between them is slow, error-prone, and easy to get out of date by lunchtime.

OpenCap’s answer is deliberate:

- **Pull** authoritative operational data from FieldCap with a one-click export.
- **Drop** the files next to the workbook (`OpenCap\`).
- **Refresh** once inside Excel — Setup, DD Tools, Costs, and motor tables update together.
- **Guide** the well on the Slidesheet with the same plan maths used for proximity dials and emailed corridor graphics.
- **Report** from the Data tab — HTML daily summary plus an attached corridor PNG — without leaving the workbook.

No second login for Excel. No separate “reporting database.” The workbook folder *is* the integration bus.

---

## Architecture at a glance

```
FieldCap (web)
    │  OData + session XHR/Fetch + DOM scrape
    ▼
Chrome extension  (.OpenCap Data Exporter)
    │  CSV pack → OpenCap\
    ▼
Slide Sheet workbook  (.xlsm)
    ├── Setup                 — job / crew / imports / AC
    ├── Data                  — ops KPIs, motors, EMAIL, reports, costs
    ├── Slidesheet            — surveys, plan, proximity, projections
    ├── DD Tools              — BHA selector & fatigue view
    ├── Pipe Tally Calculator — joint-length paste → bit depths
    └── Sidetrack (+ Points)  — sidetrack planning, shown on demand
```

Hidden sheets (`_OC_*`, `_FC_*`) hold the imported tables. Visible sheets stay calm for humans; macros do the join work.

---

## Part 1 — Browser extension: capture without retyping

The companion **Manifest V3 Chrome extension** (OpenCap Data Exporter) is the intake valve. While you are logged into FieldCap, it uses your active session to assemble a clean CSV pack for the current job.

### What gets exported

| Export | Role in the workbook |
|--------|----------------------|
| Job details | Job identity, contacts, land, rig, directional constants → Setup |
| Crew / personnel | Crew manifest names, emails, phones, roles → Setup |
| BHA equipment | Serials, hours, metres, fatigue → DD Tools + Data motors sync |
| Inventory | Tool status and hour fields for job tooling |
| Slide / rotate metres by day | ActivityLog-derived daily slide vs rotate metres per BHA |
| Ticket costs by day | Calendar daily ticket totals → Costs form / Costs sheet |

Files land as `fieldcap-job-{id}-{token}.csv` under the workbook’s `OpenCap\` folder (File System Access API), so Excel never has to hunt a Downloads pile.

### How capture stays complete

FieldCap does not expose every useful field on a single public endpoint. The extension therefore combines three techniques:

- **OData fetches** for job, crew, BHA, inventory, ActivityLogs, and tickets.
- **XHR/Fetch interception** for hour and metric fields that only appear in FieldCap’s own traffic.
- **DOM scrape** of the BHA grid when the UI holds the live truth.

Optional FieldCap UI helpers (for example dismissing the Bottom Line Verification nag) keep the browser session usable during long exports — without changing FieldCap’s data.

### The human workflow (extension → Excel)

1. Open the job in FieldCap.
2. Open the OpenCap popup, confirm Job ID, tick the CSV set you need.
3. **Fetch & Build** → write into this workbook’s `OpenCap\` directory.
4. In Excel, click **REFRESH CSVs** (Setup) or **REFRESH** (Data / DD Tools).
5. Review Setup status rows, then drill — surveys and plan stay on the Slidesheet where they belong.

That is the whole bridge: browser session → typed CSVs → VBA refresh → live dashboard.

---

## Part 2 — Standalone workbook: everything that has to live next to the hole

Once CSVs are in, the workbook is **fully usable offline**. You do not need the extension open to project the bit, paint the proximity dial, edit costs, or send EMAIL. The extension is for sync; Excel is for work.

### Setup — well database & import console

Setup is the operational front door:

- **Job identity** — job ID/name, client, UWI, surface / land location, province, AFE, licence, profile.
- **Well geometry & directional constants** — KB, ground, lat/long, mag model, declination, tolerances.
- **Equipment & rig** — rig type/name, MWD guidance, top drive / loader flags.
- **Contacts** — company men, geologist, DD/MWD coordinators, sales — with mailto links.
- **Crew manifest** — role / name / email / phone; role dropdown and sort helpers for DD vs MWD.
- **OpenCap export status** — which CSVs were found, row counts, paths.
- **Import Plan / Import AC** — bring in the survey plan and anti-collision summary without leaving Setup.
- **REFRESH CSVs / REBUILD** — refresh rebuilds from disk; rebuild is the heavier UI reset (use deliberately).

Imported rows land in hidden `_OC_*` tables so the next refresh can merge cleanly. Mud motors that look like PHX BHA motors sync into Data’s “Motors On Location.” Ticket costs overlay the Costs calendar (CSV wins on matching dates).

### DD Tools — BHA-centric equipment view

DD Tools is the equipment lens:

- Interactive **BHA selectors**.
- Per-BHA metres and hours (slide / rotate / circulate).
- Cumulative **serial-hour** tracking with fatigue emphasis (including high-hour warnings).
- Raw BHA summary for sheet formulas.
- Slide vs rotate metre breakdown fed by the ActivityLog export.

Same OpenCap pack, different job to do: Setup answers “who and where,” DD Tools answers “what is in the hole and how hard has it worked.”

---

## Part 3 — Data tab: the daily ops & reporting surface

Data is where the morning story is assembled. It is intentionally dense: depths, slide/rotate split, motor performance, wellbore position, AC concerns, motors on location, day-roll BHA mirror, and the TD calculator — then a toolbar that ships the story outward.

### Toolbar

**EMAIL · Mini Report · Full Report · Costs · REFRESH**

| Control | What it does |
|---------|----------------|
| **EMAIL** | Builds an Outlook message: HTML body from `Data!B2:F55`, plus corridor/ops **PNG attachment** rendered from the same plan/gauge maths used on the Slidesheet. Optional file paths in `I29:I33` attach as supporting docs. |
| **Mini Report** | Compact text morning summary from key cells — fast paste into chat or tickets. |
| **Full Report** | Longer structured morning summary (depths, slide %, motor info, position, etc.). |
| **Costs** | Modeless calendar form over a hidden Costs sheet — edit daily / surface batch costs without hunting grid rows. OpenCap ticket totals remain authoritative on refresh. |
| **REFRESH** | Same CSV import chain as Setup — one button when you are already on Data. |

### Reporting content the Data sheet owns

- **Last survey & position of wellbore** — lateral and up/down from the proximity / Sail frame, with total displacement; kept aligned with the Slidesheet gauge after render (`MDL_WellborePos`).
- **Motor performance & motors on location** — configuration, rev/L, hours; PDM conf list kept sorted for readable dropdowns.
- **Day roll** — BHA day rows autofilled from the BHA mirror; Pason bit/circ hour entry supported via form helpers.
- **AC info & concerns** — separation-factor context for the report body.
- **TD calculator** — Actual row tracks live bit extrapolation from the Slidesheet; Planned TD seeds from the imported plan.

### EMAIL that matches how people actually read mail

Outlook is unforgiving: huge single images get crushed; dark-mode HTML can blank tables. The current EMAIL path respects that:

- **Body** — light, fixed-width HTML table of the ops block people already know (`RangeToHTML` on `B2:F55`).
- **Attachment** — full wellbore corridor + ops PNG for anyone who wants the spatial picture at full resolution.

The corridor renderer shares plan/proximity helpers with the on-sheet gauge (`PG_*` shims), so the dial you drilled by and the picture you emailed do not disagree.

---

## Part 4 — Slidesheet: well guiding where surveys meet the plan

The Slidesheet is the directional console — measured depth marching down the page, plan targets beside you, and live projections from the last survey to the bit.

### Proximity gauge (plan vs. as-drilled)

A combined dial (plan center, survey marker, projected-to-bit marker) sits in the gauge block:

- **Gravity mode** (Inc ≥ 5°) — plan at the same MD; **UP/DN** = plan TVD − as-drilled TVD; **LT/RT** = cross-track at plan azimuth; waypoint TVD corridor (± geo half-width).
- **Magnetic mode** (Inc < 5°) — plan foot at actual TVD; horizontal distance (PRP); LT/RT of plan azimuth.

### Sail calculator & geo window

The **Sail** block is the lateral geo window the hole is flying through — not a side note to the dial:

- Waypoint table **AC14:AD33** (MD / TVD), with **Inc-to-next** and **geo window** displays (**AA** / **AB** on survey rows).
- Half-width and lateral tolerance (**AB14** / **AA14**) set the TVD corridor and left/right budget used by both the proximity gauge and the emailed corridor PNG.
- Sail **Up/down** on the survey row feeds Data’s “Position of Wellbore” so the morning report and the Slidesheet share one vertical story.

Enter the sail waypoints once; the calculator keeps geo high/low and the dial’s green band honest as MD advances.

### Target table → BURR → slide advisor (Comments)

Build guidance is driven by the **plan target table** (build targets in **T2:Y5** and related aim columns), not freehand guesswork:

1. **Active target** — the next unexhausted station ahead of the bit (`ProjActiveTargetMd` / highlight on the target block).
2. **BURR** — build-up rate required (°/30 m) to that aim (`ProjBurr`). With bit TVD available this is the constant-build (circular-arc) rate that lands target inclination at target TVD; the aim is *not* walked forward, so BURR answers “what does the next target demand,” even when the motor cannot deliver it.
3. **Metres to slide / remain to rotate** — `|BURR| / motor output × course`, then course remainder (`ProjMetersToSlide`, `ProjMetersRemainToRotate`).
4. **Required toolface** — gravity/magnetic TF mode from bit inclination; required TF text feeds the advisor.
5. **Slide advisor in Comments (column Y)** — auto lines such as  
   `Sliding 12.25m @ L90` … `BURR 2.40`  
   via `ProjSlideComment` + `RefreshSlideComments`. Effective slide metres are corrected by `|cos(required TF)|` so off-highside toolface asks for more slide. On tangent / zero-slide holds, Comments still show **TF + BURR** (`Sliding 0.00m @ …`) so the column does not go blank when metres-to-slide is zero.
6. **Slide footage below the survey** — `ProjSlideMetersBetween` sums slide intervals between the last survey and the bit, so projections account for slide the survey has not seen yet (MtrBelow).

Active-target highlighting keeps the row’s BURR aim and the lit target cell on the same station — the advisor and the table stay locked together.

That is the difference between “we have surveys” and “we know what to tell the directional hand next stand.”

### Plan import, AC, pipe tally, clear

- **Import Plan** (from Setup) loads the survey plan into `_OC_Survey` and refreshes gauge / Planned TD.
- **Import AC** parses anti-collision summary into a severity-sorted table (Setup display capped to the hottest rows).
- **Pipe Tally** form pastes joint lengths and appends bit depths onto the Slidesheet depth column — field entry without fighting the survey grid.
- **Clear Slidesheet** wipes survey/entry ranges while protecting the gauge furniture.

### Corridor graphic for the office

`MDL_CorridorImage` builds a cabinet-oblique “room” of the hole: back-wall TVD corridor (sloping geo ± band), floor lateral corridor, left-wall shadow looking down the hole, forward required vs hold projections, plus a dark ops band for day/BHA/motor/AC chips. It is the Slidesheet’s spatial argument — portable as a PNG.

---

## Part 5 — Run lifecycle & field utilities

A well is not one Slidesheet — it is a sequence of runs, day rolls, tallies, and (sometimes) a sidetrack. The workbook manages that lifecycle instead of leaving it to file copies.

### End-of-run roll (New Slidesheet / New Data)

When a BHA comes out, one macro (with a “have you finished the Data sheet?” guard) does the whole hand-off:

- **Archives the run** — exports the Data print area to a PDF named from the well identity (`{well}{run} DATA.pdf`) beside the workbook.
- **Clears the day fields** on Data for the next run.
- **Clones a fresh Slidesheet** in front of the current one — survey history preserved as a tab, new run starts clean with formulas and gauge furniture intact.

Every run of the well stays in the same file, in order, with its own archived report. No “Copy of Copy of Slidesheet FINAL v3.”

### Sidetrack mode

Dedicated **Sidetrack** and **Sidetrack Points** sheets stay hidden until needed, then toggle on with one macro when the well plan changes underfoot. Kick-off planning gets its own surface instead of being scribbled over the original plan.

### Pipe tally stand modes

The Pipe Tally Calculator adapts to how the rig actually racks pipe — **doubles, triples, or triple-half stands** — by switching column layouts on demand, plus a one-click tally clear between runs. The tally form pastes joint lengths and appends bit depths straight onto the Slidesheet.

### One-click PDF & print exports

Beyond EMAIL, the Data sheet exports its print area to a well-named PDF (or straight to the printer) on demand — the artifact for tickets, hand-overs, and end-of-well packages, generated from the same cells the morning report reads.

### Field-hardened by design

- **Protection-aware macros** — sheets stay protected against fat-finger edits; macros unprotect, work, and reprotect (`MDL_SheetProtect`).
- **Flicker-free rendering** — screen/event guards (`MDL_ScreenGuard`) keep gauge and corridor redraws from strobing a laptop on the rig floor.
- **Keyboard shortcuts** — high-frequency actions (new Slidesheet, PDF export) are bound to shortcuts for gloved-hands speed.

---

## What “standalone” really means

After a refresh, a directional can:

- Enter or clear surveys and pipe tally on the Slidesheet.
- Watch proximity and required TF update from plan + waypoints.
- Roll the day, edit costs on the calendar, sort PDM configs.
- Fire Mini/Full report text or EMAIL with HTML + corridor attachment.

…all **without** FieldCap open. When FieldCap changes (new BHA hours, new ticket day, crew swap), the extension re-exports and one REFRESH brings Excel current again.

That split — **browser for capture, workbook for decision** — is the product.

---

## Security & demo packaging

Public or shared copies of the workbook should carry **fictional job identity, land location, contacts, and crew** (see `SANITIZE_REVIEW.md`). Operational surveys and BHA serials may still appear in a field dump; scrub those separately before a public push. Real-field backups belong only under `backups/` and should stay out of git.

---

## Who it is for

| Role | What they get |
|------|----------------|
| Directional / MWD on location | Slidesheet guidance, pipe tally, day roll, motors |
| DD coordinator | EMAIL + corridor PNG, Mini/Full text, costs calendar |
| Office / sales | Consistent morning package without spreadsheet archaeology |
| Toolpush / company man | HTML body they can read in Outlook; PNG for the spatial story |

---

## Closing

PHX OpenCap does not replace FieldCap, and it does not pretend Excel is a web app. It does something more useful for the pad: **it makes FieldCap portable into the sheet that already runs the well**, and it makes that sheet speak plan, proximity, and morning report in one place.

Browser extension in. Standalone intelligence out. Same job folder. Same truth.

---

*Workbook: `Slide Sheet - Demo.xlsm` · CSV intake: `OpenCap\` · Companion extension: OpenCap Data Exporter (Chrome MV3)*

*Created by **Bruno Brottes**.*
