# 3D Wells View — 2026-11-13

## Goal

A 3D view in the PHX Slide Sheet web sheet that draws the whole planned well, the path drilled so far, and the offset wells that the anti-collision report marks as a concern. Hovering a path shows that report's separation factor and centre-to-centre distance. The picture is drawn on the machine's GPU, in a worker, and is already rendered before the view is opened.

## Approved choices

| Decision | Choice |
|---|---|
| Surface | The web sheet only: the Chrome extension in `src/extension`. Excel keeps its flat pictures. No VBA, phone, or StrapSheets code. Chosen 2026-11-13. |
| Which offset wells | The wells in the Data tab "AC Info & Concerns" table, from the same rule that fills that table. A "Show All Offset Wells" switch draws the rest, thinner and dimmer. |
| When it renders | Whenever the plan, the report, or a survey changes, the scene is rebuilt and one frame is drawn while the view is hidden, so opening it only reveals a finished picture. |
| Engine | three.js 0.187.0. WebGPU when the browser has it, WebGL2 otherwise. Wide lines where the renderer supports them; thin tubes where it does not. |
| Where the tab lives | A "Sheets \| 3D Wells" switch in the app bar. It is not a workbook sheet, so an Excel handoff cannot write an empty tab into the `.xlsm`. |

## What the view shows

- The planned path, from `_OC_Survey`.
- The drilled path, from the good surveys on Slidesheet, walked the same way as the corridor picture. A "Drilled Path" switch hides it. It is on by default.
- Each concern well, coloured by the report band already used on the Data tab: red below SF 1.5, yellow below SF 2.0, green otherwise.
- A marker at that well's closest Summary point, joined to the plan at the same measured depth.
- Hover on an offset well: the well name, then the nearest report station's printed reference MD, centre-to-centre distance and separation factor (the printed numbers, never a value invented between stations), then "Measured Here", the live 3D distance from that point to the plan.
- Hover on the plan: measured depth, TVD, and the tightest printed separation factor at the nearest report depth.
- "Show All Offset Wells" adds every other offset well that has a path in the report.
- A concern well with no detail page is listed as "No Path in Report" and is not drawn.
- Orbit, pan, zoom, and "Reset View".
- A check badge: "Checked: N stations, worst 0.xx m", or "Report and plan do not line up" when the frame check fails.

Axes are 1:1. East, North and up (up is the opposite of TVD). Positions are metres from the first plan station, so the GPU never sees map-scale coordinates.

## Data

Import AC already stores one Summary row per offset well on `_OC_AC`. The detail pages of the same PDF carry the offset path. `parseAcDetail` reads those pages (the TypeScript port of `AcPoc_ParseDetail`) and Import AC writes the lean columns to a hidden `_OC_AC_Detail` sheet: well, reference MD, offset MD, offset TVD, North, East, centre-to-centre, separation factor, warning.

Three checks run at import and again when the scene is built, the same checks as the Excel proof of concept:

1. Every concern well has a detail page.
2. Where a detail station sits on a Summary row, its separation factor agrees within 0.02 and its centre-to-centre distance within 0.05 m.
3. For each drawn station, the 3D distance from the plan (linear in measured depth) to the offset point is within 0.5 m of the printed centre-to-centre distance.

The badge text is written on `_OC_AC` and recomputed when the scene is built.

`concernWells` is the one function that chooses the Data table rows. The 3D view calls it too.

## Rendering

`src/packages/wells-3d` is pure TypeScript: the scene arrays, the pick math, and the orbit. It does not import three.js. Node tests cover it.

The extension owns the GPU code. A module worker receives the canvas with `transferControlToOffscreen`. Buffers are uploaded once per scene. A frame is drawn only when the camera, the hover, or the scene changes. `powerPreference` is `high-performance`. The pixel ratio is capped at 2. If the GPU context is lost, the worker rebuilds the meshes from the arrays it kept.

Hover picking is a straight pass over the stations and segments in the worker. At a few thousand stations that is well under a millisecond, so there is no spatial index.

The page thread still receives the pointer, so a long spreadsheet recalculation can delay the start of a drag by the length of that task. The frames themselves are drawn on the worker.

Budgets, measured and reported as measured: no frame over the display interval while orbiting during a grid recalculation; the first frame of the view within 50 ms of the switch; the hover readout within two frames.

## Verification

- Vitest, on invented wells and measures only: the detail parser, the three checks, the shared rule, the scene, and the pick.
- Playwright, in the extension's preview page: a worker-owned canvas, WebGL2 when WebGPU is absent, a wide line (or a tube, if wide lines do not paint), and a hover readout.
- One local run of the parser and the three checks against this job's plan PDF and anti-collision PDF. Nothing from that run is committed or quoted.
- Preview version raised to 0.1.56 in `version.ts`, `package.json`, and the manifest, then `pnpm build:ext`.

## Left out

Recomputing a separation factor, drawing error ellipses, remembering the camera, converting the daily corridor picture to three.js, and any Excel or phone version of this view.
