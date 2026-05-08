# DD Tools Excel Dashboard — VBA Setup

## Install

1. **Open Excel** and press `Alt+F11` to open the VBA editor.
2. In the Project pane (left), right-click **VBAProject (YourFile.xlsm)** → **Insert → Module**.
3. Copy all code from `src/excel/MDL_DDTools.bas` and paste into the new module.
4. Press `F5` or open the **Immediate Window** (`Ctrl+G`) and type:
   ```
   InitDDTools
   ```
   Press Enter. This creates the **DD Tools** sheet and three hidden data sheets.
5. Save the workbook as **`.xlsm`** (macro-enabled).

> **Seizure/distraction note:** All hidden sheets use `xlSheetVeryHidden`
> (invisible in the tab bar entirely) and every render operation runs with
> `Application.ScreenUpdating = False` and `Application.Calculation = xlCalculationManual`
> so there is zero visible flickering during data loads or BHA switches.

---

## Workflow

### How it finds the files
The **REFRESH** button scans the same folder as the workbook using `Dir()`.
It matches CSV filenames by `InStr` (case-insensitive):

| Contains in filename | → Used as |
|----------------------|-----------|
| `job-details`        | Job Details CSV |
| `bha-equipment`      | BHA Equipment CSV |
| `crew`               | Crew CSV |

The Chrome plugin names files exactly like `fieldcap-job-20786-bha-equipment.csv` — these match perfectly.
Just make sure the workbook (`.xlsm`) lives in the **same folder** as the downloaded CSVs.

### Typical session
1. Export CSVs from the FieldCap Chrome plugin → they land in your configured download folder
2. Move or save the workbook into that same folder
3. Click **REFRESH** — it finds all three files automatically, shows you their names, and asks to confirm
4. Dashboard rebuilds with one click

### New job / updated data
Click **REFRESH** again any time. The hidden data sheets are fully replaced on every refresh.
If a new BHA appears in the exported CSV, its button is created automatically.

### Rebuild without re-importing
Click **REBUILD** to re-render the dashboard from the already-imported
hidden data (useful if you resize columns or accidentally modify cells).

---

## Dashboard Layout

```
Row 1   FIELDCAP  DD TOOLS                    [IMPORT & BUILD] [REBUILD]
Row 2   JOB: Tourmaline ... | ID: 20786 | [FIELD ACTIVE]
Row 3   ████ accent divider
Row 4   METRES DRILLED | TOTAL HRS | SLIDE HRS | ROTATE HRS | SLIDE% | BHAs | CREW
Row 5   ──── border
Row 6   [BHA 1]  [BHA 2]  [BHA 3]  [BHA 4]  [BHA 5]  [BHA 6▶]  ← dynamic buttons
Row 7   ──── border
Row 9+  BHA N  |  SECTION  |  STATUS  |  MOTOR  |  GUIDANCE
        METRES / HRS SLID / HRS ROT / HRS CIRC / TOTAL HRS / BELOW ROT / ACTIVATED / COMPLETED
        #  SERIAL#  ITEM CODE  DESCRIPTION  LEN  ACCUM  OD  ID  TOP  TOTAL HRS
        ... component rows (alternating dark bands) ...
        ════════════════════
        CUMULATIVE COMPONENT HOURS  |  ALL BHAs
        ITEM CODE  DESCRIPTION  COUNT  METRES  HRS SLID  HRS ROT  TOTAL HRS
        ... rollup rows sorted by total hours descending ...
```

---

## Architecture

| Sheet        | Visibility       | Contents                         |
|--------------|------------------|----------------------------------|
| `DD Tools`   | Visible          | Dashboard, buttons, tables       |
| `_FC_Job`    | VeryHidden       | Raw job-details CSV data         |
| `_FC_Crew`   | VeryHidden       | Raw crew CSV data                |
| `_FC_BHA`    | VeryHidden       | Raw BHA equipment CSV data       |

`xlSheetVeryHidden` means the sheets cannot be un-hidden through the Excel UI
(right-click sheet tab → Show) — only via VBA. This keeps the workbook clean.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| "No data found. Use Import & Build first." | Run `InitDDTools` once, then Import & Build |
| BHA buttons have no action | Ensure the module is in a **Standard Module** (not Sheet/ThisWorkbook) |
| Colors look wrong | Requires Excel for Windows. Mac Excel has limited color/shape support |
| Buttons disappear after save | Save as `.xlsm`, not `.xlsx` |
