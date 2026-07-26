---
name: ce-plan
description: >-
  Create structured implementation plans for PHX_FieldCap features before
  touching code. Use when adding new VBA modules, redesigning the slide-sheet
  import pipeline, extending MDL_Setup.bas, adding PowerShell import scripts,
  modifying the chrome extension, or any multi-file change. Trigger phrases:
  'plan this', 'how should we build', 'what is the approach', 'break this down'.
---

# ce-plan — PHX_FieldCap

## When to use

- Before modifying `MDL_Setup.bas` (1,870 lines — plan first, then act)
- Before adding new import logic to `Import-Slidesheet.ps1` or `Import-SetupModule.ps1`
- Before designing a new Excel sheet, named range layout, or UserForm
- Before extending the Chrome extension (`background.js`)
- Any change touching 3+ files or 2+ modules

## Planning output format

Produce a structured plan with:

1. **Goal** — one sentence stating what done looks like
2. **Files to create / modify** — explicit list with reason
3. **Steps** — numbered, ordered by dependency
4. **Risk flags** — anything that could break existing workbook behavior
5. **Open questions** — decisions the user must make before work starts

## PHX_FieldCap-specific concerns to surface in every plan

- **Named ranges** — does the change affect any named ranges that VBA or formulas depend on?
- **Sheet codename vs. tab name** — always clarify which is being used
- **Import idempotency** — will re-running the PowerShell import produce duplicates?
- **Workbook compatibility** — `.xlsm` macro-enabled; avoid changes that break in Excel 2016+
- **Module scope** — `MDL_Setup` is large; isolate changes to avoid regression

## Step size rule

Each step must be small enough to verify independently. If a step cannot be tested alone, split it.

## Before starting implementation

Confirm the plan with the user before writing any code or editing any file.
