---
name: ce-work
description: >-
  Execute PHX_FieldCap development work efficiently with quality checkpoints.
  Use when implementing planned features in VBA, PowerShell, or JavaScript,
  or when the user says 'do it', 'implement this', 'build it', 'make it work'.
  Follows the plan produced by ce-plan; do not skip planning for large changes.
---

# ce-work — PHX_FieldCap

## Before writing any code

- [ ] Read the target file(s) — never edit blind
- [ ] Confirm the plan (ce-plan output or user description)
- [ ] Identify which modules/sheets are affected

## Execution order

1. **Smallest change first** — implement the core logic before edge cases
2. **One file at a time** — complete and verify each file before moving on
3. **Preserve existing behavior** — do not rename Subs, move named ranges, or change sheet structure unless that is the explicit goal

## PHX_FieldCap coding conventions

### VBA
- Declare all variables (`Option Explicit` is assumed)
- Use worksheet codenames (`wsData`, `wsSetup`) not tab names
- Name ranges via `ThisWorkbook.Names` — never hardcode `"A1:Z100"`
- Release COM objects: `Set obj = Nothing` at end of each Sub
- Error handler in every Public Sub:
```vba
Private Sub MySub()
    On Error GoTo ErrHandler
    ' ... work ...
    Exit Sub
ErrHandler:
    MsgBox "MySub: " & Err.Description, vbCritical
End Sub
```

### PowerShell
- `param()` block at top with types and defaults
- `Set-StrictMode -Version Latest` at script start
- Use `Write-Verbose` for progress, not `Write-Host`
- Follow the wizard-style progress format from the project SOP

### JavaScript (`background.js` / Chrome extension)
- `'use strict'` at top
- Async/await over raw Promise chains
- No inline event listeners — register via `addEventListener`

## Quality checkpoints

After each logical unit of work:
- [ ] Syntax check (VBA: `Debug → Compile`; PS: `$null = [System.Management.Automation.Language.Parser]::ParseFile(...)`)
- [ ] Run the affected workflow end-to-end at least once
- [ ] No new linter errors introduced
- [ ] If the user brings a live Field `.xlsm` as source of truth: run skill `adopt-field-workbook` first (Field VBA → Demo; Field cells untouched)
- [ ] If Slide Sheet VBA changed after that: run skill `dual-workbook-sync` (Demo → Field mirror + `_dual_diff_vba.ps1` PASS)

## When blocked

Stop, describe what is blocking, and ask. Do not guess at hidden behavior in a 1,870-line module.
