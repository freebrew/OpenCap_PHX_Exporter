---
name: ce-code-review
description: >-
  Review VBA modules, PowerShell scripts, and JavaScript in PHX_FieldCap for
  correctness, safety, and maintainability before committing or deploying.
  Use when the user says 'review this', 'check this', 'look for problems',
  or before any commit touching MDL_Setup.bas, Import-Slidesheet.ps1,
  Import-SetupModule.ps1, or background.js.
---

# ce-code-review — PHX_FieldCap

## Severity tiers

- **CRITICAL** — will cause data loss, workbook corruption, or silent wrong results
- **WARN** — likely to break under realistic conditions (empty range, missing sheet, re-run)
- **SUGGEST** — cleaner approach exists but current code works

Only report findings with confidence. Skip nitpicks.

## VBA checklist

- [ ] `Option Explicit` present in every module
- [ ] No unhandled `On Error Resume Next` without immediate `Err.Number` check
- [ ] No `.Select` / `.Activate` (direct object references only)
- [ ] Named ranges referenced by name, not hardcoded addresses
- [ ] All `Set obj = ...` followed by `Set obj = Nothing` before exit
- [ ] No hardcoded sheet names — use codenames or `ThisWorkbook.Sheets("name")`
- [ ] Array bounds checked before access (`If i <= UBound(arr) Then`)
- [ ] Date/number cells not treated as strings without `IsNumeric` / `IsDate` guard
- [ ] `Public` Subs that modify data have an error handler

## PowerShell checklist

- [ ] `param()` block with types — no positional `$args` access
- [ ] Paths built with `Join-Path`, not string concatenation
- [ ] File existence checked before `Import-Csv` / `Get-Content`
- [ ] No silently swallowed errors (`-ErrorAction SilentlyContinue` requires justification)
- [ ] Import is idempotent or clearly documented as non-idempotent

## JavaScript (`background.js`) checklist

- [ ] No `chrome.tabs` / `chrome.storage` calls without error handling
- [ ] No synchronous XHR
- [ ] Permissions in `manifest.json` match what the code actually uses
- [ ] No hardcoded URLs that belong in config

## Output format

```
CRITICAL [MDL_Setup.bas:342] On Error Resume Next spans 18 lines with no Err check — any failure is silently ignored.
WARN     [Import-Slidesheet.ps1:27] Path built by string concat; breaks if folder name contains spaces.
SUGGEST  [background.js:55] Promise chain can be replaced with async/await for readability.
```

One finding per line. File and line number required for CRITICAL and WARN.
