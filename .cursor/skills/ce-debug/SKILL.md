---
name: ce-debug
description: >-
  Systematically find root causes and fix bugs in PHX_FieldCap. Use when
  debugging VBA runtime errors (Type Mismatch, Subscript Out of Range, Object
  Required), broken Excel formulas, PowerShell import failures, Chrome extension
  issues, or any unexpected behavior. Trigger phrases: 'debug this', 'why is
  this failing', 'fix this bug', 'trace this error', or when a stack trace or
  error message is pasted.
---

# ce-debug — PHX_FieldCap

## Process

1. **Reproduce** — confirm the exact error message, line number, and steps to trigger
2. **Isolate** — identify the smallest unit that fails (single Sub, single formula, single PS command)
3. **Hypothesize** — form one root-cause theory before touching code
4. **Verify** — test the hypothesis; if wrong, repeat from step 3
5. **Fix** — address the root cause, not the symptom
6. **Confirm** — re-run the original failing scenario

## Common PHX_FieldCap failure modes

### VBA
| Error | Common cause |
|---|---|
| `Subscript out of range` | Sheet codename mismatch or array index off-by-one |
| `Type mismatch` | Cell value is text where number expected (check `IsNumeric`) |
| `Object required` | `Nothing` reference — check `Set` and `Is Nothing` |
| `Automation error` | Late-bound COM object (Excel, ADO) not properly released |
| `Permission denied` | File locked by another Excel instance |

### PowerShell (`Import-Slidesheet.ps1`, `Import-SetupModule.ps1`)
| Symptom | Common cause |
|---|---|
| Silent no-op | Wrong working directory; path resolution fails |
| Partial import | CSV encoding issue (UTF-8 BOM vs ANSI) |
| Duplicate rows | Import not idempotent; missing dedup key |

### Excel formulas
- Circular reference: check `Formulas → Error Checking → Circular References`
- `#REF!` after sheet rename: named ranges not updated
- `#VALUE!` in date math: cell formatted as text, not date serial

## Debugging tools in this project

- Add `Debug.Print` lines to VBA; view in Immediate Window (`Ctrl+G`)
- Use `On Error GoTo ErrHandler` with explicit `MsgBox Err.Description & " at line " & Erl`
- For PS: add `Write-Host` checkpoints and run with `-Verbose`

## Rule

Never change more than one thing between test runs. Each change must be justified by evidence from the previous run.
