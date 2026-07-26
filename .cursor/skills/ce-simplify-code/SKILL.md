---
name: ce-simplify-code
description: >-
  Simplify and refine VBA modules, PowerShell scripts, and JavaScript in
  PHX_FieldCap after features are working. Use after implementing new logic,
  after debugging sessions leave temporary scaffolding behind, or when a module
  has grown unwieldy. Trigger phrases: 'simplify this', 'clean this up',
  'refactor', 'this is getting messy', 'reduce duplication'.
---

# ce-simplify-code — PHX_FieldCap

## Guiding principle

Remove everything that does not earn its place. Simpler code is easier to debug in Excel's VBA IDE, which has no IntelliSense for custom types and no refactoring tools.

## What to target

### VBA (`MDL_Setup.bas` and others)
- **Dead Subs/Functions** — procedures never called from anywhere
- **Magic numbers** — replace `65536` with `Rows.Count`, `"Sheet3"` with the sheet codename
- **Repeated range lookups** — cache `ws.Range("A1").Value` in a variable if used 3+ times
- **Nested `If` pyramids** — flatten with early `Exit Sub` / guard clauses
- **Copy-pasted blocks** — extract to a shared helper Sub/Function
- **`Select` / `Activate` chains** — remove; reference objects directly

```vba
' Before (slop)
Sheets("Data").Select
Range("A1").Select
Selection.Value = "Hello"

' After
wsData.Range("A1").Value = "Hello"
```

### PowerShell (`Import-Slidesheet.ps1`, `Import-SetupModule.ps1`)
- Replace positional parameter access with named `param()` blocks
- Remove `Write-Host` debug lines left from development
- Collapse repeated `if ($null -eq ...)` checks into helper functions

### JavaScript (`background.js`)
- Remove `console.log` debug statements
- Consolidate repeated `chrome.storage.local.get` calls

## What NOT to simplify

- Error handling (`On Error GoTo`, `try/catch`) — never remove
- Logging to `Immediate Window` or PS `-Verbose` output — keep in dev builds
- Deliberate workarounds with comments explaining why — preserve the comment

## Process

1. Read the target file fully before editing
2. List every simplification candidate
3. Apply one category at a time (dead code → duplication → naming)
4. Verify behavior is unchanged after each batch
5. Never simplify and add features in the same pass
