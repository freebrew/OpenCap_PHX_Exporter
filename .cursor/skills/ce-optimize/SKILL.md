---
name: ce-optimize
description: >-
  Run metric-driven optimization for PHX_FieldCap performance bottlenecks.
  Use when Excel recalculation is slow, VBA loops over large ranges take too
  long, PowerShell imports are sluggish, or any measurable throughput/latency
  issue exists. Trigger phrases: 'this is slow', 'optimize this', 'speed up
  the import', 'recalc takes forever', 'performance issue'.
---

# ce-optimize — PHX_FieldCap

## Process

1. **Measure first** — establish a baseline time before changing anything
2. **Identify the bottleneck** — profile before optimizing (most slowness is in 1–2 places)
3. **One change at a time** — measure after each change; discard if no improvement
4. **Set a done condition** — define "fast enough" before starting

## Common PHX_FieldCap bottlenecks

### VBA / Excel
| Bottleneck | Fix |
|---|---|
| Loop reads cells one-by-one | Read entire range into a Variant array once, loop the array |
| `ScreenUpdating` left on | `Application.ScreenUpdating = False` before loop, restore after |
| `Calculation` mode auto | `Application.Calculation = xlCalculationManual` during data write, then `xlCalculationAutomatic` |
| `.UsedRange` on a dirty sheet | Clean phantom cells; or use explicit last-row detection |
| VLOOKUP over 10k+ rows | Replace with `Application.Match` + `Index` or load to array |

```vba
' Fast array read/write pattern
Dim data As Variant
data = wsData.Range("A1:Z1000").Value   ' one read
' ... process data() array in memory ...
wsData.Range("A1:Z1000").Value = data   ' one write
```

### PowerShell imports
| Symptom | Fix |
|---|---|
| `Import-Csv` slow on large file | Stream with `[System.IO.File]::ReadLines()` |
| Excel COM automation per-row | Batch writes using array assignment |
| Repeated file open/close | Open workbook once, write all rows, save, close |

### Measurement helpers

**VBA timer:**
```vba
Dim t As Double
t = Timer
' ... code ...
Debug.Print "Elapsed: " & Format(Timer - t, "0.000") & "s"
```

**PowerShell timer:**
```powershell
$sw = [System.Diagnostics.Stopwatch]::StartNew()
# ... code ...
Write-Verbose "Elapsed: $($sw.Elapsed.TotalSeconds)s"
```

## Rule

Never optimize code that is not measurably slow. Premature optimization in VBA typically makes code unreadable without meaningful gains.
