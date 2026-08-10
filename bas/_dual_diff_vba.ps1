# Compare Demo vs Field VBA caches. Exit 0 on PASS, 1 on illegal drift.
# Ignores Attribute VB_* lines and trailing whitespace.
# Allowlists Module11 .To / .CC / .Subject assignment lines only.
param(
    [string]$DemoCache = "",
    [string]$FieldCache = "",
    [switch]$ExportFirst
)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
if ($ExportFirst) {
    & (Join-Path $PSScriptRoot "_dual_export_caches.ps1")
}
if (-not $DemoCache) { $DemoCache = Join-Path $root "bas\cache_demo" }
if (-not $FieldCache) { $FieldCache = Join-Path $root "bas\cache_field" }
if (-not (Test-Path -LiteralPath $DemoCache)) { throw "missing demo cache: $DemoCache - run bas/_dual_export_caches.ps1" }
if (-not (Test-Path -LiteralPath $FieldCache)) { throw "missing field cache: $FieldCache - run bas/_dual_export_caches.ps1" }

function Get-NormalizedLines([string]$Path) {
    $raw = Get-Content -LiteralPath $Path -Raw
    if ($null -eq $raw) { return @() }
    $lines = $raw -split "`r?`n"
    $kept = New-Object System.Collections.Generic.List[string]
    foreach ($ln in $lines) {
        if ($ln -match '^Attribute VB_') { continue }
        [void]$kept.Add($ln.TrimEnd())
    }
    while ($kept.Count -gt 0 -and [string]::IsNullOrWhiteSpace($kept[$kept.Count - 1])) {
        $kept.RemoveAt($kept.Count - 1)
    }
    return ,$kept.ToArray()
}

function Test-IsAllowlistedMailLine([string]$Line) {
    return [bool]($Line -match '^\s*\.(To|CC|Subject)\s*=')
}

function Test-IsCodeFile([string]$Name) {
    return $Name -match '\.(bas|cls|frm)$'
}

$demoFiles = @(Get-ChildItem -LiteralPath $DemoCache -File | Where-Object { $_.Name -ne "MANIFEST.txt" })
$fieldFiles = @(Get-ChildItem -LiteralPath $FieldCache -File | Where-Object { $_.Name -ne "MANIFEST.txt" })
$demoNames = @($demoFiles | ForEach-Object { $_.Name })
$fieldNames = @($fieldFiles | ForEach-Object { $_.Name })

$failures = New-Object System.Collections.Generic.List[string]
$notes = New-Object System.Collections.Generic.List[string]

foreach ($n in ($demoNames | Where-Object { $_ -notin $fieldNames })) {
    [void]$failures.Add("ONLY IN DEMO: $n")
}
foreach ($n in ($fieldNames | Where-Object { $_ -notin $demoNames })) {
    [void]$failures.Add("ONLY IN FIELD: $n")
}

$common = @($demoNames | Where-Object { $_ -in $fieldNames } | Sort-Object)
$identical = 0
$allowlisted = 0

foreach ($name in $common) {
    if ($name -match '\.frx$') {
        [void]$notes.Add("skip binary compare: $name (compare .frm code instead)")
        continue
    }
    if (-not (Test-IsCodeFile $name)) {
        [void]$notes.Add("skip non-code: $name")
        continue
    }

    $d = Get-NormalizedLines (Join-Path $DemoCache $name)
    $f = Get-NormalizedLines (Join-Path $FieldCache $name)
    $max = [Math]::Max($d.Count, $f.Count)
    $bad = New-Object System.Collections.Generic.List[string]
    $mailOnly = $true
    $anyDiff = $false

    for ($i = 0; $i -lt $max; $i++) {
        $dl = if ($i -lt $d.Count) { $d[$i] } else { "<EOF>" }
        $fl = if ($i -lt $f.Count) { $f[$i] } else { "<EOF>" }
        if ($dl -eq $fl) { continue }
        $anyDiff = $true
        $okMail = ($name -eq "Module11.bas") -and (Test-IsAllowlistedMailLine $dl) -and (Test-IsAllowlistedMailLine $fl)
        if (-not $okMail) {
            $mailOnly = $false
            [void]$bad.Add(("  L{0}: D|{1}" -f ($i + 1), $dl))
            [void]$bad.Add(("       F|{0}" -f $fl))
            if ($bad.Count -ge 12) { break }
        }
    }

    if (-not $anyDiff) {
        $identical++
    }
    elseif ($mailOnly -and $name -eq "Module11.bas") {
        $allowlisted++
        [void]$notes.Add("ALLOWLIST Module11.bas: .To/.CC/.Subject differ (expected until sheet-driven)")
    }
    else {
        [void]$failures.Add("DIFF $name")
        foreach ($b in $bad) { [void]$failures.Add($b) }
    }
}

""
"DUAL VBA DIFF"
"  demo cache : $DemoCache"
"  field cache: $FieldCache"
"  common     : $($common.Count)"
"  identical  : $identical"
"  allowlisted: $allowlisted"
"  failures   : $($failures.Count)"
""

foreach ($n in $notes) { "NOTE: $n" }
foreach ($f in $failures) { "FAIL: $f" }

if ($failures.Count -gt 0) {
    "RESULT: FAIL - Demo and Field VBA are out of sync"
    exit 1
}

"RESULT: PASS - VBA code matches (allowlisted Module11 mail headers only)"
exit 0
