# Mirror VBA modules from Demo -> Field (in-place CodeModule replace).
# Backs up both workbooks first. Verifies dual diff at the end.
#
# CRITICAL: Field user-entered data must NOT be altered. This script only
# replaces VBComponent.CodeModule text. It must never write cells, sheets,
# named-range values, or run InitSetup / RebuildSetup / OpenCap reload.
#
# Default: mirror ALL code modules from Demo.
# Optional: -Modules MDL_CorridorImage,Module11
# Module11: preserves Field .To/.CC/.Subject unless -NoPreserveFieldMailHeaders.
param(
    [string[]]$Modules = @(),
    [switch]$NoPreserveFieldMailHeaders,
    [switch]$SkipVerify
)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
$demoPath = Join-Path $root "Slide Sheet - Demo.xlsm"
$fieldPath = Join-Path $root "Slide Sheet - 34801.xlsm"
$backups = Join-Path $root "backups"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path -LiteralPath $demoPath)) { throw "missing Demo: $demoPath" }
if (-not (Test-Path -LiteralPath $fieldPath)) { throw "missing Field: $fieldPath" }
New-Item -ItemType Directory -Force -Path $backups | Out-Null

$bakDemo = Join-Path $backups ("Slide Sheet - Demo.xlsm.bak_dual_mirror_" + $stamp)
$bakField = Join-Path $backups ("Slide Sheet - 34801.xlsm.bak_dual_mirror_" + $stamp)
Copy-Item -LiteralPath $demoPath -Destination $bakDemo -Force
Copy-Item -LiteralPath $fieldPath -Destination $bakField -Force
"backup demo : $bakDemo"
"backup field: $bakField"

function Get-ExportBody([string]$ExportPath) {
    $lines = @(Get-Content -LiteralPath $ExportPath)
    $start = 0
    while ($start -lt $lines.Count -and ($lines[$start] -match '^Attribute VB_' -or $lines[$start].Trim() -eq "")) {
        $start++
    }
    if ($start -ge $lines.Count) { return "" }
    return (($lines[$start..($lines.Count - 1)]) -join "`r`n")
}

function Get-FieldMailHeaders([object]$CodeModule) {
    $headers = @{ To = $null; CC = $null; Subject = $null }
    $n = [int]$CodeModule.CountOfLines
    if ($n -le 0) { return $headers }
    $text = $CodeModule.Lines(1, $n) -split "`r?`n"
    foreach ($ln in $text) {
        if ($ln -match '^\s*\.To\s*=') { $headers.To = $ln }
        elseif ($ln -match '^\s*\.CC\s*=') { $headers.CC = $ln }
        elseif ($ln -match '^\s*\.Subject\s*=') { $headers.Subject = $ln }
    }
    return $headers
}

function Set-FieldMailHeaders([string]$Body, [hashtable]$Headers) {
    if (-not $Headers) { return $Body }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($ln in ($Body -split "`r?`n")) {
        if ($Headers.To -and $ln -match '^\s*\.To\s*=') { [void]$lines.Add($Headers.To); continue }
        if ($Headers.CC -and $ln -match '^\s*\.CC\s*=') { [void]$lines.Add($Headers.CC); continue }
        if ($Headers.Subject -and $ln -match '^\s*\.Subject\s*=') { [void]$lines.Add($Headers.Subject); continue }
        [void]$lines.Add($ln)
    }
    return ($lines -join "`r`n")
}

$tmpDemo = Join-Path $env:TEMP ("dual_mirror_demo_" + $stamp + ".xlsm")
Copy-Item -LiteralPath $demoPath -Destination $tmpDemo -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wbDemo = $null
$wbField = $null
$exportDir = Join-Path $env:TEMP ("dual_mirror_export_" + $stamp)
New-Item -ItemType Directory -Force -Path $exportDir | Out-Null

function Open-WithRetry([object]$Excel, [string]$Path, [bool]$ReadOnly) {
    # Excel rejects COM calls (RPC_E_CALL_REJECTED) while it is still busy
    # loading/recalculating the previous workbook; wait and retry.
    for ($try = 1; $try -le 12; $try++) {
        try { return $Excel.Workbooks.Open($Path, 0, $ReadOnly) }
        catch {
            if ($_.Exception.HResult -ne 0x80010001 -or $try -eq 12) { throw }
            "open busy (attempt $try) - waiting 10s: $Path"
            Start-Sleep -Seconds 10
        }
    }
}

try {
    $wbDemo = Open-WithRetry $xl $tmpDemo $true
    $wbField = Open-WithRetry $xl $fieldPath $false
    $demoProj = $wbDemo.VBProject
    $fieldProj = $wbField.VBProject

    $beforeCount = [int]$fieldProj.VBComponents.Count
    "field before: project=$($fieldProj.Name) comps=$beforeCount"
    if ($beforeCount -lt 50) { throw "Field VBA project looks empty/corrupt before mirror - aborting" }

    $demoComps = @()
    foreach ($c in @($demoProj.VBComponents)) {
        if ($c.Type -eq 3) {
            # forms: export .frm (+ .frx) and replace code module text only
            $demoComps += $c
        }
        elseif ($c.CodeModule.CountOfLines -gt 0) {
            $demoComps += $c
        }
    }

    if ($Modules.Count -gt 0) {
        # powershell.exe -File passes "A,B" as one token; split it back apart.
        $want = @($Modules | ForEach-Object { $_ -split ',' } |
                  ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $demoComps = @($demoComps | Where-Object { $_.Name -in $want })
        $missing = @($want | Where-Object { $_ -notin @($demoComps | ForEach-Object { $_.Name }) })
        if ($missing.Count) { throw "modules not found in Demo: $($missing -join ', ')" }
    }

    foreach ($src in $demoComps) {
        $name = $src.Name
        $dst = $null
        foreach ($c in @($fieldProj.VBComponents)) {
            if ($c.Name -eq $name) { $dst = $c; break }
        }
        if (-not $dst) { throw "Field missing component: $name - inventoring modules is out of scope; add it manually first" }

        $ext = switch ($src.Type) { 1 { ".bas" } 2 { ".cls" } 3 { ".frm" } 100 { ".cls" } default { ".txt" } }
        $exportPath = Join-Path $exportDir ($name + $ext)
        $src.Export($exportPath)
        $body = Get-ExportBody $exportPath

        $mail = $null
        if ($name -eq "Module11" -and -not $NoPreserveFieldMailHeaders) {
            $mail = Get-FieldMailHeaders $dst.CodeModule
            $body = Set-FieldMailHeaders $body $mail
            "preserve Field mail headers in Module11"
        }

        $cm = $dst.CodeModule
        $n = [int]$cm.CountOfLines
        if ($n -gt 0) { $cm.DeleteLines(1, $n) }
        if ($body.Length -gt 0) { $cm.AddFromString($body) }
        "mirrored $name -> field lines=$($cm.CountOfLines)"
    }

    $afterCount = [int]$fieldProj.VBComponents.Count
    "field after: project=$($fieldProj.Name) comps=$afterCount"
    if ($afterCount -ne $beforeCount) {
        throw "Field component count changed ($beforeCount -> $afterCount) - restore $bakField"
    }

    $wbField.Save()
    "saved Field"
    $wbField.Close($false)
    $wbField = $null
    $wbDemo.Close($false)
    $wbDemo = $null
}
finally {
    if ($wbField) { try { $wbField.Close($false) } catch {} }
    if ($wbDemo) { try { $wbDemo.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item -LiteralPath $tmpDemo -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $exportDir -Recurse -Force -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Reopen verify
$xl2 = New-Object -ComObject Excel.Application
$xl2.Visible = $false
$xl2.DisplayAlerts = $false
$xl2.EnableEvents = $false
$xl2.AutomationSecurity = 1
$wb2 = $null
try {
    $wb2 = $xl2.Workbooks.Open($fieldPath, 0, $true)
    $c = [int]$wb2.VBProject.VBComponents.Count
    "reopen Field: project=$($wb2.VBProject.Name) comps=$c"
    if ($c -lt 50) { throw "Field VBA collapsed after save - restore $bakField immediately" }
}
finally {
    if ($wb2) { $wb2.Close($false) }
    $xl2.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2)
}

if (-not $SkipVerify) {
    & (Join-Path $PSScriptRoot "_dual_export_caches.ps1")
    & (Join-Path $PSScriptRoot "_dual_diff_vba.ps1")
    if ($LASTEXITCODE -ne 0) { throw "dual diff FAILED after mirror" }
}

"PASS: Demo -> Field mirror complete"
