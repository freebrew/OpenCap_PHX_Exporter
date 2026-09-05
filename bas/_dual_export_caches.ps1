# Export VBA from Demo + Field into bas/cache_demo and bas/cache_field.
# Events off; no InitSetup. Does not modify either workbook.
#
# Each workbook is exported in a fresh powershell.exe process because a second
# Excel.Application in the same process often returns VBProject.Count = 0.
param(
    [string]$DemoPath = "",
    [string]$FieldPath = ""
)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
if (-not $DemoPath) { $DemoPath = Join-Path $root "Slide Sheet - Demo.xlsm" }
if (-not $FieldPath) { $FieldPath = Join-Path $root "Slide Sheet - 35780.xlsm" }
$backups = Join-Path $root "backups"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"

if (-not (Test-Path -LiteralPath $DemoPath)) { throw "missing Demo: $DemoPath" }
if (-not (Test-Path -LiteralPath $FieldPath)) { throw "missing Field: $FieldPath" }

$worker = @'
param(
    [Parameter(Mandatory = $true)][string]$XlsmPath,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [Parameter(Mandatory = $true)][string]$Label,
    [Parameter(Mandatory = $true)][string]$Stamp
)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
    try { $_.Kill() } catch {}
}
Start-Sleep -Seconds 3

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
Get-ChildItem -LiteralPath $OutDir -Force -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $null
try {
    $miss = [Type]::Missing
    # Arg 7 = IgnoreReadOnlyRecommended; without it Open can return a stub
    # workbook whose VBProject.Count is 0.
    $wb = $xl.Workbooks.Open($XlsmPath, 0, $true, $miss, $miss, $miss, $true)
    if (-not $wb -and $xl.Workbooks.Count -ge 1) { $wb = $xl.Workbooks.Item(1) }
    Start-Sleep -Seconds 1
    $vbp = $wb.VBProject
    if ($null -eq $vbp -or [int]$vbp.VBComponents.Count -lt 1) {
        throw "$Label VBProject empty or inaccessible - enable Trust access to the VBA project object model."
    }
    $manifest = New-Object System.Collections.Generic.List[string]
    [void]$manifest.Add("# $Label VBA cache $Stamp")
    [void]$manifest.Add("# Source: $XlsmPath")
    [void]$manifest.Add("# Project: $($vbp.Name)")
    [void]$manifest.Add("")
    $count = 0
    foreach ($c in @($vbp.VBComponents)) {
        $lines = $c.CodeModule.CountOfLines
        $typeName = switch ($c.Type) { 1 { "bas" } 2 { "cls" } 3 { "frm" } 100 { "document" } default { "other$($c.Type)" } }
        [void]$manifest.Add(("{0,-28} type={1,-8} lines={2}" -f $c.Name, $typeName, $lines))
        if ($lines -eq 0 -and $c.Type -ne 3) { continue }
        $ext = switch ($c.Type) { 1 { ".bas" } 2 { ".cls" } 3 { ".frm" } 100 { ".cls" } default { ".txt" } }
        $c.Export((Join-Path $OutDir ($c.Name + $ext)))
        $count++
    }
    $manifest | Set-Content -LiteralPath (Join-Path $OutDir "MANIFEST.txt") -Encoding UTF8
    "{0}: project={1} comps={2} exported={3} -> {4}" -f $Label, $vbp.Name, $vbp.VBComponents.Count, $count, $OutDir
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_.Kill() } catch {}
    }
}
'@

$workerPath = Join-Path $env:TEMP ("dual_export_worker_" + $stamp + ".ps1")
Set-Content -LiteralPath $workerPath -Value $worker -Encoding UTF8

function Invoke-ExportWorker {
    param([string]$XlsmPath, [string]$OutDir, [string]$Label)
    # Single argument string so paths with spaces (Slide Sheet - *.xlsm) stay intact.
    $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$workerPath`" -XlsmPath `"$XlsmPath`" -OutDir `"$OutDir`" -Label `"$Label`" -Stamp `"$stamp`""
    $p = Start-Process -FilePath "powershell.exe" -ArgumentList $arg -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) { throw "export worker failed for $Label (exit $($p.ExitCode))" }
}

$demoStable = Join-Path $root "bas\cache_demo"
$fieldStable = Join-Path $root "bas\cache_field"
$demoStamp = Join-Path $root ("bas\cache_demo_" + $stamp)
$fieldStamp = Join-Path $root ("bas\cache_field_" + $stamp)

try {
    Invoke-ExportWorker -XlsmPath $DemoPath -OutDir $demoStable -Label "demo"
    Invoke-ExportWorker -XlsmPath $FieldPath -OutDir $fieldStable -Label "field"
}
finally {
    Remove-Item -LiteralPath $workerPath -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path (Join-Path $demoStable "MANIFEST.txt"))) { throw "demo cache missing MANIFEST" }
if (-not (Test-Path (Join-Path $fieldStable "MANIFEST.txt"))) { throw "field cache missing MANIFEST" }

Copy-Item -LiteralPath $demoStable -Destination $demoStamp -Recurse -Force
Copy-Item -LiteralPath $fieldStable -Destination $fieldStamp -Recurse -Force

New-Item -ItemType Directory -Force -Path $backups | Out-Null
@(
    "stamp=$stamp"
    "demo=$DemoPath"
    "field=$FieldPath"
    "demo_cache=$demoStable"
    "field_cache=$fieldStable"
    "demo_stamp=$demoStamp"
    "field_stamp=$fieldStamp"
    "saved=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
) | Set-Content -LiteralPath (Join-Path $backups "DUAL_CACHE_LATEST.txt") -Encoding UTF8

"PASS: dual caches updated ($stamp)"
