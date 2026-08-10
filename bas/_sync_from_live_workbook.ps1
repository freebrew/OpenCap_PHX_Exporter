# Pull VBA (+ workbook snapshot) FROM the live Slide Sheet into bas/.
# Use after manual Excel edits so staged modules match the source of truth.
#
# Does NOT import anything into the workbook. Events off; no InitSetup.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root "Slide Sheet - Demo.xlsm"
$bas = Join-Path $root "bas"
$liveDir = Join-Path $bas "live"
$backups = Join-Path $root "backups"
New-Item -ItemType Directory -Force -Path $liveDir, $backups | Out-Null

if (-not (Test-Path -LiteralPath $src)) { throw "missing live workbook: $src" }

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$cacheXlsm = Join-Path $backups ("Slide Sheet - Demo.xlsm.cache_manual_" + $stamp)
Copy-Item -LiteralPath $src -Destination $cacheXlsm -Force
$latest = Join-Path $backups "Slide Sheet - Demo.xlsm.cache_manual_LATEST.xlsm"
Copy-Item -LiteralPath $src -Destination $latest -Force
"workbook cache : $cacheXlsm"
"latest pointer : $latest"

$tmp = Join-Path $env:TEMP ("sync_live_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".xlsm")
Copy-Item -LiteralPath $src -Destination $tmp -Force

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $null
try {
    $wb = $xl.Workbooks.Open($tmp, 0, $true)
    if (-not $wb) { throw "failed to open temp copy" }
    $vbp = $wb.VBProject

    $vbaCache = Join-Path $backups ("vba_cache_" + $stamp)
    New-Item -ItemType Directory -Force -Path $vbaCache | Out-Null

    $exportedBas = 0
    $manifest = New-Object System.Collections.Generic.List[string]
    $manifest.Add("# VBA synced FROM live workbook $stamp")
    $manifest.Add("# Source: $src")
    $manifest.Add("# Rule: live .xlsm is source of truth; re-sync before any import.")
    $manifest.Add("")

    foreach ($c in @($vbp.VBComponents)) {
        $lines = $c.CodeModule.CountOfLines
        $typeName = switch ($c.Type) { 1 { "bas" } 2 { "cls" } 3 { "frm" } 100 { "document" } default { "other$($c.Type)" } }
        $manifest.Add(("{0,-28} type={1,-8} lines={2}" -f $c.Name, $typeName, $lines))
        "  {0,-28} {1} {2} lines" -f $c.Name, $typeName, $lines

        if ($lines -eq 0 -and $c.Type -ne 3) { continue }

        $ext = switch ($c.Type) { 1 { ".bas" } 2 { ".cls" } 3 { ".frm" } 100 { ".cls" } default { ".txt" } }
        $cachePath = Join-Path $vbaCache ($c.Name + $ext)
        $c.Export($cachePath)

        if ($c.Type -eq 1) {
            $dest = Join-Path $bas ($c.Name + ".bas")
            Copy-Item -LiteralPath $cachePath -Destination $dest -Force
            Copy-Item -LiteralPath $cachePath -Destination (Join-Path $liveDir ($c.Name + ".bas")) -Force
            $exportedBas++
        }
    }

    $manifest | Set-Content -LiteralPath (Join-Path $vbaCache "MANIFEST.txt") -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $backups "CACHE_MANUAL_LATEST.txt") -Value @"
stamp=$stamp
workbook=$cacheXlsm
latest=$latest
vba=$vbaCache
saved=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
source=$src
"@ -Encoding UTF8

    ""
    "Exported $exportedBas .bas modules -> bas\ and bas\live\"
    "Full VBA cache -> $vbaCache"
    "PASS: staged VBA now matches live workbook."
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
