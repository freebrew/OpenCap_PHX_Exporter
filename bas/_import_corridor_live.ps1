# Back up the live workbook, then import MDL_CorridorImage (new) and the patched
# Module11 (inline corridor picture under the B2:F55 table).
#
# Events stay off for the whole session so Workbook_Open / InitSetup never runs
# and field data is left alone. No cell is written by this script.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root "Slide Sheet - Demo.xlsm"
$backupDir = Join-Path $root "backups"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backup = Join-Path $backupDir ("Slide Sheet - Demo.xlsm.bak_corridor_" + $stamp)

if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
Copy-Item -LiteralPath $src -Destination $backup -Force
"backup   : $backup"

$modules = @(
    @{ Name = "MDL_PlanGauge"; Path = (Join-Path $PSScriptRoot "MDL_PlanGauge.bas") },
    @{ Name = "MDL_CorridorImage"; Path = (Join-Path $PSScriptRoot "MDL_CorridorImage.bas") },
    @{ Name = "Module11"; Path = (Join-Path $PSScriptRoot "Module11.bas") }
)
foreach ($m in $modules) {
    if (-not (Test-Path $m.Path)) { throw ("missing staged module: " + $m.Path) }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $null
try {
    # Arg 7 is IgnoreReadOnlyRecommended: without it Excel honours the workbook's
    # read-only recommendation silently (DisplayAlerts is off) and Save does nothing.
    $miss = [Type]::Missing
    $wb = $xl.Workbooks.Open($src, 0, $false, $miss, $miss, $miss, $true)
    if (-not $wb) { $wb = $xl.Workbooks.Item(1) }
    "read-only={0} ro-recommended={1} write-reserved={2}" -f $wb.ReadOnly, $wb.ReadOnlyRecommended, $wb.WriteReserved

    # If the workbook is already open elsewhere it opens read-only here and Save
    # is a silent no-op, which looks like a successful import but changes nothing.
    if ($wb.ReadOnly) { throw "workbook opened read-only - close it in Excel and re-run" }

    $proj = $wb.VBProject

    foreach ($m in $modules) {
        $existing = $null
        try { $existing = $proj.VBComponents.Item($m.Name) } catch { $existing = $null }
        if ($existing) {
            $proj.VBComponents.Remove($existing)
            "removed  : {0}" -f $m.Name
        }
        [void]$proj.VBComponents.Import($m.Path)
        $lines = $proj.VBComponents.Item($m.Name).CodeModule.CountOfLines
        "imported : {0} ({1} lines)" -f $m.Name, $lines
    }

    $wb.Save()
    "saved    : $src"

    foreach ($m in $modules) {
        "verify   : {0} = {1} lines" -f $m.Name, $proj.VBComponents.Item($m.Name).CodeModule.CountOfLines
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
