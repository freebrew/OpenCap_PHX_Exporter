# Headless run of the standalone AC corridor POC. Never opens Demo or Field.
param(
    [string]$PlanPdf = "",
    [string]$AcPdf = "",
    [int]$Capacity = 8
)
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
if ($PSVersionTable.PSEdition -eq "Core") {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path @args
    exit $LASTEXITCODE
}

$root = Split-Path -Parent $PSScriptRoot
$pocDir = Join-Path $root "docs\ac_poc"
$wbPath = Join-Path $pocDir "AC_Corridor_POC.xlsm"
if (-not (Test-Path -LiteralPath $wbPath)) { throw "missing $wbPath ; run bas/_build_ac_poc_workbook.ps1 first" }

if (-not $PlanPdf) { $PlanPdf = Join-Path $root "TEL HZ 102 CARROT 4-28-52-11 P3.pdf" }
if (-not $AcPdf) { $AcPdf = Join-Path $root "TEL HZ 102 CARROT 4-28-52-11 P3 AC.pdf" }
if (-not (Test-Path -LiteralPath $PlanPdf)) { throw "missing plan PDF: $PlanPdf" }
if (-not (Test-Path -LiteralPath $AcPdf)) { throw "missing AC PDF: $AcPdf" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $null
try {
    $wb = $xl.Workbooks.Open($wbPath, 0, $false)
    $poc = $wb.Worksheets.Item("POC")
    $poc.Range("B2").Value2 = $PlanPdf
    $poc.Range("B3").Value2 = $AcPdf
    $poc.Range("B4").Value2 = [double]$Capacity
    $poc.Range("B5").Value2 = ($pocDir.TrimEnd('\') + '\')
    $xl.Run("AcPoc_EnsureUi")
    $xl.Run("AcPoc_SetQuiet", $true)
    $xl.Run("AcPoc_RunAll")

    $log = $wb.Worksheets.Item("Log")
    $last = [int]$log.Cells($log.Rows.Count, 1).End(-4162).Row  # xlUp
    $fail = 0
    $pass = $false
    for ($r = 2; $r -le $last; $r++) {
        $st = [string]$log.Cells($r, 2).Value2
        $msg = [string]$log.Cells($r, 3).Value2
        "{0,-22} {1,-6} {2}" -f $log.Cells($r, 1).Value2, $st, $msg
        if ($st -eq "FAIL") { $fail++ }
        if ($st -eq "PASS" -and $msg -like "RUN COMPLETE*") { $pass = $true }
    }
    $wb.Save()
    if ($fail -gt 0 -or -not $pass) { throw "AC POC Log has $fail FAIL row(s) or no RUN COMPLETE" }
    "PASS run  fail=$fail"
} finally {
    if ($wb) { $wb.Close($true) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
