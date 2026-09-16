# Build the standalone AC corridor POC workbook. Never opens Demo or Field.
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
if ($PSVersionTable.PSEdition -eq "Core") {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path @args
    exit $LASTEXITCODE
}

$root = Split-Path -Parent $PSScriptRoot
$pocDir = Join-Path $root "docs\ac_poc"
$basDir = Join-Path $root "bas\poc"
$outFile = Join-Path $pocDir "AC_Corridor_POC.xlsm"
New-Item -ItemType Directory -Force -Path $pocDir | Out-Null

$modules = @(
    "MDL_AcPoc_Pdf.bas",
    "MDL_AcPoc_Main.bas",
    "MDL_AcPoc_Plan.bas",
    "MDL_AcPoc_AC.bas",
    "MDL_AcPoc_Qualify.bas",
    "MDL_AcPoc_Check.bas",
    "MDL_AcPoc_Render.bas"
)
foreach ($m in $modules) {
    $p = Join-Path $basDir $m
    if (-not (Test-Path -LiteralPath $p)) { throw "missing $p" }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
    $xl.AutomationSecurity = 1
$wb = $null
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
    $wb.Sheets.Item(1).Name = "POC"
    if (Test-Path -LiteralPath $outFile) { Remove-Item -LiteralPath $outFile -Force }
    $wb.SaveAs($outFile, 52)

    foreach ($m in $modules) {
        $p = Join-Path $basDir $m
        [void]$wb.VBProject.VBComponents.Import($p)
        "imported $m"
    }

    $tb = $null
    foreach ($c in @($wb.VBProject.VBComponents)) {
        if ($c.Name -eq "ThisWorkbook") { $tb = $c; break }
    }
    $open = @"
Private Sub Workbook_Open()
    On Error Resume Next
    AcPoc_EnsureUi
    On Error GoTo 0
End Sub
"@
    $n = [int]$tb.CodeModule.CountOfLines
    if ($n -gt 0) { $tb.CodeModule.DeleteLines(1, $n) }
    $tb.CodeModule.AddFromString($open)

    $xl.Run("AcPoc_EnsureUi")
    $wb.Save()
    "saved $outFile"
} finally {
    if ($wb) { $wb.Close($true) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
"PASS build"
