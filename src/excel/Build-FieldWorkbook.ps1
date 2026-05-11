# ============================================================
#  Build-FieldWorkbook.ps1
#  Creates FieldCap_DailyPaperwork.xlsm from scratch and
#  imports MDL_Setup.bas + MDL_DDTools.bas into it.
#
#  Usage:
#    cd src\excel
#    .\Build-FieldWorkbook.ps1
#
#  Requirements: Excel must be installed. Runs COM automation.
# ============================================================

param(
    [string]$OutPath = "$PSScriptRoot\..\..\OpenCap_FieldWorkbook.xlsm"
)

$OutPath  = (Resolve-Path -LiteralPath (Split-Path $OutPath -Parent)).Path + "\" + (Split-Path $OutPath -Leaf)
$BasDir   = $PSScriptRoot
$BasFiles = @("MDL_Setup.bas", "MDL_DDTools.bas")

Write-Host "============================================================"
Write-Host " BUILD  OpenCap Field Workbook"
Write-Host " Output : $OutPath"
Write-Host "============================================================"

# ── Start Excel ──────────────────────────────────────────────────────────────────
$xl = New-Object -ComObject Excel.Application
$xl.Visible        = $false
$xl.DisplayAlerts  = $false
$xl.EnableEvents   = $false

try {
    # ── Create workbook ──────────────────────────────────────────────────────────
    $wb = $xl.Workbooks.Add()

    # ── Trust access to VBA project (required for programmatic module import) ──
    # The user must have "Trust access to the VBA project object model" enabled in
    # Excel > Trust Center > Macro Settings. Script will still function if the
    # setting is already enabled.
    try {
        $vbp = $wb.VBProject
    } catch {
        Write-Warning "Cannot access VBA project. Enable 'Trust access to VBA project' in Excel Trust Center."
        throw
    }

    # ── Import .bas modules ──────────────────────────────────────────────────────
    foreach ($bas in $BasFiles) {
        $fullPath = Join-Path $BasDir $bas
        if (Test-Path $fullPath) {
            Write-Host "  Importing $bas ..."
            $vbp.VBComponents.Import($fullPath) | Out-Null
        } else {
            Write-Warning "  Skipping $bas (not found at $fullPath)"
        }
    }

    # ── Remove default empty Sheet1/Sheet2/Sheet3 placeholders ──────────────────
    # Keep at least one sheet; the macro will create proper sheets.
    $wb.Worksheets | ForEach-Object {
        if ($_.Name -match "^Sheet\d+$") {
            try { $_.Delete() } catch {}
        }
    }

    # ── If no sheets remain, add a placeholder ───────────────────────────────────
    if ($wb.Worksheets.Count -eq 0) {
        $wb.Worksheets.Add() | Out-Null
        $wb.Worksheets(1).Name = "Temp"
    }

    # ── Save as macro-enabled workbook ──────────────────────────────────────────
    # xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($OutPath, 52)
    Write-Host ""
    Write-Host "  Saved: $OutPath"
    Write-Host ""
    Write-Host "============================================================"
    Write-Host " NEXT STEPS"
    Write-Host "  1. Open  $OutPath  in Excel"
    Write-Host "  2. Alt+F11 to confirm modules imported"
    Write-Host "  3. In the Immediate window (Ctrl+G) run:  InitSetup"
    Write-Host "  4. Copy CSV exports into the same folder as the workbook"
    Write-Host "  5. Click  REFRESH CSVs  on the Setup sheet"
    Write-Host "============================================================"

} finally {
    try { $wb.Close($false) } catch {}
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
