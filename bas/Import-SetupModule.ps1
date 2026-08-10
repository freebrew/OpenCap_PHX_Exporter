# GetActiveObject only works in Windows PowerShell 5 (not PowerShell Core).
if ($PSVersionTable.PSEdition -eq "Core") {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"

# Prefer this repo's exported module; fall back to legacy FieldCap path.
$setupPath = Join-Path $PSScriptRoot "bas\MDL_Setup.bas"
if (-not (Test-Path -LiteralPath $setupPath)) {
    $setupPath = Join-Path $PSScriptRoot "MDL_Setup.bas"
}
if (-not (Test-Path -LiteralPath $setupPath)) {
    $setupPath = "Z:\PHX_FieldCap\src\excel\MDL_Setup.bas"
}
if (-not (Test-Path -LiteralPath $setupPath)) {
    throw "MDL_Setup.bas not found. Expected bas\MDL_Setup.bas next to this script."
}

Write-Host "Using module: $setupPath"

$xl = $null
try {
    $xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    throw "Excel is not running. Open 'Slide Sheet - Demo.xlsm' in Excel, then run this script again."
}

# Source of truth: Slide Sheet workbook (also accept OpenCap* for older workflows).
$wb = $null
foreach ($w in @($xl.Workbooks)) {
    $n = [string]$w.Name
    if ($n -like "*34801*" -or $n -like "Slide Sheet*" -or $n -like "OpenCap*") {
        $wb = $w
        break
    }
}
if (-not $wb) {
    $openNames = @($xl.Workbooks | ForEach-Object { $_.Name }) -join ", "
    if (-not $openNames) { $openNames = "(none)" }
    throw "Slide Sheet workbook not open in Excel. Open workbooks: $openNames"
}

Write-Host "Target workbook: $($wb.FullName)"

$vbp = $wb.VBProject

# â”€â”€ 1. Remove existing MDL_Setup* modules, then import fresh â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#    Prefer remove+import (more reliable than rename when VBA is busy / in break mode).
$existing = @($vbp.VBComponents | Where-Object { $_.Name -like "MDL_Setup*" -or $_.Name -eq "_MDL_Setup_OLD" })
foreach ($m in $existing) {
    Write-Host "Removing old module: $($m.Name)"
    try {
        $vbp.VBComponents.Remove($m)
    } catch {
        # Fallback: rename then remove (older approach)
        try {
            $m.Name = "_MDL_Setup_OLD"
            $vbp.VBComponents.Remove($vbp.VBComponents.Item("_MDL_Setup_OLD"))
        } catch {
            throw "Could not replace MDL_Setup. Close the VBA editor / exit break mode, then retry. Details: $($_.Exception.Message)"
        }
    }
}

Start-Sleep -Milliseconds 400

# â”€â”€ 2. Import fresh from .bas file â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#    VBComponents.Import reads the Attribute VB_Name header itself and names
#    the module correctly â€” no manual AddFromString / DeleteLines needed.
$imported = $vbp.VBComponents.Import($setupPath)
Write-Host "Imported as: $($imported.Name)  ($($imported.CodeModule.CountOfLines) lines)"

# If Import created MDL_Setup1 because name was still reserved, rename it.
if ($imported.Name -ne "MDL_Setup") {
    try {
        $imported.Name = "MDL_Setup"
        Write-Host "Renamed imported module to MDL_Setup"
    } catch {
        Write-Host "WARNING: imported as $($imported.Name). Rename it to MDL_Setup in the VBA editor."
    }
}

# â”€â”€ 4. Workbook_Open must NOT call InitSetup â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
# InitSetup / BuildSetupUI wipe Setup (manual formatting, AC table, notes).
# Refresh/rebuild stays manual via REFRESH CSVs / REBUILD buttons.
$tbMod = $vbp.VBComponents.Item("ThisWorkbook")
$cm = $tbMod.CodeModule
$all = ""
if ($cm.CountOfLines -gt 0) { $all = $cm.Lines(1, $cm.CountOfLines) }
$needsRewrite = ($all -match "InitSetup") -or ($all -notmatch "Workbook_Open")
if ($needsRewrite) {
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $wbCode  = "' Workbook_Open intentionally does not call InitSetup." + [char]13 + [char]10
    $wbCode += "' Use Setup buttons: REFRESH CSVs or REBUILD when a rebuild is wanted." + [char]13 + [char]10
    $wbCode += "Private Sub Workbook_Open()" + [char]13 + [char]10
    $wbCode += "    ' no auto InitSetup" + [char]13 + [char]10
    $wbCode += "End Sub"
    $cm.AddFromString($wbCode)
    Write-Host "ThisWorkbook updated: Workbook_Open no longer calls InitSetup."
} else {
    Write-Host "ThisWorkbook already has no-auto-InitSetup Workbook_Open; left unchanged."
}

# â”€â”€ 5. Save â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
$wb.Save()
Write-Host "Saved. Modules:"
$vbp.VBComponents | ForEach-Object {
    Write-Host "  $($_.Name)  [Type=$($_.Type)  Lines=$($_.CodeModule.CountOfLines)]"
}
