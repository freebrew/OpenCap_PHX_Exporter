<#
.SYNOPSIS
    Copies the Slides tab from "Auto Slide Sheet -34538.xlsm" into
    OpenCap_FieldWorkbook.xlsm.

    Works whether or not the destination workbook is already open in Excel.
    If Excel is already running with the file open, it attaches to that instance
    so the save goes to the live workbook the user can see.
#>
$ErrorActionPreference = "Stop"

$root       = "c:\Users\User\Desktop\Vibe_Projects\PHX_FieldCap"
$sourcePath = Join-Path $root "Auto Slide Sheet -34538.xlsm"
$destPath   = Join-Path $root "OpenCap_FieldWorkbook.xlsm"
$destLeaf   = Split-Path $destPath -Leaf

Write-Host "=== Import Slidesheet Wizard ==="
Write-Host "Source : $(Split-Path $sourcePath -Leaf)"
Write-Host "Dest   : $destLeaf"

foreach ($f in @($sourcePath, $destPath)) {
    if (-not (Test-Path $f)) { Write-Error "File not found: $f"; exit 1 }
}

# -----------------------------------------------------------------------
# Attach to a running Excel instance if one exists; otherwise create one.
# This avoids the "file already locked" problem when the workbook is open.
# -----------------------------------------------------------------------
$excel     = $null
$ownExcel  = $false   # did WE create the Excel process?

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    Write-Host "[1/4] Attached to running Excel instance."
} catch {
    $excel    = New-Object -ComObject Excel.Application
    $ownExcel = $true
    Write-Host "[1/4] Started new (hidden) Excel instance."
}

$excel.DisplayAlerts = $false
if ($ownExcel) { $excel.Visible = $false }

$srcWb = $null
$dstWb = $null

try {
    # -----------------------------------------------------------------------
    # Open source (always read-only)
    # -----------------------------------------------------------------------
    Write-Host "[2/4] Opening workbooks..."
    $srcWb = $null
    foreach ($wb in $excel.Workbooks) {
        if ($wb.FullName -eq $sourcePath) { $srcWb = $wb; break }
    }
    if ($null -eq $srcWb) {
        $srcWb = $excel.Workbooks.Open($sourcePath, 0, $true)
        Write-Host "  Opened source from disk."
    } else {
        Write-Host "  Source already open in Excel."
    }

    # -----------------------------------------------------------------------
    # Get or open destination
    # -----------------------------------------------------------------------
    $dstWb = $null
    foreach ($wb in $excel.Workbooks) {
        if ($wb.FullName -eq $destPath) { $dstWb = $wb; break }
    }
    if ($null -eq $dstWb) {
        $dstWb = $excel.Workbooks.Open($destPath, 0, $false)
        Write-Host "  Opened destination from disk."
    } else {
        Write-Host "  Destination already open in Excel (will save live copy)."
    }

    # -----------------------------------------------------------------------
    # List sheets
    # -----------------------------------------------------------------------
    Write-Host "Source sheets:"
    foreach ($sh in $srcWb.Sheets) { Write-Host "  - $($sh.Name)" }
    Write-Host "Destination sheets (before):"
    foreach ($sh in $dstWb.Sheets) { Write-Host "  - $($sh.Name)" }

    # -----------------------------------------------------------------------
    # Find the Slides sheet in source
    # -----------------------------------------------------------------------
    Write-Host "[3/4] Locating Slide tab..."
    $srcSheet = $null
    foreach ($sh in $srcWb.Sheets) {
        if ($sh.Name -like "*Slide*") { $srcSheet = $sh; break }
    }
    if ($null -eq $srcSheet) { Write-Error "No Slide sheet found in source."; exit 1 }
    Write-Host "  Using: $($srcSheet.Name)"

    # -----------------------------------------------------------------------
    # Snapshot destination names, remove any existing same-named sheet
    # -----------------------------------------------------------------------
    $namesBefore = @{}
    foreach ($sh in $dstWb.Sheets) { $namesBefore[$sh.Name] = $true }

    if ($namesBefore.ContainsKey($srcSheet.Name)) {
        Write-Host "  Removing existing $($srcSheet.Name) from destination..."
        $dstWb.Sheets.Item($srcSheet.Name).Delete()
        $namesBefore.Remove($srcSheet.Name)
    }

    # -----------------------------------------------------------------------
    # Copy after last sheet
    # -----------------------------------------------------------------------
    $afterSheet = $dstWb.Sheets.Item($dstWb.Sheets.Count)
    Write-Host "  Inserting after: $($afterSheet.Name)"
    $srcSheet.Copy([System.Reflection.Missing]::Value, $afterSheet)

    # Find the new sheet by diffing name lists
    $newSheet = $null
    foreach ($sh in $dstWb.Sheets) {
        if (-not $namesBefore.ContainsKey($sh.Name)) { $newSheet = $sh; break }
    }
    if ($null -eq $newSheet) { Write-Error "Could not locate the newly copied sheet."; exit 1 }
    Write-Host "  New sheet: $($newSheet.Name)"

    # Rename if Excel added a suffix
    if ($newSheet.Name -ne $srcSheet.Name) {
        $newSheet.Name = $srcSheet.Name
        Write-Host "  Renamed to: $($newSheet.Name)"
    }

    # -----------------------------------------------------------------------
    # Save  --  use SaveAs with explicit xlsm format (52) to guarantee format
    # -----------------------------------------------------------------------
    Write-Host "[4/4] Saving..."
    $dstWb.SaveAs($destPath, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled

    Write-Host ""
    Write-Host "SUCCESS - Sheet $($newSheet.Name) saved to:"
    Write-Host $destPath
    Write-Host ""
    Write-Host "Destination sheets (after):"
    foreach ($sh in $dstWb.Sheets) { Write-Host "  - $($sh.Name)" }

} catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    # Only close workbooks if WE opened them (i.e. they weren't already open)
    # Never close the user's live workbook from under them
    if ($srcWb -and $ownExcel) { try { $srcWb.Close($false) } catch {} }
    if ($ownExcel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
