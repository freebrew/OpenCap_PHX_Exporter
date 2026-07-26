# GetActiveObject only works in Windows PowerShell 5 (not PowerShell Core).
if ($PSVersionTable.PSEdition -eq "Core") {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"

$setupPath = "C:\Users\User\Desktop\Vibe_Projects\PHX_FieldCap\src\excel\MDL_Setup.bas"

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks | Where-Object { $_.Name -like "OpenCap*" }
if (-not $wb) { throw "OpenCap workbook not open in Excel" }

$vbp = $wb.VBProject

# ── 1. Rename any existing MDL_Setup* copies to free the "MDL_Setup" name ─────────
#    Renaming first (rather than removing) lets VBA release the name before we
#    import, so the new module lands as "MDL_Setup" with no numeric suffix.
$existing = @($vbp.VBComponents | Where-Object { $_.Name -like "MDL_Setup*" })
foreach ($m in $existing) {
    Write-Host "Staging old module for removal: $($m.Name)"
    $m.Name = "_MDL_Setup_OLD"
}

# ── 2. Import fresh from .bas file ─────────────────────────────────────────────────
#    VBComponents.Import reads the Attribute VB_Name header itself and names
#    the module correctly — no manual AddFromString / DeleteLines needed.
$imported = $vbp.VBComponents.Import($setupPath)
Write-Host "Imported as: $($imported.Name)  ($($imported.CodeModule.CountOfLines) lines)"

# ── 3. Remove the staged old copy ─────────────────────────────────────────────────
$old = $vbp.VBComponents | Where-Object { $_.Name -eq "_MDL_Setup_OLD" }
if ($old) {
    $vbp.VBComponents.Remove($old)
    Write-Host "Old module removed."
}

# ── 4. Re-write ThisWorkbook event ────────────────────────────────────────────────
$tbMod = $vbp.VBComponents | Where-Object { $_.Name -eq "ThisWorkbook" }
$lc = $tbMod.CodeModule.CountOfLines
if ($lc -gt 0) { $tbMod.CodeModule.DeleteLines(1, $lc) }
$wbCode  = "Private Sub Workbook_Open()" + [char]13 + [char]10
$wbCode += "    On Error Resume Next" + [char]13 + [char]10
$wbCode += "    MDL_Setup.InitSetup" + [char]13 + [char]10
$wbCode += "    On Error GoTo 0" + [char]13 + [char]10
$wbCode += "End Sub"
$tbMod.CodeModule.AddFromString($wbCode)
Write-Host "ThisWorkbook updated."

# ── 5. Save ───────────────────────────────────────────────────────────────────────
$wb.SaveAs($wb.FullName, 52)
Write-Host "Saved. Modules:"
$vbp.VBComponents | ForEach-Object {
    Write-Host "  $($_.Name)  [Type=$($_.Type)  Lines=$($_.CodeModule.CountOfLines)]"
}
