$ErrorActionPreference = 'Stop'

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $xl.Workbooks | Where-Object { $_.Name -like 'OpenCap*' }
if (-not $wb) { throw 'Workbook not open' }

$vbp = $wb.VBProject

# ── 1. Remove all MDL_Setup copies ───────────────────────────────────────────
foreach ($n in @('MDL_Setup','MDL_Setup1','MDL_Setup2','MDL_Setup3')) {
    $m = $vbp.VBComponents | Where-Object { $_.Name -eq $n }
    if ($m) { $vbp.VBComponents.Remove($m); Write-Host "Removed: $n" }
}

# ── 2. Add fresh MDL_Setup from code text ────────────────────────────────────
$setupPath = 'C:\Users\User\Desktop\Vibe_Projects\PHX_FieldCap\src\excel\MDL_Setup.bas'
$raw = [System.IO.File]::ReadAllText($setupPath, [System.Text.Encoding]::Default)
$lines = $raw -split "`n"
$si = 0
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -notmatch '^Attribute\s+VB_' -and $lines[$i] -notmatch '^VERSION\s+') {
        $si = $i; break
    }
}
$clean = ($lines[$si..($lines.Count-1)]) -join "`n"
$mod = $vbp.VBComponents.Add(1)
$mod.Name = 'MDL_Setup'
$mod.CodeModule.AddFromString($clean)
Write-Host "MDL_Setup inserted ($($clean.Length) chars)"

# ── 3. Fix ThisWorkbook: remove OnTime, use direct call with guard ───────────
$tbMod = $vbp.VBComponents | Where-Object { $_.Name -eq 'ThisWorkbook' }
$lc = $tbMod.CodeModule.CountOfLines
if ($lc -gt 0) { $tbMod.CodeModule.DeleteLines(1, $lc) }

$wbCode  = 'Private Sub Workbook_Open()' + [char]13 + [char]10
$wbCode += '    On Error Resume Next' + [char]13 + [char]10
$wbCode += '    MDL_Setup.InitSetup' + [char]13 + [char]10
$wbCode += '    On Error GoTo 0' + [char]13 + [char]10
$wbCode += 'End Sub'
$tbMod.CodeModule.AddFromString($wbCode)
Write-Host 'ThisWorkbook updated (direct call, guard protects).'

# ── 4. Save ──────────────────────────────────────────────────────────────────
$wb.Save()
Write-Host 'Saved. Modules:'
$vbp.VBComponents | ForEach-Object { Write-Host "  $($_.Name)  [Type=$($_.Type)]" }
