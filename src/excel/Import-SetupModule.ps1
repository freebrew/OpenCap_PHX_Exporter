# GetActiveObject (COM interop) only exists in .NET Framework / Windows PowerShell 5.
# If running in PowerShell 7 (Core), relaunch automatically with powershell.exe.
if ($PSVersionTable.PSEdition -eq 'Core') {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $xl.Workbooks | Where-Object { $_.Name -like 'OpenCap*' }
if (-not $wb) { throw 'Workbook not open' }

$vbp = $wb.VBProject

# ── 1. Remove MDL_Setup and any stray Module* copies (previous failed imports) ─
$setupPath = 'C:\Users\User\Desktop\Vibe_Projects\PHX_FieldCap\src\excel\MDL_Setup.bas'
$targetCode = [System.IO.File]::ReadAllText($setupPath).Substring(0, [Math]::Min(200, [System.IO.File]::ReadAllText($setupPath).Length))

foreach ($m in @($vbp.VBComponents)) {
    $remove = $false
    if ($m.Name -match '^MDL_Setup') { $remove = $true }
    # Also remove any generic Module* that contains MDL_Setup code (ghost imports)
    elseif ($m.Name -match '^Module\d+' -and $m.Type -eq 1) {
        $firstLines = ''
        if ($m.CodeModule.CountOfLines -gt 0) {
            $firstLines = $m.CodeModule.Lines(1, [Math]::Min(5, $m.CodeModule.CountOfLines))
        }
        if ($firstLines -match 'MDL_Setup|OpenCap Field Workbook') { $remove = $true }
    }
    if ($remove) { $vbp.VBComponents.Remove($m); Write-Host "Removed: $($m.Name)" }
}

# ── 2. Import MDL_Setup directly from .bas file ──────────────────────────────
# VBComponents.Import reads the Attribute VB_Name header and handles naming
# automatically — avoids the 0x800A802D rename error after Add().
$mod = $vbp.VBComponents.Import($setupPath)
Write-Host "MDL_Setup imported ($($mod.CodeModule.CountOfLines) lines)"

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
