$ErrorActionPreference = 'Stop'

$root     = 'C:\Users\User\Desktop\Vibe_Projects\PHX_FieldCap'
$outFile  = Join-Path $root 'OpenCap_FieldWorkbook.xlsm'
$basSetup = Join-Path $root 'src\excel\MDL_Setup.bas'
$basDDT   = Join-Path $root 'src\excel\MDL_DDTools.bas'

# Start Excel
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false

# Create new workbook
$wb = $xl.Workbooks.Add()

# Delete extra sheets (keep just one)
while ($wb.Sheets.Count -gt 1) {
    $wb.Sheets($wb.Sheets.Count).Delete()
}
$wb.Sheets(1).Name = 'Sheet1'

# Save as macro-enabled first (required before importing VBA)
if (Test-Path $outFile) { Remove-Item $outFile -Force }
$wb.SaveAs($outFile, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
Write-Host "Workbook saved: $outFile"

# Import MDL_Setup
$mod1 = $wb.VBProject.VBComponents.Import($basSetup)
Write-Host "Imported: $($mod1.Name)"

# Import MDL_DDTools
try {
    $mod2 = $wb.VBProject.VBComponents.Import($basDDT)
    Write-Host "Imported: $($mod2.Name)"
} catch {
    Write-Host "MDL_DDTools import failed: $($_.Exception.Message)"
    Write-Host "Skipping MDL_DDTools for now."
}

# Set up ThisWorkbook code
$tbMod = $wb.VBProject.VBComponents | Where-Object { $_.Name -eq 'ThisWorkbook' }
$wbCode  = 'Private Sub Workbook_Open()' + [char]13 + [char]10
$wbCode += '    On Error Resume Next' + [char]13 + [char]10
$wbCode += '    MDL_Setup.InitSetup' + [char]13 + [char]10
$wbCode += '    On Error GoTo 0' + [char]13 + [char]10
$wbCode += 'End Sub'
$tbMod.CodeModule.AddFromString($wbCode)
Write-Host 'ThisWorkbook.Workbook_Open added.'

# Add Worksheet_Change event to the Setup sheet (auto-sort crew on role change)
# We inject into Sheet1's code-behind since InitSetup renames it to "Setup"
$shMod = $wb.VBProject.VBComponents | Where-Object { $_.Name -eq 'Sheet1' }
$shCode  = 'Private Sub Worksheet_Change(ByVal Target As Range)' + [char]13 + [char]10
$shCode += '    Dim rTop As Long, rBot As Long' + [char]13 + [char]10
$shCode += '    On Error Resume Next' + [char]13 + [char]10
$shCode += '    rTop = Me.Range("OC_CrewTop").Row' + [char]13 + [char]10
$shCode += '    rBot = Me.Range("OC_CrewBot").Row' + [char]13 + [char]10
$shCode += '    On Error GoTo 0' + [char]13 + [char]10
$shCode += '    If rTop = 0 Or rBot = 0 Then Exit Sub' + [char]13 + [char]10
$shCode += '    If Target.Column <> 10 Then Exit Sub' + [char]13 + [char]10
$shCode += '    If Target.Row < rTop Or Target.Row > rBot Then Exit Sub' + [char]13 + [char]10
$shCode += '    Application.EnableEvents = False' + [char]13 + [char]10
$shCode += '    MDL_Setup.SortCrewByRole' + [char]13 + [char]10
$shCode += '    Application.EnableEvents = True' + [char]13 + [char]10
$shCode += 'End Sub'
$shMod.CodeModule.AddFromString($shCode)
Write-Host 'Sheet1.Worksheet_Change event added.'

# Save
$wb.Save()
Write-Host ''
Write-Host 'BUILD COMPLETE. Modules:'
$wb.VBProject.VBComponents | ForEach-Object { Write-Host "  $($_.Name) [Type=$($_.Type)]" }
Write-Host ''
Write-Host 'Close and reopen the workbook to trigger Workbook_Open > InitSetup.'
Write-Host 'Or use Alt+F8 > RebuildSetup to test immediately.'
