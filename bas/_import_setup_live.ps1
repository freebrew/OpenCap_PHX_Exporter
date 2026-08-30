# Headless import of bas\MDL_Setup.bas into the live Demo workbook.
# - Backs up Demo into backups/ first
# - Works on a %TEMP% copy: opened straight from Z:\ (network drive) the
#   VBProject loads empty (macro trust policy); local copies load fully.
# - Events off so Workbook_Open / InitSetup cannot run
# - Remove + Import (established MDL_Setup pattern), then compile check
# - Saves temp copy, copies back over Demo, reopen verify on a fresh temp copy
#
# GetActiveObject-free: creates its own Excel COM instance.
if ($PSVersionTable.PSEdition -eq "Core") {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    exit $LASTEXITCODE
}
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$book = Join-Path $root "Slide Sheet - Demo.xlsm"
$basFile = Join-Path $PSScriptRoot "MDL_Setup.bas"
$backups = Join-Path $root "backups"
if (-not (Test-Path -LiteralPath $book)) { throw "missing workbook: $book" }
if (-not (Test-Path -LiteralPath $basFile)) { throw "missing module: $basFile" }
New-Item -ItemType Directory -Force -Path $backups | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$bak = Join-Path $backups ("Slide Sheet - Demo.xlsm.bak_setup_import_" + $stamp)
Copy-Item -LiteralPath $book -Destination $bak -Force
"backup: $bak"

$tmpBook = Join-Path $env:TEMP ("setup_import_" + $stamp + ".xlsm")
Copy-Item -LiteralPath $book -Destination $tmpBook -Force

. (Join-Path $PSScriptRoot "VbaCompile.ps1")

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
# msoAutomationSecurityLow: with ForceDisable (3) the VBProject loads empty.
# EnableEvents = False is what keeps Workbook_Open / InitSetup from firing.
$xl.AutomationSecurity = 1
$wb = $null
$beforeCount = 0
try {
    $wb = $xl.Workbooks.Open($tmpBook, 0, $false)
    # VBProject can come back as an empty stub right after Open while Excel
    # is still loading the project - poll until components appear.
    $vbp = $null
    for ($try = 1; $try -le 15; $try++) {
        try {
            $vbp = $wb.VBProject
            if ($vbp -and [int]$vbp.VBComponents.Count -gt 0) { break }
        } catch {
            if ($try -eq 15) { throw "VBProject not accessible - enable 'Trust access to the VBA project object model'." }
        }
        "waiting for VBProject to load (attempt $try)..."
        Start-Sleep -Seconds 2
    }
    $beforeCount = [int]$vbp.VBComponents.Count
    "before: project=$($vbp.Name) comps=$beforeCount"
    if ($beforeCount -lt 50) { throw "Demo VBA project looks empty/corrupt ($beforeCount comps) - aborting before any change" }

    $existing = @($vbp.VBComponents | Where-Object { $_.Name -eq "MDL_Setup" })
    foreach ($m in $existing) {
        $vbp.VBComponents.Remove($m)
        "removed  MDL_Setup"
    }
    Start-Sleep -Milliseconds 400

    $imported = $vbp.VBComponents.Import($basFile)
    if ($imported.Name -ne "MDL_Setup") { $imported.Name = "MDL_Setup" }
    "imported MDL_Setup ($($imported.CodeModule.CountOfLines) lines)"

    # ThisWorkbook must not auto-run InitSetup (guard only; never rewrite here)
    $tb = $vbp.VBComponents.Item("ThisWorkbook").CodeModule
    if ($tb.CountOfLines -gt 0 -and ($tb.Lines(1, $tb.CountOfLines) -match "InitSetup")) {
        Write-Warning "ThisWorkbook references InitSetup - review before shipping"
    }

    $res = Test-VbaCompile -Excel $xl
    Show-CompileResult $res $xl
    if (-not $res.Ok) { throw "VBA compile failed after import - workbook NOT saved; restore not needed" }

    $wb.Save()
    "saved temp copy"
    $wb.Close($false)
    $wb = $null
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    try { $xl.Quit() } catch {}
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Copy-Item -LiteralPath $tmpBook -Destination $book -Force
"copied back over Demo"

# Reopen verify (fresh temp copy of the updated Demo)
$tmpVerify = Join-Path $env:TEMP ("setup_verify_" + $stamp + ".xlsm")
Copy-Item -LiteralPath $book -Destination $tmpVerify -Force
$xl2 = New-Object -ComObject Excel.Application
$xl2.Visible = $false
$xl2.DisplayAlerts = $false
$xl2.EnableEvents = $false
$xl2.AutomationSecurity = 1
$wb2 = $null
try {
    $wb2 = $xl2.Workbooks.Open($tmpVerify, 0, $true)
    $c = [int]$wb2.VBProject.VBComponents.Count
    "reopen: project=$($wb2.VBProject.Name) comps=$c"
    if ($c -ne $beforeCount) { throw "component count changed ($beforeCount -> $c) - restore $bak" }
}
finally {
    if ($wb2) { $wb2.Close($false) }
    $xl2.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2)
    Remove-Item -LiteralPath $tmpBook, $tmpVerify -Force -ErrorAction SilentlyContinue
}
"PASS: MDL_Setup imported into Demo"
