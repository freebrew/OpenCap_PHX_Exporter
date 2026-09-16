# Render EMAIL PNG samples from a temp copy of Demo. Does not open Field,
# does not write Demo/Field cells, does not attach to the user's Excel.
# Output: docs/corridor_email.png, docs/corridor_live.png, docs/daily_corridor_band.png
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root "Slide Sheet - Demo.xlsm"
if (-not (Test-Path -LiteralPath $src)) { throw "missing $src" }

$tmp = Join-Path $env:TEMP ("corr_demo_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".xlsm")
Copy-Item -LiteralPath $src -Destination $tmp -Force
Write-Host ("temp copy: {0}" -f $tmp)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1
$wb = $null
try {
    $wb = $xl.Workbooks.Open($tmp, 0, $false)
    if (-not $wb) { $wb = $xl.Workbooks.Item(1) }
    # VBProject can lag behind Workbooks.Open; poll until components are visible.
    $comps = 0
    for ($try = 1; $try -le 15; $try++) {
        try { $comps = $wb.VBProject.VBComponents.Count } catch { $comps = 0 }
        if ($comps -gt 0) { break }
        Start-Sleep -Seconds 2
    }
    if ($comps -le 0) { throw "VBProject never became available" }
    Write-Host ("opened comps={0}" -f $comps)
    $xl.Run("CorridorTraceReset")

    $full = Join-Path $env:TEMP "daily_report_demo.png"
    $band = Join-Path $env:TEMP "daily_corridor_band_demo.png"
    if (Test-Path -LiteralPath $full) { Remove-Item -LiteralPath $full -Force }
    if (Test-Path -LiteralPath $band) { Remove-Item -LiteralPath $band -Force }

    Write-Host "-> RenderCorridorPng all"
    $png = $xl.Run("MDL_CorridorImage.RenderCorridorPng", $full, "all")
    if ([string]::IsNullOrWhiteSpace($png)) {
        Write-Host ("FAIL all: {0}" -f $xl.Run("MDL_CorridorImage.CorridorLastError"))
        Write-Host ("scene: {0}" -f $xl.Run("MDL_CorridorImage.CorridorSceneInfo"))
    }
    else {
        $img = [System.Drawing.Image]::FromFile($png)
        Write-Host ("OK all {0}x{1} {2} bytes" -f $img.Width, $img.Height, (Get-Item $png).Length)
        $img.Dispose()
        Copy-Item -LiteralPath $png -Destination (Join-Path $root "docs\corridor_email.png") -Force
        Copy-Item -LiteralPath $png -Destination (Join-Path $root "docs\corridor_live.png") -Force
    }

    Write-Host "-> RenderCorridorPng corridor"
    $b = $xl.Run("MDL_CorridorImage.RenderCorridorPng", $band, "corridor")
    if ([string]::IsNullOrWhiteSpace($b)) {
        Write-Host ("FAIL band: {0}" -f $xl.Run("MDL_CorridorImage.CorridorLastError"))
    }
    else {
        $img = [System.Drawing.Image]::FromFile($b)
        Write-Host ("OK band {0}x{1}" -f $img.Width, $img.Height)
        $img.Dispose()
        Copy-Item -LiteralPath $b -Destination (Join-Path $root "docs\daily_corridor_band.png") -Force
    }

    $wb.Close($false)
    $wb = $null
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
Write-Host "done"
