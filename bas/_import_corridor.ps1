# Import the corridor renderer and the patched MDL_PlanGauge (public shims added)
# into the live workbook, then render one PNG so the result can be eyeballed
# against bas\_proto_tunnel_render.ps1.
#
# Events stay off throughout so Workbook_Open / InitSetup cannot touch field data.
# Requires: Excel > Trust Center > Macro Settings > Trust access to the VBA project
# object model.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$book = Join-Path $root "Slide Sheet - Demo.xlsm"

$modules = @(
    @{ name = "MDL_PlanGauge"; file = Join-Path $root "bas\live\MDL_PlanGauge.bas" }
    @{ name = "MDL_CorridorImage"; file = Join-Path $root "bas\MDL_CorridorImage.bas" }
)
foreach ($m in $modules) {
    if (-not (Test-Path -LiteralPath $m.file)) { throw ("missing " + $m.file) }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
# msoAutomationSecurityLow: macros must be enabled to call RenderCorridorPng.
# EnableEvents = False is what keeps Workbook_Open / InitSetup from firing.
$xl.AutomationSecurity = 1
$wb = $null
try {
    $wb = $xl.Workbooks.Open($book, $false, $false)
    if (-not $wb) { $wb = $xl.Workbooks.Item(1) }

    try { $vbp = $wb.VBProject } catch { throw "VBProject not accessible - enable 'Trust access to the VBA project object model'." }

    foreach ($m in $modules) {
        $existing = $null
        foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $m.name) { $existing = $c; break } }
        if ($existing) {
            $vbp.VBComponents.Remove($existing)
            "removed  {0}" -f $m.name
        }
        $vbp.VBComponents.Import($m.file) | Out-Null
        "imported {0}" -f $m.name
    }

    $wb.Save()
    "saved workbook"

    $png = $xl.Run("RenderCorridorPng", "")
    if ([string]::IsNullOrWhiteSpace($png)) {
        "RENDER FAILED - RenderCorridorPng returned empty"
    }
    else {
        "PNG      {0}" -f $png
        if (Test-Path -LiteralPath $png) {
            $img = [System.Drawing.Image]::FromFile($png)
            "size     {0} x {1} px, {2:N0} bytes" -f $img.Width, $img.Height, (Get-Item $png).Length
            $img.Dispose()
            $dest = Join-Path $root "docs\corridor_live.png"
            Copy-Item -LiteralPath $png -Destination $dest -Force
            "copied   {0}" -f $dest
        }
    }

    # the render must not have altered the sheet, so leave without saving again
    $wb.Close($false)
    $wb = $null
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
