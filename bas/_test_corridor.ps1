# Debug loop for the corridor renderer.
#
# Works on a throwaway copy in %TEMP%, never the live workbook, so it is safe to
# run while the real Slide Sheet is open in Excel. Imports the current bas files,
# renders, and reports the scene summary or the trapped error.
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$book = Join-Path $root "Slide Sheet - Demo.xlsm"
$tmp = Join-Path $env:TEMP ("corr_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".xlsm")
Copy-Item -LiteralPath $book -Destination $tmp -Force

$modules = @(
    @{ name = "MDL_PlanGauge"; file = Join-Path $root "bas\MDL_PlanGauge.bas" }
    @{ name = "MDL_CorridorImage"; file = Join-Path $root "bas\MDL_CorridorImage.bas" }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 1      # macros must run; EnableEvents=False blocks Workbook_Open
$wb = $null
try {
    $wb = $xl.Workbooks.Open($tmp, $false, $false)
    if (-not $wb) { $wb = $xl.Workbooks.Item(1) }
    try { $vbp = $wb.VBProject } catch { throw "VBProject not accessible - enable 'Trust access to the VBA project object model'." }

    foreach ($m in $modules) {
        foreach ($c in $vbp.VBComponents) { if ($c.Name -eq $m.name) { $vbp.VBComponents.Remove($c); break } }
        $vbp.VBComponents.Import($m.file) | Out-Null
    }
    "imported {0} modules into the temp copy" -f $modules.Count

    # Does On Error still work, or is the VBE breaking on every error? Without this
    # answer a failure and a hang look identical.
    "probe    : {0}" -f $xl.Run("CorridorProbe")
    $xl.Run("CorridorTraceReset")
    "trace    : {0}" -f (Join-Path $env:TEMP "corridor_trace.txt")

    # Each stage is announced before it runs, so if one blocks the last line printed
    # names it.
    foreach ($stage in @("scene", "canvas", "room", "section", "day", "metrics", "ops", "export")) {
        "-> {0}" -f $stage
        try {
            $r = $xl.Run("CorridorDiag", $stage)
            "   {0}" -f $r
            if ($r -like "FAIL*") { break }
        }
        catch {
            # A compile error parks Excel on a modal dialog; once something else has
            # dismissed it the call lands here and the VBE still has the bad line
            # selected, which is the only way to find out where it is.
            "   CALL ERROR: {0}" -f $_.Exception.Message
            try {
                $cp = $xl.VBE.ActiveCodePane
                "   pane   : {0}" -f $cp.CodeModule.Name
                $sl = 0; $sc = 0; $el = 0; $ec = 0
                $cp.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
                "   line   : {0}" -f $sl
                $from = [Math]::Max(1, $sl - 3)
                $to = [Math]::Min($cp.CodeModule.CountOfLines, $sl + 3)
                for ($n = $from; $n -le $to; $n++) {
                    "   {0}{1,5}: {2}" -f $(if ($n -eq $sl) { ">" } else { " " }), $n, $cp.CodeModule.Lines($n, 1)
                }
            }
            catch { "   could not read code pane: {0}" -f $_.Exception.Message }
            break
        }
    }

    $out = Join-Path $root "docs\daily_report_live.png"
    if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }

    "-> full render"
    $png = $xl.Run("RenderCorridorPng", $out)
    if ([string]::IsNullOrWhiteSpace($png)) {
        "RESULT   : FAILED"
        "error    : {0}" -f $xl.Run("CorridorLastError")
    }
    else {
        $img = [System.Drawing.Image]::FromFile($png)
        "RESULT   : {0}" -f $png
        "size     : {0} x {1} px, {2:N0} bytes" -f $img.Width, $img.Height, (Get-Item $png).Length
        $img.Dispose()
        Copy-Item -LiteralPath $png -Destination (Join-Path $root "docs\corridor_live.png") -Force
    }

    $band = Join-Path $root "docs\daily_corridor_band.png"
    if (Test-Path -LiteralPath $band) { Remove-Item -LiteralPath $band -Force }
    "-> corridor band"
    $bandPng = $xl.Run("RenderCorridorPng", $band, "corridor")
    if ([string]::IsNullOrWhiteSpace($bandPng)) {
        "BAND     : FAILED"
        "error    : {0}" -f $xl.Run("CorridorLastError")
    }
    else {
        $img = [System.Drawing.Image]::FromFile($bandPng)
        "BAND     : {0} ({1}x{2})" -f $bandPng, $img.Width, $img.Height
        $img.Dispose()
    }

    # the drawing is deleted by the renderer, so nothing should need saving
    $wb.Close($false)
    $wb = $null
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
