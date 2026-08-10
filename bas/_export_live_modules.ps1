# Export the live VBA modules we are about to touch, so patches are made against
# the workbook's actual code rather than possibly-stale copies in bas\.
# Exports to bas\live\<Name>.bas
if ($PSVersionTable.PSEdition -eq "Core") {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    exit $LASTEXITCODE
}
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root "Slide Sheet - Demo.xlsm"
$dst = Join-Path $PSScriptRoot "live"
New-Item -ItemType Directory -Force -Path $dst | Out-Null

$want = @("MDL_PlanGauge", "Module11", "Module21", "MDL_WellborePos")

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 3
$tmp = Join-Path $env:TEMP ("wbexp_" + [Guid]::NewGuid().ToString("N") + ".xlsm")
Copy-Item -LiteralPath $src -Destination $tmp -Force
$wb = $xl.Workbooks.Open($tmp, $false, $true)
if (-not $wb) { $wb = $xl.Workbooks.Item(1) }

try {
    $vbp = $wb.VBProject
    Write-Host ("components: {0}" -f $vbp.VBComponents.Count)
    $names = @()
    foreach ($c in $vbp.VBComponents) { $names += ("{0} (type {1}, {2} lines)" -f $c.Name, $c.Type, $c.CodeModule.CountOfLines) }
    $names | Sort-Object | ForEach-Object { Write-Host ("  " + $_) }

    Write-Host ""
    foreach ($n in $want) {
        $comp = $null
        foreach ($c in $vbp.VBComponents) { if ([string]$c.Name -eq $n) { $comp = $c; break } }
        if (-not $comp) { Write-Host ("  MISSING: {0}" -f $n); continue }
        $path = Join-Path $dst ($n + ".bas")
        $comp.Export($path)
        Write-Host ("  exported {0} -> {1} ({2} lines)" -f $n, $path, $comp.CodeModule.CountOfLines)
    }
}
finally {
    try { $wb.Close($false) } catch {}
    try { $xl.Quit() } catch {}
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
