param(
    [string]$Entry = "remotion/src/index.jsx",
    [string]$OutDir = "remotion/out"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$assets = @(
    "screenshot-1280-export",
    "screenshot-1280-operations",
    "screenshot-1280-diagnostics",
    "screenshot-640-export",
    "screenshot-640-operations",
    "screenshot-640-diagnostics",
    "small-promo-export",
    "small-promo-operations",
    "small-promo-diagnostics",
    "marquee-export",
    "marquee-operations",
    "marquee-diagnostics"
)

if (-not (Test-Path $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir | Out-Null
}
Get-ChildItem -Path $OutDir -Filter "*.jpg" -ErrorAction SilentlyContinue | Remove-Item -Force

$total = $assets.Count
for ($i = 0; $i -lt $total; $i++) {
    $asset = $assets[$i]
    $current = $i + 1
    $percent = [math]::Round(($current / $total) * 100)
    $output = Join-Path $OutDir "$asset.jpg"

    Write-Host ("[STEP 1/1] Rendering asset [ {0}/{1} ] [ {2}% ] {3}" -f $current, $total, $percent, $asset)

    npx remotion still $Entry $asset $output --public-dir=remotion/public --image-format=jpeg --jpeg-quality=95 --overwrite=true --log=warn
}

Write-Host "------------------------------------------------------------"
Write-Host "RUN COMPLETE | Status: SUCCESS"
Write-Host ("Rendered : {0} assets" -f $total)
Write-Host ("Output   : {0}" -f (Resolve-Path $OutDir))
Write-Host "------------------------------------------------------------"

