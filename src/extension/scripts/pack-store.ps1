# Build a Chrome Web Store zip from the current src/extension/dist build.
#   powershell -NoProfile -File src/extension/scripts/pack-store.ps1
# The store assigns its own extension id and rejects a manifest "key", so the
# packed manifest drops it. dist itself is left as built (the unpacked preview keeps its id).
param()
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$extRoot = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $extRoot)
$built = Join-Path $extRoot "dist"
$outDir = Join-Path $repoRoot "dist"

$manifestPath = Join-Path $built "manifest.json"
if (-not (Test-Path -LiteralPath $manifestPath)) { throw "no build at $built (run the extension build first)" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ([int]$manifest.manifest_version -ne 3) { throw "store zip requires Manifest V3" }
$ver = [string]$manifest.version
$src = Get-Content -LiteralPath (Join-Path $extRoot "src/version.ts") -Raw
if ($src -notmatch "PREVIEW_VERSION = `"$([regex]::Escape($ver))`"") { throw "dist is $ver but src/version.ts differs; rebuild first" }
$descLen = ([string]$manifest.description).Length
if ($descLen -gt 132) { throw "manifest description is $descLen chars (Chrome Web Store max 132)" }
foreach ($icon in $manifest.icons.PSObject.Properties) {
    if (-not (Test-Path -LiteralPath (Join-Path $built $icon.Value))) { throw "missing icon: $($icon.Value)" }
}

$stage = Join-Path ([System.IO.Path]::GetTempPath()) ("phx-store-" + [guid]::NewGuid().ToString("N"))
try {
    Copy-Item -LiteralPath $built -Destination $stage -Recurse
    Get-ChildItem -LiteralPath $stage -Recurse -File |
        Where-Object { $_.Extension -in ".map", ".pem", ".crx" } |
        Remove-Item -Force
    $env:PHX_STORE_MANIFEST = Join-Path $stage "manifest.json"
    node -e "const fs=require('fs');const p=process.env.PHX_STORE_MANIFEST;const m=JSON.parse(fs.readFileSync(p,'utf8'));delete m.key;fs.writeFileSync(p,JSON.stringify(m,null,2)+'\n')"
    if ($LASTEXITCODE -ne 0) { throw "could not strip key from manifest" }

    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    $zipPath = Join-Path $outDir "OpenCap_Slidesheet_v$ver.zip"
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($stage, $zipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
}
finally {
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$probe = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    $names = @($probe.Entries | ForEach-Object { $_.FullName.Replace("\", "/") })
    $entry = $probe.GetEntry("manifest.json")
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $packed = $reader.ReadToEnd() | ConvertFrom-Json
    $reader.Dispose()
}
finally {
    $probe.Dispose()
}
if ($names -notcontains "manifest.json") { throw "zip root missing manifest.json" }
if ($packed.PSObject.Properties.Name -contains "key") { throw "packed manifest still has key" }
$expected = @(Get-ChildItem -LiteralPath $built -Recurse -File | Where-Object { $_.Extension -notin ".map", ".pem", ".crx" }).Count
$files = @($names | Where-Object { -not $_.EndsWith("/") }).Count
if ($files -ne $expected) { throw "zip has $files files, dist has $expected" }

Write-Output "PACKED $zipPath"
Write-Output ("version  " + $ver)
Write-Output ("entries  " + $names.Count)
Write-Output ("MB       " + [math]::Round((Get-Item -LiteralPath $zipPath).Length / 1MB, 2))
