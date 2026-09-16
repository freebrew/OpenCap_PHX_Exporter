# Build a Chrome Web Store / Edge Add-ons zip (manifest at zip root).
#   powershell -NoProfile -File src/chrome-extension/pack-store.ps1
param()
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$extRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $extRoot)
$distDir = Join-Path $repoRoot "dist"
$manifestPath = Join-Path $extRoot "manifest.json"

$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if ([int]$manifest.manifest_version -ne 3) { throw "store zip requires Manifest V3" }
$ver = [string]$manifest.version
if ($ver -notmatch '^\d+\.\d+\.\d+$') { throw "invalid manifest version: $ver" }
$descLen = ([string]$manifest.description).Length
if ($descLen -gt 132) { throw "manifest description is $descLen chars (Chrome Web Store max 132)" }

$include = @(
    "manifest.json",
    "browser-api.js",
    "background.js",
    "content.js",
    "injected-spy.js",
    "popup.html",
    "popup.js",
    "icons/icon16.png",
    "icons/icon48.png",
    "icons/icon128.png"
)

foreach ($rel in $include) {
    $p = Join-Path $extRoot $rel
    if (-not (Test-Path -LiteralPath $p)) { throw "missing store file: $rel" }
}

New-Item -ItemType Directory -Force -Path $distDir | Out-Null
$zipName = "OpenCap_DataExporter_v$ver.zip"
$zipPath = Join-Path $distDir $zipName
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($rel in $include) {
        $src = Join-Path $extRoot $rel
        $entry = $rel.Replace("\", "/")
        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip, $src, $entry, [System.IO.Compression.CompressionLevel]::Optimal)
    }
}
finally {
    $zip.Dispose()
}

$probe = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    $names = @($probe.Entries | ForEach-Object { $_.FullName })
}
finally {
    $probe.Dispose()
}
if ($names -notcontains "manifest.json") { throw "zip root missing manifest.json" }
if ($names | Where-Object { $_ -match '\.(ps1|pem|md|crx)$' }) { throw "zip contains non-store files" }

Write-Output "PACKED $zipPath"
Write-Output ("version  " + $ver)
Write-Output ("bytes    " + (Get-Item -LiteralPath $zipPath).Length)
Write-Output "entries"
$names | ForEach-Object { Write-Output ("  " + $_) }
