param(
  [ValidateSet("Install", "Repair", "Rollback", "Uninstall")]
  [string]$Action = "Install",
  [string]$InstallRoot,
  [string]$RegistryChrome,
  [string]$RegistryEdge,
  [string]$Version,
  [string]$ExtensionId
)
$ErrorActionPreference = "Stop"
if (-not $InstallRoot) { $InstallRoot = Join-Path $env:LOCALAPPDATA "PHX\OutlookHost" }
if (-not $RegistryChrome) { $RegistryChrome = "HKCU:\Software\Google\Chrome\NativeMessagingHosts\com.phx.outlook" }
if (-not $RegistryEdge) { $RegistryEdge = "HKCU:\Software\Microsoft\Edge\NativeMessagingHosts\com.phx.outlook" }

function Read-InstalledVersion {
  if ($Version) { return $Version }
  $path = Join-Path $PSScriptRoot "host-version.json"
  $parsed = Get-Content -Raw -Path $path | ConvertFrom-Json
  return [string]$parsed.version
}

function Read-ExtensionId {
  if ($ExtensionId) { return $ExtensionId.Trim() }
  return (Get-Content -Raw -Path (Join-Path $PSScriptRoot "extension-id.txt")).Trim()
}

function Set-HostRegistration($manifestPath) {
  foreach ($key in @($RegistryChrome, $RegistryEdge)) {
    $parent = Split-Path $key
    $leaf = Split-Path $key -Leaf
    if (-not (Test-Path $parent)) { New-Item -Path $parent -Force | Out-Null }
    if (-not (Test-Path $key)) { New-Item -Path $parent -Name $leaf -Force | Out-Null }
    Set-ItemProperty -Path $key -Name "(default)" -Value $manifestPath
  }
}

function Remove-HostRegistration {
  foreach ($key in @($RegistryChrome, $RegistryEdge)) {
    if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force }
  }
}

if ($Action -eq "Uninstall") {
  Remove-HostRegistration
  if (Test-Path $InstallRoot) { Remove-Item -LiteralPath $InstallRoot -Recurse -Force }
  Write-Output "Uninstalled $InstallRoot"
  return
}

if ($Action -eq "Rollback") {
  $previous = Join-Path $InstallRoot "previous.txt"
  if (-not (Test-Path $previous)) { throw "There is no previous helper to restore." }
  $manifestPath = (Get-Content -Raw -Path $previous).Trim()
  if (-not (Test-Path $manifestPath)) { throw "The previous helper manifest is missing." }
  Set-HostRegistration $manifestPath
  Write-Output "Restored $manifestPath"
  return
}

$resolvedVersion = Read-InstalledVersion
$resolvedId = Read-ExtensionId
if ($resolvedId -notmatch "^[a-p]{32}$") { throw "Extension id is not a Chrome extension id." }
$dest = Join-Path $InstallRoot $resolvedVersion
New-Item -ItemType Directory -Force -Path $dest | Out-Null
foreach ($name in @("phx-outlook-host.ps1", "phx-outlook-host.cmd", "install-host.ps1", "install-host.cmd", "repair-host.cmd", "uninstall-host.cmd")) {
  Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $dest $name) -Force
}
Set-Content -LiteralPath (Join-Path $dest "extension-id.txt") -Value $resolvedId -NoNewline
Set-Content -LiteralPath (Join-Path $dest "expected-origin.txt") -Value "chrome-extension://$resolvedId/" -NoNewline
Set-Content -LiteralPath (Join-Path $dest "host-version.json") -Value "{`"version`":`"$resolvedVersion`"}" -NoNewline
$cmdPath = (Join-Path $dest "phx-outlook-host.cmd").Replace("\", "\\").Replace('"', '')
$manifestPath = Join-Path $dest "com.phx.outlook.json"
$manifest = @"
{
  "name": "com.phx.outlook",
  "description": "PHX Classic Outlook compose host",
  "path": "$cmdPath",
  "type": "stdio",
  "allowed_origins": ["chrome-extension://$resolvedId/"]
}
"@
Set-Content -LiteralPath $manifestPath -Value $manifest -NoNewline
$previousFile = Join-Path $InstallRoot "previous.txt"
if (Test-Path $RegistryChrome) {
  $current = [string](Get-ItemProperty -Path $RegistryChrome)."(default)"
  if ($current -and $current -ne $manifestPath) { Set-Content -LiteralPath $previousFile -Value $current -NoNewline }
}
Set-HostRegistration $manifestPath
Write-Output "$Action $manifestPath"
