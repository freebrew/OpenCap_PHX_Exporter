$ErrorActionPreference = "Stop"
$root = Join-Path $env:TEMP ("phx-host-install-" + [guid]::NewGuid().ToString("N"))
$chrome = "HKCU:\Software\PHX\OutlookHostTest\" + [guid]::NewGuid().ToString("N") + "\chrome"
$edge = "HKCU:\Software\PHX\OutlookHostTest\" + [guid]::NewGuid().ToString("N") + "\edge"
$installer = Join-Path $PSScriptRoot "..\install-host.ps1"
$id = "glkcmpmiomoppdkmdcjaepneimgbipdl"
$source = Get-Content -Raw -Path $installer
foreach ($required in @(
  "HKCU:\Software\Google\Chrome\NativeMessagingHosts\com.phx.outlook",
  "HKCU:\Software\Microsoft\Edge\NativeMessagingHosts\com.phx.outlook",
  "LOCALAPPDATA"
)) {
  if ($source -notlike "*$required*") { throw "Installer is missing $required" }
}
try {
  & $installer -Action Install -InstallRoot $root -RegistryChrome $chrome -RegistryEdge $edge -Version "1.0.0" -ExtensionId $id | Out-Null
  $first = Join-Path $root "1.0.0\com.phx.outlook.json"
  if (-not (Test-Path $first)) { throw "Install did not write the manifest." }
  $manifest = Get-Content -Raw -Path $first | ConvertFrom-Json
  if ($manifest.allowed_origins[0] -ne "chrome-extension://$id/") { throw "Allowed origin was not written." }
  if ((Get-Item $chrome).GetValue("") -ne $first) { throw "Chrome registration does not point at 1.0.0." }
  & $installer -Action Install -InstallRoot $root -RegistryChrome $chrome -RegistryEdge $edge -Version "1.0.1" -ExtensionId $id | Out-Null
  $second = Join-Path $root "1.0.1\com.phx.outlook.json"
  if (-not (Test-Path $first) -or -not (Test-Path $second)) { throw "Upgrade removed a version folder." }
  if ((Get-Item $chrome).GetValue("") -ne $second) { throw "Upgrade did not switch the registration." }
  & $installer -Action Rollback -InstallRoot $root -RegistryChrome $chrome -RegistryEdge $edge | Out-Null
  if ((Get-Item $chrome).GetValue("") -ne $first) { throw "Rollback did not restore 1.0.0." }
  & $installer -Action Repair -InstallRoot $root -RegistryChrome $chrome -RegistryEdge $edge -Version "1.0.1" -ExtensionId $id | Out-Null
  if ((Get-Item $chrome).GetValue("") -ne $second) { throw "Repair did not rewrite 1.0.1." }
  & $installer -Action Uninstall -InstallRoot $root -RegistryChrome $chrome -RegistryEdge $edge | Out-Null
  if ((Test-Path $root) -or (Test-Path $chrome) -or (Test-Path $edge)) { throw "Uninstall left files or registration behind." }
  Write-Output "native host install ok"
} finally {
  if (Test-Path $root) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
  foreach ($key in @($chrome, $edge)) {
    if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue }
  }
}
