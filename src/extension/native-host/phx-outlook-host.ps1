# Chrome native-messaging host for Classic Outlook compose. Stdout is the protocol, so nothing else may write to it.
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$script:MaxBytes = 64MB
if ($env:PHX_MAIL_MAX_BYTES) { $script:MaxBytes = [int64]$env:PHX_MAIL_MAX_BYTES }
$script:Runs = @{}

function Get-ExpectedOrigin {
  if ($env:PHX_MAIL_ORIGIN) { return $env:PHX_MAIL_ORIGIN.Trim() }
  $path = Join-Path $PSScriptRoot "expected-origin.txt"
  if (Test-Path $path) { return (Get-Content -Raw -Path $path).Trim() }
  return ""
}

function Read-HostMessage {
  $stdin = $script:InputReader
  try { $length = $stdin.ReadInt32() } catch { return $null }
  if ($length -le 0 -or $length -gt 8388608) { throw "Message length is not allowed." }
  $bytes = $stdin.ReadBytes($length)
  if ($bytes.Length -ne $length) { throw "Message was truncated." }
  return [Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json
}

function Write-HostMessage($fields) {
  $json = ConvertTo-Json -InputObject $fields -Compress -Depth 6
  $bytes = [Text.Encoding]::UTF8.GetBytes($json)
  $script:OutputWriter.Write([int32]$bytes.Length)
  $script:OutputWriter.Write($bytes)
  $script:OutputWriter.Flush()
}

function Read-Field($message, $name) {
  $property = $message.PSObject.Properties[$name]
  if ($null -eq $property -or $null -eq $property.Value) { return "" }
  return [string]$property.Value
}

function Test-RunId($runId) { return $runId -match "^[a-f0-9]{32}$" }

function Get-RunDir($runId) {
  $root = [IO.Path]::GetFullPath((Join-Path $env:TEMP "phx-mail"))
  $full = [IO.Path]::GetFullPath((Join-Path $root $runId))
  if (-not $full.StartsWith($root)) { throw "Run folder is outside the helper temp directory." }
  return $full
}

function Clear-Run($runId) {
  if ($script:Runs.ContainsKey($runId)) {
    Remove-Item -LiteralPath $script:Runs[$runId].Dir -Recurse -Force -ErrorAction SilentlyContinue
    $script:Runs.Remove($runId)
  }
}

function Convert-SafeName($name) {
  $leaf = [IO.Path]::GetFileName($name)
  $clean = [regex]::Replace($leaf, "[^A-Za-z0-9._ -]", "_")
  if ($clean.Length -gt 80) { $clean = $clean.Substring(0, 80) }
  if ($clean -eq "" -or $clean -eq "." -or $clean -eq "..") { return "file.bin" }
  return $clean
}

function Read-RegistryDefault($path) {
  try { return [string](Get-ItemProperty -Path $path -ErrorAction Stop)."(default)" } catch { return "" }
}

function Test-ClassicRegistered {
  return (Test-Path "HKLM:\SOFTWARE\Classes\Outlook.Application\CLSID") -or (Test-Path "HKCU:\Software\Classes\Outlook.Application\CLSID")
}

function Test-OutlookRunning {
  try {
    $null = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    return $true
  } catch { return $false }
}

function Get-OutlookApp {
  try { return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch {}
  if (-not (Test-ClassicRegistered)) { return $null }
  try { return New-Object -ComObject Outlook.Application } catch {}
  $null = Start-Process -FilePath "outlook.exe" -WindowStyle Minimized -ErrorAction SilentlyContinue
  $deadline = (Get-Date).AddSeconds(30)
  while ((Get-Date) -lt $deadline) {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch {}
    Start-Sleep -Milliseconds 500
  }
  return $null
}

function Send-Probe {
  $mailto = Read-RegistryDefault "HKCU:\Software\Classes\mailto\shell\open\command"
  if ($mailto -eq "") { $mailto = Read-RegistryDefault "HKLM:\SOFTWARE\Classes\mailto\shell\open\command" }
  $client = Read-RegistryDefault "HKCU:\Software\Clients\Mail"
  if ($client -eq "") { $client = Read-RegistryDefault "HKLM:\SOFTWARE\Clients\Mail" }
  $version = "0"
  $versionPath = Join-Path $PSScriptRoot "host-version.json"
  if (Test-Path $versionPath) {
    $parsed = Get-Content -Raw -Path $versionPath | ConvertFrom-Json
    if ($parsed.version) { $version = [string]$parsed.version }
  }
  Write-HostMessage @{
    ok = $true
    command = "probe"
    hostVersion = $version
    protocol = 1
    outlookRegistered = [bool](Test-ClassicRegistered)
    outlookRunning = [bool](Test-OutlookRunning)
    mailtoCommand = $mailto
    defaultClient = $client
  }
}

function Start-Compose($message) {
  $runId = Read-Field $message "runId"
  if (-not (Test-RunId $runId)) { throw "Run id is not allowed." }
  if ($script:Runs.ContainsKey($runId)) { throw "That compose is already open." }
  $subject = Read-Field $message "subject"
  $html = Read-Field $message "html"
  if ($subject.Length -gt 200) { throw "Subject is too long." }
  if ($html.Length -gt 2000000) { throw "Message body is too long." }
  $dir = Get-RunDir $runId
  New-Item -ItemType Directory -Force -Path $dir | Out-Null
  $script:Runs[$runId] = @{
    Dir = $dir
    Subject = $subject
    To = Read-Field $message "to"
    Cc = Read-Field $message "cc"
    Html = $html
    Files = @{}
    Bytes = [int64]0
  }
  Write-HostMessage @{ ok = $true; command = "compose.begin" }
}

function Add-Chunk($message) {
  $runId = Read-Field $message "runId"
  if (-not $script:Runs.ContainsKey($runId)) { throw "Start the compose before sending a file." }
  $run = $script:Runs[$runId]
  $fileId = Read-Field $message "fileId"
  if ($fileId -notmatch "^[a-f0-9]{8,32}$") { throw "File id is not allowed." }
  $role = Read-Field $message "role"
  if ($role -ne "attachment" -and $role -ne "visible-png" -and $role -ne "inline-png") { throw "File role is not allowed." }
  $index = [int](Read-Field $message "index")
  $count = [int](Read-Field $message "count")
  if ($count -lt 1 -or $count -gt 512 -or $index -lt 0 -or $index -ge $count) { throw "Chunk position is not allowed." }
  if (-not $run.Files.ContainsKey($fileId)) {
    if ($run.Files.Count -ge 8) { throw "Too many files." }
    $run.Files[$fileId] = @{ Next = 0; Count = $count; Role = $role; Name = (Convert-SafeName (Read-Field $message "name")); Parts = New-Object System.Collections.Generic.List[byte[]] }
  }
  $file = $run.Files[$fileId]
  if ($file.Count -ne $count -or $index -ne $file.Next) { throw "Chunks arrived out of order." }
  $data = Read-Field $message "data"
  $decoded = [Convert]::FromBase64String($data)
  if ($decoded.Length -gt 262144) { throw "Chunk is too large." }
  $run.Bytes += $decoded.Length
  if ($run.Bytes -gt $script:MaxBytes) { throw "The message is over the size limit." }
  $file.Parts.Add($decoded)
  $file.Next += 1
  if ($file.Next -eq $file.Count) {
    $stream = New-Object IO.MemoryStream
    foreach ($part in $file.Parts) { $stream.Write($part, 0, $part.Length) }
    $path = Join-Path $run.Dir $file.Name
    [IO.File]::WriteAllBytes($path, $stream.ToArray())
    $file.Path = $path
  }
  Write-HostMessage @{ ok = $true; command = "file.chunk" }
}

function Show-Compose($message) {
  $runId = Read-Field $message "runId"
  if (-not $script:Runs.ContainsKey($runId)) { throw "There is no compose to open." }
  $run = $script:Runs[$runId]
  foreach ($file in $run.Files.Values) {
    if (-not $file.Path) { throw "A file is missing chunks." }
  }
  $created = $false
  $reply = $null
  try {
    if ($env:PHX_MAIL_SKIP_OUTLOOK -eq "1") {
      $reply = @{ ok = $true; command = "compose.commit"; displayed = $false; skipped = $true }
    } else {
      $app = Get-OutlookApp
      if ($null -eq $app) {
        $reply = @{ ok = $false; code = "outlook-unavailable"; error = "Classic Outlook is not available. New Outlook cannot open this compose." }
      } else {
        $mail = $app.CreateItem(0)
        $created = $true
        $mail.To = $run.To
        $mail.CC = $run.Cc
        $mail.Subject = $run.Subject
        foreach ($file in $run.Files.Values) {
          if ($file.Role -eq "inline-png") { continue }
          $null = $mail.Attachments.Add($file.Path)
        }
        foreach ($file in $run.Files.Values) {
          if ($file.Role -ne "inline-png") { continue }
          $attachment = $mail.Attachments.Add($file.Path)
          try {
            $attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "daily_report.png")
            $attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", $true)
          } catch {}
        }
        $mail.HTMLBody = $run.Html
        $mail.Display()
        $reply = @{ ok = $true; command = "compose.commit"; displayed = $true }
      }
    }
  } catch {
    if ($created) {
      $reply = @{ ok = $false; code = "indeterminate"; created = $true; error = $_.Exception.Message }
    } else {
      $reply = @{ ok = $false; code = "outlook-unavailable"; error = $_.Exception.Message }
    }
  } finally {
    Clear-Run $runId
  }
  Write-HostMessage $reply
}

function Receive-Command($message) {
  $origin = Read-Field $message "origin"
  $expected = Get-ExpectedOrigin
  if ($expected -eq "" -or $origin -ne $expected) {
    Write-HostMessage @{ ok = $false; code = "origin"; error = "Caller origin is not allowed." }
    return
  }
  $command = Read-Field $message "command"
  switch ($command) {
    "probe" { Send-Probe; return }
    "compose.begin" { Start-Compose $message; return }
    "file.chunk" { Add-Chunk $message; return }
    "compose.commit" { Show-Compose $message; return }
    "cancel" { Clear-Run (Read-Field $message "runId"); Write-HostMessage @{ ok = $true; command = "cancel" }; return }
    default { throw "Command is not allowed." }
  }
}

$script:InputReader = New-Object IO.BinaryReader ([Console]::OpenStandardInput())
$script:OutputWriter = New-Object IO.BinaryWriter ([Console]::OpenStandardOutput())
while ($true) {
  try { $message = Read-HostMessage } catch { break }
  if ($null -eq $message) { break }
  try {
    Receive-Command $message
  } catch {
    $runId = Read-Field $message "runId"
    if (Test-RunId $runId) { Clear-Run $runId }
    Write-HostMessage @{ ok = $false; code = "schema"; error = $_.Exception.Message }
  }
}
