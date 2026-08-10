# Safe VBA compile checker.
# Dot-source this file, then call Test-VbaCompile.
# It never uses SendKeys against the VBE editor: the compile-error dialog is
# located by window class, its text is read from the child controls, and the OK
# button is dismissed with BM_CLICK. Stray keystrokes previously leaked into the
# code pane and corrupted source lines.

Add-Type -Namespace Win32Dlg -Name Api -MemberDefinition @"
[DllImport("user32.dll")]
public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
[DllImport("user32.dll")]
public static extern bool EnumChildWindows(IntPtr hWnd, EnumWindowsProc lpEnumFunc, IntPtr lParam);
[DllImport("user32.dll", CharSet = CharSet.Auto)]
public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
[DllImport("user32.dll", CharSet = CharSet.Auto)]
public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
[DllImport("user32.dll")]
public static extern bool IsWindowVisible(IntPtr hWnd);
[DllImport("user32.dll", CharSet = CharSet.Auto)]
public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
[DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
"@ -ErrorAction SilentlyContinue

function Get-WindowClass([IntPtr]$hwnd) {
    $sb = New-Object System.Text.StringBuilder 256
    [void][Win32Dlg.Api]::GetClassName($hwnd, $sb, $sb.Capacity)
    $sb.ToString()
}

function Get-WindowText2([IntPtr]$hwnd) {
    $sb = New-Object System.Text.StringBuilder 1024
    [void][Win32Dlg.Api]::GetWindowText($hwnd, $sb, $sb.Capacity)
    $sb.ToString()
}

function Find-VbaDialog {
    $found = [IntPtr]::Zero
    $cb = [Win32Dlg.Api+EnumWindowsProc] {
        param($hwnd, $lparam)
        if (-not [Win32Dlg.Api]::IsWindowVisible($hwnd)) { return $true }
        if ((Get-WindowClass $hwnd) -ne "#32770") { return $true }
        $t = Get-WindowText2 $hwnd
        if ($t -like "*Microsoft Visual Basic*") {
            $script:__dlgHwnd = $hwnd
            return $false
        }
        return $true
    }
    $script:__dlgHwnd = [IntPtr]::Zero
    [void][Win32Dlg.Api]::EnumWindows($cb, [IntPtr]::Zero)
    return $script:__dlgHwnd
}

function Get-DialogInfo([IntPtr]$hwnd) {
    $script:__texts = @()
    $script:__okBtn = [IntPtr]::Zero
    $cb = [Win32Dlg.Api+EnumWindowsProc] {
        param($child, $lparam)
        $cls = Get-WindowClass $child
        $txt = Get-WindowText2 $child
        if ($cls -eq "Static" -and $txt.Trim()) { $script:__texts += $txt.Trim() }
        if ($cls -eq "Button" -and $txt -match "OK") {
            if ($script:__okBtn -eq [IntPtr]::Zero) { $script:__okBtn = $child }
        }
        return $true
    }
    [void][Win32Dlg.Api]::EnumChildWindows($hwnd, $cb, [IntPtr]::Zero)
    [pscustomobject]@{
        Text  = ($script:__texts -join " | ")
        OkBtn = $script:__okBtn
    }
}

function Close-Dialog([IntPtr]$hwnd, [IntPtr]$okBtn) {
    $BM_CLICK = 0x00F5
    $WM_CLOSE = 0x0010
    if ($okBtn -ne [IntPtr]::Zero) {
        [void][Win32Dlg.Api]::SendMessage($okBtn, $BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero)
    } else {
        [void][Win32Dlg.Api]::SendMessage($hwnd, $WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
    }
    Start-Sleep -Milliseconds 200
}

function Reset-VbaProject {
    param([Parameter(Mandatory = $true)] $Excel)
    # Leave break/run mode so modules become editable again.
    try {
        $ctl = $Excel.VBE.CommandBars.FindControl(1, 2571)
        if ($ctl -and $ctl.Enabled) { $ctl.Execute(); Start-Sleep -Milliseconds 250; return $true }
    } catch {}
    foreach ($bar in @($Excel.VBE.CommandBars)) {
        try {
            foreach ($ctl in @($bar.Controls)) {
                if ($ctl.Caption -match '^&?Reset$' -and $ctl.Enabled) {
                    $ctl.Execute()
                    Start-Sleep -Milliseconds 250
                    return $true
                }
            }
        } catch {}
    }
    return $false
}

function Test-VbaCompile {
    param(
        [Parameter(Mandatory = $true)] $Excel,
        [int]$TimeoutMs = 6000
    )

    # Dismiss any dialog left over from a previous run, then leave break mode.
    $stale = Find-VbaDialog
    if ($stale -ne [IntPtr]::Zero) {
        $si = Get-DialogInfo $stale
        Close-Dialog $stale $si.OkBtn
    }
    [void](Reset-VbaProject -Excel $Excel)

    # Compile via the VBE command bar control (id 578 = Debug > Compile).
    # This must run on this thread: a worker thread has no access to $Excel and
    # the compile would silently never happen, reporting a false pass.
    # A disabled Compile control means the project is already fully compiled.
    $invoked = $false
    $alreadyCompiled = $false
    try {
        $ctl = $Excel.VBE.CommandBars.FindControl(1, 578)
        if (-not $ctl.Enabled) {
            $alreadyCompiled = $true
            $invoked = $true
        } else {
            $ctl.Execute()
            $invoked = $true
        }
    } catch {
        Write-Host ("compile invoke failed: {0}" -f $_.Exception.Message)
    }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $hwnd = [IntPtr]::Zero
    if (-not $alreadyCompiled) {
        while ($sw.ElapsedMilliseconds -lt $TimeoutMs) {
            $hwnd = Find-VbaDialog
            if ($hwnd -ne [IntPtr]::Zero) { break }
            Start-Sleep -Milliseconds 150
        }
    }

    $result = [pscustomobject]@{
        Ok       = $true
        Invoked  = $invoked
        Message  = ""
        Module   = ""
        Line     = 0
        LineText = ""
        Token    = ""
    }

    if ($hwnd -ne [IntPtr]::Zero) {
        $info = Get-DialogInfo $hwnd
        $result.Ok = $false
        $result.Message = $info.Text
        Close-Dialog $hwnd $info.OkBtn
    }

    Start-Sleep -Milliseconds 200
    try {
        $pane = $Excel.VBE.ActiveCodePane
        if ($pane) {
            $m = $pane.CodeModule
            $sl = 0; $sc = 0; $el = 0; $ec = 0
            $pane.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
            $result.Module = $m.Parent.Name
            $result.Line = $sl
            if ($sl -ge 1 -and $sl -le $m.CountOfLines) {
                $result.LineText = $m.Lines($sl, 1)
                if ($el -eq $sl -and $ec -gt $sc) {
                    $len = [Math]::Min($ec - $sc, $result.LineText.Length - ($sc - 1))
                    if ($len -gt 0) {
                        $result.Token = $result.LineText.Substring($sc - 1, $len)
                    }
                }
            }
        }
    } catch {}

    return $result
}

function Show-CompileResult($res, $Excel) {
    if (-not $res.Invoked) {
        Write-Host "COMPILE: NOT INVOKED (could not execute Debug>Compile)" -ForegroundColor Yellow
        return
    }
    if ($res.Ok) {
        Write-Host "COMPILE: OK" -ForegroundColor Green
        return
    }
    Write-Host "COMPILE FAILED" -ForegroundColor Red
    Write-Host ("  Message: {0}" -f $res.Message)
    Write-Host ("  Module : {0}" -f $res.Module)
    Write-Host ("  Line   : {0}" -f $res.Line)
    Write-Host ("  Text   : {0}" -f $res.LineText)
    if ($res.Token) { Write-Host ("  Token  : >>{0}<<" -f $res.Token) }
    try {
        $m = $Excel.VBE.ActiveCodePane.CodeModule
        Write-Host "  Context:"
        for ($i = [Math]::Max(1, $res.Line - 6); $i -le [Math]::Min($m.CountOfLines, $res.Line + 6); $i++) {
            $mark = if ($i -eq $res.Line) { ">>" } else { "  " }
            Write-Host ("  {0}{1,5}|{2}" -f $mark, $i, $m.Lines($i, 1))
        }
    } catch {}
}
