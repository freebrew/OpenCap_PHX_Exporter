if ($PSVersionTable.PSEdition -eq "Core") {
    & powershell.exe -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path
    exit $LASTEXITCODE
}
$ErrorActionPreference = "Stop"
. "z:\PHX_FieldCap\bas\VbaCompile.ps1"

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $null
foreach ($w in @($xl.Workbooks)) { if ($w.Name -like "*35780*" -or $w.Name -like "Slide Sheet*") { $wb = $w; break } }
Write-Host ("Workbook: {0}" -f $wb.Name)

$res = Test-VbaCompile -Excel $xl
Show-CompileResult $res $xl

if ($res.Ok) {
    Write-Host "`nRun-level check:"
    foreach ($p in @("EnsureDataBhaMirrorUnlocked")) {
        try {
            $xl.Run($p)
            Write-Host ("  {0}: OK" -f $p)
        } catch {
            Write-Host ("  {0}: FAIL {1}" -f $p, $_.Exception.Message)
        }
    }
}
