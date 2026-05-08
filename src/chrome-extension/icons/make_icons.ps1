# Generates simple PNG icons from an SVG using .NET Drawing
# Run once: pwsh -File make_icons.ps1

Add-Type -AssemblyName System.Drawing

function Make-Icon {
    param([int]$Size, [string]$OutPath)

    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode    = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    $radius = [int]($Size * 0.2)
    $rect   = New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)
    $path   = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($rect.X, $rect.Y, $radius * 2, $radius * 2, 180, 90)
    $path.AddArc($rect.Right - $radius * 2, $rect.Y, $radius * 2, $radius * 2, 270, 90)
    $path.AddArc($rect.Right - $radius * 2, $rect.Bottom - $radius * 2, $radius * 2, $radius * 2, 0, 90)
    $path.AddArc($rect.X, $rect.Bottom - $radius * 2, $radius * 2, $radius * 2, 90, 90)
    $path.CloseFigure()

    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(37, 99, 235))
    $g.FillPath($brush, $path)

    $fontSize = [float]($Size * 0.45)
    $font  = New-Object System.Drawing.Font("Arial", $fontSize, [System.Drawing.FontStyle]::Bold)
    $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $fmt   = New-Object System.Drawing.StringFormat
    $fmt.Alignment = [System.Drawing.StringAlignment]::Center
    $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center

    $textRect = New-Object System.Drawing.RectangleF(0, 0, $Size, $Size)
    $g.DrawString("FC", $font, $white, $textRect, $fmt)

    $bmp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Created: $OutPath"
}

$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Make-Icon -Size 16  -OutPath "$dir\icon16.png"
Make-Icon -Size 48  -OutPath "$dir\icon48.png"
Make-Icon -Size 128 -OutPath "$dir\icon128.png"
Write-Host "Icons generated."
