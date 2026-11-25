# Directional-well mark for the extension toolbar, manifest icons, and Chrome Web Store assets.
# Renders at 1024x1024 and downscales with HighQualityBicubic for crisp edges.
# Run from anywhere: powershell -NoProfile -File src/extension/scripts/make-well-icon.ps1

Add-Type -AssemblyName System.Drawing

function New-RoundedRectPath([single]$X, [single]$Y, [single]$W, [single]$H, [single]$R) {
    $d = $R * 2
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddArc($X, $Y, $d, $d, 180, 90)
    $p.AddArc($X + $W - $d, $Y, $d, $d, 270, 90)
    $p.AddArc($X + $W - $d, $Y + $H - $d, $d, $d, 0, 90)
    $p.AddArc($X, $Y + $H - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function Draw-WellMark([System.Drawing.Graphics]$g, [single]$X, [single]$Y, [single]$Size, [bool]$AddShadow) {
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

    $radius = [single]([Math]::Max(2, $Size * 0.22))

    if ($AddShadow) {
        $shOffY = [single]($Size * 0.03)
        $shPath = New-RoundedRectPath ($X + $Size * 0.01) ($Y + $shOffY) $Size $Size $radius
        $shBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(72, 2, 10, 18))
        $g.FillPath($shBrush, $shPath)
        $shBrush.Dispose()
        $shPath.Dispose()
    }

    $path = New-RoundedRectPath $X $Y $Size $Size $radius
    $rect = New-Object System.Drawing.RectangleF($X, $Y, $Size, $Size)
    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        [System.Drawing.Color]::FromArgb(255, 22, 56, 84),
        [System.Drawing.Color]::FromArgb(255, 11, 30, 48),
        45.0
    )
    $g.FillPath($bg, $path)

    # Subtle inner rim so the dark tile reads cleanly on both dark and light store backgrounds.
    $rimW = [single]([Math]::Max(1.0, $Size * 0.018))
    $rim = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(80, 148, 226, 218), $rimW)
    $rimPath = New-RoundedRectPath ($X + $rimW / 2) ($Y + $rimW / 2) ($Size - $rimW) ($Size - $rimW) ($radius - $rimW / 2)
    $g.DrawPath($rim, $rimPath)
    $rimPath.Dispose()
    $rim.Dispose()

    $x0 = $X + $Size * 0.34
    $y0 = $Y + $Size * 0.22
    $y1 = $Y + $Size * 0.46
    $x2 = $X + $Size * 0.68
    $y2 = $Y + $Size * 0.70
    $pts = [System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF([single]$x0, [single]$y0)),
        (New-Object System.Drawing.PointF([single]$x0, [single]$y1)),
        (New-Object System.Drawing.PointF([single](($x0 + $x2) / 2), [single]($y1 + $Size * 0.16))),
        (New-Object System.Drawing.PointF([single]$x2, [single]$y2))
    )

    # Soft glow behind wellbore path.
    $glowPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(48, 94, 234, 212), [single]([Math]::Max(2.0, $Size * 0.14)))
    $glowPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $glowPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $glowPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $g.DrawCurve($glowPen, $pts, 0.6)
    $glowPen.Dispose()

    # Main wellbore path.
    $penW = [single]([Math]::Max(1.6, $Size * 0.088))
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 94, 234, 212), $penW)
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $g.DrawCurve($pen, $pts, 0.6)

    # Surface / rig floor tick.
    $tickW = [single]([Math]::Max(1.0, $Size * 0.044))
    $tick = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 203, 224, 236), $tickW)
    $tick.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $tick.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $g.DrawLine($tick, [single]($x0 - $Size * 0.12), [single]$y0, [single]($x0 + $Size * 0.12), [single]$y0)

    # Bit / target dot.
    $bit = [single]([Math]::Max(2.0, $Size * 0.145))
    $bitGlow = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(65, 248, 250, 252))
    $g.FillEllipse($bitGlow, [single]($x2 - $bit * 0.7), [single]($y2 - $bit * 0.7), [single]($bit * 1.4), [single]($bit * 1.4))
    $bitGlow.Dispose()

    $bitBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
    $g.FillEllipse($bitBrush, [single]($x2 - $bit / 2), [single]($y2 - $bit / 2), [single]$bit, [single]$bit)

    $path.Dispose()
    $bg.Dispose()
    $pen.Dispose()
    $tick.Dispose()
    $bitBrush.Dispose()
}

function Save-Resized([System.Drawing.Bitmap]$Master, [int]$TargetSize, [string]$OutPath) {
    $bmp = New-Object System.Drawing.Bitmap($TargetSize, $TargetSize, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.DrawImage($Master, 0, 0, $TargetSize, $TargetSize)
    $g.Dispose()
    $bmp.Save($OutPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Output $OutPath
}

$extRoot = Split-Path $PSScriptRoot -Parent
$dir = Join-Path $extRoot "public\icons"
$storeDir = Join-Path $extRoot "store-assets"
New-Item -ItemType Directory -Force -Path $dir, $storeDir | Out-Null

# 1) Full-bleed 1024x1024 master for extension toolbar/manifest icons + upscaled promo asset.
$master = New-Object System.Drawing.Bitmap(1024, 1024, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$gm = [System.Drawing.Graphics]::FromImage($master)
$gm.Clear([System.Drawing.Color]::Transparent)
Draw-WellMark $gm 0 0 1024 $false
$gm.Dispose()

foreach ($size in 16, 32, 48, 128) {
    Save-Resized $master $size (Join-Path $dir "icon$size.png")
}
Save-Resized $master 512 (Join-Path $storeDir "icon-512.png")
# 2) Chrome Web Store icon: full-bleed 128x128 PNG matching sibling items in the publisher dashboard.
Save-Resized $master 128 (Join-Path $storeDir "store-icon-128.png")
$master.Dispose()
