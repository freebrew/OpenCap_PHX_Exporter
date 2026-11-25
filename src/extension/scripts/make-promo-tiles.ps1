# Generate Chrome Web Store Small Promo Tile (440x280) and Marquee Promo Tile (1400x560)
# as high-quality JPEGs (24-bit RGB, no alpha) using the upscaled icon and real Demo screenshots.
# Run: powershell -NoProfile -File src/extension/scripts/make-promo-tiles.ps1

Add-Type -AssemblyName System.Drawing

function New-RoundedRect([single]$X, [single]$Y, [single]$W, [single]$H, [single]$R) {
    $d = $R * 2
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddArc($X, $Y, $d, $d, 180, 90)
    $p.AddArc($X + $W - $d, $Y, $d, $d, 270, 90)
    $p.AddArc($X + $W - $d, $Y + $H - $d, $d, $d, 0, 90)
    $p.AddArc($X, $Y + $H - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function Save-Jpeg([System.Drawing.Bitmap]$Bmp, [string]$OutPath, [long]$Quality = 95) {
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" } | Select-Object -First 1
    $params = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, $Quality)
    $Bmp.Save($OutPath, $codec, $params)
    $params.Dispose()
    Write-Output $OutPath
}

function Draw-GridBackdrop([System.Drawing.Graphics]$g, [int]$W, [int]$H) {
    $rect = New-Object System.Drawing.Rectangle(0, 0, $W, $H)
    $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        [System.Drawing.Color]::FromArgb(255, 10, 25, 41),
        [System.Drawing.Color]::FromArgb(255, 16, 48, 66),
        22.0
    )
    $g.FillRectangle($bg, $rect)
    $bg.Dispose()

    # Subtle radial teal glow top-right.
    $glowPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $glowPath.AddEllipse([int]($W * 0.45), [int](-$H * 0.35), [int]($W * 0.8), [int]($H * 1.4))
    $pg = New-Object System.Drawing.Drawing2D.PathGradientBrush($glowPath)
    $pg.CenterColor = [System.Drawing.Color]::FromArgb(52, 45, 212, 191)
    $pg.SurroundColors = [System.Drawing.Color[]]@([System.Drawing.Color]::FromArgb(0, 16, 48, 66))
    $g.FillPath($pg, $glowPath)
    $pg.Dispose()
    $glowPath.Dispose()

    # Subtle engineering grid.
    $gridPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(18, 148, 226, 218), 1.0)
    $step = 32
    for ($x = $step; $x -lt $W; $x += $step) { $g.DrawLine($gridPen, $x, 0, $x, $H) }
    for ($y = $step; $y -lt $H; $y += $step) { $g.DrawLine($gridPen, 0, $y, $W, $y) }
    $gridPen.Dispose()
}

$extRoot = Split-Path $PSScriptRoot -Parent
$storeDir = Join-Path $extRoot "store-assets"
$iconPath = Join-Path $storeDir "icon-512.png"
$shotDataPath = Join-Path $storeDir "screenshot-3-slidesheet.jpg"
$shot3dPath = Join-Path $storeDir "screenshot-4-wells-3d.jpg"

$iconImg = [System.Drawing.Image]::FromFile($iconPath)
$shotDataImg = [System.Drawing.Image]::FromFile($shotDataPath)
$shot3dImg = [System.Drawing.Image]::FromFile($shot3dPath)

try {
    # ============================================================
    # 1) Small Promo Tile: 440 x 280 (no alpha)
    # ============================================================
    $small = New-Object System.Drawing.Bitmap(440, 280, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $gs = [System.Drawing.Graphics]::FromImage($small)
    $gs.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $gs.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $gs.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $gs.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    Draw-GridBackdrop $gs 440 280

    # Upscaled directional-well icon on top center-left.
    $gs.DrawImage($iconImg, 32, 34, 84, 84)

    $fEyebrow = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $fTitle = New-Object System.Drawing.Font("Segoe UI", 22, [System.Drawing.FontStyle]::Bold)
    $fSub = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Regular)
    $fPill = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

    $bTeal = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 94, 234, 212))
    $bWhite = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
    $bMuted = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 198, 218, 232))

    $gs.DrawString("DIRECTIONAL DRILLING", $fEyebrow, $bTeal, 132, 42)
    $gs.DrawString("OpenCap", $fTitle, $bWhite, 130, 58)
    $gs.DrawString("Slidesheet", $fTitle, $bWhite, 130, 88)

    $subRect = New-Object System.Drawing.RectangleF(32, 140, 376, 62)
    $gs.DrawString("Full directional workbook in Chrome with FieldCap sync, proximity gauges, 3D well paths & daily reports.", $fSub, $bMuted, $subRect)

    # Feature pills along bottom.
    $pills = @("FieldCap Sync", "3D Wells", "Daily Reports")
    $px = 32
    foreach ($label in $pills) {
        $sz = $gs.MeasureString($label, $fPill)
        $pw = [single]($sz.Width + 18)
        $ph = [single]26
        $pp = New-RoundedRect $px 224 $pw $ph 13
        $pb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(46, 94, 234, 212))
        $pPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(120, 94, 234, 212), 1.0)
        $gs.FillPath($pb, $pp)
        $gs.DrawPath($pPen, $pp)
        $gs.DrawString($label, $fPill, $bTeal, ($px + 9), 229)
        $pb.Dispose(); $pPen.Dispose(); $pp.Dispose()
        $px += [int]$pw + 8
    }

    $fEyebrow.Dispose(); $fTitle.Dispose(); $fSub.Dispose(); $fPill.Dispose()
    $bTeal.Dispose(); $bWhite.Dispose(); $bMuted.Dispose()
    $gs.Dispose()
    Save-Jpeg $small (Join-Path $storeDir "small-promo-440x280.jpg") 95
    $small.Dispose()

    # ============================================================
    # 2) Marquee Promo Tile: 1400 x 560 (no alpha)
    # ============================================================
    $marq = New-Object System.Drawing.Bitmap(1400, 560, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $gm = [System.Drawing.Graphics]::FromImage($marq)
    $gm.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $gm.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $gm.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $gm.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    Draw-GridBackdrop $gm 1400 560

    # Left column: upscaled icon + title + benefits.
    $gm.DrawImage($iconImg, 64, 62, 108, 108)

    $fmEye = New-Object System.Drawing.Font("Segoe UI", 12.5, [System.Drawing.FontStyle]::Bold)
    $fmTitle = New-Object System.Drawing.Font("Segoe UI", 31, [System.Drawing.FontStyle]::Bold)
    $fmSub = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Regular)
    $fmBul = New-Object System.Drawing.Font("Segoe UI", 14.5, [System.Drawing.FontStyle]::Bold)

    $bTeal2 = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 94, 234, 212))
    $bWhite2 = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 248, 250, 252))
    $bMuted2 = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 203, 222, 236))

    $gm.DrawString("DIRECTIONAL DRILLING WORKBOOK", $fmEye, $bTeal2, 188, 78)
    $gm.DrawString("OpenCap Slidesheet", $fmTitle, $bWhite2, 184, 102)

    $descRect = New-Object System.Drawing.RectangleF(64, 200, 580, 96)
    $gm.DrawString("Run your directional slide sheet right in Chrome — surveys, BHA hours, proximity gauges, 3D well trajectories, and daily reports in one workspace.", $fmSub, $bMuted2, $descRect)

    $bullets = @(
        "Live FieldCap sync & local Excel handoff",
        "Interactive 3D Wells & anti-collision review",
        "One-click morning reports & Outlook drafts"
    )
    $by = 336
    foreach ($b in $bullets) {
        $dot = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 94, 234, 212))
        $gm.FillEllipse($dot, 66, $by + 8, 10, 10)
        $dot.Dispose()
        $gm.DrawString($b, $fmBul, $bWhite2, 88, $by)
        $by += 44
    }

    # Right column: layered real screenshots (Slidesheet behind, 3D Wells forward).
    $card1Rect = New-Object System.Drawing.Rectangle(700, 48, 540, 338)
    $sh1 = New-RoundedRect ($card1Rect.X + 8) ($card1Rect.Y + 12) $card1Rect.Width $card1Rect.Height 14
    $shBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(110, 2, 8, 16))
    $gm.FillPath($shBrush, $sh1)
    $sh1.Dispose()

    $clip1 = New-RoundedRect $card1Rect.X $card1Rect.Y $card1Rect.Width $card1Rect.Height 14
    $gm.SetClip($clip1)
    $gm.DrawImage($shotDataImg, $card1Rect)
    $gm.ResetClip()
    $borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(140, 94, 234, 212), 2.0)
    $gm.DrawPath($borderPen, $clip1)
    $clip1.Dispose()

    # Forward card: 3D Wells screenshot.
    $card2Rect = New-Object System.Drawing.Rectangle(815, 192, 540, 338)
    $sh2 = New-RoundedRect ($card2Rect.X + 8) ($card2Rect.Y + 12) $card2Rect.Width $card2Rect.Height 14
    $gm.FillPath($shBrush, $sh2)
    $sh2.Dispose()

    $clip2 = New-RoundedRect $card2Rect.X $card2Rect.Y $card2Rect.Width $card2Rect.Height 14
    $gm.SetClip($clip2)
    $gm.DrawImage($shot3dImg, $card2Rect)
    $gm.ResetClip()
    $gm.DrawPath($borderPen, $clip2)
    $clip2.Dispose()

    $shBrush.Dispose(); $borderPen.Dispose()
    $fmEye.Dispose(); $fmTitle.Dispose(); $fmSub.Dispose(); $fmBul.Dispose()
    $bTeal2.Dispose(); $bWhite2.Dispose(); $bMuted2.Dispose()
    $gm.Dispose()
    Save-Jpeg $marq (Join-Path $storeDir "marquee-1400x560.jpg") 95
    $marq.Dispose()
}
finally {
    $iconImg.Dispose()
    $shotDataImg.Dispose()
    $shot3dImg.Dispose()
}
