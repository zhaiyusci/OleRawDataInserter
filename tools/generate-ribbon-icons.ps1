$ErrorActionPreference = 'Stop'

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$assetsDir = Join-Path $projectRoot 'assets'
New-Item -ItemType Directory -Force -Path $assetsDir | Out-Null

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

function New-RoundedRectPath {
    param(
        [float]$X,
        [float]$Y,
        [float]$Width,
        [float]$Height,
        [float]$Radius
    )

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $diameter = $Radius * 2
    $path.AddArc($X, $Y, $diameter, $diameter, 180, 90)
    $path.AddArc($X + $Width - $diameter, $Y, $diameter, $diameter, 270, 90)
    $path.AddArc($X + $Width - $diameter, $Y + $Height - $diameter, $diameter, $diameter, 0, 90)
    $path.AddArc($X, $Y + $Height - $diameter, $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return $path
}

function New-Brush([int]$A, [int]$R, [int]$G, [int]$B) {
    return New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb($A, $R, $G, $B))
}

function New-Pen([int]$A, [int]$R, [int]$G, [int]$B, [float]$Width = 2) {
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb($A, $R, $G, $B)), $Width
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    return $pen
}

function Draw-Panel {
    param($Graphics, [System.Drawing.Color]$Accent)
    $shadow = New-Brush 32 24 20 16
    $panel = New-Brush 255 255 250 240
    $border = New-Pen 255 62 49 35 2
    $accentBrush = New-Object System.Drawing.SolidBrush $Accent

    $Graphics.FillEllipse($shadow, 8, 9, 48, 48)
    $path = New-RoundedRectPath 8 7 48 48 12
    $Graphics.FillPath($panel, $path)
    $Graphics.DrawPath($border, $path)
    $Graphics.FillRectangle($accentBrush, 12, 11, 40, 7)

    $shadow.Dispose()
    $panel.Dispose()
    $border.Dispose()
    $accentBrush.Dispose()
    $path.Dispose()
}

function Draw-InsertFigurePackage {
    param($Graphics)
    Draw-Panel $Graphics ([System.Drawing.Color]::FromArgb(232, 154, 54))

    $folder = New-Brush 255 225 164 73
    $folderDark = New-Pen 255 123 82 28 2
    $paper = New-Brush 255 246 240 225
    $line = New-Pen 255 33 103 132 3
    $dot = New-Brush 255 33 103 132

    $Graphics.FillRectangle($folder, 16, 23, 31, 19)
    $Graphics.FillRectangle($folder, 18, 19, 13, 7)
    $Graphics.DrawRectangle($folderDark, 16, 23, 31, 19)
    $Graphics.FillRectangle($paper, 22, 28, 25, 19)
    $Graphics.DrawLine($line, 25, 41, 31, 35)
    $Graphics.DrawLine($line, 31, 35, 37, 38)
    $Graphics.DrawLine($line, 37, 38, 44, 30)
    $Graphics.FillEllipse($dot, 23, 39, 5, 5)
    $Graphics.FillEllipse($dot, 29, 33, 5, 5)
    $Graphics.FillEllipse($dot, 35, 36, 5, 5)
    $Graphics.FillEllipse($dot, 42, 28, 5, 5)

    $folder.Dispose()
    $folderDark.Dispose()
    $paper.Dispose()
    $line.Dispose()
    $dot.Dispose()
}

function Draw-ImageFiles {
    param($Graphics)
    Draw-Panel $Graphics ([System.Drawing.Color]::FromArgb(35, 132, 159))

    $framePen = New-Pen 255 42 82 94 2
    $sky = New-Brush 255 212 238 242
    $mountain = New-Brush 255 49 131 111
    $sun = New-Brush 255 232 177 65
    $clipPen = New-Pen 255 119 83 51 3

    $Graphics.FillRectangle($sky, 17, 18, 30, 25)
    $Graphics.DrawRectangle($framePen, 17, 18, 30, 25)
    $mountainPoints = @(
        [System.Drawing.PointF]::new(19, 41),
        [System.Drawing.PointF]::new(29, 30),
        [System.Drawing.PointF]::new(36, 37),
        [System.Drawing.PointF]::new(43, 27),
        [System.Drawing.PointF]::new(47, 41)
    )
    $Graphics.FillPolygon($mountain, $mountainPoints)
    $Graphics.FillEllipse($sun, 37, 21, 6, 6)
    $Graphics.DrawArc($clipPen, 42, 31, 11, 17, 70, 260)
    $Graphics.DrawArc($clipPen, 38, 27, 17, 24, 75, 255)

    $framePen.Dispose()
    $sky.Dispose()
    $mountain.Dispose()
    $sun.Dispose()
    $clipPen.Dispose()
}

function Draw-ManageFiles {
    param($Graphics)
    Draw-Panel $Graphics ([System.Drawing.Color]::FromArgb(64, 126, 92))

    $imagePen = New-Pen 255 42 72 57 2
    $imageFill = New-Brush 255 222 239 225
    $sliderPen = New-Pen 255 37 50 43 3
    $knob = New-Brush 255 232 154 54

    $Graphics.FillRectangle($imageFill, 15, 18, 25, 23)
    $Graphics.DrawRectangle($imagePen, 15, 18, 25, 23)
    $Graphics.DrawLine($imagePen, 18, 38, 25, 30)
    $Graphics.DrawLine($imagePen, 25, 30, 31, 35)
    $Graphics.DrawLine($imagePen, 31, 35, 38, 25)
    $Graphics.DrawLine($sliderPen, 44, 20, 52, 20)
    $Graphics.DrawLine($sliderPen, 44, 31, 52, 31)
    $Graphics.DrawLine($sliderPen, 44, 42, 52, 42)
    $Graphics.FillEllipse($knob, 46, 16, 8, 8)
    $Graphics.FillEllipse($knob, 41, 27, 8, 8)
    $Graphics.FillEllipse($knob, 48, 38, 8, 8)

    $imagePen.Dispose()
    $imageFill.Dispose()
    $sliderPen.Dispose()
    $knob.Dispose()
}

function Draw-UsageHelp {
    param($Graphics)
    Draw-Panel $Graphics ([System.Drawing.Color]::FromArgb(113, 97, 73))

    $page = New-Brush 255 255 246 226
    $pagePen = New-Pen 255 78 63 47 2
    $rulePen = New-Pen 255 158 138 102 1.5
    $qPen = New-Pen 255 34 112 139 4
    $qBrush = New-Brush 255 34 112 139

    $leftPath = New-RoundedRectPath 15 19 16 27 3
    $rightPath = New-RoundedRectPath 33 19 16 27 3
    $Graphics.FillPath($page, $leftPath)
    $Graphics.FillPath($page, $rightPath)
    $Graphics.DrawPath($pagePen, $leftPath)
    $Graphics.DrawPath($pagePen, $rightPath)
    $Graphics.DrawLine($pagePen, 32, 20, 32, 46)
    $Graphics.DrawLine($rulePen, 18, 27, 28, 27)
    $Graphics.DrawLine($rulePen, 18, 33, 28, 33)
    $Graphics.DrawLine($rulePen, 18, 39, 27, 39)
    $Graphics.DrawArc($qPen, 37, 24, 8, 8, 200, 260)
    $Graphics.DrawLine($qPen, 41, 32, 41, 36)
    $Graphics.FillEllipse($qBrush, 39, 40, 4, 4)

    $page.Dispose()
    $pagePen.Dispose()
    $rulePen.Dispose()
    $qPen.Dispose()
    $qBrush.Dispose()
    $leftPath.Dispose()
    $rightPath.Dispose()
}

function New-RibbonIcon {
    param(
        [string]$FileName,
        [scriptblock]$Draw
    )

    $bitmap = New-Object System.Drawing.Bitmap 64, 64
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $graphics.Clear([System.Drawing.Color]::Transparent)
    & $Draw $graphics
    $path = Join-Path $assetsDir $FileName
    $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
    return $path
}

$created = @()
$created += New-RibbonIcon 'insert-figure-package-icon.png' { param($g) Draw-InsertFigurePackage $g }
$created += New-RibbonIcon 'insert-image-files-icon.png' { param($g) Draw-ImageFiles $g }
$created += New-RibbonIcon 'manage-image-ole-files-icon.png' { param($g) Draw-ManageFiles $g }
$created += New-RibbonIcon 'usage-help-icon.png' { param($g) Draw-UsageHelp $g }

$created | ForEach-Object { "Generated: $_" }
