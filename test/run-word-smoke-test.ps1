param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'

$sampleDir = Join-Path $ProjectRoot 'test\sample-plot'
$supportDir = Join-Path $sampleDir 'support-extra'
$modulePath = Join-Path $ProjectRoot 'src\RawDataOleInserter.bas'
$dialogPath = Join-Path $ProjectRoot 'src\ImageSupportFilesDialog.frm'
$logPath = Join-Path $ProjectRoot 'test\word-smoke-test.log'

function Write-Step {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value $line
}

Set-Content -LiteralPath $logPath -Encoding UTF8 -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting smoke test"

Write-Step 'Preparing sample plot folder'
New-Item -ItemType Directory -Force -Path $sampleDir | Out-Null
New-Item -ItemType Directory -Force -Path $supportDir | Out-Null

Add-Type -AssemblyName System.Drawing
$bitmap = New-Object System.Drawing.Bitmap 640, 360
$bitmap.SetResolution(96, 96)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.Clear([System.Drawing.Color]::FromArgb(248, 244, 236))
$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(44, 88, 120)), 5
$graphics.DrawLine($pen, 60, 280, 580, 80)
$font = New-Object System.Drawing.Font 'Georgia', 28
$brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(50, 50, 46))
$graphics.DrawString('plot.png smoke test', $font, $brush, 110, 150)
$bitmap.Save((Join-Path $sampleDir 'plot.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$bitmap.Save((Join-Path $sampleDir 'display.jpg'), [System.Drawing.Imaging.ImageFormat]::Jpeg)
$graphics.Dispose()
$pen.Dispose()
$font.Dispose()
$brush.Dispose()
$bitmap.Dispose()

Set-Content -LiteralPath (Join-Path $sampleDir 'plot.py') -Encoding UTF8 -Value "print('sample plot source')"
Set-Content -LiteralPath (Join-Path $sampleDir 'data.csv') -Encoding UTF8 -Value "x,y`n1,2`n2,4"
Set-Content -LiteralPath (Join-Path $supportDir 'data.csv') -Encoding UTF8 -Value "x,y`n3,9`n4,16"

$word = $null
$doc = $null

try {
    Write-Step 'Starting Word.Application'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1

    Write-Step 'Creating document'
    $doc = $word.Documents.Add()
    Write-Step 'Importing VBA module'
    $doc.VBProject.VBComponents.Import($modulePath) | Out-Null
    Write-Step 'Importing support-files dialog'
    $doc.VBProject.VBComponents.Import($dialogPath) | Out-Null
    Write-Step 'Adding smoke-test runner macro'
    $runner = $doc.VBProject.VBComponents.Add(1)
    $escapedSampleDir = $sampleDir.Replace('"', '""')
    $escapedDisplayImage = (Join-Path $sampleDir 'display.jpg').Replace('"', '""')
    $escapedDataPath = (Join-Path $sampleDir 'data.csv').Replace('"', '""')
    $escapedDuplicateDataPath = (Join-Path $supportDir 'data.csv').Replace('"', '""')
    $escapedPlotPyPath = (Join-Path $sampleDir 'plot.py').Replace('"', '""')
    $escapedLogPath = $logPath.Replace('"', '""')
    $runner.CodeModule.AddFromString(@"
Public Sub SmokeTest()
    Dim supportFiles As Collection
    Dim existingPicture As InlineShape
    Dim existingFloatingPicture As Shape
    Dim floatingAnchor As Range

    On Error GoTo Failed
    InsertPlotFolder "$escapedSampleDir"
    Set supportFiles = New Collection
    supportFiles.Add "$escapedDataPath"
    supportFiles.Add "$escapedDuplicateDataPath"
    supportFiles.Add "$escapedPlotPyPath"
    InsertImageAndSupportFilesAsOle "$escapedDisplayImage", supportFiles
    Set existingPicture = Selection.InlineShapes.AddPicture("$escapedDisplayImage", False, True)
    existingPicture.Width = 180
    existingPicture.Height = 101.25
    Set supportFiles = New Collection
    supportFiles.Add "$escapedDataPath"
    supportFiles.Add "$escapedDuplicateDataPath"
    AttachSupportFilesToInlineImage existingPicture, supportFiles
    Set floatingAnchor = ActiveDocument.Range(ActiveDocument.Content.End - 1, ActiveDocument.Content.End - 1)
    Set existingFloatingPicture = ActiveDocument.Shapes.AddPicture("$escapedDisplayImage", False, True, 72, 144, 160, 90, floatingAnchor)
    existingFloatingPicture.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    existingFloatingPicture.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    existingFloatingPicture.Left = 72
    existingFloatingPicture.Top = 144
    existingFloatingPicture.WrapFormat.Type = wdWrapSquare
    Set supportFiles = New Collection
    supportFiles.Add "$escapedDataPath"
    AttachSupportFilesToFloatingImage existingFloatingPicture, supportFiles
    Open "$escapedLogPath" For Append As #1
    Print #1, Format`$(Now, "yyyy-mm-dd hh:nn:ss") & " VBA SmokeTest OK"
    Close #1
    Exit Sub
Failed:
    Open "$escapedLogPath" For Append As #1
    Print #1, Format`$(Now, "yyyy-mm-dd hh:nn:ss") & " VBA SmokeTest ERROR " & Err.Number & ": " & Err.Description
    Close #1
End Sub
"@)

    Write-Step 'Running InsertPlotFolder macro'
    $word.Run('SmokeTest')

    Write-Step 'Macro returned; checking document'
    $shapeCount = $doc.InlineShapes.Count
    if ($shapeCount -ne 3) {
        throw "Expected 3 inline shape(s), found $shapeCount"
    }
    $attachedShape = $doc.InlineShapes.Item(3)
    if ([math]::Abs($attachedShape.Width - 180) -gt 0.5 -or [math]::Abs($attachedShape.Height - 101.25) -gt 0.5) {
        throw "Expected attached image size 180 x 101.25 pt, found $([math]::Round($attachedShape.Width, 2)) x $([math]::Round($attachedShape.Height, 2)) pt"
    }
    $floatingShapeCount = $doc.Shapes.Count
    if ($floatingShapeCount -ne 1) {
        throw "Expected 1 floating shape, found $floatingShapeCount"
    }
    $attachedFloatingShape = $doc.Shapes.Item(1)
    if ([math]::Abs($attachedFloatingShape.Width - 160) -gt 0.5 -or [math]::Abs($attachedFloatingShape.Height - 90) -gt 0.5) {
        throw "Expected attached floating image size 160 x 90 pt, found $([math]::Round($attachedFloatingShape.Width, 2)) x $([math]::Round($attachedFloatingShape.Height, 2)) pt"
    }
    if ([math]::Abs($attachedFloatingShape.Left - 72) -gt 0.5 -or [math]::Abs($attachedFloatingShape.Top - 144) -gt 0.5) {
        throw "Expected attached floating image position 72 x 144 pt, found $([math]::Round($attachedFloatingShape.Left, 2)) x $([math]::Round($attachedFloatingShape.Top, 2)) pt"
    }

    [pscustomobject]@{
        Status = 'OK'
        SampleFolder = $sampleDir
        InlineShapes = $shapeCount
        AttachedWidth = [math]::Round($attachedShape.Width, 2)
        AttachedHeight = [math]::Round($attachedShape.Height, 2)
        FloatingShapes = $floatingShapeCount
        FloatingWidth = [math]::Round($attachedFloatingShape.Width, 2)
        FloatingHeight = [math]::Round($attachedFloatingShape.Height, 2)
        FloatingLeft = [math]::Round($attachedFloatingShape.Left, 2)
        FloatingTop = [math]::Round($attachedFloatingShape.Top, 2)
    } | Format-List

    Write-Step "Completed smoke test with $shapeCount inline shape(s) and $floatingShapeCount floating shape(s)"
}
finally {
    if ($doc -ne $null) {
        $doc.Close($false)
    }
    if ($word -ne $null) {
        $word.Quit()
    }
}
