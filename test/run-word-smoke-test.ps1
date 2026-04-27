param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'

$sampleDir = Join-Path $ProjectRoot 'test\sample-plot'
$modulePath = Join-Path $ProjectRoot 'src\RawDataOleInserter.bas'
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
$graphics.Dispose()
$pen.Dispose()
$font.Dispose()
$brush.Dispose()
$bitmap.Dispose()

Set-Content -LiteralPath (Join-Path $sampleDir 'plot.py') -Encoding UTF8 -Value "print('sample plot source')"
Set-Content -LiteralPath (Join-Path $sampleDir 'data.csv') -Encoding UTF8 -Value "x,y`n1,2`n2,4"

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
    Write-Step 'Adding smoke-test runner macro'
    $runner = $doc.VBProject.VBComponents.Add(1)
    $escapedSampleDir = $sampleDir.Replace('"', '""')
    $escapedLogPath = $logPath.Replace('"', '""')
    $runner.CodeModule.AddFromString(@"
Public Sub SmokeTest()
    On Error GoTo Failed
    InsertPlotFolder "$escapedSampleDir"
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

    [pscustomobject]@{
        Status = 'OK'
        SampleFolder = $sampleDir
        InlineShapes = $shapeCount
    } | Format-List

    Write-Step "Completed smoke test with $shapeCount inline shape(s)"
}
finally {
    if ($doc -ne $null) {
        $doc.Close($false)
    }
    if ($word -ne $null) {
        $word.Quit()
    }
}
