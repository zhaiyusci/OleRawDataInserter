param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'

$sampleDir = Join-Path $ProjectRoot 'test\sample-plot'
$supportDir = Join-Path $sampleDir 'support-extra'
$supportFolderDir = Join-Path $sampleDir 'support-folder'
$supportNestedDir = Join-Path $supportFolderDir 'nested'
$managedWorkspaceDir = Join-Path $sampleDir 'managed-workspace'
$managedWorkspaceNestedDir = Join-Path $managedWorkspaceDir 'support-folder\nested'
$packageWorkspaceDir = Join-Path $sampleDir 'package-workspace'
$packageWorkspaceNestedDir = Join-Path $packageWorkspaceDir 'nested'
$unicodeWorkspaceName = -join ([char[]](0x4E2D, 0x6587, 0x8D44, 0x6599))
$unicodeNestedName = (-join ([char[]](0x6570, 0x636E))) + ' ' + (-join ([char[]](0x6587, 0x4EF6, 0x5939)))
$unicodeRootFileName = (-join ([char[]](0x7ED3, 0x679C))) + ' ' + (-join ([char[]](0x6570, 0x636E))) + '.csv'
$unicodeNestedFileName = (-join ([char[]](0x8BF4, 0x660E))) + '.txt'
$unicodeWorkspaceDir = Join-Path $sampleDir $unicodeWorkspaceName
$unicodeWorkspaceNestedDir = Join-Path $unicodeWorkspaceDir $unicodeNestedName
$unicodeExtractDir = Join-Path $sampleDir 'unicode-extract'
$modulePath = Join-Path $ProjectRoot 'src\RawDataOleInserter.bas'
$dialogPath = Join-Path $ProjectRoot 'src\ImageSupportFilesDialog.frm'
$usageDialogPath = Join-Path $ProjectRoot 'src\UsageHelpDialog.frm'
$zipToolPath = Join-Path $ProjectRoot 'dist\FigurePackageZipTool.exe'
$logPath = Join-Path $ProjectRoot 'test\word-smoke-test.log'

function Write-Step {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Encoding UTF8 -Value $line
}

Set-Content -LiteralPath $logPath -Encoding UTF8 -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Starting smoke test"

if (-not (Test-Path -LiteralPath $zipToolPath)) {
    throw "Missing zip helper: $zipToolPath"
}
$env:OLE_RAW_DATA_ZIP_TOOL = $zipToolPath

Write-Step 'Preparing sample plot folder'
New-Item -ItemType Directory -Force -Path $sampleDir | Out-Null
New-Item -ItemType Directory -Force -Path $supportDir | Out-Null
New-Item -ItemType Directory -Force -Path $supportNestedDir | Out-Null
New-Item -ItemType Directory -Force -Path $managedWorkspaceNestedDir | Out-Null
New-Item -ItemType Directory -Force -Path $packageWorkspaceNestedDir | Out-Null
New-Item -ItemType Directory -Force -Path $unicodeWorkspaceNestedDir | Out-Null

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
Set-Content -LiteralPath (Join-Path $sampleDir 'notes.txt') -Encoding UTF8 -Value "managed OLE support note"
Set-Content -LiteralPath (Join-Path $supportDir 'data.csv') -Encoding UTF8 -Value "x,y`n3,9`n4,16"
Set-Content -LiteralPath (Join-Path $supportNestedDir 'metadata.json') -Encoding UTF8 -Value '{"source":"folder support"}'
Set-Content -LiteralPath (Join-Path $managedWorkspaceDir 'plot.py') -Encoding UTF8 -Value "print('managed workspace source')"
Set-Content -LiteralPath (Join-Path $managedWorkspaceDir 'notes.txt') -Encoding UTF8 -Value "managed workspace support note"
Set-Content -LiteralPath (Join-Path $managedWorkspaceNestedDir 'metadata.json') -Encoding UTF8 -Value '{"source":"managed workspace"}'
Set-Content -LiteralPath (Join-Path $packageWorkspaceDir 'data.csv') -Encoding UTF8 -Value "x,y`n5,25`n6,36"
Set-Content -LiteralPath (Join-Path $packageWorkspaceNestedDir 'metadata.json') -Encoding UTF8 -Value '{"source":"package workspace"}'
Set-Content -LiteralPath (Join-Path $unicodeWorkspaceDir $unicodeRootFileName) -Encoding UTF8 -Value "x,y`n7,49`n8,64"
Set-Content -LiteralPath (Join-Path $unicodeWorkspaceNestedDir $unicodeNestedFileName) -Encoding UTF8 -Value "unicode filename smoke test"

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
    Write-Step 'Importing usage help dialog'
    $doc.VBProject.VBComponents.Import($usageDialogPath) | Out-Null
    Write-Step 'Adding smoke-test runner macro'
    $runner = $doc.VBProject.VBComponents.Add(1)
    $escapedSampleDir = $sampleDir.Replace('"', '""')
    $escapedDisplayImage = (Join-Path $sampleDir 'display.jpg').Replace('"', '""')
    $escapedDataPath = (Join-Path $sampleDir 'data.csv').Replace('"', '""')
    $escapedDuplicateDataPath = (Join-Path $supportDir 'data.csv').Replace('"', '""')
    $escapedSupportFolderDir = $supportFolderDir.Replace('"', '""')
    $escapedManagedWorkspaceDir = $managedWorkspaceDir.Replace('"', '""')
    $escapedPackageWorkspaceDir = $packageWorkspaceDir.Replace('"', '""')
    $escapedUnicodeWorkspaceDir = $unicodeWorkspaceDir.Replace('"', '""')
    $escapedPlotPyPath = (Join-Path $sampleDir 'plot.py').Replace('"', '""')
    $escapedNotesPath = (Join-Path $sampleDir 'notes.txt').Replace('"', '""')
    $escapedLogPath = $logPath.Replace('"', '""')
    $runner.CodeModule.AddFromString(@"
Public Sub SmokeTest()
    Dim supportFiles As Collection
    Dim keepEntries As Collection
    Dim initialZipPath As String
    Dim managedZipPath As String
    Dim managedEntries As Collection
    Dim unicodeZipPath As String
    Dim unicodeEntries As Collection
    Dim unicodeExtractDir As String
    Dim unicodeRootEntry As String
    Dim unicodeNestedName As String
    Dim unicodeNestedFileName As String
    Dim unicodeNestedEntry As String
    Dim firstOle As InlineShape
    Dim managedInline As InlineShape
    Dim existingPicture As InlineShape
    Dim linkedPicture As InlineShape
    Dim existingFloatingPicture As Shape
    Dim floatingAnchor As Range
    Dim supportDialog As ImageSupportFilesDialog
    Dim fso As Object

    On Error GoTo Failed
    LogSmokeStep "$escapedLogPath", "VBA step: dialog check"
    Set supportDialog = New ImageSupportFilesDialog
    If supportDialog.Controls("cmdAddFiles").Caption <> "Open Folder..." Then Err.Raise vbObjectError + 907, "SmokeTest", "Support dialog did not switch to system folder mode"
    Unload supportDialog
    LogSmokeStep "$escapedLogPath", "VBA step: insert plot folder"
    InsertPlotFolder "$escapedSampleDir"
    initialZipPath = ExtractZipFromInlineOleObject(ActiveDocument.InlineShapes.Item(1))
    AssertNoRawRootZipEntry initialZipPath
    LogSmokeStep "$escapedLogPath", "VBA step: manage existing ole"
    Set firstOle = ActiveDocument.InlineShapes.Item(1)
    LogSmokeStep "$escapedLogPath", "VBA step: call manage existing ole"
    ManageWorkingFolderInInlineOle firstOle, "$escapedManagedWorkspaceDir", ExtractDisplayImageFromInlineOleObject(firstOle)
    WaitForSmokeSeconds 2
    LogSmokeStep "$escapedLogPath", "VBA step: extract managed ole zip"
    Set managedInline = GetSmokeInlineShape(1)
    LogSmokeStep "$escapedLogPath", "VBA step: call extract managed ole zip"
    managedZipPath = ExtractZipFromInlineOleObject(managedInline)
    Set managedEntries = GetZipEntryNames(managedZipPath)
    If Not CollectionContainsText(managedEntries, "plot.py") Then Err.Raise vbObjectError + 900, "SmokeTest", "Managed OLE zip lost plot.py"
    If Not CollectionContainsText(managedEntries, "notes.txt") Then Err.Raise vbObjectError + 901, "SmokeTest", "Managed OLE zip did not add notes.txt"
    If CollectionContainsText(managedEntries, "data.csv") Then Err.Raise vbObjectError + 902, "SmokeTest", "Managed OLE zip did not remove data.csv"
    If Not CollectionContainsText(managedEntries, "support-folder/nested/metadata.json") Then Err.Raise vbObjectError + 905, "SmokeTest", "Managed OLE zip did not add folder contents"
    LogSmokeStep "$escapedLogPath", "VBA step: insert image package"
    InsertImageAndWorkingFolderAsOle "$escapedDisplayImage", "$escapedPackageWorkspaceDir"
    Set managedEntries = GetZipEntryNames(ExtractZipFromInlineOleObject(ActiveDocument.InlineShapes.Item(2)))
    If Not CollectionContainsText(managedEntries, "nested/metadata.json") Then Err.Raise vbObjectError + 906, "SmokeTest", "Image + folder zip did not add workspace folder contents"
    LogSmokeStep "$escapedLogPath", "VBA step: insert unicode package"
    InsertImageAndWorkingFolderAsOle "$escapedDisplayImage", "$escapedUnicodeWorkspaceDir"
    unicodeZipPath = ExtractZipFromInlineOleObject(ActiveDocument.InlineShapes.Item(3))
    unicodeRootEntry = SmokeUnicodeChar(&H7ED3) & SmokeUnicodeChar(&H679C) & " " & SmokeUnicodeChar(&H6570) & SmokeUnicodeChar(&H636E) & ".csv"
    unicodeNestedName = SmokeUnicodeChar(&H6570) & SmokeUnicodeChar(&H636E) & " " & SmokeUnicodeChar(&H6587) & SmokeUnicodeChar(&H4EF6) & SmokeUnicodeChar(&H5939)
    unicodeNestedFileName = SmokeUnicodeChar(&H8BF4) & SmokeUnicodeChar(&H660E) & ".txt"
    unicodeNestedEntry = unicodeNestedName & "/" & unicodeNestedFileName
    LogSmokeStep "$escapedLogPath", "VBA step: list unicode package"
    Set unicodeEntries = GetZipEntryNames(unicodeZipPath)
    If Not CollectionContainsText(unicodeEntries, unicodeRootEntry) Then Err.Raise vbObjectError + 908, "SmokeTest", "UTF-8 zip did not preserve Chinese root file name"
    If Not CollectionContainsText(unicodeEntries, unicodeNestedEntry) Then Err.Raise vbObjectError + 909, "SmokeTest", "UTF-8 zip did not preserve Chinese nested file name"
    Set fso = CreateObject("Scripting.FileSystemObject")
    unicodeExtractDir = fso.BuildPath(fso.GetSpecialFolder(2), "OleRawDataInserterUnicodeExtract_" & Format(Now, "yyyymmdd_hhnnss"))
    LogSmokeStep "$escapedLogPath", "VBA step: extract unicode package"
    ExtractZipToWorkingFolder unicodeZipPath, unicodeExtractDir
    If Not fso.FileExists(fso.BuildPath(unicodeExtractDir, unicodeRootEntry)) Then Err.Raise vbObjectError + 910, "SmokeTest", "UTF-8 zip did not extract Chinese root file name"
    If Not fso.FileExists(fso.BuildPath(fso.BuildPath(unicodeExtractDir, unicodeNestedName), unicodeNestedFileName)) Then Err.Raise vbObjectError + 911, "SmokeTest", "UTF-8 zip did not extract Chinese nested file name"
    On Error Resume Next
    fso.DeleteFolder unicodeExtractDir, True
    On Error GoTo Failed
    LogSmokeStep "$escapedLogPath", "VBA step: attach inline picture"
    Selection.EndKey Unit:=wdStory
    Set existingPicture = Selection.InlineShapes.AddPicture("$escapedDisplayImage", False, True)
    existingPicture.Width = 180
    existingPicture.Height = 101.25
    AttachWorkingFolderToInlineImage existingPicture, "$escapedPackageWorkspaceDir"
    LogSmokeStep "$escapedLogPath", "VBA step: attach linked picture"
    Selection.EndKey Unit:=wdStory
    Set linkedPicture = Selection.InlineShapes.AddPicture("$escapedDisplayImage", True, False)
    linkedPicture.Width = 120
    linkedPicture.Height = 67.5
    AttachWorkingFolderToInlineImage linkedPicture, "$escapedPackageWorkspaceDir"
    LogSmokeStep "$escapedLogPath", "VBA step: attach floating picture"
    Set floatingAnchor = ActiveDocument.Range(ActiveDocument.Content.End - 1, ActiveDocument.Content.End - 1)
    Set existingFloatingPicture = ActiveDocument.Shapes.AddPicture("$escapedDisplayImage", False, True, 72, 144, 160, 90, floatingAnchor)
    existingFloatingPicture.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    existingFloatingPicture.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    existingFloatingPicture.Left = 72
    existingFloatingPicture.Top = 144
    existingFloatingPicture.WrapFormat.Type = wdWrapSquare
    AttachWorkingFolderToFloatingImage existingFloatingPicture, "$escapedPackageWorkspaceDir"
    Open "$escapedLogPath" For Append As #1
    Print #1, Format`$(Now, "yyyy-mm-dd hh:nn:ss") & " VBA SmokeTest OK"
    Close #1
    Exit Sub
Failed:
    Open "$escapedLogPath" For Append As #1
    Print #1, Format`$(Now, "yyyy-mm-dd hh:nn:ss") & " VBA SmokeTest ERROR " & Err.Number & " [" & Err.Source & "]: " & Err.Description
    Close #1
End Sub

Private Sub LogSmokeStep(ByVal logPath As String, ByVal message As String)
    Open logPath For Append As #1
    Print #1, Format`$(Now, "yyyy-mm-dd hh:nn:ss") & " " & message
    Close #1
End Sub

Private Sub WaitForSmokeSeconds(ByVal seconds As Double)
    Dim startTime As Double

    startTime = Timer
    Do
        DoEvents
    Loop While SmokeElapsedSeconds(startTime) < seconds
End Sub

Private Function SmokeElapsedSeconds(ByVal startTime As Double) As Double
    If Timer >= startTime Then
        SmokeElapsedSeconds = Timer - startTime
    Else
        SmokeElapsedSeconds = (86400# - startTime) + Timer
    End If
End Function

Private Function CollectionContainsText(ByVal values As Collection, ByVal expected As String) As Boolean
    Dim value As Variant

    For Each value In values
        If LCase`$(CStr(value)) = LCase`$(expected) Then
            CollectionContainsText = True
            Exit Function
        End If
    Next value
End Function

Private Function GetSmokeInlineShape(ByVal index As Long) As InlineShape
    On Error GoTo Failed
    Set GetSmokeInlineShape = ActiveDocument.InlineShapes.Item(index)
    Exit Function
Failed:
    Err.Raise Err.Number, "GetSmokeInlineShape", "index " & CStr(index) & ": " & Err.Description
End Function

Private Sub AssertNoRawRootZipEntry(ByVal zipPath As String)
    Dim fso As Object
    Dim wsh As Object
    Dim listPath As String
    Dim command As String
    Dim exitCode As Long
    Dim fileNum As Integer
    Dim lineText As String
    Dim normalizedLine As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    listPath = fso.BuildPath(fso.GetSpecialFolder(2), "OleRawDataInserterSmokeZipList_" & Format(Now, "yyyymmdd_hhnnss") & ".txt")
    command = "cmd.exe /c tar.exe -tf " & QuoteForTestCommandLine(zipPath) & " > " & QuoteForTestCommandLine(listPath)
    exitCode = wsh.Run(command, 0, True)
    If exitCode <> 0 Then Err.Raise vbObjectError + 903, "SmokeTest", "tar.exe failed to list initial zip entries"

    fileNum = FreeFile
    Open listPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        normalizedLine = Trim(Replace(lineText, "\", "/"))
        If normalizedLine = "." Or normalizedLine = "./" Then
            Close #fileNum
            fso.DeleteFile listPath, True
            Err.Raise vbObjectError + 904, "SmokeTest", "Initial zip contains a raw root entry: " & lineText
        End If
    Loop
    Close #fileNum
    fso.DeleteFile listPath, True
End Sub

Private Function QuoteForTestCommandLine(ByVal value As String) As String
    QuoteForTestCommandLine = Chr(34) & Replace(value, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function

Private Function SmokeUnicodeChar(ByVal code As Long) As String
    If code > 32767 Then code = code - 65536
    SmokeUnicodeChar = ChrW(code)
End Function
"@)

    Write-Step 'Running InsertPlotFolder macro'
    $word.Run('SmokeTest')

    $logContent = Get-Content -LiteralPath $logPath -Encoding UTF8 -Raw
    if ($logContent -match 'VBA SmokeTest ERROR') {
        throw ($logContent -split "`r?`n" | Where-Object { $_ -match 'VBA SmokeTest ERROR' } | Select-Object -Last 1)
    }

    Write-Step 'Macro returned; checking document'
    $shapeCount = $doc.InlineShapes.Count
    if ($shapeCount -ne 5) {
        throw "Expected 5 inline shape(s), found $shapeCount"
    }
    $attachedShape = $doc.InlineShapes.Item(4)
    if ([math]::Abs($attachedShape.Width - 180) -gt 0.5 -or [math]::Abs($attachedShape.Height - 101.25) -gt 0.5) {
        throw "Expected attached image size 180 x 101.25 pt, found $([math]::Round($attachedShape.Width, 2)) x $([math]::Round($attachedShape.Height, 2)) pt"
    }
    $linkedAttachedShape = $doc.InlineShapes.Item(5)
    if ([math]::Abs($linkedAttachedShape.Width - 120) -gt 0.5 -or [math]::Abs($linkedAttachedShape.Height - 67.5) -gt 0.5) {
        throw "Expected linked attached image size 120 x 67.5 pt, found $([math]::Round($linkedAttachedShape.Width, 2)) x $([math]::Round($linkedAttachedShape.Height, 2)) pt"
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
    Remove-Item -LiteralPath (Join-Path $sampleDir 'notes.txt') -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $supportFolderDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $managedWorkspaceDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $packageWorkspaceDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $unicodeWorkspaceDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $unicodeExtractDir -Recurse -Force -ErrorAction SilentlyContinue
}
