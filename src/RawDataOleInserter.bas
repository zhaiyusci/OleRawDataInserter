Attribute VB_Name = "RawDataOleInserter"
Option Explicit

Private Type SizeInPoints
    Width As Double
    Height As Double
End Type

#If VBA7 Then
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
#Else
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
#End If

#If VBA7 Then
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputBuf As GdiplusStartupInput, ByVal outputBuf As LongPtr) As Long
    Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef image As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
    Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
    Private Declare PtrSafe Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal image As LongPtr, ByRef resolution As Single) As Long
    Private Declare PtrSafe Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal image As LongPtr, ByRef resolution As Single) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputBuf As GdiplusStartupInput, ByVal outputBuf As Long) As Long
    Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef image As Long) As Long
    Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, ByRef width As Long) As Long
    Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, ByRef height As Long) As Long
    Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal image As Long, ByRef resolution As Single) As Long
    Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal image As Long, ByRef resolution As Single) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const ZIP_COPY_FLAGS As Long = 4 + 16
Private Const DEFAULT_DPI As Double = 96#

Public Sub InsertPlotFolderAsOle()
    Dim folderPath As String

    folderPath = PickPlotFolder()
    If Len(folderPath) = 0 Then Exit Sub

    InsertPlotFolder folderPath
End Sub

Public Sub InsertImageSupportFilesAsOle()
    ImageSupportFilesDialog.Show
End Sub

Public Sub ShowUsageHelp()
    MsgBox BuildUsageHelpText(), vbInformation, DecodeEscapedText("Figure Package \u4F7F\u7528\u8BF4\u660E")
End Sub

Public Sub InsertPlotImageAsOle()
    Dim plotPath As String

    plotPath = PickPlotImage()
    If Len(plotPath) = 0 Then Exit Sub

    InsertPlotImageFile plotPath
End Sub

Public Sub InsertPlotFolder(ByVal folderPath As String)
    Dim fso As Object
    Dim plotPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = NormalizeFolderPath(folderPath)

    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder does not exist: " & folderPath, vbCritical
        Exit Sub
    End If

    plotPath = fso.BuildPath(folderPath, "plot.png")
    If Not fso.FileExists(plotPath) Then
        MsgBox "plot.png is not found: " & folderPath, vbCritical
        Exit Sub
    End If

    InsertImageAndRawDataAsOle plotPath, folderPath
End Sub

Public Sub InsertPlotImageFile(ByVal plotPath As String)
    Dim fso As Object
    Dim folderPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    plotPath = Trim$(plotPath)

    If Not fso.FileExists(plotPath) Then
        MsgBox "plot.png does not exist: " & plotPath, vbCritical
        Exit Sub
    End If

    If LCase$(fso.GetFileName(plotPath)) <> "plot.png" Then
        MsgBox "Please choose plot.png.", vbExclamation
        Exit Sub
    End If

    folderPath = fso.GetParentFolderName(plotPath)
    InsertImageAndRawDataAsOle plotPath, folderPath
End Sub

Public Sub InsertImageAndRawDataAsOle(ByVal imagePath As String, ByVal rawDataPath As String)
    Dim fso As Object
    Dim zipPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    imagePath = Trim$(imagePath)
    rawDataPath = Trim$(rawDataPath)

    If Not fso.FileExists(imagePath) Then
        MsgBox "Image file does not exist: " & imagePath, vbCritical
        Exit Sub
    End If

    If Not fso.FileExists(rawDataPath) And Not fso.FolderExists(rawDataPath) Then
        MsgBox "Raw data path does not exist: " & rawDataPath, vbCritical
        Exit Sub
    End If

    zipPath = BuildTempZipPath(fso, rawDataPath)
    CreateZipFromPath rawDataPath, zipPath, True

    InsertImageZipAsOle imagePath, zipPath
End Sub

Public Sub InsertImageAndSupportFilesAsOle(ByVal imagePath As String, ByVal supportFiles As Collection)
    Dim fso As Object
    Dim zipPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    imagePath = Trim$(imagePath)

    If Not fso.FileExists(imagePath) Then
        MsgBox "Image file does not exist: " & imagePath, vbCritical
        Exit Sub
    End If

    If Not IsSupportedDisplayImage(fso.GetExtensionName(imagePath)) Then
        MsgBox "Please choose a PNG, JPG, JPEG, TIF, or TIFF image.", vbExclamation
        Exit Sub
    End If

    If supportFiles Is Nothing Then
        MsgBox "No support files were selected.", vbExclamation
        Exit Sub
    End If

    If supportFiles.Count = 0 Then
        MsgBox "No support files were selected.", vbExclamation
        Exit Sub
    End If

    zipPath = BuildTempZipPathFromBase(fso, fso.GetBaseName(imagePath) & "_support")
    CreateZipFromSupportFiles supportFiles, zipPath, fso.GetBaseName(imagePath)

    InsertImageZipAsOle imagePath, zipPath
End Sub

Private Sub InsertImageZipAsOle(ByVal imagePath As String, ByVal zipPath As String)
    Dim fso As Object
    Dim iconPath As String
    Dim oleObject As InlineShape
    Dim imgSize As SizeInPoints
    Dim imageWidth As Double
    Dim imageHeight As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(zipPath) Then
        MsgBox "Support zip file does not exist: " & zipPath, vbCritical
        Exit Sub
    End If

    iconPath = fso.BuildPath(GetTempFolderPath(fso), "OleRawDataInserterTransparent.ico")
    CreateTransparentIcon iconPath

    imgSize = GetImagePrintSize(imagePath)
    imageWidth = imgSize.Width
    imageHeight = imgSize.Height
    FitSizeToTextArea imageWidth, imageHeight

    Set oleObject = Selection.InlineShapes.AddOLEObject( _
        FileName:=zipPath, _
        LinkToFile:=False, _
        DisplayAsIcon:=True, _
        IconFileName:=iconPath, _
        IconIndex:=0, _
        IconLabel:=vbNullString)

    oleObject.Width = imageWidth
    oleObject.Height = imageHeight

    On Error Resume Next
    oleObject.Borders.Enable = False
    oleObject.Line.Visible = msoFalse
    oleObject.Fill.Visible = msoTrue
    oleObject.Fill.UserPicture imagePath
    On Error GoTo 0
End Sub

Private Function PickPlotFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing plot.png"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickPlotFolder = vbNullString
        Else
            PickPlotFolder = .SelectedItems(1)
        End If
    End With
End Function

Private Function PickPlotImage() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Choose plot.png"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "plot.png", "plot.png"
        .Filters.Add "PNG files", "*.png"
        .Filters.Add "All files", "*.*"
        If .Show <> -1 Then
            PickPlotImage = vbNullString
        Else
            PickPlotImage = .SelectedItems(1)
        End If
    End With
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    Do While Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop
    NormalizeFolderPath = folderPath
End Function

Private Function BuildTempZipPath(ByVal fso As Object, ByVal folderPath As String) As String
    Dim baseName As String

    If fso.FolderExists(folderPath) Then
        baseName = fso.GetFolder(folderPath).Name
    Else
        baseName = fso.GetBaseName(folderPath)
    End If

    BuildTempZipPath = BuildTempZipPathFromBase(fso, baseName)
End Function

Private Function BuildTempZipPathFromBase(ByVal fso As Object, ByVal baseName As String) As String
    Dim stamp As String

    stamp = Format$(Now, "yyyymmdd_hhnnss")
    BuildTempZipPathFromBase = fso.BuildPath(GetTempFolderPath(fso), baseName & "_" & stamp & ".zip")
End Function

Private Function GetTempFolderPath(ByVal fso As Object) As String
    Dim tempPath As String

    tempPath = Environ$("LOCALAPPDATA")
    If Len(tempPath) > 0 Then
        tempPath = fso.BuildPath(tempPath, "Temp")
    Else
        tempPath = Environ$("TEMP")
    End If

    GetTempFolderPath = tempPath
End Function

Private Function IsSupportedDisplayImage(ByVal extensionName As String) As Boolean
    Select Case LCase$(extensionName)
        Case "png", "jpg", "jpeg", "tif", "tiff"
            IsSupportedDisplayImage = True
        Case Else
            IsSupportedDisplayImage = False
    End Select
End Function

Private Function BuildUsageHelpText() As String
    Dim escaped As String

    escaped = _
        "Figure Package \u4F7F\u7528\u8BF4\u660E\n\n" & _
        "\u65B9\u5F0F\u4E00\uFF1APlot \u6587\u4EF6\u5939\n" & _
        "1. \u5728\u4E00\u4E2A\u6587\u4EF6\u5939\u4E2D\u51C6\u5907 plot.py\u3001\u539F\u59CB\u6570\u636E\uFF0C\u4EE5\u53CA\u5176\u4ED6\u8F85\u52A9\u811A\u672C\u6216\u56FE\u7247\u3002\n" & _
        "2. \u8FD0\u884C plot.py \u751F\u6210\u6700\u7EC8\u56FE\u7247 plot.png\u3002plot.png \u5FC5\u987B\u548C plot.py \u653E\u5728\u540C\u4E00\u4E2A\u6587\u4EF6\u5939\u91CC\u3002\n" & _
        "3. \u5728 Word \u4E2D\u70B9\u51FB Figure Package > Insert Figure Package\uFF0C\u9009\u62E9\u8FD9\u4E2A\u6587\u4EF6\u5939\u3002\n" & _
        "4. \u63D2\u4EF6\u4F1A\u628A\u8BE5\u6587\u4EF6\u5939\u4E2D\u7684\u539F\u59CB\u6570\u636E\u548C\u811A\u672C\u6253\u5305\u6210 zip\uFF0C\u5E76\u4F5C\u4E3A OLE \u5BF9\u8C61\u5D4C\u5165\u5F53\u524D\u6587\u6863\uFF1BWord \u4E2D\u663E\u793A\u7684\u662F plot.png\u3002\n" & _
        "5. \u6253\u5305\u65F6\u4E0D\u4F1A\u5305\u542B\u9876\u5C42 plot.png\u3001plot.svg\u3001plot.pdf\uFF1B\u4F1A\u5305\u542B plot.py\u3001\u5176\u4ED6\u6570\u636E\u6587\u4EF6\u3001\u811A\u672C\u3001\u56FE\u7247\u548C\u5B50\u6587\u4EF6\u5939\u3002\n\n" & _
        "\u65B9\u5F0F\u4E8C\uFF1A\u56FE\u7247 + \u652F\u6301\u6587\u4EF6\n" & _
        "1. \u5728 Word \u4E2D\u70B9\u51FB Figure Package > Insert Image + Files\uFF0C\u6253\u5F00\u786E\u8BA4\u7A97\u53E3\u3002\n" & _
        "2. \u5728\u7A97\u53E3\u4E2D\u70B9\u51FB Choose image...\uFF0C\u9009\u62E9\u8981\u663E\u793A\u7684\u56FE\u7247\uFF0C\u652F\u6301 png\u3001jpg\u3001jpeg\u3001tif\u3001tiff\u3002\n" & _
        "3. \u70B9\u51FB Add files... \u591A\u9009\u9700\u8981\u5D4C\u5165\u7684\u652F\u6301\u6587\u4EF6\uFF0C\u6587\u4EF6\u7C7B\u578B\u4E0D\u9650\uFF1B\u7A97\u53E3\u4E2D\u4F1A\u5217\u51FA\u5DF2\u9009\u62E9\u7684\u6587\u4EF6\u3002\n" & _
        "4. \u5982\u679C\u9009\u9519\u4E86\uFF0C\u53EF\u4EE5\u7528 Remove selected \u6216 Clear \u8C03\u6574\u5217\u8868\u3002\n" & _
        "5. \u786E\u8BA4\u56FE\u7247\u548C\u652F\u6301\u6587\u4EF6\u5217\u8868\u65E0\u8BEF\u540E\uFF0C\u70B9\u51FB Insert\u3002\u63D2\u4EF6\u4F1A\u628A\u652F\u6301\u6587\u4EF6\u6253\u5305\u6210 zip\uFF0C\u5E76\u4F5C\u4E3A OLE \u5BF9\u8C61\u5D4C\u5165\u5F53\u524D\u6587\u6863\uFF1BWord \u4E2D\u663E\u793A\u7684\u662F\u6240\u9009\u56FE\u7247\u3002\n" & _
        "6. \u5982\u679C\u652F\u6301\u6587\u4EF6\u540C\u540D\uFF0Czip \u5185\u4F1A\u81EA\u52A8\u6539\u540D\uFF0C\u907F\u514D\u8986\u76D6\u3002\n\n" & _
        "\u56FE\u7247\u5C3A\u5BF8\uFF1A\u59CB\u7EC8\u4FDD\u6301\u9AD8\u5BBD\u6BD4\uFF1B\u5C0F\u4E8E\u7248\u5FC3\u65F6\u4FDD\u7559\u539F\u59CB\u5370\u5237\u5C3A\u5BF8\uFF1B\u5927\u4E8E\u7248\u5FC3\u65F6\u7B49\u6BD4\u7F29\u5C0F\u5230\u80FD\u653E\u8FDB\u7248\u5FC3\u3002"

    BuildUsageHelpText = DecodeEscapedText(escaped)
End Function

Private Function DecodeEscapedText(ByVal escaped As String) As String
    Dim result As String
    Dim i As Long
    Dim code As Long
    Dim token As String

    i = 1
    Do While i <= Len(escaped)
        If Mid$(escaped, i, 2) = "\n" Then
            result = result & vbCrLf
            i = i + 2
        ElseIf Mid$(escaped, i, 2) = "\u" And i + 5 <= Len(escaped) Then
            token = Mid$(escaped, i + 2, 4)
            code = CLng("&H" & token)
            result = result & UnicodeChar(code)
            i = i + 6
        Else
            result = result & Mid$(escaped, i, 1)
            i = i + 1
        End If
    Loop

    DecodeEscapedText = result
End Function

Private Function UnicodeChar(ByVal code As Long) As String
    If code > 32767 Then code = code - 65536
    UnicodeChar = ChrW(code)
End Function

Private Sub CreateZipFromSupportFiles(ByVal supportFiles As Collection, ByVal zipPath As String, ByVal baseName As String)
    Dim fso As Object
    Dim stagingFolder As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    stagingFolder = BuildTempStagingFolderPath(fso, baseName)

    On Error GoTo Failed
    StageSupportFiles fso, supportFiles, stagingFolder
    CreateZipFromPath stagingFolder, zipPath, False
    DeleteFolderIfExists fso, stagingFolder
    Exit Sub

Failed:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    On Error Resume Next
    DeleteFolderIfExists fso, stagingFolder
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
    On Error GoTo 0
    Err.Raise errNumber, errSource, errDescription
End Sub

Private Function BuildTempStagingFolderPath(ByVal fso As Object, ByVal baseName As String) As String
    Dim stamp As String

    stamp = Format$(Now, "yyyymmdd_hhnnss")
    BuildTempStagingFolderPath = fso.BuildPath(GetTempFolderPath(fso), "OleRawDataInserter_" & baseName & "_" & stamp)
End Function

Private Sub StageSupportFiles(ByVal fso As Object, ByVal supportFiles As Collection, ByVal stagingFolder As String)
    Dim usedNames As Object
    Dim filePath As Variant
    Dim targetName As String

    If supportFiles Is Nothing Then
        Err.Raise vbObjectError + 520, "StageSupportFiles", "No support files were selected."
    End If

    If supportFiles.Count = 0 Then
        Err.Raise vbObjectError + 521, "StageSupportFiles", "No support files were selected."
    End If

    DeleteFolderIfExists fso, stagingFolder
    fso.CreateFolder stagingFolder
    Set usedNames = CreateObject("Scripting.Dictionary")

    For Each filePath In supportFiles
        If Not fso.FileExists(CStr(filePath)) Then
            Err.Raise vbObjectError + 522, "StageSupportFiles", "Support file does not exist: " & CStr(filePath)
        End If

        targetName = GetUniqueStagedFileName(fso, usedNames, fso.GetFileName(CStr(filePath)))
        fso.CopyFile CStr(filePath), fso.BuildPath(stagingFolder, targetName), True
    Next filePath
End Sub

Private Function GetUniqueStagedFileName(ByVal fso As Object, ByVal usedNames As Object, ByVal fileName As String) As String
    Dim baseName As String
    Dim extensionName As String
    Dim candidate As String
    Dim index As Long

    baseName = fso.GetBaseName(fileName)
    extensionName = fso.GetExtensionName(fileName)
    candidate = fileName
    index = 2

    Do While usedNames.Exists(LCase$(candidate))
        If Len(extensionName) > 0 Then
            candidate = baseName & "_" & CStr(index) & "." & extensionName
        Else
            candidate = baseName & "_" & CStr(index)
        End If
        index = index + 1
    Loop

    usedNames.Add LCase$(candidate), True
    GetUniqueStagedFileName = candidate
End Function

Private Sub DeleteFolderIfExists(ByVal fso As Object, ByVal folderPath As String)
    If Len(folderPath) > 0 Then
        If fso.FolderExists(folderPath) Then fso.DeleteFolder folderPath, True
    End If
End Sub

Private Sub CreateZipFromPath(ByVal sourcePath As String, ByVal zipPath As String, Optional ByVal excludeGeneratedPlotFiles As Boolean = True)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    On Error Resume Next
    CreateZipWithTar sourcePath, zipPath, excludeGeneratedPlotFiles
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear

    CreateZipWithPowerShell sourcePath, zipPath, excludeGeneratedPlotFiles
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    If CreateZipWithShell(sourcePath, zipPath, excludeGeneratedPlotFiles) Then Exit Sub
    Err.Raise vbObjectError + 514, "CreateZipFromFolder", "Failed to create zip file: " & zipPath
End Sub

Private Sub CreateZipWithTar(ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean)
    Dim fso As Object
    Dim wsh As Object
    Dim command As String
    Dim exitCode As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    If fso.FolderExists(sourcePath) Then
        command = "cmd.exe /c tar.exe -a -cf " & QuoteForCommandLine(zipPath)
        If excludeGeneratedPlotFiles Then
            command = command & " --exclude=./plot.png --exclude=./plot.svg --exclude=./plot.pdf"
        End If
        command = command & " -C " & QuoteForCommandLine(sourcePath) & " ."
    Else
        command = "cmd.exe /c tar.exe -a -cf " & QuoteForCommandLine(zipPath) & " -C " & QuoteForCommandLine(fso.GetParentFolderName(sourcePath)) & " " & QuoteForCommandLine(fso.GetFileName(sourcePath))
    End If

    Set wsh = CreateObject("WScript.Shell")
    exitCode = wsh.Run(command, 0, True)

    If exitCode <> 0 Then
        Err.Raise vbObjectError + 515, "CreateZipWithTar", "tar.exe failed with exit code " & exitCode
    End If

    If Dir$(zipPath) = vbNullString Then
        Err.Raise vbObjectError + 516, "CreateZipWithTar", "tar.exe did not create zip file: " & zipPath
    End If
End Sub

Private Function CreateZipWithShell(ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean) As Boolean
    Dim fso As Object
    Dim shellApp As Object
    Dim zipFolder As Object
    Dim sourceFolder As Object
    Dim sourceFile As Object
    Dim sourceItem As Object

    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    CreateEmptyZip zipPath

    Set shellApp = CreateObject("Shell.Application")
    Set zipFolder = shellApp.NameSpace(zipPath)

    If zipFolder Is Nothing Then
        GoTo Failed
    End If

    If fso.FolderExists(sourcePath) Then
        Set sourceFolder = shellApp.NameSpace(sourcePath)
        If sourceFolder Is Nothing Then GoTo Failed
        For Each sourceItem In sourceFolder.Items
            If Not excludeGeneratedPlotFiles Or Not IsGeneratedPlotFile(sourceItem.Name) Then
                zipFolder.CopyHere sourceItem, ZIP_COPY_FLAGS
            End If
        Next sourceItem
    Else
        Set sourceFolder = shellApp.NameSpace(fso.GetParentFolderName(sourcePath))
        If sourceFolder Is Nothing Then GoTo Failed
        Set sourceFile = sourceFolder.ParseName(fso.GetFileName(sourcePath))
        If sourceFile Is Nothing Then GoTo Failed
        zipFolder.CopyHere sourceFile, ZIP_COPY_FLAGS
    End If

    If Not WaitForZipToSettle(zipPath, 30) Then
        GoTo Failed
    End If

    CreateZipWithShell = True
    Exit Function

Failed:
    On Error Resume Next
    If Not fso Is Nothing Then
        If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
    End If
    CreateZipWithShell = False
End Function

Private Sub CreateZipWithPowerShell(ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean)
    Dim fso As Object
    Dim wsh As Object
    Dim command As String
    Dim scriptPath As String
    Dim exitCode As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    scriptPath = fso.BuildPath(GetTempFolderPath(fso), "OleRawDataInserterZip.ps1")
    WriteZipScript scriptPath, sourcePath, zipPath, excludeGeneratedPlotFiles

    Set wsh = CreateObject("WScript.Shell")
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & QuoteForCommandLine(scriptPath)
    exitCode = wsh.Run(command, 0, True)

    If exitCode <> 0 Then
        Err.Raise vbObjectError + 512, "CreateZipWithPowerShell", "PowerShell Compress-Archive failed with exit code " & exitCode
    End If

    If Dir$(zipPath) = vbNullString Then
        Err.Raise vbObjectError + 513, "CreateZipWithPowerShell", "PowerShell did not create zip file: " & zipPath
    End If
End Sub

Private Function PowerShellQuote(ByVal value As String) As String
    PowerShellQuote = "'" & Replace(value, "'", "''") & "'"
End Function

Private Function QuoteForCommandLine(ByVal value As String) As String
    QuoteForCommandLine = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Sub WriteZipScript(ByVal scriptPath As String, ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean)
    Dim fileNum As Integer
    Dim scriptText As String

    scriptText = _
        "$ErrorActionPreference = 'Stop'" & vbCrLf & _
        "$source = Get-Item -LiteralPath " & PowerShellQuote(sourcePath) & " -Force" & vbCrLf & _
        "$destination = " & PowerShellQuote(zipPath) & vbCrLf & _
        "if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Force }" & vbCrLf & _
        "if ($source.PSIsContainer) {" & vbCrLf

    If excludeGeneratedPlotFiles Then
        scriptText = scriptText & _
            "    $generated = @('plot.png', 'plot.svg', 'plot.pdf')" & vbCrLf & _
            "    $items = Get-ChildItem -LiteralPath $source.FullName -Force | Where-Object { $generated -notcontains $_.Name.ToLowerInvariant() }" & vbCrLf
    Else
        scriptText = scriptText & _
            "    $items = Get-ChildItem -LiteralPath $source.FullName -Force" & vbCrLf
    End If

    scriptText = scriptText & _
        "    if (-not $items) { throw 'Source folder is empty.' }" & vbCrLf & _
        "    Compress-Archive -Path $items.FullName -DestinationPath $destination -Force" & vbCrLf & _
        "} else {" & vbCrLf & _
        "    Compress-Archive -LiteralPath $source.FullName -DestinationPath $destination -Force" & vbCrLf & _
        "}"

    fileNum = FreeFile
    Open scriptPath For Output As #fileNum
    Print #fileNum, scriptText
    Close #fileNum
End Sub

Private Function IsGeneratedPlotFile(ByVal fileName As String) As Boolean
    Select Case LCase$(fileName)
        Case "plot.png", "plot.svg", "plot.pdf"
            IsGeneratedPlotFile = True
        Case Else
            IsGeneratedPlotFile = False
    End Select
End Function

Private Sub CreateEmptyZip(ByVal zipPath As String)
    Dim fileNum As Integer
    Dim endOfCentralDirectory As String

    endOfCentralDirectory = Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, Chr$(0))

    fileNum = FreeFile
    Open zipPath For Binary Access Write As #fileNum
    Put #fileNum, , endOfCentralDirectory
    Close #fileNum
End Sub

Private Function WaitForZipToSettle(ByVal zipPath As String, ByVal timeoutSeconds As Double) As Boolean
    Dim fso As Object
    Dim startTime As Double
    Dim lastSize As Double
    Dim stableTicks As Long
    Dim currentSize As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    startTime = Timer
    lastSize = -1

    Do
        If fso.FileExists(zipPath) Then
            currentSize = fso.GetFile(zipPath).Size
            If currentSize > 22 And currentSize = lastSize Then
                stableTicks = stableTicks + 1
                If stableTicks >= 4 Then
                    WaitForZipToSettle = True
                    Exit Function
                End If
            Else
                stableTicks = 0
                lastSize = currentSize
            End If
        End If

        DoEvents
        Sleep 250
    Loop While ElapsedSeconds(startTime) < timeoutSeconds
End Function

Private Function ElapsedSeconds(ByVal startTime As Double) As Double
    If Timer >= startTime Then
        ElapsedSeconds = Timer - startTime
    Else
        ElapsedSeconds = (86400# - startTime) + Timer
    End If
End Function

Private Function GetImagePrintSize(ByVal imgPath As String) As SizeInPoints
    Dim result As SizeInPoints
    Dim gdiInput As GdiplusStartupInput
#If VBA7 Then
    Dim token As LongPtr
    Dim image As LongPtr
#Else
    Dim token As Long
    Dim image As Long
#End If
    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim dpiX As Single
    Dim dpiY As Single

    result.Width = 360#
    result.Height = 240#

    gdiInput.GdiplusVersion = 1
    If GdiplusStartup(token, gdiInput, 0) <> 0 Then
        GetImagePrintSize = result
        Exit Function
    End If

    If GdipLoadImageFromFile(StrPtr(imgPath), image) = 0 Then
        GdipGetImageWidth image, pixelWidth
        GdipGetImageHeight image, pixelHeight
        GdipGetImageHorizontalResolution image, dpiX
        GdipGetImageVerticalResolution image, dpiY

        If dpiX <= 0 Then dpiX = DEFAULT_DPI
        If dpiY <= 0 Then dpiY = DEFAULT_DPI

        result.Width = (pixelWidth / dpiX) * 72#
        result.Height = (pixelHeight / dpiY) * 72#

        GdipDisposeImage image
    End If

    GdiplusShutdown token
    GetImagePrintSize = result
End Function

Private Sub FitSizeToTextArea(ByRef imageWidthPoints As Double, ByRef imageHeightPoints As Double)
    Dim textAreaWidth As Double
    Dim textAreaHeight As Double
    Dim fitScale As Double
    Dim widthScale As Double
    Dim heightScale As Double

    Call GetCurrentTextAreaSize(textAreaWidth, textAreaHeight)

    If imageWidthPoints <= 0 Or imageHeightPoints <= 0 Then Exit Sub
    If textAreaWidth <= 0 Or textAreaHeight <= 0 Then Exit Sub

    widthScale = textAreaWidth / imageWidthPoints
    heightScale = textAreaHeight / imageHeightPoints
    fitScale = 1#

    If widthScale < fitScale Then fitScale = widthScale
    If heightScale < fitScale Then fitScale = heightScale

    If fitScale < 1# Then
        imageWidthPoints = imageWidthPoints * fitScale
        imageHeightPoints = imageHeightPoints * fitScale
    End If
End Sub

Private Sub GetCurrentTextAreaSize(ByRef textAreaWidth As Double, ByRef textAreaHeight As Double)
    Dim setup As Object

    On Error GoTo Failed
    Set setup = Selection.Sections(1).PageSetup

    textAreaWidth = setup.PageWidth - setup.LeftMargin - setup.RightMargin - setup.Gutter
    textAreaHeight = setup.PageHeight - setup.TopMargin - setup.BottomMargin

    If textAreaWidth <= 0 Then textAreaWidth = setup.PageWidth - setup.LeftMargin - setup.RightMargin
    If textAreaHeight <= 0 Then textAreaHeight = setup.PageHeight - setup.TopMargin - setup.BottomMargin
    Exit Sub

Failed:
    textAreaWidth = 0
    textAreaHeight = 0
End Sub

Private Sub CreateTransparentIcon(ByVal iconPath As String)
    Dim fileNum As Integer
    Dim icoData() As Byte

    icoData = BuildTransparentIconBytes()

    fileNum = FreeFile
    Open iconPath For Binary Access Write As #fileNum
    Put #fileNum, , icoData
    Close #fileNum
End Sub

Private Function BuildTransparentIconBytes() As Byte()
    Const ICON_WIDTH As Long = 16
    Const ICON_HEIGHT As Long = 16
    Const BITMAP_HEADER_SIZE As Long = 40
    Const PIXEL_BYTES As Long = ICON_WIDTH * ICON_HEIGHT * 4
    Const MASK_BYTES As Long = 64
    Const IMAGE_BYTES As Long = BITMAP_HEADER_SIZE + PIXEL_BYTES + MASK_BYTES

    Dim bytes() As Byte
    Dim pos As Long
    Dim i As Long

    ReDim bytes(0 To 21 + IMAGE_BYTES)

    AppendUInt16 bytes, pos, 0
    AppendUInt16 bytes, pos, 1
    AppendUInt16 bytes, pos, 1

    AppendByte bytes, pos, ICON_WIDTH
    AppendByte bytes, pos, ICON_HEIGHT
    AppendByte bytes, pos, 0
    AppendByte bytes, pos, 0
    AppendUInt16 bytes, pos, 1
    AppendUInt16 bytes, pos, 32
    AppendUInt32 bytes, pos, IMAGE_BYTES
    AppendUInt32 bytes, pos, 22

    AppendUInt32 bytes, pos, BITMAP_HEADER_SIZE
    AppendUInt32 bytes, pos, ICON_WIDTH
    AppendUInt32 bytes, pos, ICON_HEIGHT * 2
    AppendUInt16 bytes, pos, 1
    AppendUInt16 bytes, pos, 32
    AppendUInt32 bytes, pos, 0
    AppendUInt32 bytes, pos, PIXEL_BYTES
    AppendUInt32 bytes, pos, 0
    AppendUInt32 bytes, pos, 0
    AppendUInt32 bytes, pos, 0
    AppendUInt32 bytes, pos, 0

    For i = 1 To PIXEL_BYTES
        AppendByte bytes, pos, 0
    Next i

    For i = 1 To MASK_BYTES
        AppendByte bytes, pos, &HFF
    Next i

    BuildTransparentIconBytes = bytes
End Function

Private Sub AppendByte(ByRef bytes() As Byte, ByRef pos As Long, ByVal value As Long)
    bytes(pos) = CByte(value And &HFF&)
    pos = pos + 1
End Sub

Private Sub AppendUInt16(ByRef bytes() As Byte, ByRef pos As Long, ByVal value As Long)
    AppendByte bytes, pos, value
    AppendByte bytes, pos, value \ &H100&
End Sub

Private Sub AppendUInt32(ByRef bytes() As Byte, ByRef pos As Long, ByVal value As Long)
    AppendByte bytes, pos, value
    AppendByte bytes, pos, value \ &H100&
    AppendByte bytes, pos, value \ &H10000
    AppendByte bytes, pos, value \ &H1000000
End Sub
