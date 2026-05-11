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
Private Const CFB_END_OF_CHAIN As Long = -2
Private Const CFB_FREE_SECTOR As Long = -1

Public Sub InsertPlotFolderAsOle()
    Dim folderPath As String

    folderPath = PickPlotFolder()
    If Len(folderPath) = 0 Then Exit Sub

    InsertPlotFolder folderPath
End Sub

Public Sub InsertImageSupportFilesAsOle()
    ImageSupportFilesDialog.Show
End Sub

Public Sub AttachSupportFilesToSelectedImageAsOle()
    Dim inlineImage As InlineShape
    Dim inlineOle As InlineShape
    Dim floatingImage As Shape
    Dim floatingOle As Shape
    Dim dialog As ImageSupportFilesDialog

    If TryGetSelectedInlineImage(inlineImage) Then
        Set dialog = New ImageSupportFilesDialog
        dialog.ConfigureForSelectedInlineImage inlineImage
        dialog.Show
        Exit Sub
    End If

    If TryGetSelectedInlineOlePackage(inlineOle) Then
        Set dialog = New ImageSupportFilesDialog
        dialog.ConfigureForSelectedInlineOle inlineOle
        dialog.Show
        Exit Sub
    End If

    If TryGetSelectedFloatingImage(floatingImage) Then
        Set dialog = New ImageSupportFilesDialog
        dialog.ConfigureForSelectedFloatingImage floatingImage
        dialog.Show
        Exit Sub
    End If

    If TryGetSelectedFloatingOlePackage(floatingOle) Then
        Set dialog = New ImageSupportFilesDialog
        dialog.ConfigureForSelectedFloatingOle floatingOle
        dialog.Show
        Exit Sub
    End If

    MsgBox T("msg.selectExisting"), vbExclamation
End Sub

Public Sub ShowUsageHelp()
    UsageHelpDialog.ShowUsageText T("usage.text")
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
        MsgBox TF("msg.folderMissing", folderPath), vbCritical
        Exit Sub
    End If

    plotPath = fso.BuildPath(folderPath, "plot.png")
    If Not fso.FileExists(plotPath) Then
        MsgBox TF("msg.plotMissingInFolder", folderPath), vbCritical
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
        MsgBox TF("msg.plotFileMissing", plotPath), vbCritical
        Exit Sub
    End If

    If LCase$(fso.GetFileName(plotPath)) <> "plot.png" Then
        MsgBox T("msg.choosePlot"), vbExclamation
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
        MsgBox TF("msg.imageMissing", imagePath), vbCritical
        Exit Sub
    End If

    If Not fso.FileExists(rawDataPath) And Not fso.FolderExists(rawDataPath) Then
        MsgBox TF("msg.rawDataMissing", rawDataPath), vbCritical
        Exit Sub
    End If

    zipPath = BuildTempZipPath(fso, rawDataPath)
    CreateZipFromPath rawDataPath, zipPath, True

    InsertImageZipAsOle imagePath, zipPath
End Sub

Public Sub AttachSupportFilesToInlineImage(ByVal targetInlineShape As Object, ByVal supportFiles As Collection)
    Dim fso As Object
    Dim imagePath As String
    Dim zipPath As String
    Dim targetRange As Range
    Dim imageWidth As Double
    Dim imageHeight As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetInlineShape Is Nothing Then
        MsgBox T("msg.selectedImageGone"), vbExclamation
        Exit Sub
    End If

    ValidateSupportFiles supportFiles

    imagePath = ExtractImageFromInlineShape(targetInlineShape, "selected_image")
    zipPath = BuildTempZipPathFromBase(fso, fso.GetBaseName(imagePath) & "_support")
    CreateZipFromSupportFiles supportFiles, zipPath, fso.GetBaseName(imagePath)

    imageWidth = targetInlineShape.Width
    imageHeight = targetInlineShape.Height
    Set targetRange = targetInlineShape.Range.Duplicate
    targetRange.Collapse wdCollapseStart

    targetInlineShape.Delete
    targetRange.Select
    InsertImageZipAsOleWithSize imagePath, zipPath, imageWidth, imageHeight
End Sub

Public Sub AttachWorkingFolderToInlineImage(ByVal targetInlineShape As Object, ByVal workingFolderPath As String)
    Dim fso As Object
    Dim imagePath As String
    Dim zipPath As String
    Dim targetRange As Range
    Dim imageWidth As Double
    Dim imageHeight As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetInlineShape Is Nothing Then
        MsgBox T("msg.selectedImageGone"), vbExclamation
        Exit Sub
    End If

    imagePath = ExtractImageFromInlineShape(targetInlineShape, "selected_image")
    zipPath = BuildTempZipPathFromBase(fso, fso.GetBaseName(imagePath) & "_support")
    CreateZipFromWorkingFolder workingFolderPath, zipPath, False

    imageWidth = targetInlineShape.Width
    imageHeight = targetInlineShape.Height
    Set targetRange = targetInlineShape.Range.Duplicate
    targetRange.Collapse wdCollapseStart

    targetInlineShape.Delete
    targetRange.Select
    InsertImageZipAsOleWithSize imagePath, zipPath, imageWidth, imageHeight
End Sub

Public Sub AttachSupportFilesToFloatingImage(ByVal targetShape As Object, ByVal supportFiles As Collection)
    Dim fso As Object
    Dim imagePath As String
    Dim zipPath As String
    Dim anchorRange As Range
    Dim convertedInline As InlineShape
    Dim oleInline As InlineShape
    Dim oleShape As Shape
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim imageLeft As Single
    Dim imageTop As Single
    Dim relativeHorizontalPosition As Long
    Dim relativeVerticalPosition As Long
    Dim wrapType As Long
    Dim layoutInCell As Long
    Dim lockAnchor As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetShape Is Nothing Then
        MsgBox T("msg.selectedImageGone"), vbExclamation
        Exit Sub
    End If

    If Not IsFloatingPictureShape(targetShape) Then
        MsgBox T("msg.selectPictureOnly"), vbExclamation
        Exit Sub
    End If

    ValidateSupportFiles supportFiles

    imageWidth = targetShape.Width
    imageHeight = targetShape.Height
    imageLeft = targetShape.Left
    imageTop = targetShape.Top
    relativeHorizontalPosition = targetShape.RelativeHorizontalPosition
    relativeVerticalPosition = targetShape.RelativeVerticalPosition
    wrapType = targetShape.WrapFormat.Type
    layoutInCell = targetShape.LayoutInCell
    lockAnchor = targetShape.LockAnchor

    imagePath = TryGetLinkedImageSourceFromShape(targetShape)

    ' Convert the selected floating picture temporarily so the OpenXML only contains that image.
    Set convertedInline = targetShape.ConvertToInlineShape
    If Len(imagePath) = 0 Then
        imagePath = ExtractImageFromInlineShape(convertedInline, "selected_image")
    End If
    zipPath = BuildTempZipPathFromBase(fso, fso.GetBaseName(imagePath) & "_support")
    CreateZipFromSupportFiles supportFiles, zipPath, fso.GetBaseName(imagePath)

    Set anchorRange = convertedInline.Range.Duplicate
    anchorRange.Collapse wdCollapseStart
    convertedInline.Delete
    anchorRange.Select
    Set oleInline = AddImageZipOleInline(imagePath, zipPath)
    oleInline.Width = imageWidth
    oleInline.Height = imageHeight

    On Error Resume Next
    Set oleShape = oleInline.ConvertToShape
    If Not oleShape Is Nothing Then
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.RelativeHorizontalPosition = relativeHorizontalPosition
        oleShape.RelativeVerticalPosition = relativeVerticalPosition
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
        oleShape.WrapFormat.Type = wrapType
        oleShape.LayoutInCell = layoutInCell
        oleShape.LockAnchor = lockAnchor
        oleShape.LockAspectRatio = msoFalse
        oleShape.Fill.Visible = msoTrue
        oleShape.Fill.UserPicture imagePath
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
    End If
    On Error GoTo 0
End Sub

Public Sub AttachWorkingFolderToFloatingImage(ByVal targetShape As Object, ByVal workingFolderPath As String)
    Dim fso As Object
    Dim imagePath As String
    Dim zipPath As String
    Dim anchorRange As Range
    Dim convertedInline As InlineShape
    Dim oleInline As InlineShape
    Dim oleShape As Shape
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim imageLeft As Single
    Dim imageTop As Single
    Dim relativeHorizontalPosition As Long
    Dim relativeVerticalPosition As Long
    Dim wrapType As Long
    Dim layoutInCell As Long
    Dim lockAnchor As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetShape Is Nothing Then
        MsgBox T("msg.selectedImageGone"), vbExclamation
        Exit Sub
    End If

    If Not IsFloatingPictureShape(targetShape) Then
        MsgBox T("msg.selectPictureOnly"), vbExclamation
        Exit Sub
    End If

    imageWidth = targetShape.Width
    imageHeight = targetShape.Height
    imageLeft = targetShape.Left
    imageTop = targetShape.Top
    relativeHorizontalPosition = targetShape.RelativeHorizontalPosition
    relativeVerticalPosition = targetShape.RelativeVerticalPosition
    wrapType = targetShape.WrapFormat.Type
    layoutInCell = targetShape.LayoutInCell
    lockAnchor = targetShape.LockAnchor

    imagePath = TryGetLinkedImageSourceFromShape(targetShape)

    Set convertedInline = targetShape.ConvertToInlineShape
    If Len(imagePath) = 0 Then
        imagePath = ExtractImageFromInlineShape(convertedInline, "selected_image")
    End If
    zipPath = BuildTempZipPathFromBase(fso, fso.GetBaseName(imagePath) & "_support")
    CreateZipFromWorkingFolder workingFolderPath, zipPath, False

    Set anchorRange = convertedInline.Range.Duplicate
    anchorRange.Collapse wdCollapseStart
    convertedInline.Delete
    anchorRange.Select
    Set oleInline = AddImageZipOleInline(imagePath, zipPath)
    oleInline.Width = imageWidth
    oleInline.Height = imageHeight

    On Error Resume Next
    Set oleShape = oleInline.ConvertToShape
    If Not oleShape Is Nothing Then
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.RelativeHorizontalPosition = relativeHorizontalPosition
        oleShape.RelativeVerticalPosition = relativeVerticalPosition
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
        oleShape.WrapFormat.Type = wrapType
        oleShape.LayoutInCell = layoutInCell
        oleShape.LockAnchor = lockAnchor
        oleShape.LockAspectRatio = msoFalse
        oleShape.Fill.Visible = msoTrue
        oleShape.Fill.UserPicture imagePath
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
    End If
    On Error GoTo 0
End Sub

Public Function ExtractZipFromInlineOleObject(ByVal targetInlineShape As Object) As String
    Dim targetRange As Range
    Dim openXml As String
    Dim stage As String

    On Error GoTo Failed

    If targetInlineShape Is Nothing Then
        Err.Raise vbObjectError + 540, "ExtractZipFromInlineOleObject", "The selected OLE object is no longer available."
    End If

    stage = "read inline shape range"
    Set targetRange = ReadInlineShapeRangeWithRetry(targetInlineShape, "ExtractZipFromInlineOleObject")
    stage = "read WordOpenXML"
    openXml = ReadRangeWordOpenXmlWithRetry(targetRange, "ExtractZipFromInlineOleObject")
    stage = "extract zip from OLE XML"
    ExtractZipFromInlineOleObject = ExtractZipFromOleOpenXml(openXml, "selected_ole")
    Exit Function

Failed:
    Err.Raise Err.Number, "ExtractZipFromInlineOleObject", stage & ": " & Err.Description
End Function

Public Function ExtractZipFromFloatingOleObject(ByRef targetShape As Object) As String
    Dim convertedInline As InlineShape
    Dim restoredShape As Shape
    Dim zipPath As String
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim imageLeft As Single
    Dim imageTop As Single
    Dim relativeHorizontalPosition As Long
    Dim relativeVerticalPosition As Long
    Dim wrapType As Long
    Dim layoutInCell As Long
    Dim lockAnchor As Boolean

    If targetShape Is Nothing Then
        Err.Raise vbObjectError + 540, "ExtractZipFromFloatingOleObject", "The selected OLE object is no longer available."
    End If

    imageWidth = targetShape.Width
    imageHeight = targetShape.Height
    imageLeft = targetShape.Left
    imageTop = targetShape.Top
    relativeHorizontalPosition = targetShape.RelativeHorizontalPosition
    relativeVerticalPosition = targetShape.RelativeVerticalPosition
    wrapType = targetShape.WrapFormat.Type
    layoutInCell = targetShape.LayoutInCell
    lockAnchor = targetShape.LockAnchor

    Set convertedInline = targetShape.ConvertToInlineShape
    zipPath = ExtractZipFromOleOpenXml(ReadRangeWordOpenXmlWithRetry(convertedInline.Range, "ExtractZipFromFloatingOleObject"), "selected_ole")
    Set restoredShape = convertedInline.ConvertToShape
    restoredShape.Width = imageWidth
    restoredShape.Height = imageHeight
    restoredShape.RelativeHorizontalPosition = relativeHorizontalPosition
    restoredShape.RelativeVerticalPosition = relativeVerticalPosition
    restoredShape.Left = imageLeft
    restoredShape.Top = imageTop
    restoredShape.WrapFormat.Type = wrapType
    restoredShape.LayoutInCell = layoutInCell
    restoredShape.LockAnchor = lockAnchor
    Set targetShape = restoredShape

    ExtractZipFromFloatingOleObject = zipPath
End Function

Public Function ExtractDisplayImageFromInlineOleObject(ByVal targetInlineShape As Object) As String
    If targetInlineShape Is Nothing Then
        Err.Raise vbObjectError + 564, "ExtractDisplayImageFromInlineOleObject", "The selected OLE object is no longer available."
    End If

    ExtractDisplayImageFromInlineOleObject = ExtractImageFromInlineShape(targetInlineShape, "selected_ole_display")
End Function

Public Function ExtractDisplayImageFromFloatingOleObject(ByRef targetShape As Object) As String
    Dim convertedInline As InlineShape
    Dim restoredShape As Shape
    Dim imagePath As String
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim imageLeft As Single
    Dim imageTop As Single
    Dim relativeHorizontalPosition As Long
    Dim relativeVerticalPosition As Long
    Dim wrapType As Long
    Dim layoutInCell As Long
    Dim lockAnchor As Boolean

    If targetShape Is Nothing Then
        Err.Raise vbObjectError + 565, "ExtractDisplayImageFromFloatingOleObject", "The selected OLE object is no longer available."
    End If

    imageWidth = targetShape.Width
    imageHeight = targetShape.Height
    imageLeft = targetShape.Left
    imageTop = targetShape.Top
    relativeHorizontalPosition = targetShape.RelativeHorizontalPosition
    relativeVerticalPosition = targetShape.RelativeVerticalPosition
    wrapType = targetShape.WrapFormat.Type
    layoutInCell = targetShape.LayoutInCell
    lockAnchor = targetShape.LockAnchor

    Set convertedInline = targetShape.ConvertToInlineShape
    imagePath = ExtractImageFromInlineShape(convertedInline, "selected_ole_display")
    Set restoredShape = convertedInline.ConvertToShape
    restoredShape.Width = imageWidth
    restoredShape.Height = imageHeight
    restoredShape.RelativeHorizontalPosition = relativeHorizontalPosition
    restoredShape.RelativeVerticalPosition = relativeVerticalPosition
    restoredShape.Left = imageLeft
    restoredShape.Top = imageTop
    restoredShape.WrapFormat.Type = wrapType
    restoredShape.LayoutInCell = layoutInCell
    restoredShape.LockAnchor = lockAnchor
    Set targetShape = restoredShape

    ExtractDisplayImageFromFloatingOleObject = imagePath
End Function

Private Function ReadInlineShapeRangeWithRetry(ByVal targetInlineShape As Object, ByVal errorSource As String) As Range
    Dim attempt As Long
    Dim targetRange As Range
    Dim lastErrNumber As Long
    Dim lastErrDescription As String

    For attempt = 1 To 12
        On Error Resume Next
        Set targetRange = targetInlineShape.Range
        lastErrNumber = Err.Number
        lastErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0

        If lastErrNumber = 0 And Not targetRange Is Nothing Then
            Set ReadInlineShapeRangeWithRetry = targetRange
            Exit Function
        End If

        DoEvents
        Sleep 250
    Next attempt

    If lastErrNumber = 0 Then
        lastErrNumber = vbObjectError + 561
        lastErrDescription = "Could not access the selected OLE object range."
    End If
    Err.Raise lastErrNumber, errorSource, lastErrDescription
End Function

Private Function ReadRangeWordOpenXmlWithRetry(ByVal sourceRange As Range, ByVal errorSource As String) As String
    Dim attempt As Long
    Dim openXml As String
    Dim lastErrNumber As Long
    Dim lastErrDescription As String

    For attempt = 1 To 12
        On Error Resume Next
        openXml = CStr(sourceRange.WordOpenXML)
        lastErrNumber = Err.Number
        lastErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0

        If lastErrNumber = 0 And Len(openXml) > 0 Then
            ReadRangeWordOpenXmlWithRetry = openXml
            Exit Function
        End If

        DoEvents
        Sleep 250
    Next attempt

    If lastErrNumber = 0 Then
        lastErrNumber = vbObjectError + 560
        lastErrDescription = "Could not read the selected OLE object XML."
    End If
    Err.Raise lastErrNumber, errorSource, lastErrDescription
End Function

Public Function GetZipEntryNames(ByVal zipPath As String) As Collection
    Dim entries As Collection
    Dim fso As Object
    Dim listPath As String
    Dim listText As String
    Dim lines As Variant
    Dim lineItem As Variant
    Dim lineText As String
    Dim normalizedEntry As String

    Set entries = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(zipPath) Then
        Err.Raise vbObjectError + 541, "GetZipEntryNames", "Zip file does not exist: " & zipPath
    End If

    listPath = BuildTempFilePath(fso, "OLEPackagerZipList", ".txt")
    RunZipTool "list " & QuoteForCommandLine(zipPath) & " " & QuoteForCommandLine(listPath), "GetZipEntryNames"

    If fso.FileExists(listPath) Then
        listText = ReadUtf8TextFile(listPath)
        listText = Replace(Replace(listText, vbCrLf, vbLf), vbCr, vbLf)
        lines = Split(listText, vbLf)
        For Each lineItem In lines
            lineText = CStr(lineItem)
            If Not IsZipDirectoryEntry(lineText) Then
                normalizedEntry = NormalizeZipEntryName(lineText)
                If Len(normalizedEntry) > 0 Then entries.Add normalizedEntry
            End If
        Next lineItem
    End If
    DeleteFileIfExists fso, listPath

    Set GetZipEntryNames = entries
End Function

Private Function IsZipDirectoryEntry(ByVal entryName As String) As Boolean
    entryName = Trim$(Replace(entryName, "\", "/"))
    IsZipDirectoryEntry = (Len(entryName) > 0 And Right$(entryName, 1) = "/")
End Function

Public Sub ManageFilesInInlineOle(ByVal targetInlineShape As Object, ByVal existingZipPath As String, ByVal keepEntryNames As Collection, ByVal newFiles As Collection)
    Dim fso As Object
    Dim imagePath As String
    Dim newZipPath As String
    Dim targetRange As Range
    Dim imageWidth As Double
    Dim imageHeight As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetInlineShape Is Nothing Then
        MsgBox T("msg.selectedOleGone"), vbExclamation
        Exit Sub
    End If

    imagePath = ExtractImageFromInlineShape(targetInlineShape, "selected_ole_display")
    newZipPath = BuildTempZipPathFromBase(fso, "managed_ole_support")
    CreateZipFromManagedFiles existingZipPath, keepEntryNames, newFiles, newZipPath, fso.GetBaseName(newZipPath)

    imageWidth = targetInlineShape.Width
    imageHeight = targetInlineShape.Height
    Set targetRange = targetInlineShape.Range.Duplicate
    targetRange.Collapse wdCollapseStart

    targetInlineShape.Delete
    targetRange.Select
    InsertImageZipAsOleWithSize imagePath, newZipPath, imageWidth, imageHeight
End Sub

Public Sub ManageWorkingFolderInInlineOle(ByVal targetInlineShape As Object, ByVal workingFolderPath As String, Optional ByVal displayImagePath As String = vbNullString)
    Dim fso As Object
    Dim imagePath As String
    Dim newZipPath As String
    Dim targetRange As Range
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim stage As String

    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetInlineShape Is Nothing Then
        MsgBox T("msg.selectedOleGone"), vbExclamation
        Exit Sub
    End If

    stage = "resolve display image"
    imagePath = Trim$(displayImagePath)
    If Len(imagePath) = 0 Or Not fso.FileExists(imagePath) Then
        imagePath = ExtractImageFromInlineShape(targetInlineShape, "selected_ole_display")
    End If
    newZipPath = BuildTempZipPathFromBase(fso, "managed_ole_support")

    stage = "create UTF-8 zip"
    CreateZipFromWorkingFolder workingFolderPath, newZipPath, True

    stage = "capture original size"
    imageWidth = targetInlineShape.Width
    imageHeight = targetInlineShape.Height
    Set targetRange = targetInlineShape.Range.Duplicate
    targetRange.Collapse wdCollapseStart

    stage = "replace selected OLE"
    targetInlineShape.Delete
    targetRange.Select
    InsertImageZipAsOleWithSize imagePath, newZipPath, imageWidth, imageHeight
    Exit Sub

Failed:
    Err.Raise Err.Number, "ManageWorkingFolderInInlineOle", stage & ": " & Err.Description
End Sub

Public Sub ManageFilesInFloatingOle(ByVal targetShape As Object, ByVal existingZipPath As String, ByVal keepEntryNames As Collection, ByVal newFiles As Collection)
    Dim fso As Object
    Dim imagePath As String
    Dim newZipPath As String
    Dim anchorRange As Range
    Dim convertedInline As InlineShape
    Dim oleInline As InlineShape
    Dim oleShape As Shape
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim imageLeft As Single
    Dim imageTop As Single
    Dim relativeHorizontalPosition As Long
    Dim relativeVerticalPosition As Long
    Dim wrapType As Long
    Dim layoutInCell As Long
    Dim lockAnchor As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetShape Is Nothing Then
        MsgBox T("msg.selectedOleGone"), vbExclamation
        Exit Sub
    End If

    imageWidth = targetShape.Width
    imageHeight = targetShape.Height
    imageLeft = targetShape.Left
    imageTop = targetShape.Top
    relativeHorizontalPosition = targetShape.RelativeHorizontalPosition
    relativeVerticalPosition = targetShape.RelativeVerticalPosition
    wrapType = targetShape.WrapFormat.Type
    layoutInCell = targetShape.LayoutInCell
    lockAnchor = targetShape.LockAnchor

    Set convertedInline = targetShape.ConvertToInlineShape
    imagePath = ExtractImageFromInlineShape(convertedInline, "selected_ole_display")
    newZipPath = BuildTempZipPathFromBase(fso, "managed_ole_support")
    CreateZipFromManagedFiles existingZipPath, keepEntryNames, newFiles, newZipPath, fso.GetBaseName(newZipPath)

    Set anchorRange = convertedInline.Range.Duplicate
    anchorRange.Collapse wdCollapseStart
    convertedInline.Delete
    anchorRange.Select
    Set oleInline = AddImageZipOleInline(imagePath, newZipPath)
    oleInline.Width = imageWidth
    oleInline.Height = imageHeight

    On Error Resume Next
    Set oleShape = oleInline.ConvertToShape
    If Not oleShape Is Nothing Then
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.RelativeHorizontalPosition = relativeHorizontalPosition
        oleShape.RelativeVerticalPosition = relativeVerticalPosition
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
        oleShape.WrapFormat.Type = wrapType
        oleShape.LayoutInCell = layoutInCell
        oleShape.LockAnchor = lockAnchor
        oleShape.LockAspectRatio = msoFalse
        oleShape.Fill.Visible = msoTrue
        oleShape.Fill.UserPicture imagePath
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
    End If
    On Error GoTo 0
End Sub

Public Sub ManageWorkingFolderInFloatingOle(ByVal targetShape As Object, ByVal workingFolderPath As String, Optional ByVal displayImagePath As String = vbNullString)
    Dim fso As Object
    Dim imagePath As String
    Dim newZipPath As String
    Dim anchorRange As Range
    Dim convertedInline As InlineShape
    Dim oleInline As InlineShape
    Dim oleShape As Shape
    Dim imageWidth As Double
    Dim imageHeight As Double
    Dim imageLeft As Single
    Dim imageTop As Single
    Dim relativeHorizontalPosition As Long
    Dim relativeVerticalPosition As Long
    Dim wrapType As Long
    Dim layoutInCell As Long
    Dim lockAnchor As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    If targetShape Is Nothing Then
        MsgBox T("msg.selectedOleGone"), vbExclamation
        Exit Sub
    End If

    imageWidth = targetShape.Width
    imageHeight = targetShape.Height
    imageLeft = targetShape.Left
    imageTop = targetShape.Top
    relativeHorizontalPosition = targetShape.RelativeHorizontalPosition
    relativeVerticalPosition = targetShape.RelativeVerticalPosition
    wrapType = targetShape.WrapFormat.Type
    layoutInCell = targetShape.LayoutInCell
    lockAnchor = targetShape.LockAnchor

    imagePath = Trim$(displayImagePath)
    Set convertedInline = targetShape.ConvertToInlineShape
    If Len(imagePath) = 0 Or Not fso.FileExists(imagePath) Then
        imagePath = ExtractImageFromInlineShape(convertedInline, "selected_ole_display")
    End If
    newZipPath = BuildTempZipPathFromBase(fso, "managed_ole_support")
    CreateZipFromWorkingFolder workingFolderPath, newZipPath, True

    Set anchorRange = convertedInline.Range.Duplicate
    anchorRange.Collapse wdCollapseStart
    convertedInline.Delete
    anchorRange.Select
    Set oleInline = AddImageZipOleInline(imagePath, newZipPath)
    oleInline.Width = imageWidth
    oleInline.Height = imageHeight

    On Error Resume Next
    Set oleShape = oleInline.ConvertToShape
    If Not oleShape Is Nothing Then
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.RelativeHorizontalPosition = relativeHorizontalPosition
        oleShape.RelativeVerticalPosition = relativeVerticalPosition
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
        oleShape.WrapFormat.Type = wrapType
        oleShape.LayoutInCell = layoutInCell
        oleShape.LockAnchor = lockAnchor
        oleShape.LockAspectRatio = msoFalse
        oleShape.Fill.Visible = msoTrue
        oleShape.Fill.UserPicture imagePath
        oleShape.Width = imageWidth
        oleShape.Height = imageHeight
        oleShape.Left = imageLeft
        oleShape.Top = imageTop
    End If
    On Error GoTo 0
End Sub

Public Sub InsertImageAndSupportFilesAsOle(ByVal imagePath As String, ByVal supportFiles As Collection)
    Dim fso As Object
    Dim zipPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    imagePath = Trim$(imagePath)

    If Not fso.FileExists(imagePath) Then
        MsgBox TF("msg.imageMissing", imagePath), vbCritical
        Exit Sub
    End If

    If Not IsSupportedDisplayImage(fso.GetExtensionName(imagePath)) Then
        MsgBox T("msg.chooseSupportedImage"), vbExclamation
        Exit Sub
    End If

    If supportFiles Is Nothing Then
        MsgBox T("msg.noSupportFiles"), vbExclamation
        Exit Sub
    End If

    If supportFiles.Count = 0 Then
        MsgBox T("msg.noSupportFiles"), vbExclamation
        Exit Sub
    End If

    zipPath = BuildTempZipPathFromBase(fso, fso.GetBaseName(imagePath) & "_support")
    CreateZipFromSupportFiles supportFiles, zipPath, fso.GetBaseName(imagePath)

    InsertImageZipAsOle imagePath, zipPath
End Sub

Public Sub InsertImageAndWorkingFolderAsOle(ByVal imagePath As String, ByVal workingFolderPath As String)
    Dim fso As Object
    Dim zipPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    imagePath = Trim$(imagePath)

    If Not fso.FileExists(imagePath) Then
        MsgBox TF("msg.imageMissing", imagePath), vbCritical
        Exit Sub
    End If

    If Not IsSupportedDisplayImage(fso.GetExtensionName(imagePath)) Then
        MsgBox T("msg.chooseSupportedImage"), vbExclamation
        Exit Sub
    End If

    zipPath = BuildTempZipPathFromBase(fso, fso.GetBaseName(imagePath) & "_support")
    CreateZipFromWorkingFolder workingFolderPath, zipPath, False

    InsertImageZipAsOle imagePath, zipPath
End Sub

Public Sub ExtractZipToWorkingFolder(ByVal zipPath As String, ByVal targetFolder As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    DeleteFolderIfExists fso, targetFolder
    fso.CreateFolder targetFolder
    ExtractZipToFolder zipPath, targetFolder
End Sub

Private Sub InsertImageZipAsOle(ByVal imagePath As String, ByVal zipPath As String)
    Dim imgSize As SizeInPoints
    Dim imageWidth As Double
    Dim imageHeight As Double

    imgSize = GetImagePrintSize(imagePath)
    imageWidth = imgSize.Width
    imageHeight = imgSize.Height
    FitSizeToTextArea imageWidth, imageHeight

    InsertImageZipAsOleWithSize imagePath, zipPath, imageWidth, imageHeight
End Sub

Private Sub InsertImageZipAsOleWithSize(ByVal imagePath As String, ByVal zipPath As String, ByVal imageWidth As Double, ByVal imageHeight As Double)
    Dim oleObject As InlineShape

    Set oleObject = AddImageZipOleInline(imagePath, zipPath)
    oleObject.Width = imageWidth
    oleObject.Height = imageHeight
End Sub

Private Function AddImageZipOleInline(ByVal imagePath As String, ByVal zipPath As String) As InlineShape
    Dim fso As Object
    Dim iconPath As String
    Dim oleObject As InlineShape

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(zipPath) Then
        Err.Raise vbObjectError + 530, "AddImageZipOleInline", "Support zip file does not exist: " & zipPath
    End If

    iconPath = fso.BuildPath(GetTempFolderPath(fso), "OLEPackagerTransparent.ico")
    CreateTransparentIcon iconPath

    Set oleObject = Selection.InlineShapes.AddOLEObject( _
        FileName:=zipPath, _
        LinkToFile:=False, _
        DisplayAsIcon:=True, _
        IconFileName:=iconPath, _
        IconIndex:=0, _
        IconLabel:=vbNullString)

    On Error Resume Next
    oleObject.Borders.Enable = False
    oleObject.Line.Visible = msoFalse
    oleObject.Fill.Visible = msoTrue
    oleObject.Fill.UserPicture imagePath
    On Error GoTo 0

    Set AddImageZipOleInline = oleObject
End Function

Private Function PickPlotFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = T("picker.plotFolder")
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
        .Title = T("picker.plotImage")
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

Private Function NormalizeZipEntryName(ByVal entryName As String) As String
    entryName = Trim$(Replace(entryName, "\", "/"))

    Do While Left$(entryName, 2) = "./"
        entryName = Mid$(entryName, 3)
    Loop
    Do While Left$(entryName, 1) = "/"
        entryName = Mid$(entryName, 2)
    Loop
    Do While Right$(entryName, 1) = "/"
        entryName = Left$(entryName, Len(entryName) - 1)
    Loop

    If entryName = "." Then entryName = vbNullString
    NormalizeZipEntryName = entryName
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

Private Function TryGetSelectedInlineImage(ByRef inlineImage As InlineShape) As Boolean
    On Error Resume Next
    If Selection.InlineShapes.Count > 0 Then
        Set inlineImage = Selection.InlineShapes(1)
    End If
    On Error GoTo 0

    If Not inlineImage Is Nothing Then
        If Not IsInlinePictureShape(inlineImage) Then Set inlineImage = Nothing
    End If

    TryGetSelectedInlineImage = Not inlineImage Is Nothing
End Function

Private Function TryGetSelectedInlineOlePackage(ByRef inlineOle As InlineShape) As Boolean
    On Error Resume Next
    If Selection.InlineShapes.Count > 0 Then
        Set inlineOle = Selection.InlineShapes(1)
    End If
    On Error GoTo 0

    If Not inlineOle Is Nothing Then
        If Not IsInlineOlePackage(inlineOle) Then Set inlineOle = Nothing
    End If

    TryGetSelectedInlineOlePackage = Not inlineOle Is Nothing
End Function

Private Function TryGetSelectedFloatingImage(ByRef floatingImage As Shape) As Boolean
    On Error Resume Next
    If Selection.ShapeRange.Count > 0 Then
        Set floatingImage = Selection.ShapeRange(1)
    End If
    On Error GoTo 0

    If Not floatingImage Is Nothing Then
        If Not IsFloatingPictureShape(floatingImage) Then Set floatingImage = Nothing
    End If

    TryGetSelectedFloatingImage = Not floatingImage Is Nothing
End Function

Private Function TryGetSelectedFloatingOlePackage(ByRef floatingOle As Shape) As Boolean
    On Error Resume Next
    If Selection.ShapeRange.Count > 0 Then
        Set floatingOle = Selection.ShapeRange(1)
    End If
    On Error GoTo 0

    If Not floatingOle Is Nothing Then
        If Not IsFloatingOlePackage(floatingOle) Then Set floatingOle = Nothing
    End If

    TryGetSelectedFloatingOlePackage = Not floatingOle Is Nothing
End Function

Private Function IsInlinePictureShape(ByVal inlineImage As Object) As Boolean
    On Error Resume Next
    IsInlinePictureShape = (inlineImage.Type = wdInlineShapePicture Or inlineImage.Type = wdInlineShapeLinkedPicture)
    On Error GoTo 0
End Function

Private Function IsFloatingPictureShape(ByVal floatingImage As Object) As Boolean
    On Error Resume Next
    IsFloatingPictureShape = (floatingImage.Type = msoPicture Or floatingImage.Type = msoLinkedPicture)
    On Error GoTo 0
End Function

Private Function IsInlineOlePackage(ByVal inlineOle As Object) As Boolean
    On Error Resume Next
    IsInlineOlePackage = (inlineOle.Type = wdInlineShapeEmbeddedOLEObject Or inlineOle.Type = wdInlineShapeLinkedOLEObject)
    If IsInlineOlePackage Then
        IsInlineOlePackage = HasOlePackageOpenXml(CStr(inlineOle.Range.WordOpenXML))
    End If
    On Error GoTo 0
End Function

Private Function IsFloatingOlePackage(ByVal floatingOle As Object) As Boolean
    On Error Resume Next
    IsFloatingOlePackage = (floatingOle.Type = msoEmbeddedOLEObject Or floatingOle.Type = msoLinkedOLEObject)
    If IsFloatingOlePackage Then
        IsFloatingOlePackage = (LCase$(CStr(floatingOle.OLEFormat.ProgID)) = "package" Or LCase$(CStr(floatingOle.OLEFormat.ClassType)) = "package")
    End If
    On Error GoTo 0
End Function

Private Function HasOlePackageOpenXml(ByVal openXml As String) As Boolean
    HasOlePackageOpenXml = (InStr(1, openXml, "application/vnd.openxmlformats-officedocument.oleObject", vbTextCompare) > 0 _
        And InStr(1, openXml, "Ole10Native", vbTextCompare) > 0)
End Function

Private Function BuildUsageHelpText() As String
    BuildUsageHelpText = T("usage.text")
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

Private Sub ValidateSupportFiles(ByVal supportFiles As Collection)
    If supportFiles Is Nothing Then
        Err.Raise vbObjectError + 520, "ValidateSupportFiles", "No support files or folders were selected."
    End If

    If supportFiles.Count = 0 Then
        Err.Raise vbObjectError + 521, "ValidateSupportFiles", "No support files or folders were selected."
    End If
End Sub

Private Function ExtractImageFromOpenXml(ByVal openXml As String, ByVal baseName As String) As String
    Dim xmlDoc As Object
    Dim imageParts As Object
    Dim imagePart As Object
    Dim binaryNode As Object
    Dim partName As String
    Dim contentType As String
    Dim extensionName As String
    Dim fso As Object
    Dim outputPath As String

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False

    If Not xmlDoc.LoadXML(openXml) Then
        Err.Raise vbObjectError + 531, "ExtractImageFromOpenXml", "Could not read the selected image XML."
    End If

    xmlDoc.setProperty "SelectionNamespaces", "xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'"
    Set imageParts = xmlDoc.SelectNodes("//pkg:part[starts-with(@pkg:contentType, 'image/')]")
    If imageParts Is Nothing Then
        Err.Raise vbObjectError + 532, "ExtractImageFromOpenXml", "Could not find image data in the selected picture."
    End If
    If imageParts.Length = 0 Then
        Err.Raise vbObjectError + 532, "ExtractImageFromOpenXml", "Could not find image data in the selected picture."
    End If

    Set imagePart = imageParts.Item(0)
    Set binaryNode = imagePart.SelectSingleNode("pkg:binaryData")
    If binaryNode Is Nothing Then
        Set binaryNode = imagePart.SelectSingleNode("*[local-name()='binaryData']")
    End If
    If binaryNode Is Nothing Then
        Err.Raise vbObjectError + 533, "ExtractImageFromOpenXml", "Could not find embedded image bytes in the selected picture."
    End If

    partName = CStr(imagePart.getAttribute("pkg:name"))
    contentType = CStr(imagePart.getAttribute("pkg:contentType"))
    extensionName = GetImageExtensionFromOpenXmlPart(partName, contentType)

    Set fso = CreateObject("Scripting.FileSystemObject")
    outputPath = BuildTempFilePath(fso, SanitizeFileName(baseName), "." & extensionName)
    WriteBase64ToFile CStr(binaryNode.Text), outputPath
    ExtractImageFromOpenXml = outputPath
End Function

Private Function ExtractImageFromInlineShape(ByVal inlineShape As Object, ByVal baseName As String) As String
    Dim imagePath As String
    Dim sourceRange As Range
    Dim openXml As String

    imagePath = TryGetLinkedImageSourceFromInlineShape(inlineShape)
    If Len(imagePath) > 0 Then
        ExtractImageFromInlineShape = imagePath
        Exit Function
    End If

    Set sourceRange = ReadInlineShapeRangeWithRetry(inlineShape, "ExtractImageFromInlineShape")

    On Error Resume Next
    openXml = ReadRangeWordOpenXmlWithRetry(sourceRange, "ExtractImageFromInlineShape")
    If Err.Number = 0 Then
        imagePath = ExtractImageFromOpenXml(openXml, baseName)
    End If
    If Err.Number = 0 And Len(imagePath) > 0 Then
        On Error GoTo 0
        ExtractImageFromInlineShape = imagePath
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ExtractImageFromInlineShape = ExtractDisplayImageFromRangeViaHtml(sourceRange, baseName)
End Function

Private Function TryGetLinkedImageSourceFromInlineShape(ByVal inlineShape As Object) As String
    Dim fso As Object
    Dim sourcePath As String

    On Error Resume Next
    sourcePath = CStr(inlineShape.LinkFormat.SourceFullName)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(sourcePath) > 0 Then
        If fso.FileExists(sourcePath) Then TryGetLinkedImageSourceFromInlineShape = sourcePath
    End If
End Function

Private Function TryGetLinkedImageSourceFromShape(ByVal targetShape As Object) As String
    Dim fso As Object
    Dim sourcePath As String

    On Error Resume Next
    sourcePath = CStr(targetShape.LinkFormat.SourceFullName)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(sourcePath) > 0 Then
        If fso.FileExists(sourcePath) Then TryGetLinkedImageSourceFromShape = sourcePath
    End If
End Function

Private Function ExtractDisplayImageFromRangeViaHtml(ByVal sourceRange As Range, ByVal baseName As String) As String
    Dim fso As Object
    Dim tempDoc As Document
    Dim tempFolderPath As String
    Dim htmlPath As String
    Dim exportedImagePath As String
    Dim outputPath As String
    Dim extensionName As String
    Dim errorMessage As String

    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFolderPath = BuildUniqueTempFolderPath(fso, "OLEPackagerImageExport")
    fso.CreateFolder tempFolderPath
    htmlPath = fso.BuildPath(tempFolderPath, "image.html")

    sourceRange.Select
    On Error Resume Next
    Selection.CopyAsPicture
    If Err.Number <> 0 Then
        Err.Clear
        sourceRange.Copy
    End If
    On Error GoTo Failed

    Set tempDoc = Application.Documents.Add(Visible:=False)
    On Error Resume Next
    tempDoc.Range.PasteSpecial DataType:=wdPasteEnhancedMetafile
    If Err.Number <> 0 Then
        Err.Clear
        tempDoc.Range.Paste
    End If
    On Error GoTo Failed
    tempDoc.SaveAs2 FileName:=htmlPath, FileFormat:=wdFormatFilteredHTML, AddToRecentFiles:=False

    exportedImagePath = FindLargestImageFile(fso, tempFolderPath)
    If Len(exportedImagePath) = 0 Then
        Err.Raise vbObjectError + 534, "ExtractDisplayImageFromRangeViaHtml", "Word did not export an image file for the selected picture."
    End If

    extensionName = LCase$(fso.GetExtensionName(exportedImagePath))
    If Len(extensionName) = 0 Then extensionName = "png"
    outputPath = BuildTempFilePath(fso, SanitizeFileName(baseName) & "_export", "." & extensionName)
    fso.CopyFile exportedImagePath, outputPath, True
    ExtractDisplayImageFromRangeViaHtml = outputPath

    On Error Resume Next
    If Not tempDoc Is Nothing Then tempDoc.Close SaveChanges:=False
    If fso.FolderExists(tempFolderPath) Then fso.DeleteFolder tempFolderPath, True
    On Error GoTo 0
    Exit Function

Failed:
    errorMessage = Err.Description
    On Error Resume Next
    If Not tempDoc Is Nothing Then tempDoc.Close SaveChanges:=False
    If Not fso Is Nothing Then
        If Len(tempFolderPath) > 0 Then
            If fso.FolderExists(tempFolderPath) Then fso.DeleteFolder tempFolderPath, True
        End If
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 534, "ExtractDisplayImageFromRangeViaHtml", "Could not export image data from the selected picture. " & errorMessage
End Function

Private Function BuildUniqueTempFolderPath(ByVal fso As Object, ByVal prefix As String) As String
    Dim basePath As String
    Dim candidatePath As String
    Dim index As Long

    basePath = fso.BuildPath(GetTempFolderPath(fso), prefix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000)))
    candidatePath = basePath
    index = 2
    Do While fso.FolderExists(candidatePath)
        candidatePath = basePath & "_" & CStr(index)
        index = index + 1
    Loop

    BuildUniqueTempFolderPath = candidatePath
End Function

Private Function FindLargestImageFile(ByVal fso As Object, ByVal folderPath As String) As String
    Dim bestPath As String
    Dim bestSize As Double

    FindLargestImageFileRecursive fso, folderPath, bestPath, bestSize
    FindLargestImageFile = bestPath
End Function

Private Sub FindLargestImageFileRecursive(ByVal fso As Object, ByVal folderPath As String, ByRef bestPath As String, ByRef bestSize As Double)
    Dim folder As Object
    Dim subFolder As Object
    Dim fileItem As Object

    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set folder = fso.GetFolder(folderPath)

    For Each fileItem In folder.Files
        If IsImageFileExtension(fso.GetExtensionName(fileItem.Name)) Then
            If CDbl(fileItem.Size) > bestSize Then
                bestPath = CStr(fileItem.Path)
                bestSize = CDbl(fileItem.Size)
            End If
        End If
    Next fileItem

    For Each subFolder In folder.SubFolders
        FindLargestImageFileRecursive fso, CStr(subFolder.Path), bestPath, bestSize
    Next subFolder
End Sub

Private Function IsImageFileExtension(ByVal extensionName As String) As Boolean
    Select Case LCase$(extensionName)
        Case "png", "jpg", "jpeg", "tif", "tiff", "bmp", "gif", "emf", "wmf"
            IsImageFileExtension = True
        Case Else
            IsImageFileExtension = False
    End Select
End Function

Private Function GetImageExtensionFromOpenXmlPart(ByVal partName As String, ByVal contentType As String) As String
    Dim fso As Object
    Dim extensionName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    extensionName = LCase$(fso.GetExtensionName(partName))
    If Len(extensionName) > 0 Then
        GetImageExtensionFromOpenXmlPart = extensionName
        Exit Function
    End If

    Select Case LCase$(contentType)
        Case "image/png"
            GetImageExtensionFromOpenXmlPart = "png"
        Case "image/jpeg", "image/jpg"
            GetImageExtensionFromOpenXmlPart = "jpg"
        Case "image/tiff"
            GetImageExtensionFromOpenXmlPart = "tif"
        Case "image/x-emf"
            GetImageExtensionFromOpenXmlPart = "emf"
        Case Else
            GetImageExtensionFromOpenXmlPart = "png"
    End Select
End Function

Private Function SanitizeFileName(ByVal value As String) As String
    Dim invalidChars As String
    Dim i As Long
    Dim ch As String

    invalidChars = "\/:*?""<>|"
    SanitizeFileName = Trim$(value)
    If Len(SanitizeFileName) = 0 Then SanitizeFileName = "selected_image"

    For i = 1 To Len(invalidChars)
        ch = Mid$(invalidChars, i, 1)
        SanitizeFileName = Replace(SanitizeFileName, ch, "_")
    Next i
End Function

Private Sub WriteBase64ToFile(ByVal base64Text As String, ByVal outputPath As String)
    Dim xmlDoc As Object
    Dim node As Object
    Dim stream As Object

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xmlDoc.createElement("base64")
    node.DataType = "bin.base64"
    node.Text = base64Text

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write node.nodeTypedValue
    stream.SaveToFile outputPath, 2
    stream.Close
End Sub

Private Function ExtractZipFromOleOpenXml(ByVal openXml As String, ByVal baseName As String) As String
    Dim xmlDoc As Object
    Dim oleParts As Object
    Dim olePart As Object
    Dim binaryNode As Object
    Dim fso As Object
    Dim olePath As String

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False

    If Not xmlDoc.LoadXML(openXml) Then
        Err.Raise vbObjectError + 543, "ExtractZipFromOleOpenXml", "Could not read the selected OLE object XML."
    End If

    xmlDoc.setProperty "SelectionNamespaces", "xmlns:pkg='http://schemas.microsoft.com/office/2006/xmlPackage'"
    Set oleParts = xmlDoc.SelectNodes("//pkg:part[@pkg:contentType='application/vnd.openxmlformats-officedocument.oleObject']")
    If oleParts Is Nothing Then
        Err.Raise vbObjectError + 544, "ExtractZipFromOleOpenXml", "Could not find embedded OLE data in the selected object."
    End If
    If oleParts.Length = 0 Then
        Err.Raise vbObjectError + 544, "ExtractZipFromOleOpenXml", "Could not find embedded OLE data in the selected object."
    End If

    Set olePart = oleParts.Item(0)
    Set binaryNode = olePart.SelectSingleNode("pkg:binaryData")
    If binaryNode Is Nothing Then
        Err.Raise vbObjectError + 545, "ExtractZipFromOleOpenXml", "Could not find embedded OLE bytes in the selected object."
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    olePath = BuildTempFilePath(fso, SanitizeFileName(baseName), ".bin")
    WriteBase64ToFile CStr(binaryNode.Text), olePath
    ExtractZipFromOleOpenXml = ExtractZipFromOlePackageFile(olePath, baseName)
End Function

Private Function ExtractZipFromOlePackageFile(ByVal olePath As String, ByVal baseName As String) As String
    Dim fso As Object
    Dim oleBytes() As Byte
    Dim nativeBytes() As Byte
    Dim zipStart As Long
    Dim zipSize As Long
    Dim zipPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    oleBytes = ReadBinaryFile(olePath)
    nativeBytes = ExtractOle10NativeStream(oleBytes)
    zipStart = FindZipStart(nativeBytes)
    If zipStart < 0 Then
        Err.Raise vbObjectError + 546, "ExtractZipFromOlePackageFile", "Could not find zip data inside the selected OLE object."
    End If

    If zipStart >= 4 Then
        zipSize = ReadUInt32AsLong(nativeBytes, zipStart - 4)
    End If
    If zipSize <= 0 Or zipStart + zipSize > ByteArrayLength(nativeBytes) Then
        zipSize = ByteArrayLength(nativeBytes) - zipStart
    End If

    zipPath = BuildTempFilePath(fso, SanitizeFileName(baseName), ".zip")
    WriteByteRangeToFile nativeBytes, zipStart, zipSize, zipPath
    ExtractZipFromOlePackageFile = zipPath
End Function

Private Function ExtractOle10NativeStream(ByRef oleBytes() As Byte) As Byte()
    Dim sectorSize As Long
    Dim miniSectorSize As Long
    Dim numFatSectors As Long
    Dim firstDirSector As Long
    Dim miniStreamCutoff As Long
    Dim firstMiniFatSector As Long
    Dim fat() As Long
    Dim miniFat() As Long
    Dim dirBytes() As Byte
    Dim rootStart As Long
    Dim rootSize As Long
    Dim nativeStart As Long
    Dim nativeSize As Long
    Dim rootStream() As Byte

    If ByteArrayLength(oleBytes) < 512 Then
        Err.Raise vbObjectError + 547, "ExtractOle10NativeStream", "The embedded OLE data is too small."
    End If
    If oleBytes(0) <> &HD0 Or oleBytes(1) <> &HCF Or oleBytes(2) <> &H11 Or oleBytes(3) <> &HE0 Then
        Err.Raise vbObjectError + 548, "ExtractOle10NativeStream", "The embedded OLE data is not a compound file."
    End If

    sectorSize = 2 ^ ReadUInt16AsLong(oleBytes, 30)
    miniSectorSize = 2 ^ ReadUInt16AsLong(oleBytes, 32)
    numFatSectors = ReadUInt32AsLong(oleBytes, 44)
    firstDirSector = ReadInt32LE(oleBytes, 48)
    miniStreamCutoff = ReadUInt32AsLong(oleBytes, 56)
    firstMiniFatSector = ReadInt32LE(oleBytes, 60)

    fat = BuildFatTable(oleBytes, sectorSize, numFatSectors)
    dirBytes = ReadRegularStreamBytes(oleBytes, fat, firstDirSector, sectorSize, -1)
    FindOleDirectoryStreams dirBytes, rootStart, rootSize, nativeStart, nativeSize

    If nativeSize < miniStreamCutoff Then
        rootStream = ReadRegularStreamBytes(oleBytes, fat, rootStart, sectorSize, rootSize)
        miniFat = BuildMiniFatTable(oleBytes, fat, firstMiniFatSector, sectorSize)
        ExtractOle10NativeStream = ReadMiniStreamBytes(rootStream, miniFat, nativeStart, miniSectorSize, nativeSize)
    Else
        ExtractOle10NativeStream = ReadRegularStreamBytes(oleBytes, fat, nativeStart, sectorSize, nativeSize)
    End If
End Function

Private Function BuildFatTable(ByRef oleBytes() As Byte, ByVal sectorSize As Long, ByVal numFatSectors As Long) As Long()
    Dim fatSectorIds As Collection
    Dim fat() As Long
    Dim firstDifatSector As Long
    Dim numDifatSectors As Long
    Dim difatSector As Long
    Dim i As Long
    Dim j As Long
    Dim pos As Long
    Dim fatSectorId As Variant
    Dim sectorBytes() As Byte
    Dim nextDifatSector As Long

    If numFatSectors <= 0 Then
        Err.Raise vbObjectError + 549, "BuildFatTable", "The OLE compound file has no FAT sectors."
    End If

    Set fatSectorIds = New Collection
    For i = 0 To 108
        fatSectorId = ReadInt32LE(oleBytes, 76 + (i * 4))
        If CLng(fatSectorId) >= 0 Then fatSectorIds.Add CLng(fatSectorId)
    Next i

    firstDifatSector = ReadInt32LE(oleBytes, 68)
    numDifatSectors = ReadUInt32AsLong(oleBytes, 72)
    difatSector = firstDifatSector
    For i = 1 To numDifatSectors
        If difatSector < 0 Then Exit For
        sectorBytes = GetCfbSectorBytes(oleBytes, difatSector, sectorSize)
        For j = 0 To 126
            fatSectorId = ReadInt32LE(sectorBytes, j * 4)
            If CLng(fatSectorId) >= 0 Then fatSectorIds.Add CLng(fatSectorId)
        Next j
        nextDifatSector = ReadInt32LE(sectorBytes, sectorSize - 4)
        difatSector = nextDifatSector
    Next i

    If fatSectorIds.Count < numFatSectors Then
        Err.Raise vbObjectError + 550, "BuildFatTable", "The OLE compound file FAT sector list is incomplete."
    End If

    ReDim fat(0 To (numFatSectors * (sectorSize \ 4)) - 1)
    pos = 0
    For i = 0 To numFatSectors - 1
        sectorBytes = GetCfbSectorBytes(oleBytes, CLng(fatSectorIds(i + 1)), sectorSize)
        For j = 0 To sectorSize - 4 Step 4
            fat(pos) = ReadInt32LE(sectorBytes, j)
            pos = pos + 1
        Next j
    Next i

    BuildFatTable = fat
End Function

Private Function BuildMiniFatTable(ByRef oleBytes() As Byte, ByRef fat() As Long, ByVal firstMiniFatSector As Long, ByVal sectorSize As Long) As Long()
    Dim miniFatBytes() As Byte
    Dim miniFat() As Long
    Dim i As Long

    If firstMiniFatSector < 0 Then
        ReDim miniFat(0 To 0)
        BuildMiniFatTable = miniFat
        Exit Function
    End If

    miniFatBytes = ReadRegularStreamBytes(oleBytes, fat, firstMiniFatSector, sectorSize, -1)
    ReDim miniFat(0 To (ByteArrayLength(miniFatBytes) \ 4) - 1)
    For i = 0 To UBound(miniFat)
        miniFat(i) = ReadInt32LE(miniFatBytes, i * 4)
    Next i

    BuildMiniFatTable = miniFat
End Function

Private Function ReadRegularStreamBytes(ByRef oleBytes() As Byte, ByRef fat() As Long, ByVal startSector As Long, ByVal sectorSize As Long, ByVal streamSize As Long) As Byte()
    Dim chain As Collection
    Dim result() As Byte
    Dim sectorId As Variant
    Dim sectorBytes() As Byte
    Dim pos As Long
    Dim resultSize As Long

    Set chain = GetFatSectorChain(fat, startSector)
    If chain.Count = 0 Then
        ReDim result(0 To 0)
        ReadRegularStreamBytes = result
        Exit Function
    End If

    resultSize = chain.Count * sectorSize
    ReDim result(0 To resultSize - 1)
    pos = 0
    For Each sectorId In chain
        sectorBytes = GetCfbSectorBytes(oleBytes, CLng(sectorId), sectorSize)
        CopyBytes sectorBytes, 0, result, pos, sectorSize
        pos = pos + sectorSize
    Next sectorId

    If streamSize >= 0 And streamSize < resultSize Then
        ReDim Preserve result(0 To streamSize - 1)
    End If
    ReadRegularStreamBytes = result
End Function

Private Function ReadMiniStreamBytes(ByRef rootStream() As Byte, ByRef miniFat() As Long, ByVal startMiniSector As Long, ByVal miniSectorSize As Long, ByVal streamSize As Long) As Byte()
    Dim chain As Collection
    Dim result() As Byte
    Dim miniSectorId As Variant
    Dim pos As Long
    Dim sourceOffset As Long
    Dim resultSize As Long

    Set chain = GetFatSectorChain(miniFat, startMiniSector)
    If chain.Count = 0 Or streamSize <= 0 Then
        ReDim result(0 To 0)
        ReadMiniStreamBytes = result
        Exit Function
    End If

    resultSize = chain.Count * miniSectorSize
    ReDim result(0 To resultSize - 1)
    pos = 0
    For Each miniSectorId In chain
        sourceOffset = CLng(miniSectorId) * miniSectorSize
        CopyBytes rootStream, sourceOffset, result, pos, miniSectorSize
        pos = pos + miniSectorSize
    Next miniSectorId

    If streamSize < resultSize Then
        ReDim Preserve result(0 To streamSize - 1)
    End If
    ReadMiniStreamBytes = result
End Function

Private Function GetFatSectorChain(ByRef fat() As Long, ByVal startSector As Long) As Collection
    Dim chain As Collection
    Dim sectorId As Long
    Dim guard As Long

    Set chain = New Collection
    sectorId = startSector
    Do While sectorId >= 0 And sectorId <= UBound(fat)
        chain.Add sectorId
        sectorId = fat(sectorId)
        guard = guard + 1
        If sectorId = CFB_END_OF_CHAIN Or sectorId = CFB_FREE_SECTOR Then Exit Do
        If guard > UBound(fat) + 1 Then
            Err.Raise vbObjectError + 551, "GetFatSectorChain", "The OLE compound file has a circular sector chain."
        End If
    Loop

    Set GetFatSectorChain = chain
End Function

Private Sub FindOleDirectoryStreams(ByRef dirBytes() As Byte, ByRef rootStart As Long, ByRef rootSize As Long, ByRef nativeStart As Long, ByRef nativeSize As Long)
    Dim offset As Long
    Dim entryName As String
    Dim entryType As Long
    Dim nameLength As Long
    Dim foundNative As Boolean

    For offset = 0 To ByteArrayLength(dirBytes) - 128 Step 128
        nameLength = ReadUInt16AsLong(dirBytes, offset + 64)
        entryType = CLng(dirBytes(offset + 66))
        entryName = ReadUtf16String(dirBytes, offset, nameLength)

        If offset = 0 Then
            rootStart = ReadInt32LE(dirBytes, offset + 116)
            rootSize = ReadUInt32AsLong(dirBytes, offset + 120)
        End If

        If entryType = 2 And InStr(1, entryName, "Ole10Native", vbTextCompare) > 0 Then
            nativeStart = ReadInt32LE(dirBytes, offset + 116)
            nativeSize = ReadUInt32AsLong(dirBytes, offset + 120)
            foundNative = True
        End If
    Next offset

    If Not foundNative Then
        Err.Raise vbObjectError + 552, "FindOleDirectoryStreams", "The selected OLE object does not contain an Ole10Native package stream."
    End If
End Sub

Private Function FindZipStart(ByRef bytes() As Byte) As Long
    Dim i As Long

    For i = 0 To ByteArrayLength(bytes) - 4
        If bytes(i) = &H50 And bytes(i + 1) = &H4B Then
            If (bytes(i + 2) = &H3 And bytes(i + 3) = &H4) Or (bytes(i + 2) = &H5 And bytes(i + 3) = &H6) Then
                FindZipStart = i
                Exit Function
            End If
        End If
    Next i

    FindZipStart = -1
End Function

Private Function GetCfbSectorBytes(ByRef bytes() As Byte, ByVal sectorId As Long, ByVal sectorSize As Long) As Byte()
    Dim offset As Long

    If sectorId < 0 Then
        Err.Raise vbObjectError + 553, "GetCfbSectorBytes", "Invalid OLE sector id."
    End If
    offset = (sectorId + 1) * sectorSize
    GetCfbSectorBytes = CopyByteRange(bytes, offset, sectorSize)
End Function

Private Function ReadBinaryFile(ByVal filePath As String) As Byte()
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile filePath
    ReadBinaryFile = stream.Read
    stream.Close
End Function

Private Sub WriteByteRangeToFile(ByRef bytes() As Byte, ByVal startOffset As Long, ByVal byteCount As Long, ByVal outputPath As String)
    Dim stream As Object
    Dim outputBytes() As Byte

    outputBytes = CopyByteRange(bytes, startOffset, byteCount)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write outputBytes
    stream.SaveToFile outputPath, 2
    stream.Close
End Sub

Private Function CopyByteRange(ByRef bytes() As Byte, ByVal startOffset As Long, ByVal byteCount As Long) As Byte()
    Dim result() As Byte
    Dim i As Long

    If byteCount <= 0 Then
        ReDim result(0 To 0)
        CopyByteRange = result
        Exit Function
    End If

    ReDim result(0 To byteCount - 1)
    For i = 0 To byteCount - 1
        result(i) = bytes(startOffset + i)
    Next i
    CopyByteRange = result
End Function

Private Sub CopyBytes(ByRef source() As Byte, ByVal sourceOffset As Long, ByRef destination() As Byte, ByVal destinationOffset As Long, ByVal byteCount As Long)
    Dim i As Long

    For i = 0 To byteCount - 1
        destination(destinationOffset + i) = source(sourceOffset + i)
    Next i
End Sub

Private Function ByteArrayLength(ByRef bytes() As Byte) As Long
    On Error GoTo EmptyArray
    ByteArrayLength = UBound(bytes) - LBound(bytes) + 1
    Exit Function

EmptyArray:
    ByteArrayLength = 0
End Function

Private Function ReadUInt16AsLong(ByRef bytes() As Byte, ByVal offset As Long) As Long
    ReadUInt16AsLong = CLng(bytes(offset)) + (CLng(bytes(offset + 1)) * 256&)
End Function

Private Function ReadUInt32AsLong(ByRef bytes() As Byte, ByVal offset As Long) As Long
    Dim value As Double

    value = CDbl(bytes(offset)) _
        + (CDbl(bytes(offset + 1)) * 256#) _
        + (CDbl(bytes(offset + 2)) * 65536#) _
        + (CDbl(bytes(offset + 3)) * 16777216#)
    If value > 2147483647# Then
        Err.Raise vbObjectError + 554, "ReadUInt32AsLong", "OLE package value is too large."
    End If
    ReadUInt32AsLong = CLng(value)
End Function

Private Function ReadInt32LE(ByRef bytes() As Byte, ByVal offset As Long) As Long
    Dim value As Double

    value = CDbl(bytes(offset)) _
        + (CDbl(bytes(offset + 1)) * 256#) _
        + (CDbl(bytes(offset + 2)) * 65536#) _
        + (CDbl(bytes(offset + 3)) * 16777216#)
    If value >= 2147483648# Then value = value - 4294967296#
    ReadInt32LE = CLng(value)
End Function

Private Function ReadUtf16String(ByRef bytes() As Byte, ByVal offset As Long, ByVal byteCount As Long) As String
    Dim i As Long
    Dim code As Long
    Dim result As String

    If byteCount <= 2 Then
        ReadUtf16String = vbNullString
        Exit Function
    End If

    For i = 0 To byteCount - 3 Step 2
        code = CLng(bytes(offset + i)) + (CLng(bytes(offset + i + 1)) * 256&)
        If code <> 0 Then result = result & UnicodeChar(code)
    Next i
    ReadUtf16String = result
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
    ValidateSupportFiles supportFiles
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

Private Sub CreateZipFromWorkingFolder(ByVal workingFolderPath As String, ByVal zipPath As String, ByVal allowEmpty As Boolean)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    workingFolderPath = Trim$(workingFolderPath)

    If Len(workingFolderPath) = 0 Or Not fso.FolderExists(workingFolderPath) Then
        Err.Raise vbObjectError + 557, "CreateZipFromWorkingFolder", "Package contents folder does not exist: " & workingFolderPath
    End If

    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    If FolderHasAnyFiles(fso.GetFolder(workingFolderPath)) Then
        CreateZipFromPath workingFolderPath, zipPath, False
    ElseIf allowEmpty Then
        CreateEmptyZip zipPath
    Else
        Err.Raise vbObjectError + 558, "CreateZipFromWorkingFolder", "Package contents folder is empty: " & workingFolderPath
    End If
End Sub

Private Sub CreateZipFromManagedFiles(ByVal existingZipPath As String, ByVal keepEntryNames As Collection, ByVal newFiles As Collection, ByVal zipPath As String, ByVal baseName As String)
    Dim fso As Object
    Dim stagingFolder As String
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    stagingFolder = BuildTempStagingFolderPath(fso, baseName)

    On Error GoTo Failed
    If Not fso.FileExists(existingZipPath) Then
        Err.Raise vbObjectError + 555, "CreateZipFromManagedFiles", "Existing embedded zip file does not exist: " & existingZipPath
    End If

    DeleteFolderIfExists fso, stagingFolder
    fso.CreateFolder stagingFolder
    ExtractZipToFolder existingZipPath, stagingFolder
    PruneStagedZipEntries fso, stagingFolder, keepEntryNames
    StageAdditionalSupportFiles fso, newFiles, stagingFolder

    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
    If FolderHasAnyFiles(fso.GetFolder(stagingFolder)) Then
        CreateZipFromPath stagingFolder, zipPath, False
    Else
        CreateEmptyZip zipPath
    End If

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

Private Sub ExtractZipToFolder(ByVal zipPath As String, ByVal targetFolder As String)
    RunZipTool "extract " & QuoteForCommandLine(zipPath) & " " & QuoteForCommandLine(targetFolder), "ExtractZipToFolder"
End Sub

Private Sub PruneStagedZipEntries(ByVal fso As Object, ByVal stagingFolder As String, ByVal keepEntryNames As Collection)
    Dim keepSet As Object

    Set keepSet = CreateObject("Scripting.Dictionary")
    AddKeepEntries keepSet, keepEntryNames
    PruneFilesInFolder fso, fso.GetFolder(stagingFolder), stagingFolder, keepSet
    DeleteEmptySubfolders fso.GetFolder(stagingFolder)
End Sub

Private Sub AddKeepEntries(ByVal keepSet As Object, ByVal keepEntryNames As Collection)
    Dim entryName As Variant
    Dim normalizedEntry As String

    If keepEntryNames Is Nothing Then Exit Sub
    For Each entryName In keepEntryNames
        normalizedEntry = NormalizeZipEntryName(CStr(entryName))
        If Len(normalizedEntry) > 0 Then keepSet(LCase$(normalizedEntry)) = True
    Next entryName
End Sub

Private Sub PruneFilesInFolder(ByVal fso As Object, ByVal folder As Object, ByVal rootFolderPath As String, ByVal keepSet As Object)
    Dim fileItem As Object
    Dim subFolder As Object
    Dim relativeName As String

    For Each fileItem In folder.Files
        relativeName = GetRelativeZipEntryName(rootFolderPath, CStr(fileItem.Path))
        If Not keepSet.Exists(LCase$(relativeName)) Then
            fso.DeleteFile CStr(fileItem.Path), True
        End If
    Next fileItem

    For Each subFolder In folder.SubFolders
        PruneFilesInFolder fso, subFolder, rootFolderPath, keepSet
    Next subFolder
End Sub

Private Sub StageAdditionalSupportFiles(ByVal fso As Object, ByVal supportFiles As Collection, ByVal stagingFolder As String)
    Dim usedNames As Object
    Dim existingFile As Object
    Dim existingFolder As Object
    Dim supportPath As Variant

    If supportFiles Is Nothing Then Exit Sub

    Set usedNames = CreateObject("Scripting.Dictionary")
    For Each existingFile In fso.GetFolder(stagingFolder).Files
        usedNames(LCase$(CStr(existingFile.Name))) = True
    Next existingFile
    For Each existingFolder In fso.GetFolder(stagingFolder).SubFolders
        usedNames(LCase$(CStr(existingFolder.Name))) = True
    Next existingFolder

    For Each supportPath In supportFiles
        StageSupportPath fso, CStr(supportPath), stagingFolder, usedNames, "StageAdditionalSupportFiles"
    Next supportPath
End Sub

Private Function FolderHasAnyFiles(ByVal folder As Object) As Boolean
    Dim subFolder As Object

    If folder.Files.Count > 0 Then
        FolderHasAnyFiles = True
        Exit Function
    End If

    For Each subFolder In folder.SubFolders
        If FolderHasAnyFiles(subFolder) Then
            FolderHasAnyFiles = True
            Exit Function
        End If
    Next subFolder
End Function

Private Sub DeleteEmptySubfolders(ByVal folder As Object)
    Dim subFolder As Object
    Dim subFoldersToCheck As Collection
    Dim item As Variant

    Set subFoldersToCheck = New Collection
    For Each subFolder In folder.SubFolders
        subFoldersToCheck.Add subFolder
    Next subFolder

    For Each item In subFoldersToCheck
        DeleteEmptySubfolders item
        If item.Files.Count = 0 And item.SubFolders.Count = 0 Then item.Delete True
    Next item
End Sub

Private Function GetRelativeZipEntryName(ByVal rootFolderPath As String, ByVal filePath As String) As String
    Dim rootWithSlash As String
    Dim relativePath As String

    rootWithSlash = rootFolderPath
    If Right$(rootWithSlash, 1) <> "\" Then rootWithSlash = rootWithSlash & "\"
    relativePath = Mid$(filePath, Len(rootWithSlash) + 1)
    GetRelativeZipEntryName = NormalizeZipEntryName(Replace(relativePath, "\", "/"))
End Function

Private Function BuildTempStagingFolderPath(ByVal fso As Object, ByVal baseName As String) As String
    Dim stamp As String

    stamp = Format$(Now, "yyyymmdd_hhnnss")
    BuildTempStagingFolderPath = fso.BuildPath(GetTempFolderPath(fso), "OLEPackager_" & baseName & "_" & stamp)
End Function

Private Function BuildTempFilePath(ByVal fso As Object, ByVal prefix As String, ByVal extensionName As String) As String
    Dim stamp As String

    stamp = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000))
    BuildTempFilePath = fso.BuildPath(GetTempFolderPath(fso), prefix & "_" & stamp & extensionName)
End Function

Private Sub StageSupportFiles(ByVal fso As Object, ByVal supportFiles As Collection, ByVal stagingFolder As String)
    Dim usedNames As Object
    Dim supportPath As Variant

    If supportFiles Is Nothing Then
        Err.Raise vbObjectError + 520, "StageSupportFiles", "No support files or folders were selected."
    End If

    If supportFiles.Count = 0 Then
        Err.Raise vbObjectError + 521, "StageSupportFiles", "No support files or folders were selected."
    End If

    DeleteFolderIfExists fso, stagingFolder
    fso.CreateFolder stagingFolder
    Set usedNames = CreateObject("Scripting.Dictionary")

    For Each supportPath In supportFiles
        StageSupportPath fso, CStr(supportPath), stagingFolder, usedNames, "StageSupportFiles"
    Next supportPath
End Sub

Private Sub StageSupportPath(ByVal fso As Object, ByVal supportPath As String, ByVal stagingFolder As String, ByVal usedNames As Object, ByVal errorSource As String)
    Dim targetName As String
    Dim targetPath As String
    Dim sourceFolder As Object

    supportPath = Trim$(supportPath)

    If fso.FileExists(supportPath) Then
        targetName = GetUniqueStagedFileName(fso, usedNames, fso.GetFileName(supportPath))
        fso.CopyFile supportPath, fso.BuildPath(stagingFolder, targetName), True
    ElseIf fso.FolderExists(supportPath) Then
        Set sourceFolder = fso.GetFolder(supportPath)
        targetName = CStr(sourceFolder.Name)
        If Len(targetName) = 0 Then targetName = "folder"
        targetName = GetUniqueStagedFolderName(usedNames, targetName)
        targetPath = fso.BuildPath(stagingFolder, targetName)
        CopySupportFolderToStaging fso, CStr(sourceFolder.Path), targetPath
    Else
        Err.Raise vbObjectError + 522, errorSource, "Support file or folder does not exist: " & supportPath
    End If
End Sub

Private Function GetUniqueStagedFolderName(ByVal usedNames As Object, ByVal folderName As String) As String
    Dim candidate As String
    Dim index As Long

    candidate = folderName
    index = 2

    Do While usedNames.Exists(LCase$(candidate))
        candidate = folderName & "_" & CStr(index)
        index = index + 1
    Loop

    usedNames.Add LCase$(candidate), True
    GetUniqueStagedFolderName = candidate
End Function

Private Sub CopySupportFolderToStaging(ByVal fso As Object, ByVal sourceFolderPath As String, ByVal targetFolderPath As String)
    Dim sourceFolder As Object
    Dim fileItem As Object
    Dim subFolder As Object

    Set sourceFolder = fso.GetFolder(sourceFolderPath)
    If Not fso.FolderExists(targetFolderPath) Then fso.CreateFolder targetFolderPath

    For Each fileItem In sourceFolder.Files
        fso.CopyFile CStr(fileItem.Path), fso.BuildPath(targetFolderPath, CStr(fileItem.Name)), True
    Next fileItem

    For Each subFolder In sourceFolder.SubFolders
        CopySupportFolderToStaging fso, CStr(subFolder.Path), fso.BuildPath(targetFolderPath, CStr(subFolder.Name))
    Next subFolder
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

Private Sub DeleteFileIfExists(ByVal fso As Object, ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) > 0 Then
        If fso.FileExists(filePath) Then fso.DeleteFile filePath, True
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub CreateZipFromPath(ByVal sourcePath As String, ByVal zipPath As String, Optional ByVal excludeGeneratedPlotFiles As Boolean = True)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    CreateZipWithZipTool sourcePath, zipPath, excludeGeneratedPlotFiles
End Sub

Private Sub CreateZipWithTar(ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean)
    Dim fso As Object
    Dim wsh As Object
    Dim command As String
    Dim itemArguments As String
    Dim exitCode As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    If fso.FolderExists(sourcePath) Then
        itemArguments = BuildTarFolderItemArguments(fso, sourcePath, excludeGeneratedPlotFiles)
        If Len(itemArguments) = 0 Then
            CreateEmptyZip zipPath
            Exit Sub
        End If
        command = "cmd.exe /c tar.exe -a -cf " & QuoteForCommandLine(zipPath) & " -C " & QuoteForCommandLine(sourcePath) & " -- " & itemArguments
    Else
        command = "cmd.exe /c tar.exe -a -cf " & QuoteForCommandLine(zipPath) & " -C " & QuoteForCommandLine(fso.GetParentFolderName(sourcePath)) & " -- " & QuoteForCommandLine(fso.GetFileName(sourcePath))
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

Private Function BuildTarFolderItemArguments(ByVal fso As Object, ByVal sourceFolderPath As String, ByVal excludeGeneratedPlotFiles As Boolean) As String
    Dim sourceFolder As Object
    Dim subFolder As Object
    Dim fileItem As Object
    Dim arguments As String

    Set sourceFolder = fso.GetFolder(sourceFolderPath)

    For Each subFolder In sourceFolder.SubFolders
        arguments = arguments & " " & QuoteForCommandLine(subFolder.Name)
    Next subFolder

    For Each fileItem In sourceFolder.Files
        If Not excludeGeneratedPlotFiles Or Not IsGeneratedPlotFile(fileItem.Name) Then
            arguments = arguments & " " & QuoteForCommandLine(fileItem.Name)
        End If
    Next fileItem

    BuildTarFolderItemArguments = Trim$(arguments)
End Function

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

Private Sub CreateZipWithZipTool(ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    RunZipTool "create " & QuoteForCommandLine(sourcePath) & " " & QuoteForCommandLine(zipPath) & " " & ZipToolBoolean(excludeGeneratedPlotFiles), "CreateZipWithZipTool"

    If Dir$(zipPath) = vbNullString Then
        Err.Raise vbObjectError + 513, "CreateZipWithZipTool", "Zip tool did not create zip file: " & zipPath
    End If
End Sub

Private Sub RunZipTool(ByVal arguments As String, ByVal errorSource As String)
    Dim fso As Object
    Dim wsh As Object
    Dim toolPath As String
    Dim command As String
    Dim exitCode As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    toolPath = GetZipToolPath(fso)
    command = QuoteForCommandLine(toolPath) & " " & arguments

    Set wsh = CreateObject("WScript.Shell")
    exitCode = wsh.Run(command, 0, True)
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 562, errorSource, "OLEPackagerZipTool failed with exit code " & exitCode & ": " & command
    End If
End Sub

Private Function GetZipToolPath(ByVal fso As Object) As String
    Dim envPath As String
    Dim candidatePath As String

    envPath = Trim$(Environ$("OLE_RAW_DATA_ZIP_TOOL"))
    If Len(envPath) > 0 Then
        If fso.FileExists(envPath) Then
            GetZipToolPath = envPath
            Exit Function
        End If
    End If

    candidatePath = fso.BuildPath(Application.StartupPath, "OLEPackagerZipTool.exe")
    If fso.FileExists(candidatePath) Then
        GetZipToolPath = candidatePath
        Exit Function
    End If

    On Error Resume Next
    candidatePath = fso.BuildPath(ThisDocument.Path, "OLEPackagerZipTool.exe")
    If Err.Number = 0 Then
        If fso.FileExists(candidatePath) Then
            GetZipToolPath = candidatePath
            On Error GoTo 0
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0

    Err.Raise vbObjectError + 563, "GetZipToolPath", "Could not find OLEPackagerZipTool.exe. Reinstall the add-in."
End Function

Private Function ZipToolBoolean(ByVal value As Boolean) As String
    If value Then
        ZipToolBoolean = "true"
    Else
        ZipToolBoolean = "false"
    End If
End Function

Private Sub CreateZipWithPowerShell(ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean)
    Dim fso As Object
    Dim wsh As Object
    Dim command As String
    Dim scriptPath As String
    Dim logPath As String
    Dim exitCode As Long
    Dim stage As String

    On Error GoTo Failed

    stage = "prepare zip script"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True

    scriptPath = BuildTempFilePath(fso, "OLEPackagerZip", ".ps1")
    logPath = BuildTempFilePath(fso, "OLEPackagerZip", ".log")
    WriteZipScript scriptPath, sourcePath, zipPath, excludeGeneratedPlotFiles

    stage = "run zip script"
    Set wsh = CreateObject("WScript.Shell")
    command = BuildPowerShellScriptCommand(scriptPath, logPath)
    exitCode = wsh.Run(command, 0, True)

    stage = "check zip script result"
    If exitCode <> 0 Then
        Err.Raise vbObjectError + 512, "CreateZipWithPowerShell", "PowerShell ZipArchive failed with exit code " & exitCode & ". Script kept at: " & scriptPath & ". Log: " & ReadFailureLogIfExists(logPath)
    End If

    If Dir$(zipPath) = vbNullString Then
        Err.Raise vbObjectError + 513, "CreateZipWithPowerShell", "PowerShell did not create zip file: " & zipPath
    End If
    DeleteFileIfExists fso, scriptPath
    DeleteFileIfExists fso, logPath
    Exit Sub

Failed:
    Err.Raise Err.Number, "CreateZipWithPowerShell", stage & ": " & Err.Description
End Sub

Private Function PowerShellQuote(ByVal value As String) As String
    PowerShellQuote = "'" & Replace(value, "'", "''") & "'"
End Function

Private Function PowerShellBoolean(ByVal value As Boolean) As String
    If value Then
        PowerShellBoolean = "$true"
    Else
        PowerShellBoolean = "$false"
    End If
End Function

Private Function QuoteForCommandLine(ByVal value As String) As String
    QuoteForCommandLine = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function BuildPowerShellScriptCommand(ByVal scriptPath As String, Optional ByVal logPath As String = vbNullString) As String
    BuildPowerShellScriptCommand = "cmd.exe /c powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & QuoteForCommandLine(scriptPath)
    If Len(logPath) > 0 Then BuildPowerShellScriptCommand = BuildPowerShellScriptCommand & " > " & QuoteForCommandLine(logPath) & " 2>&1"
End Function

Private Function ReadFailureLogIfExists(ByVal logPath As String) As String
    Dim fso As Object
    Dim fileHandle As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(logPath) = 0 Or Not fso.FileExists(logPath) Then
        ReadFailureLogIfExists = logPath
        Exit Function
    End If

    Set fileHandle = fso.OpenTextFile(logPath, 1, False)
    ReadFailureLogIfExists = logPath & " " & Left$(fileHandle.ReadAll, 500)
    fileHandle.Close
    Err.Clear
    On Error GoTo 0
End Function

Private Sub AddScriptLine(ByRef scriptText As String, ByVal lineText As String)
    If Len(scriptText) > 0 Then scriptText = scriptText & vbCrLf
    scriptText = scriptText & lineText
End Sub

Private Sub WriteZipScript(ByVal scriptPath As String, ByVal sourcePath As String, ByVal zipPath As String, ByVal excludeGeneratedPlotFiles As Boolean)
    Dim scriptText As String

    AddScriptLine scriptText, "$ErrorActionPreference = 'Stop'"
    AddScriptLine scriptText, "Add-Type -AssemblyName System.IO.Compression"
    AddScriptLine scriptText, "Add-Type -AssemblyName System.IO.Compression.FileSystem"
    AddScriptLine scriptText, "$sourcePath = " & PowerShellQuote(sourcePath)
    AddScriptLine scriptText, "$destination = " & PowerShellQuote(zipPath)
    AddScriptLine scriptText, "$excludeGenerated = " & PowerShellBoolean(excludeGeneratedPlotFiles)
    AddScriptLine scriptText, "$utf8 = [System.Text.Encoding]::UTF8"
    AddScriptLine scriptText, "$generated = @('plot.png', 'plot.svg', 'plot.pdf')"
    AddScriptLine scriptText, "$source = Get-Item -LiteralPath $sourcePath -Force"
    AddScriptLine scriptText, "$destinationFolder = [System.IO.Path]::GetDirectoryName($destination)"
    AddScriptLine scriptText, "if ($destinationFolder -and -not (Test-Path -LiteralPath $destinationFolder)) { New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null }"
    AddScriptLine scriptText, "if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Force }"
    AddScriptLine scriptText, "function Get-ZipEntryName([string]$fullName, [string]$root) {"
    AddScriptLine scriptText, "    $relative = $fullName.Substring($root.Length).TrimStart([char[]]@('\', '/'))"
    AddScriptLine scriptText, "    return $relative.Replace('\', '/')"
    AddScriptLine scriptText, "}"
    AddScriptLine scriptText, "function Add-EntryFile($archive, $file, [string]$entryName) {"
    AddScriptLine scriptText, "    if ([string]::IsNullOrWhiteSpace($entryName)) { return }"
    AddScriptLine scriptText, "    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $entryName, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null"
    AddScriptLine scriptText, "}"
    AddScriptLine scriptText, "function Add-EmptyDirectoryEntry($archive, [string]$entryName) {"
    AddScriptLine scriptText, "    if ([string]::IsNullOrWhiteSpace($entryName)) { return }"
    AddScriptLine scriptText, "    if (-not $entryName.EndsWith('/')) { $entryName = $entryName + '/' }"
    AddScriptLine scriptText, "    $archive.CreateEntry($entryName) | Out-Null"
    AddScriptLine scriptText, "}"
    AddScriptLine scriptText, "$stream = [System.IO.File]::Open($destination, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)"
    AddScriptLine scriptText, "try {"
    AddScriptLine scriptText, "    $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Create, $false, $utf8)"
    AddScriptLine scriptText, "    try {"
    AddScriptLine scriptText, "        if ($source.PSIsContainer) {"
    AddScriptLine scriptText, "            $root = $source.FullName.TrimEnd([char[]]@('\', '/')) + [System.IO.Path]::DirectorySeparatorChar"
    AddScriptLine scriptText, "            foreach ($file in Get-ChildItem -LiteralPath $source.FullName -Force -File -Recurse) {"
    AddScriptLine scriptText, "                if ($excludeGenerated -and $file.DirectoryName.TrimEnd([char[]]@('\', '/')).Equals($source.FullName.TrimEnd([char[]]@('\', '/')), [System.StringComparison]::OrdinalIgnoreCase) -and ($generated -contains $file.Name.ToLowerInvariant())) { continue }"
    AddScriptLine scriptText, "                Add-EntryFile $archive $file (Get-ZipEntryName $file.FullName $root)"
    AddScriptLine scriptText, "            }"
    AddScriptLine scriptText, "            foreach ($directory in Get-ChildItem -LiteralPath $source.FullName -Force -Directory -Recurse) {"
    AddScriptLine scriptText, "                $children = @(Get-ChildItem -LiteralPath $directory.FullName -Force)"
    AddScriptLine scriptText, "                if ($children.Count -eq 0) { Add-EmptyDirectoryEntry $archive (Get-ZipEntryName $directory.FullName $root) }"
    AddScriptLine scriptText, "            }"
    AddScriptLine scriptText, "        } else {"
    AddScriptLine scriptText, "            Add-EntryFile $archive $source $source.Name"
    AddScriptLine scriptText, "        }"
    AddScriptLine scriptText, "    } finally {"
    AddScriptLine scriptText, "        $archive.Dispose()"
    AddScriptLine scriptText, "    }"
    AddScriptLine scriptText, "} finally {"
    AddScriptLine scriptText, "    $stream.Dispose()"
    AddScriptLine scriptText, "}"

    WriteUtf8TextFile scriptPath, scriptText
End Sub

Private Sub WriteZipListScript(ByVal scriptPath As String, ByVal zipPath As String, ByVal listPath As String)
    Dim scriptText As String

    AddScriptLine scriptText, "$ErrorActionPreference = 'Stop'"
    AddScriptLine scriptText, "Add-Type -AssemblyName System.IO.Compression"
    AddScriptLine scriptText, "$zipPath = " & PowerShellQuote(zipPath)
    AddScriptLine scriptText, "$listPath = " & PowerShellQuote(listPath)
    AddScriptLine scriptText, "$utf8 = [System.Text.Encoding]::UTF8"
    AddScriptLine scriptText, "$stream = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)"
    AddScriptLine scriptText, "try {"
    AddScriptLine scriptText, "    $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false, $utf8)"
    AddScriptLine scriptText, "    try {"
    AddScriptLine scriptText, "        $entries = @($archive.Entries | ForEach-Object { $_.FullName })"
    AddScriptLine scriptText, "    } finally {"
    AddScriptLine scriptText, "        $archive.Dispose()"
    AddScriptLine scriptText, "    }"
    AddScriptLine scriptText, "} finally {"
    AddScriptLine scriptText, "    $stream.Dispose()"
    AddScriptLine scriptText, "}"
    AddScriptLine scriptText, "[System.IO.File]::WriteAllLines($listPath, [string[]]$entries, $utf8)"

    WriteUtf8TextFile scriptPath, scriptText
End Sub

Private Sub WriteZipExtractScript(ByVal scriptPath As String, ByVal zipPath As String, ByVal targetFolder As String)
    Dim scriptText As String

    AddScriptLine scriptText, "$ErrorActionPreference = 'Stop'"
    AddScriptLine scriptText, "Add-Type -AssemblyName System.IO.Compression"
    AddScriptLine scriptText, "Add-Type -AssemblyName System.IO.Compression.FileSystem"
    AddScriptLine scriptText, "$zipPath = " & PowerShellQuote(zipPath)
    AddScriptLine scriptText, "$targetFolder = " & PowerShellQuote(targetFolder)
    AddScriptLine scriptText, "$utf8 = [System.Text.Encoding]::UTF8"
    AddScriptLine scriptText, "[System.IO.Directory]::CreateDirectory($targetFolder) | Out-Null"
    AddScriptLine scriptText, "$root = [System.IO.Path]::GetFullPath($targetFolder).TrimEnd([char[]]@([System.IO.Path]::DirectorySeparatorChar, [System.IO.Path]::AltDirectorySeparatorChar))"
    AddScriptLine scriptText, "$rootPrefix = $root + [System.IO.Path]::DirectorySeparatorChar"
    AddScriptLine scriptText, "$stream = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)"
    AddScriptLine scriptText, "try {"
    AddScriptLine scriptText, "    $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false, $utf8)"
    AddScriptLine scriptText, "    try {"
    AddScriptLine scriptText, "        foreach ($entry in $archive.Entries) {"
    AddScriptLine scriptText, "            $entryName = $entry.FullName.Replace('\', '/')"
    AddScriptLine scriptText, "            if ([string]::IsNullOrWhiteSpace($entryName)) { continue }"
    AddScriptLine scriptText, "            if ($entryName.StartsWith('/') -or $entryName.Contains(':')) { throw ('Unsafe zip entry name: ' + $entry.FullName) }"
    AddScriptLine scriptText, "            $relativePath = $entryName.Replace('/', [System.IO.Path]::DirectorySeparatorChar)"
    AddScriptLine scriptText, "            $targetPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($root, $relativePath))"
    AddScriptLine scriptText, "            if (-not $targetPath.Equals($root, [System.StringComparison]::OrdinalIgnoreCase) -and -not $targetPath.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { throw ('Unsafe zip entry path: ' + $entry.FullName) }"
    AddScriptLine scriptText, "            if ($entryName.EndsWith('/')) {"
    AddScriptLine scriptText, "                [System.IO.Directory]::CreateDirectory($targetPath) | Out-Null"
    AddScriptLine scriptText, "            } else {"
    AddScriptLine scriptText, "                $parent = [System.IO.Path]::GetDirectoryName($targetPath)"
    AddScriptLine scriptText, "                if ($parent) { [System.IO.Directory]::CreateDirectory($parent) | Out-Null }"
    AddScriptLine scriptText, "                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)"
    AddScriptLine scriptText, "            }"
    AddScriptLine scriptText, "        }"
    AddScriptLine scriptText, "    } finally {"
    AddScriptLine scriptText, "        $archive.Dispose()"
    AddScriptLine scriptText, "    }"
    AddScriptLine scriptText, "} finally {"
    AddScriptLine scriptText, "    $stream.Dispose()"
    AddScriptLine scriptText, "}"

    WriteUtf8TextFile scriptPath, scriptText
End Sub

Private Sub WriteUtf8TextFile(ByVal filePath As String, ByVal text As String)
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
End Sub

Private Function ReadUtf8TextFile(ByVal filePath As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadUtf8TextFile = stream.ReadText(-1)
    stream.Close
    Set stream = Nothing
End Function

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
