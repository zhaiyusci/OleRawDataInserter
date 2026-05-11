VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImageSupportFilesDialog 
   Caption         =   "Insert Image + Package Folder"
   ClientHeight    =   5940
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9000
   OleObjectBlob   =   "ImageSupportFilesDialog.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ImageSupportFilesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private attachToSelectedImage As Boolean
Private manageExistingOle As Boolean
Private targetInlineImage As Object
Private targetFloatingImage As Object
Private targetInlineOle As Object
Private targetFloatingOle As Object
Private existingZipPath As String
Private existingDisplayImagePath As String
Private workingFolderPath As String
Private openedWorkingFolder As Boolean

Private Const DIALOG_WIDTH As Single = 500
Private Const DIALOG_HEIGHT As Single = 330

Public Sub ConfigureForSelectedInlineImage(ByVal inlineImage As Object)
    attachToSelectedImage = True
    manageExistingOle = False
    Set targetInlineImage = inlineImage
    PrepareEmptyWorkingFolder "selected_image_support"
    ConfigureForSelectedImageMode
End Sub

Public Sub ConfigureForSelectedFloatingImage(ByVal floatingImage As Object)
    attachToSelectedImage = True
    manageExistingOle = False
    Set targetFloatingImage = floatingImage
    PrepareEmptyWorkingFolder "selected_image_support"
    ConfigureForSelectedImageMode
End Sub

Public Sub ConfigureForSelectedInlineOle(ByVal inlineOle As Object)
    attachToSelectedImage = False
    manageExistingOle = True
    Set targetInlineOle = inlineOle
    existingDisplayImagePath = ExtractDisplayImageFromInlineOleObject(inlineOle)
    existingZipPath = ExtractZipFromInlineOleObject(inlineOle)
    PrepareWorkingFolderFromExistingZip existingZipPath, "selected_ole_support"
    ConfigureForSelectedOleMode
End Sub

Public Sub ConfigureForSelectedFloatingOle(ByVal floatingOle As Object)
    attachToSelectedImage = False
    manageExistingOle = True
    Set targetFloatingOle = floatingOle
    existingDisplayImagePath = ExtractDisplayImageFromFloatingOleObject(targetFloatingOle)
    existingZipPath = ExtractZipFromFloatingOleObject(targetFloatingOle)
    PrepareWorkingFolderFromExistingZip existingZipPath, "selected_ole_support"
    ConfigureForSelectedOleMode
End Sub

Private Sub UserForm_Initialize()
    ConfigureSystemFileManagerControls
    UpdateDialogState
End Sub

Private Sub UserForm_Activate()
    EnsureWorkingFolder
End Sub

Private Sub cmdBrowseImage_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = T("dialog.fileDialog.chooseImage")
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add T("dialog.filter.images"), "*.png;*.jpg;*.jpeg;*.tif;*.tiff"
        .Filters.Add T("dialog.filter.png"), "*.png"
        .Filters.Add T("dialog.filter.jpeg"), "*.jpg;*.jpeg"
        .Filters.Add T("dialog.filter.tiff"), "*.tif;*.tiff"
        .Filters.Add T("dialog.filter.all"), "*.*"
        If .Show = -1 Then txtImagePath.Text = .SelectedItems(1)
    End With

    UpdateDialogState
End Sub

Private Sub cmdAddFiles_Click()
    EnsureWorkingFolder
    OpenWorkingFolder
End Sub

Private Sub cmdAddFolder_Click()
    EnsureWorkingFolder
    OpenWorkingFolder
End Sub

Private Sub cmdRemoveSelected_Click()
    EnsureWorkingFolder
    OpenWorkingFolder
End Sub

Private Sub cmdClearFiles_Click()
    EnsureWorkingFolder
    If MsgBox(T("msg.clearFolder"), vbQuestion + vbYesNo) = vbYes Then
        ClearWorkingFolder
        OpenWorkingFolder
    End If
    UpdateDialogState
End Sub

Private Sub cmdInsert_Click()
    Dim imagePath As String

    EnsureWorkingFolder

    If Not attachToSelectedImage And Not manageExistingOle Then
        imagePath = Trim$(txtImagePath.Text)
        If Len(imagePath) = 0 Then
            MsgBox T("msg.chooseImage"), vbExclamation
            Exit Sub
        End If

        If Not IsDialogSupportedImage(imagePath) Then
            MsgBox T("msg.chooseSupportedImage"), vbExclamation
            Exit Sub
        End If
    End If

    If Not manageExistingOle Then
        If Not WorkingFolderHasAnyFiles Then
            MsgBox T("msg.packageFolderEmpty"), vbExclamation
            OpenWorkingFolder
            Exit Sub
        End If
    End If

    If manageExistingOle Then
        If Not targetInlineOle Is Nothing Then
            ManageWorkingFolderInInlineOle targetInlineOle, workingFolderPath, existingDisplayImagePath
        ElseIf Not targetFloatingOle Is Nothing Then
            ManageWorkingFolderInFloatingOle targetFloatingOle, workingFolderPath, existingDisplayImagePath
        Else
            MsgBox T("msg.selectedOleGone"), vbExclamation
            Exit Sub
        End If
    ElseIf attachToSelectedImage Then
        If Not targetInlineImage Is Nothing Then
            AttachWorkingFolderToInlineImage targetInlineImage, workingFolderPath
        ElseIf Not targetFloatingImage Is Nothing Then
            AttachWorkingFolderToFloatingImage targetFloatingImage, workingFolderPath
        Else
            MsgBox T("msg.selectedImageGone"), vbExclamation
            Exit Sub
        End If
    Else
        InsertImageAndWorkingFolderAsOle imagePath, workingFolderPath
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtImagePath_Change()
    UpdateDialogState
End Sub

Private Sub ConfigureSystemFileManagerControls()
    Me.Width = DIALOG_WIDTH
    Me.Height = DIALOG_HEIGHT

    Caption = T("dialog.insert.caption")
    lblIntro.Caption = T("dialog.insert.intro")
    lblIntro.Left = 18
    lblIntro.Top = 15
    lblIntro.Width = 444
    lblIntro.Height = 42
    lblIntro.WordWrap = True

    lblImage.Caption = T("dialog.displayImage")
    lblImage.Left = 18
    lblImage.Top = 66
    lblImage.Width = 180
    lblImage.Height = 12

    txtImagePath.Left = 18
    txtImagePath.Top = 84
    txtImagePath.Width = 342
    txtImagePath.Height = 18
    txtImagePath.Enabled = True

    cmdBrowseImage.Caption = T("dialog.chooseImage")
    cmdBrowseImage.Left = 372
    cmdBrowseImage.Top = 83
    cmdBrowseImage.Width = 90
    cmdBrowseImage.Height = 21
    cmdBrowseImage.Enabled = True

    lblSupportFiles.Caption = T("dialog.packageFolder")
    lblSupportFiles.Left = 18
    lblSupportFiles.Top = 122
    lblSupportFiles.Width = 180
    lblSupportFiles.Height = 12

    lstSupportFiles.Visible = False
    lstSupportFiles.Left = 18
    lstSupportFiles.Top = 144
    lstSupportFiles.Width = 444
    lstSupportFiles.Height = 18

    cmdAddFiles.Caption = T("dialog.openFolder")
    cmdAddFiles.Visible = True
    cmdAddFiles.Left = 18
    cmdAddFiles.Top = 144
    cmdAddFiles.Width = 108
    cmdAddFiles.Height = 24

    cmdAddFolder.Visible = False
    cmdRemoveSelected.Visible = False

    cmdClearFiles.Caption = T("dialog.clearFolder")
    cmdClearFiles.Visible = True
    cmdClearFiles.Left = 138
    cmdClearFiles.Top = 144
    cmdClearFiles.Width = 90
    cmdClearFiles.Height = 24

    lblSummary.Left = 18
    lblSummary.Top = 186
    lblSummary.Width = 444
    lblSummary.Height = 54
    lblSummary.WordWrap = True
    lblSummary.BackStyle = 1
    lblSummary.BackColor = RGB(248, 248, 248)
    lblSummary.BorderStyle = 1

    cmdInsert.Caption = T("dialog.insertButton")
    cmdInsert.Left = 324
    cmdInsert.Top = 258
    cmdInsert.Width = 66
    cmdInsert.Height = 24
    cmdInsert.Default = True

    cmdCancel.Left = 402
    cmdCancel.Top = 258
    cmdCancel.Width = 60
    cmdCancel.Height = 24
    cmdCancel.Cancel = True
End Sub

Private Sub ConfigureForSelectedImageMode()
    Caption = T("dialog.attach.caption")
    lblIntro.Caption = T("dialog.attach.intro")
    lblImage.Caption = T("dialog.selectedImage")
    txtImagePath.Text = T("dialog.selectedImage")
    txtImagePath.Enabled = False
    cmdBrowseImage.Enabled = False
    UpdateDialogState
End Sub

Private Sub ConfigureForSelectedOleMode()
    Caption = T("dialog.manage.caption")
    lblIntro.Caption = T("dialog.manage.intro")
    lblImage.Caption = T("dialog.selectedOle")
    txtImagePath.Text = T("dialog.selectedOle")
    txtImagePath.Enabled = False
    cmdBrowseImage.Enabled = False
    cmdInsert.Caption = T("dialog.rebuildButton")
    UpdateDialogState
End Sub

Private Sub EnsureWorkingFolder()
    If Len(workingFolderPath) = 0 Then
        PrepareEmptyWorkingFolder "image_support"
    End If
End Sub

Private Sub PrepareEmptyWorkingFolder(ByVal baseName As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    workingFolderPath = BuildDialogWorkingFolderPath(fso, baseName)
    fso.CreateFolder workingFolderPath
    openedWorkingFolder = False
    UpdateDialogState
End Sub

Private Sub PrepareWorkingFolderFromExistingZip(ByVal zipPath As String, ByVal baseName As String)
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    workingFolderPath = BuildDialogWorkingFolderPath(fso, baseName)
    ExtractZipToWorkingFolder zipPath, workingFolderPath
    openedWorkingFolder = False
    UpdateDialogState
End Sub

Private Function BuildDialogWorkingFolderPath(ByVal fso As Object, ByVal baseName As String) As String
    Dim rootPath As String
    Dim candidatePath As String
    Dim index As Long

    rootPath = fso.BuildPath(GetDialogTempFolderPath(fso), "OLEPackager_Work_" & SanitizeDialogFileName(baseName) & "_" & Format$(Now, "yyyymmdd_hhnnss"))
    candidatePath = rootPath
    index = 2
    Do While fso.FolderExists(candidatePath)
        candidatePath = rootPath & "_" & CStr(index)
        index = index + 1
    Loop

    BuildDialogWorkingFolderPath = candidatePath
End Function

Private Function GetDialogTempFolderPath(ByVal fso As Object) As String
    Dim tempPath As String

    tempPath = Environ$("LOCALAPPDATA")
    If Len(tempPath) > 0 Then
        tempPath = fso.BuildPath(tempPath, "Temp")
    Else
        tempPath = Environ$("TEMP")
    End If

    GetDialogTempFolderPath = tempPath
End Function

Private Function SanitizeDialogFileName(ByVal value As String) As String
    Dim invalidChars As String
    Dim i As Long
    Dim ch As String

    invalidChars = "\/:*?""<>|"
    SanitizeDialogFileName = Trim$(value)
    If Len(SanitizeDialogFileName) = 0 Then
        SanitizeDialogFileName = "package"
    End If

    For i = 1 To Len(invalidChars)
        ch = Mid$(invalidChars, i, 1)
        SanitizeDialogFileName = Replace(SanitizeDialogFileName, ch, "_")
    Next i
End Function

Private Sub OpenWorkingFolder()
    Dim fso As Object
    Dim wsh As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(workingFolderPath) = 0 Or Not fso.FolderExists(workingFolderPath) Then
        PrepareEmptyWorkingFolder "image_support"
    End If

    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "explorer.exe " & QuoteDialogCommandLine(workingFolderPath), 1, False
    openedWorkingFolder = True
    UpdateDialogState
End Sub

Private Sub ClearWorkingFolder()
    Dim fso As Object
    Dim folder As Object
    Dim fileItem As Object
    Dim subFolder As Object
    Dim pathsToDelete As Collection
    Dim item As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(workingFolderPath) = 0 Or Not fso.FolderExists(workingFolderPath) Then
        Exit Sub
    End If

    Set pathsToDelete = New Collection
    Set folder = fso.GetFolder(workingFolderPath)
    For Each fileItem In folder.Files
        pathsToDelete.Add CStr(fileItem.Path)
    Next fileItem
    For Each subFolder In folder.SubFolders
        pathsToDelete.Add CStr(subFolder.Path)
    Next subFolder

    For Each item In pathsToDelete
        If fso.FileExists(CStr(item)) Then
            fso.DeleteFile CStr(item), True
        ElseIf fso.FolderExists(CStr(item)) Then
            fso.DeleteFolder CStr(item), True
        End If
    Next item
End Sub

Private Function WorkingFolderHasAnyFiles() As Boolean
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(workingFolderPath) = 0 Then
        Exit Function
    End If
    If Not fso.FolderExists(workingFolderPath) Then
        Exit Function
    End If
    WorkingFolderHasAnyFiles = DialogFolderHasAnyFiles(fso.GetFolder(workingFolderPath))
End Function

Private Function DialogFolderHasAnyFiles(ByVal folder As Object) As Boolean
    Dim subFolder As Object

    If folder.Files.Count > 0 Then
        DialogFolderHasAnyFiles = True
        Exit Function
    End If

    For Each subFolder In folder.SubFolders
        If DialogFolderHasAnyFiles(subFolder) Then
            DialogFolderHasAnyFiles = True
            Exit Function
        End If
    Next subFolder
End Function

Private Function QuoteDialogCommandLine(ByVal value As String) As String
    QuoteDialogCommandLine = Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Sub UpdateDialogState()
    Dim hasImage As Boolean
    Dim folderText As String
    Dim actionText As String

    hasImage = attachToSelectedImage Or manageExistingOle Or Len(Trim$(txtImagePath.Text)) > 0
    cmdInsert.Enabled = hasImage

    If Len(workingFolderPath) > 0 Then
        folderText = workingFolderPath
    Else
        folderText = T("dialog.folderPending")
    End If

    If manageExistingOle Then
        actionText = T("dialog.afterRebuild")
    ElseIf attachToSelectedImage Then
        actionText = T("dialog.afterInsert")
    Else
        actionText = T("dialog.afterChooseInsert")
    End If

    lblSummary.Caption = TF("dialog.summary", folderText, actionText)
End Sub

Private Function IsDialogSupportedImage(ByVal imagePath As String) As Boolean
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Select Case LCase$(fso.GetExtensionName(imagePath))
        Case "png", "jpg", "jpeg", "tif", "tiff"
            IsDialogSupportedImage = True
        Case Else
            IsDialogSupportedImage = False
    End Select
End Function
