VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImageSupportFilesDialog 
   Caption         =   "Insert Image + Support Files"
   ClientHeight    =   5436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8664.001
   OleObjectBlob   =   "ImageSupportFilesDialog.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ImageSupportFilesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const EMBEDDED_PREFIX As String = "[embedded] "
Private Const NEW_PREFIX As String = "[new] "

Private attachToSelectedImage As Boolean
Private manageExistingOle As Boolean
Private targetInlineImage As Object
Private targetFloatingImage As Object
Private targetInlineOle As Object
Private targetFloatingOle As Object
Private existingZipPath As String

Public Sub ConfigureForSelectedInlineImage(ByVal inlineImage As Object)
    attachToSelectedImage = True
    manageExistingOle = False
    Set targetInlineImage = inlineImage
    ConfigureForSelectedImageMode
End Sub

Public Sub ConfigureForSelectedFloatingImage(ByVal floatingImage As Object)
    attachToSelectedImage = True
    manageExistingOle = False
    Set targetFloatingImage = floatingImage
    ConfigureForSelectedImageMode
End Sub

Public Sub ConfigureForSelectedInlineOle(ByVal inlineOle As Object)
    attachToSelectedImage = False
    manageExistingOle = True
    Set targetInlineOle = inlineOle
    existingZipPath = ExtractZipFromInlineOleObject(inlineOle)
    ConfigureForSelectedOleMode
    LoadEmbeddedEntries existingZipPath
End Sub

Public Sub ConfigureForSelectedFloatingOle(ByVal floatingOle As Object)
    attachToSelectedImage = False
    manageExistingOle = True
    Set targetFloatingOle = floatingOle
    existingZipPath = ExtractZipFromFloatingOleObject(targetFloatingOle)
    ConfigureForSelectedOleMode
    LoadEmbeddedEntries existingZipPath
End Sub

Private Sub UserForm_Initialize()
    UpdateDialogState
End Sub

Private Sub cmdBrowseImage_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Choose the image to display"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Image files", "*.png;*.jpg;*.jpeg;*.tif;*.tiff"
        .Filters.Add "PNG files", "*.png"
        .Filters.Add "JPEG files", "*.jpg;*.jpeg"
        .Filters.Add "TIFF files", "*.tif;*.tiff"
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then txtImagePath.Text = .SelectedItems(1)
    End With

    UpdateDialogState
End Sub

Private Sub cmdAddFiles_Click()
    Dim i As Long

    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Choose support files to embed"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then
            For i = 1 To .SelectedItems.Count
                AddSupportFileIfMissing CStr(.SelectedItems(i))
            Next i
        End If
    End With

    UpdateDialogState
End Sub

Private Sub cmdRemoveSelected_Click()
    Dim i As Long

    For i = lstSupportFiles.ListCount - 1 To 0 Step -1
        If lstSupportFiles.Selected(i) Then lstSupportFiles.RemoveItem i
    Next i

    UpdateDialogState
End Sub

Private Sub cmdClearFiles_Click()
    lstSupportFiles.Clear
    UpdateDialogState
End Sub

Private Sub cmdInsert_Click()
    Dim supportFiles As Collection
    Dim keepEntries As Collection
    Dim newFiles As Collection
    Dim imagePath As String
    Dim itemText As String
    Dim i As Long

    If Not attachToSelectedImage And Not manageExistingOle Then
        imagePath = Trim$(txtImagePath.Text)
        If Len(imagePath) = 0 Then
            MsgBox "Please choose an image.", vbExclamation
            Exit Sub
        End If

        If Not IsDialogSupportedImage(imagePath) Then
            MsgBox "Please choose a PNG, JPG, JPEG, TIF, or TIFF image.", vbExclamation
            Exit Sub
        End If
    End If

    If lstSupportFiles.ListCount = 0 And Not manageExistingOle Then
        MsgBox "Please add at least one support file.", vbExclamation
        Exit Sub
    End If

    If manageExistingOle Then
        Set keepEntries = New Collection
        Set newFiles = New Collection
        For i = 0 To lstSupportFiles.ListCount - 1
            itemText = CStr(lstSupportFiles.List(i))
            If Left$(itemText, Len(EMBEDDED_PREFIX)) = EMBEDDED_PREFIX Then
                keepEntries.Add Mid$(itemText, Len(EMBEDDED_PREFIX) + 1)
            ElseIf Left$(itemText, Len(NEW_PREFIX)) = NEW_PREFIX Then
                newFiles.Add Mid$(itemText, Len(NEW_PREFIX) + 1)
            Else
                newFiles.Add itemText
            End If
        Next i

        If Not targetInlineOle Is Nothing Then
            ManageFilesInInlineOle targetInlineOle, existingZipPath, keepEntries, newFiles
        ElseIf Not targetFloatingOle Is Nothing Then
            ManageFilesInFloatingOle targetFloatingOle, existingZipPath, keepEntries, newFiles
        Else
            MsgBox "The selected OLE object is no longer available.", vbExclamation
            Exit Sub
        End If
    ElseIf attachToSelectedImage Then
        Set supportFiles = New Collection
        For i = 0 To lstSupportFiles.ListCount - 1
            supportFiles.Add CStr(lstSupportFiles.List(i))
        Next i

        If Not targetInlineImage Is Nothing Then
            AttachSupportFilesToInlineImage targetInlineImage, supportFiles
        ElseIf Not targetFloatingImage Is Nothing Then
            AttachSupportFilesToFloatingImage targetFloatingImage, supportFiles
        Else
            MsgBox "The selected image is no longer available.", vbExclamation
            Exit Sub
        End If
    Else
        Set supportFiles = New Collection
        For i = 0 To lstSupportFiles.ListCount - 1
            supportFiles.Add CStr(lstSupportFiles.List(i))
        Next i
        InsertImageAndSupportFilesAsOle imagePath, supportFiles
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtImagePath_Change()
    UpdateDialogState
End Sub

Private Sub AddSupportFileIfMissing(ByVal filePath As String)
    Dim i As Long
    Dim displayText As String

    If manageExistingOle Then
        displayText = NEW_PREFIX & filePath
    Else
        displayText = filePath
    End If

    For i = 0 To lstSupportFiles.ListCount - 1
        If LCase$(CStr(lstSupportFiles.List(i))) = LCase$(displayText) Then Exit Sub
    Next i

    lstSupportFiles.AddItem displayText
End Sub

Private Sub ConfigureForSelectedImageMode()
    Caption = "Attach Files to Selected Image"
    lblIntro.Caption = "Review the support files before attaching them to the selected image. The selected image will be replaced by an OLE object while keeping its size and position."
    lblImage.Caption = "Selected image in current document"
    txtImagePath.Text = "Selected image in current document"
    txtImagePath.Enabled = False
    cmdBrowseImage.Enabled = False
    UpdateDialogState
End Sub

Private Sub ConfigureForSelectedOleMode()
    Caption = "Manage Files in Selected OLE Object"
    lblIntro.Caption = "Review the files already embedded in the selected OLE object. Remove rows to delete files, or add new files, then click Insert to rebuild the OLE object while keeping its display, size, and position."
    lblImage.Caption = "Selected OLE object in current document"
    txtImagePath.Text = "Selected OLE object in current document"
    txtImagePath.Enabled = False
    cmdBrowseImage.Enabled = False
    UpdateDialogState
End Sub

Private Sub LoadEmbeddedEntries(ByVal zipPath As String)
    Dim entries As Collection
    Dim entryName As Variant

    Set entries = GetZipEntryNames(zipPath)
    lstSupportFiles.Clear
    For Each entryName In entries
        lstSupportFiles.AddItem EMBEDDED_PREFIX & CStr(entryName)
    Next entryName
    UpdateDialogState
End Sub

Private Sub UpdateDialogState()
    Dim hasImage As Boolean

    hasImage = attachToSelectedImage Or manageExistingOle Or Len(Trim$(txtImagePath.Text)) > 0
    lblSupportFiles.Caption = "Support files (" & CStr(lstSupportFiles.ListCount) & " selected)"
    cmdInsert.Enabled = (hasImage And (manageExistingOle Or lstSupportFiles.ListCount > 0))

    If cmdInsert.Enabled Then
        If manageExistingOle Then
            lblSummary.Caption = "Ready. Click Insert to rebuild the selected OLE object with the files currently shown in the list."
        ElseIf attachToSelectedImage Then
            lblSummary.Caption = "Ready. Click Insert to attach the selected support files while preserving the selected image size and position."
        Else
            lblSummary.Caption = "Ready. Click Insert to embed the selected support files and display the chosen image."
        End If
    Else
        If manageExistingOle Then
            lblSummary.Caption = "The package can be rebuilt even if the list is empty."
        ElseIf attachToSelectedImage Then
            lblSummary.Caption = "Add at least one support file."
        Else
            lblSummary.Caption = "Choose an image and at least one support file."
        End If
    End If
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
