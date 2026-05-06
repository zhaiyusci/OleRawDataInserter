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
    Dim imagePath As String
    Dim i As Long

    imagePath = Trim$(txtImagePath.Text)
    If Len(imagePath) = 0 Then
        MsgBox "Please choose an image.", vbExclamation
        Exit Sub
    End If

    If Not IsDialogSupportedImage(imagePath) Then
        MsgBox "Please choose a PNG, JPG, JPEG, TIF, or TIFF image.", vbExclamation
        Exit Sub
    End If

    If lstSupportFiles.ListCount = 0 Then
        MsgBox "Please add at least one support file.", vbExclamation
        Exit Sub
    End If

    Set supportFiles = New Collection
    For i = 0 To lstSupportFiles.ListCount - 1
        supportFiles.Add CStr(lstSupportFiles.List(i))
    Next i

    InsertImageAndSupportFilesAsOle imagePath, supportFiles
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

    For i = 0 To lstSupportFiles.ListCount - 1
        If LCase$(CStr(lstSupportFiles.List(i))) = LCase$(filePath) Then Exit Sub
    Next i

    lstSupportFiles.AddItem filePath
End Sub

Private Sub UpdateDialogState()
    lblSupportFiles.Caption = "Support files (" & CStr(lstSupportFiles.ListCount) & " selected)"
    cmdInsert.Enabled = (Len(Trim$(txtImagePath.Text)) > 0 And lstSupportFiles.ListCount > 0)

    If cmdInsert.Enabled Then
        lblSummary.Caption = "Ready. Click Insert to embed the selected support files and display the chosen image."
    Else
        lblSummary.Caption = "Choose an image and at least one support file."
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
