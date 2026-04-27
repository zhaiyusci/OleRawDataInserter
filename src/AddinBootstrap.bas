Attribute VB_Name = "AddinBootstrap"
Option Explicit

Private Const ADDIN_NAME As String = "OleRawDataInserter"
Private Const BUTTON_TAG As String = "OleRawDataInserter.InsertPlotFolderAsOle"
Private Const BUTTON_CAPTION As String = "Insert Raw Data OLE"

Public Sub AutoExec()
    InstallAddinButton
End Sub

Public Sub AutoExit()
    RemoveAddinButton
End Sub

Public Sub InstallAddinButton()
    Dim bar As CommandBar
    Dim button As CommandBarButton

    On Error Resume Next
    Set bar = Application.CommandBars("Add-Ins")
    If bar Is Nothing Then Set bar = Application.CommandBars("Menu Bar")
    On Error GoTo 0

    If bar Is Nothing Then Exit Sub

    RemoveAddinButton

    Set button = bar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    With button
        .Caption = BUTTON_CAPTION
        .Tag = BUTTON_TAG
        .Style = msoButtonCaption
        .OnAction = "'" & ADDIN_NAME & ".dotm'!InsertPlotFolderAsOle"
        .TooltipText = "Package the selected plot folder as an embedded OLE zip object."
    End With
End Sub

Public Sub RemoveAddinButton()
    Dim bar As CommandBar
    Dim control As CommandBarControl

    On Error Resume Next
    Set bar = Application.CommandBars("Add-Ins")
    If bar Is Nothing Then Set bar = Application.CommandBars("Menu Bar")

    If Not bar Is Nothing Then
        For Each control In bar.Controls
            If control.Tag = BUTTON_TAG Then control.Delete
        Next control
    End If
    On Error GoTo 0
End Sub
