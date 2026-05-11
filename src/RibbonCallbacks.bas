Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Private ribbonUi As IRibbonUI

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set ribbonUi = ribbon
End Sub

Public Sub RibbonGetLabel(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = RibbonText(control.Id, "label")
End Sub

Public Sub RibbonGetScreentip(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = RibbonText(control.Id, "screentip")
End Sub

Public Sub RibbonGetSupertip(ByVal control As IRibbonControl, ByRef returnedVal)
    returnedVal = RibbonText(control.Id, "supertip")
End Sub

Public Sub RibbonInsertPlotFolder(ByVal control As IRibbonControl)
    InsertPlotFolderAsOle
End Sub

Public Sub RibbonInsertImageSupportFiles(ByVal control As IRibbonControl)
    InsertImageSupportFilesAsOle
End Sub

Public Sub RibbonAttachSupportFilesToSelectedImage(ByVal control As IRibbonControl)
    AttachSupportFilesToSelectedImageAsOle
End Sub

Public Sub RibbonShowUsage(ByVal control As IRibbonControl)
    ShowUsageHelp
End Sub

Private Function RibbonText(ByVal controlId As String, ByVal part As String) As String
    Select Case controlId
        Case "tabOLEPackager"
            RibbonText = T("app.name")
        Case "grpInsert"
            RibbonText = T("ribbon.group.insert")
        Case "btnInsertPlotFolder"
            RibbonText = T("ribbon.insertPackage." & part)
        Case "btnInsertImageSupportFiles"
            RibbonText = T("ribbon.insertImageFiles." & part)
        Case "btnAttachSupportFilesToSelectedImage"
            RibbonText = T("ribbon.manage." & part)
        Case "btnShowUsage"
            RibbonText = T("ribbon.usage." & part)
    End Select
End Function
