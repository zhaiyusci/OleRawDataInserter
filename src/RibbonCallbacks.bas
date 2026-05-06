Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Private ribbonUi As IRibbonUI

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set ribbonUi = ribbon
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
