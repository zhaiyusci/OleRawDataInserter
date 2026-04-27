Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Private ribbonUi As IRibbonUI

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set ribbonUi = ribbon
End Sub

Public Sub RibbonInsertPlotFolder(ByVal control As IRibbonControl)
    InsertPlotFolderAsOle
End Sub
