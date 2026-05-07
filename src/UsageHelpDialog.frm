VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UsageHelpDialog 
   Caption         =   "Figure Package Usage"
   ClientHeight    =   8040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10980
   OleObjectBlob   =   "UsageHelpDialog.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UsageHelpDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowUsageText(ByVal usageText As String)
    txtUsage.Text = usageText
    txtUsage.SelStart = 0
    Me.Show
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
