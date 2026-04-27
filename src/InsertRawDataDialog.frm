VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InsertRawDataDialog
   Caption         =   "Insert plot.png Package"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   StartUpPosition =   1
   Begin MSForms.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   360
      Left            =   6480
      TabIndex        =   4
      Top             =   2280
      Width           =   960
   End
   Begin MSForms.CommandButton cmdInsert
      Caption         =   "Insert"
      Height          =   360
      Left            =   5280
      TabIndex        =   3
      Top             =   2280
      Width           =   960
   End
   Begin MSForms.CommandButton cmdBrowsePlot
      Caption         =   "Browse..."
      Height          =   330
      Left            =   5880
      TabIndex        =   2
      Top             =   960
      Width           =   1560
   End
   Begin MSForms.TextBox txtPlotPath
      Height          =   330
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   5160
   End
   Begin MSForms.Label lblHelp
      Caption         =   "Choose plot.png. The add-in will automatically package all source files in the same folder as embedded raw data."
      Height          =   600
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   6960
      WordWrap        =   -1
   End
   Begin MSForms.Label lblPlot
      Caption         =   "plot.png"
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2760
   End
End
Attribute VB_Name = "InsertRawDataDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowsePlot_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Choose plot.png"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "plot.png", "plot.png"
        .Filters.Add "PNG files", "*.png"
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then txtPlotPath.Text = .SelectedItems(1)
    End With
End Sub

Private Sub cmdInsert_Click()
    If Len(Trim$(txtPlotPath.Text)) = 0 Then
        MsgBox "Please choose plot.png.", vbExclamation
        Exit Sub
    End If

    InsertPlotImageFile txtPlotPath.Text
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
