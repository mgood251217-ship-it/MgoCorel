VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportLoading 
   Caption         =   "Loading Export"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmExportLoading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lblTitle As Object
Private lblPath As Object
Private lblStatus As Object
Private lblPercent As Object
Private lblBarBack As Object
Private lblBar As Object

Private Sub UserForm_Initialize()
    Me.Caption = "MgoCorel - Export"
    Me.Width = 420
    Me.Height = 150
    Me.StartUpPosition = 1

    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle", True)
    With lblTitle
        .Left = 15
        .Top = 12
        .Width = 380
        .Height = 20
        .Caption = "Menyiapkan export..."
        .Font.Bold = True
        .Font.Size = 11
    End With

    Set lblPath = Me.Controls.Add("Forms.Label.1", "lblPath", True)
    With lblPath
        .Left = 15
        .Top = 38
        .Width = 380
        .Height = 32
        .Caption = ""
        .WordWrap = True
    End With

    Set lblBarBack = Me.Controls.Add("Forms.Label.1", "lblBarBack", True)
    With lblBarBack
        .Left = 15
        .Top = 78
        .Width = 350
        .Height = 16
        .BackColor = RGB(220, 220, 220)
        .Caption = ""
    End With

    Set lblBar = Me.Controls.Add("Forms.Label.1", "lblBar", True)
    With lblBar
        .Left = 15
        .Top = 78
        .Width = 0
        .Height = 16
        .BackColor = RGB(0, 120, 215)
        .Caption = ""
    End With

    Set lblPercent = Me.Controls.Add("Forms.Label.1", "lblPercent", True)
    With lblPercent
        .Left = 370
        .Top = 78
        .Width = 35
        .Height = 16
        .Caption = "0%"
    End With

    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus", True)
    With lblStatus
        .Left = 15
        .Top = 102
        .Width = 390
        .Height = 20
        .Caption = "Menunggu..."
    End With
End Sub

Public Sub SetExportInfo( _
    ByVal current As Long, _
    ByVal total As Long, _
    ByVal filePath As String)

    Dim progress As Double
    Dim barWidth As Double

    If total <= 0 Then Exit Sub

    progress = current / total
    barWidth = 350# * progress

    If barWidth > 350 Then barWidth = 350
    If barWidth < 0 Then barWidth = 0

    lblTitle.Caption = "Export label " & current & " dari " & total
    lblPath.Caption = filePath
    lblBar.Width = barWidth
    lblPercent.Caption = Format$(progress * 100#, "0") & "%"
    lblStatus.Caption = "Sedang export..."

    DoEvents
End Sub

Public Sub SetFolder(ByVal folderPath As String)
    lblPath.Caption = "Folder export: " & folderPath
    lblStatus.Caption = "Menyiapkan..."
    DoEvents
End Sub

