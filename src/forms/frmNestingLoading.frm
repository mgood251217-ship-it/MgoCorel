VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNestingLoading 
   Caption         =   "Nesting Loading"
   ClientHeight    =   570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4580
   OleObjectBlob   =   "frmNestingLoading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNestingLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim contentBottom As Single

    ApplyButtonTheme Me
    Me.Font.Name = "Calibri"
    Me.Font.Size = 12
    cmdBatal.Font.Name = "Calibri"
    cmdBatal.Font.Size = 12
    cmdBatal.Height = 24
    cmdBatal.Width = 75
    lblStatus.Font.Name = "Calibri"
    lblStatus.Font.Size = 12
    lblStatus.Width = 90
    lblPercent.Font.Name = "Calibri"
    lblPercent.Font.Size = 12
    lblPercent.Left = lblStatus.Left + lblStatus.Width + 8
    lblPercent.Width = 50

    contentBottom = cmdBatal.Top + cmdBatal.Height
    If lblStatus.Top + lblStatus.Height > contentBottom Then
        contentBottom = lblStatus.Top + lblStatus.Height
    End If
    Me.Height = contentBottom + (Me.Height - Me.InsideHeight) + 8
End Sub

Private Sub cmdBatal_Click()

    NestingCancelRequested = True

    cmdBatal.Enabled = False

    lblStatus.Caption = _
        "Membatalkan nesting..."

End Sub

Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer)

    If CloseMode = 0 Then

        NestingCancelRequested = True

        Cancel = True

        cmdBatal.Enabled = False

        lblStatus.Caption = _
            "Membatalkan nesting..."

    End If

End Sub

