VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNestingLoading 
   Caption         =   "Nesting Loading"
   ClientHeight    =   570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "frmNestingLoading.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNestingLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

