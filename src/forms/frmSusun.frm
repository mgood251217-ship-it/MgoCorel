VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSusun 
   Caption         =   "Susun Objek"
   ClientHeight    =   2745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   OleObjectBlob   =   "frmSusun.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSusun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub UserForm_Initialize()
    ApplyButtonTheme Me

    txtGap.value = "2"
    chkCutLine.value = False
    optPage.value = True
    optObject.value = False
End Sub

Private Sub cmdSusun_Click()
    Dim gapMm As Double

    gapMm = Val(txtGap.value)

    If gapMm < 0 Then
        MsgBox "Gap tidak boleh kurang dari 0.", vbExclamation, "MgoCorel"
        txtGap.SetFocus
        Exit Sub
    End If

    ArrangeSelection _
        gapMm, _
        chkCutLine.value, _
        optPage.value

    Unload Me
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

