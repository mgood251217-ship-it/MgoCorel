VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDots 
   Caption         =   "Pengaturan Dots"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "frmDots.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public TopCount As Long
Public BottomCount As Long
Public LeftCount As Long
Public RightCount As Long
Public Confirmed As Boolean

Private Sub UserForm_Initialize()
    If TopCount <= 0 Then TopCount = 2
    If BottomCount <= 0 Then BottomCount = 2
    If LeftCount <= 0 Then LeftCount = 2
    If RightCount <= 0 Then RightCount = 2

    txtAtas.value = CStr(TopCount)
    txtBawah.value = CStr(BottomCount)
    txtKiri.value = CStr(LeftCount)
    txtKanan.value = CStr(RightCount)

    Confirmed = False
End Sub

Private Sub cmdSimpan_Click()
    If Not IsValidCount(txtAtas.value) Then
        MsgBox "Jumlah dots Atas harus minimal 1.", vbExclamation, "MgoCorel"
        txtAtas.SetFocus
        Exit Sub
    End If

    If Not IsValidCount(txtBawah.value) Then
        MsgBox "Jumlah dots Bawah harus minimal 1.", vbExclamation, "MgoCorel"
        txtBawah.SetFocus
        Exit Sub
    End If

    If Not IsValidCount(txtKiri.value) Then
        MsgBox "Jumlah dots Kiri harus minimal 1.", vbExclamation, "MgoCorel"
        txtKiri.SetFocus
        Exit Sub
    End If

    If Not IsValidCount(txtKanan.value) Then
        MsgBox "Jumlah dots Kanan harus minimal 1.", vbExclamation, "MgoCorel"
        txtKanan.SetFocus
        Exit Sub
    End If

    TopCount = CLng(txtAtas.value)
    BottomCount = CLng(txtBawah.value)
    LeftCount = CLng(txtKiri.value)
    RightCount = CLng(txtKanan.value)

    Confirmed = True
    Me.Hide
End Sub

Private Sub cmdBatal_Click()
    Confirmed = False
    Me.Hide
End Sub

Private Function IsValidCount(ByVal value As String) As Boolean
    If Not IsNumeric(value) Then Exit Function

    If CLng(Val(value)) < 1 Then Exit Function

    IsValidCount = True
End Function
