VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCanvas 
   Caption         =   "Atur Canvas"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4635
   OleObjectBlob   =   "frmCanvas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public CanvasTop As Double
Public CanvasBottom As Double
Public CanvasLeft As Double
Public CanvasRight As Double
Public Confirmed As Boolean

Private Sub UserForm_Initialize()
    ApplyButtonTheme Me

    If CanvasTop <= 0 Then CanvasTop = GetCanvasExtra() / 2#
    If CanvasBottom <= 0 Then CanvasBottom = GetCanvasExtra() / 2#
    If CanvasLeft <= 0 Then CanvasLeft = GetCanvasExtra() / 2#
    If CanvasRight <= 0 Then CanvasRight = GetCanvasExtra() / 2#

    txtAtas.value = Format$(CanvasTop, "0.##")
    txtBawah.value = Format$(CanvasBottom, "0.##")
    txtKiri.value = Format$(CanvasLeft, "0.##")
    txtKanan.value = Format$(CanvasRight, "0.##")

    Confirmed = False
End Sub

Private Sub cmdSimpan_Click()
    If Not IsValidCanvas(txtAtas.value) Then
        MsgBox "Canvas Atas harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtAtas.SetFocus
        Exit Sub
    End If

    If Not IsValidCanvas(txtBawah.value) Then
        MsgBox "Canvas Bawah harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtBawah.SetFocus
        Exit Sub
    End If

    If Not IsValidCanvas(txtKiri.value) Then
        MsgBox "Canvas Kiri harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtKiri.SetFocus
        Exit Sub
    End If

    If Not IsValidCanvas(txtKanan.value) Then
        MsgBox "Canvas Kanan harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtKanan.SetFocus
        Exit Sub
    End If

    CanvasTop = CDbl(txtAtas.value)
    CanvasBottom = CDbl(txtBawah.value)
    CanvasLeft = CDbl(txtKiri.value)
    CanvasRight = CDbl(txtKanan.value)

    Confirmed = True

    Me.Hide
End Sub

Private Sub cmdBatal_Click()
    Confirmed = False
    Me.Hide
End Sub

Private Function IsValidCanvas(ByVal value As String) As Boolean
    If Not IsNumeric(value) Then Exit Function

    If CDbl(value) <= 0 Then Exit Function

    IsValidCanvas = True
End Function
