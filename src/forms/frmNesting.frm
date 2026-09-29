VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNesting 
   Caption         =   "Nesting"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "frmNesting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    txtPanjang.value = "1"
    txtLebar.value = "1"
    txtGap.value = "2"

    cmbRotasi.Clear

    cmbRotasi.AddItem "Tanpa Rotasi"
    cmbRotasi.AddItem "Rotasi 90°"
    cmbRotasi.AddItem "Rotasi 180°"
    cmbRotasi.AddItem "Rotasi Bebas"

    cmbRotasi.ListIndex = 0

End Sub

Private Sub cmdNesting_Click()

    Dim boxWidthM As Double
    Dim boxHeightM As Double
    Dim gapMm As Double
    Dim rotationMode As Long
    Dim success As Boolean

    If Not IsNumeric(txtPanjang.value) Then
        MsgBox "Panjang nesting harus berupa angka.", _
               vbExclamation, "MgoCorel"
        txtPanjang.SetFocus
        Exit Sub
    End If

    If CDbl(txtPanjang.value) <= 0 Then
        MsgBox "Panjang nesting harus lebih dari 0.", _
               vbExclamation, "MgoCorel"
        txtPanjang.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtLebar.value) Then
        MsgBox "Lebar nesting harus berupa angka.", _
               vbExclamation, "MgoCorel"
        txtLebar.SetFocus
        Exit Sub
    End If

    If CDbl(txtLebar.value) <= 0 Then
        MsgBox "Lebar nesting harus lebih dari 0.", _
               vbExclamation, "MgoCorel"
        txtLebar.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtGap.value) Then
        MsgBox "Gap harus berupa angka.", _
               vbExclamation, "MgoCorel"
        txtGap.SetFocus
        Exit Sub
    End If

    If CDbl(txtGap.value) < 0 Then
        MsgBox "Gap tidak boleh kurang dari 0.", _
               vbExclamation, "MgoCorel"
        txtGap.SetFocus
        Exit Sub
    End If

    If cmbRotasi.ListIndex < 0 Then
        MsgBox "Pilih mode rotasi.", _
               vbExclamation, "MgoCorel"
        cmbRotasi.SetFocus
        Exit Sub
    End If

    boxWidthM = CDbl(txtPanjang.value)
    boxHeightM = CDbl(txtLebar.value)
    gapMm = CDbl(txtGap.value)
    rotationMode = cmbRotasi.ListIndex

    Me.Hide
    DoEvents

    success = ArrangeNesting( _
        boxWidthM, _
        boxHeightM, _
        gapMm, _
        rotationMode)

    If Not success Then
        Me.Show vbModal
    End If

End Sub

Private Sub cmdBatal_Click()
    Me.Hide
End Sub

