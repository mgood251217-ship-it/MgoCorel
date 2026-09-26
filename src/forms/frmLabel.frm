VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLabel 
   Caption         =   "UserForm1"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmLabel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Initialize()
    LoadProducts
    LoadFinishings
    LoadOperators

    txtQuantity.Value = "1"
    txtDeadline.Value = Format$(Date, "dd/mm/yyyy")
End Sub

Private Sub LoadProducts()
    Dim item As Variant

    cmbProduk.Clear

    For Each item In GetProducts()
        cmbProduk.AddItem item(0)
    Next item
End Sub

Private Sub LoadFinishings()
    Dim item As Variant

    cmbFinishing.Clear

    For Each item In GetFinishings()
        cmbFinishing.AddItem item(0)
    Next item
End Sub

Private Sub LoadOperators()
    Dim item As Variant

    cmbOperator.Clear

    For Each item In GetOperators()
        cmbOperator.AddItem item
    Next item
End Sub

Private Sub cmdSimpan_Click()
    Dim sr As ShapeRange
    Dim oneShape As ShapeRange
    Dim shp As Shape
    Dim i As Long
    Dim widthCm As Double
    Dim heightCm As Double

    If Trim$(txtNoInv.Value) = "" Then
        MsgBox "No Inv wajib diisi.", vbExclamation, "MgoCorel"
        txtNoInv.SetFocus
        Exit Sub
    End If

    If Trim$(txtNama.Value) = "" Then
        MsgBox "Nama wajib diisi.", vbExclamation, "MgoCorel"
        txtNama.SetFocus
        Exit Sub
    End If

    If cmbProduk.ListIndex = -1 Then
        MsgBox "Produk wajib dipilih.", vbExclamation, "MgoCorel"
        cmbProduk.SetFocus
        Exit Sub
    End If

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", vbExclamation, "MgoCorel"
        Exit Sub
    End If

    If cmbFinishing.ListIndex = -1 Then
        MsgBox "Finishing wajib dipilih.", vbExclamation, "MgoCorel"
        cmbFinishing.SetFocus
        Exit Sub
    End If

    If Val(txtQuantity.Value) <= 0 Then
        MsgBox "Quantity harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtQuantity.SetFocus
        Exit Sub
    End If

    If Trim$(txtDeadline.Value) = "" Then
        MsgBox "Deadline wajib diisi.", vbExclamation, "MgoCorel"
        txtDeadline.SetFocus
        Exit Sub
    End If

    If cmbOperator.ListIndex = -1 Then
        MsgBox "Operator wajib dipilih.", vbExclamation, "MgoCorel"
        cmbOperator.SetFocus
        Exit Sub
    End If

    For i = 1 To sr.Count
        Set shp = sr(i)

        widthCm = GetShapeWidthCm(shp)
        heightCm = GetShapeHeightCm(shp)

        Set oneShape = CreateShapeRange
        oneShape.Add shp

        CreateLabelCanvas oneShape

        CreateLabelTexts _
            oneShape, _
            txtNama.Value, _
            cmbProduk.Value, _
            widthCm, _
            heightCm, _
            cmbFinishing.Value, _
            txtQuantity.Value, _
            txtDeadline.Value, _
            cmbOperator.Value, _
            txtNoInv.Value
    Next i

    Unload Me
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub