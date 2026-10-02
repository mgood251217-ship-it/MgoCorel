VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHitung 
   Caption         =   "Hitung Harga Spanduk"
   ClientHeight    =   6555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "frmHitung.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Private Sub UserForm_Initialize()
    ApplyButtonTheme Me

    LoadProducts
    LoadFinishings

    PrepareDetailList

    lblHargaMeter.caption = "Rp 0"
    lblSubtotal.caption = "Rp 0"
    lblTotal.caption = "Rp 0"

    If cmbProduk.ListCount > 0 Then
        cmbProduk.ListIndex = 0
    End If

    If cmbFinishing.ListCount > 0 Then
        cmbFinishing.ListIndex = 0
    End If
End Sub

Public Sub PrepareDetailList()
    With lstDetail
        .Clear
        .ColumnCount = 6
        .columnWidths = "80 pt;80 pt;80 pt;40 pt;90 pt;100 pt"
        .AddItem "Produk"
        .List(0, 1) = "Finishing"
        .List(0, 2) = "Ukuran"
        .List(0, 3) = "Qty"
        .List(0, 4) = "Satuan"
        .List(0, 5) = "Jumlah"
    End With
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

Private Sub cmdHitung_Click()
    CalculateSelectionPrice Me
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub
