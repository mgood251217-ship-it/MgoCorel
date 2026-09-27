VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHitung 
   Caption         =   "UserForm1"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900.001
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
    LoadProducts
    LoadFinishings

    lstDetail.ColumnCount = 6
    lstDetail.ColumnWidths = "85 pt;85 pt;70 pt;40 pt;90 pt;90 pt"

    lblHargaMeter.Caption = "Harga / m² : Rp 0"
    lblSubtotal.Caption = "Subtotal : Rp 0"
    lblTotal.Caption = "TOTAL : Rp 0"

    If cmbProduk.ListCount > 0 Then
        cmbProduk.ListIndex = 0
    End If

    If cmbFinishing.ListCount > 0 Then
        cmbFinishing.ListIndex = 0
    End If
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
