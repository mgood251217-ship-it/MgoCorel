VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportOption 
   Caption         =   "Pengaturan Export"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2550
   OleObjectBlob   =   "frmExportOption.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExportOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedImageType As Long
Public SelectedDPI As Double
Public Confirmed As Boolean

Private Sub UserForm_Initialize()
    optRGB.value = True
    optCMYK.value = False
    txtDPI.value = "100"
    Confirmed = False
End Sub

Private Sub cmdExport_Click()
    Dim dpi As Double

    dpi = Val(txtDPI.value)

    If dpi <= 0 Then
        MsgBox "DPI harus lebih dari 0.", _
               vbExclamation, _
               "MgoCorel"
        txtDPI.SetFocus
        Exit Sub
    End If

    If optRGB.value Then
        SelectedImageType = cdrRGBColorImage
    ElseIf optCMYK.value Then
        SelectedImageType = cdrCMYKColorImage
    Else
        MsgBox "Pilih mode warna terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    SelectedDPI = dpi
    Confirmed = True

    Me.Hide
End Sub

Private Sub cmdBatal_Click()
    Confirmed = False
    Me.Hide
End Sub

