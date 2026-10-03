VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRectangleNesting 
   Caption         =   "Rectangle Nesting"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3960
   OleObjectBlob   =   "frmRectangleNesting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRectangleNesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub UserForm_Initialize()
    ApplyButtonTheme Me

    txtPanjang.value = "1"
    txtLebar.value = "1"
    txtGap.value = "2"
    chkRotate.value = False

    UpdateInfo
End Sub

Private Sub txtPanjang_Change()
    UpdateInfo
End Sub

Private Sub txtLebar_Change()
    UpdateInfo
End Sub

Private Sub txtGap_Change()
    UpdateInfo
End Sub

Private Sub UpdateInfo()
    Dim sr As ShapeRange

    Set sr = ActiveSelectionRange

    lblInfo.caption = _
        "Objek : " & sr.count
End Sub

Private Sub cmdNesting_Click()
    Dim boxWidthM As Double
    Dim boxHeightM As Double
    Dim gapMm As Double
    Dim autoRotate As Boolean
    Dim success As Boolean

    If Not IsNumeric(txtPanjang.value) Or CDbl(txtPanjang.value) <= 0 Then
        MsgBox "Panjang nesting harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtPanjang.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtLebar.value) Or CDbl(txtLebar.value) <= 0 Then
        MsgBox "Lebar nesting harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtLebar.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(txtGap.value) Or CDbl(txtGap.value) < 0 Then
        MsgBox "Gap tidak boleh kurang dari 0.", vbExclamation, "MgoCorel"
        txtGap.SetFocus
        Exit Sub
    End If

    boxWidthM = CDbl(txtPanjang.value)
    boxHeightM = CDbl(txtLebar.value)
    gapMm = CDbl(txtGap.value)
    autoRotate = chkRotate.value

    success = ArrangeRectangleNesting( _
        boxWidthM, _
        boxHeightM, _
        gapMm, _
        autoRotate)

    If success Then
        Me.Hide
    End If
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

