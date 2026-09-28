VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLabel 
   Caption         =   "Pelabelan Spanduk"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4095
   OleObjectBlob   =   "frmLabel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mDotsTop As Long
Private mDotsBottom As Long
Private mDotsLeft As Long
Private mDotsRight As Long

Private Sub UserForm_Initialize()
    LoadSystems
    LoadProducts
    LoadFinishings
    LoadOperators
    LoadLabelState Me

    InitializeDotCounts
    UpdateDotsButton
End Sub

Private Sub LoadSystems()
    cmbSystem.Clear
    cmbSystem.AddItem "Offline"
    cmbSystem.AddItem "Online"
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

Private Sub InitializeDotCounts()
    Dim sr As ShapeRange
    Dim shp As Shape
    Dim widthM As Double
    Dim heightM As Double

    mDotsTop = 2
    mDotsBottom = 2
    mDotsLeft = 2
    mDotsRight = 2

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then Exit Sub

    Set shp = sr(1)

    widthM = GetShapeWidthM(shp)
    heightM = GetShapeHeightM(shp)

    mDotsTop = GetDefaultDotCount(widthM)
    mDotsBottom = GetDefaultDotCount(widthM)
    mDotsLeft = GetDefaultDotCount(heightM)
    mDotsRight = GetDefaultDotCount(heightM)
End Sub

Private Function GetDefaultDotCount(ByVal lengthM As Double) As Long
    Dim whole As Long

    whole = Int(lengthM)

    If lengthM > whole Then
        whole = whole + 1
    End If

    GetDefaultDotCount = whole + 1

    If GetDefaultDotCount < 2 Then
        GetDefaultDotCount = 2
    End If
End Function

Private Function IsDotFinishing() As Boolean
    If StrComp(Trim$(cmbFinishing.value), "SESTAND", vbTextCompare) = 0 Or _
       StrComp(Trim$(cmbFinishing.value), "MATIK", vbTextCompare) = 0 Then
        IsDotFinishing = True
    Else
        IsDotFinishing = False
    End If
End Function

Private Sub UpdateDotsButton()
    cmdAturDots.Enabled = IsDotFinishing()
End Sub

Private Sub cmbFinishing_Change()
    UpdateDotsButton
End Sub

Private Sub cmdAturDots_Click()
    Dim formDots As frmDots

    If Not IsDotFinishing() Then Exit Sub

    Set formDots = New frmDots

    formDots.TopCount = mDotsTop
    formDots.BottomCount = mDotsBottom
    formDots.LeftCount = mDotsLeft
    formDots.RightCount = mDotsRight

    formDots.Show vbModal

    If formDots.Confirmed Then
        mDotsTop = formDots.TopCount
        mDotsBottom = formDots.BottomCount
        mDotsLeft = formDots.LeftCount
        mDotsRight = formDots.RightCount
    End If

    Unload formDots
    Set formDots = Nothing
End Sub

Private Sub cmdSimpan_Click()
    Dim sr As ShapeRange
    Dim oneShape As ShapeRange
    Dim shp As Shape
    Dim canvas As Shape
    Dim labelGroup As Shape
    Dim i As Long
    Dim widthM As Double
    Dim heightM As Double
    Dim dotsTop As Long
    Dim dotsBottom As Long
    Dim dotsLeft As Long
    Dim dotsRight As Long

    If cmbSystem.ListIndex = -1 Then
        MsgBox "System wajib dipilih.", vbExclamation, "MgoCorel"
        cmbSystem.SetFocus
        Exit Sub
    End If

    If Trim$(txtNoInv.value) = "" Then
        MsgBox "No Inv wajib diisi.", vbExclamation, "MgoCorel"
        txtNoInv.SetFocus
        Exit Sub
    End If

    If Trim$(txtNama.value) = "" Then
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

    If Val(txtQuantity.value) <= 0 Then
        MsgBox "Quantity harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtQuantity.SetFocus
        Exit Sub
    End If

    If Trim$(txtDeadline.value) = "" Then
        MsgBox "Deadline wajib diisi.", vbExclamation, "MgoCorel"
        txtDeadline.SetFocus
        Exit Sub
    End If

    If cmbOperator.ListIndex = -1 Then
        MsgBox "Operator wajib dipilih.", vbExclamation, "MgoCorel"
        cmbOperator.SetFocus
        Exit Sub
    End If

    dotsTop = 0
    dotsBottom = 0
    dotsLeft = 0
    dotsRight = 0

    If IsDotFinishing() Then
        dotsTop = mDotsTop
        dotsBottom = mDotsBottom
        dotsLeft = mDotsLeft
        dotsRight = mDotsRight
    End If

    SaveLabelState Me

    For i = 1 To sr.Count
        Set shp = sr(i)

        widthM = GetShapeWidthM(shp)
        heightM = GetShapeHeightM(shp)

        Set oneShape = CreateShapeRange
        oneShape.Add shp

        Set canvas = CreateLabelCanvas(oneShape)

        Set labelGroup = CreateLabelTexts( _
            oneShape, _
            canvas, _
            cmbSystem.value, _
            txtNama.value, _
            cmbProduk.value, _
            widthM, _
            heightM, _
            cmbFinishing.value, _
            txtQuantity.value, _
            txtDeadline.value, _
            cmbOperator.value, _
            txtNoInv.value, _
            dotsTop, _
            dotsBottom, _
            dotsLeft, _
            dotsRight _
        )
    Next i

    Unload Me
End Sub

Private Sub cmdClear_Click()
    ClearLabelState Me
    InitializeDotCounts
    UpdateDotsButton
End Sub

Private Sub cmdBatal_Click()
    SaveLabelState Me
    Unload Me
End Sub

