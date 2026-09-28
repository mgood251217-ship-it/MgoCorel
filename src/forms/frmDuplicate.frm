VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuplicate 
   Caption         =   "Duplicate Quantity"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "frmDuplicate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    txtQuantity.value = "10"
    txtGap.value = "2"
    UpdateInfo
End Sub

Private Sub txtQuantity_Change()
    UpdateInfo
End Sub

Private Sub txtGap_Change()
    UpdateInfo
End Sub

Private Sub UpdateInfo()
    Dim sr As ShapeRange
    Dim shp As Shape
    Dim quantity As Long
    Dim totalCount As Long
    Dim gapMm As Double
    Dim gapDoc As Double
    Dim pageWidth As Double
    Dim pageHeight As Double
    Dim sourceWidth As Double
    Dim sourceHeight As Double
    Dim columns As Long
    Dim rows As Long
    Dim capacity As Long
    Dim totalPages As Long

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        lblInfo.Caption = _
            "Objek      : 0" & vbCrLf & _
            "Muat /page : -" & vbCrLf & _
            "Kolom      : -" & vbCrLf & _
            "Baris       : -" & vbCrLf & _
            "Total page : -"
        Exit Sub
    End If

    Set shp = sr(1)

    quantity = Val(txtQuantity.value)
    gapMm = Val(txtGap.value)

    If quantity < 1 Then
        lblInfo.Caption = _
            "Objek      : " & sr.Count & vbCrLf & _
            "Muat /page : -" & vbCrLf & _
            "Kolom      : -" & vbCrLf & _
            "Baris       : -" & vbCrLf & _
            "Total page : -"
        Exit Sub
    End If

    totalCount = sr.Count * quantity

    If gapMm < 0 Then
        gapMm = 0
    End If

    pageWidth = ActivePage.SizeWidth
    pageHeight = ActivePage.SizeHeight

    sourceWidth = shp.SizeWidth
    sourceHeight = shp.SizeHeight

    If sourceWidth > pageWidth Or _
       sourceHeight > pageHeight Then

        lblInfo.Caption = _
            "Objek      : " & totalCount & vbCrLf & _
            "Muat /page : 0" & vbCrLf & _
            "Kolom      : 0" & vbCrLf & _
            "Baris       : 0" & vbCrLf & _
            "Total page : -"

        Exit Sub
    End If

    gapDoc = Application.ConvertUnits( _
        gapMm, _
        cdrMillimeter, _
        ActiveDocument.Unit _
    )

    columns = Int( _
        (pageWidth + gapDoc) / _
        (sourceWidth + gapDoc) _
    )

    rows = Int( _
        (pageHeight + gapDoc) / _
        (sourceHeight + gapDoc) _
    )

    If columns < 1 Then
        columns = 1
    End If

    If rows < 1 Then
        rows = 1
    End If

    capacity = columns * rows

    totalPages = _
        (totalCount + capacity - 1) \ _
        capacity

    lblInfo.Caption = _
        "Objek      : " & totalCount & vbCrLf & _
        "Muat /page : " & capacity & vbCrLf & _
        "Kolom      : " & columns & vbCrLf & _
        "Baris       : " & rows & vbCrLf & _
        "Total page : " & totalPages
End Sub

Private Sub cmdDuplicate_Click()
    Dim quantity As Long
    Dim gapMm As Double

    quantity = CLng(Val(txtQuantity.value))
    gapMm = Val(txtGap.value)

    If quantity < 1 Then
        MsgBox "Quantity harus lebih dari 0.", _
               vbExclamation, _
               "MgoCorel"
        txtQuantity.SetFocus
        Exit Sub
    End If

    If gapMm < 0 Then
        MsgBox "Gap tidak boleh kurang dari 0.", _
               vbExclamation, _
               "MgoCorel"
        txtGap.SetFocus
        Exit Sub
    End If

    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    DuplicateAndImpose _
        quantity, _
        gapMm

    Unload Me
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

