Attribute VB_Name = "modHitung"
Option Explicit

Public Sub Plugin_ShowHitungForm()
    frmHitung.Show vbModal
End Sub

Public Sub CalculateSelectionPrice(ByVal form As Object)
    Dim sr As ShapeRange
    Dim shp As Shape
    Dim products As Collection
    Dim finishings As Collection
    Dim item As Variant
    Dim groups As Object
    Dim key As String
    Dim data As Variant
    Dim keys As Variant
    Dim i As Long
    Dim widthM As Double
    Dim heightM As Double
    Dim areaM2 As Double
    Dim quantity As Double
    Dim productPrice As Double
    Dim finishingPrice As Double
    Dim pricePerM2 As Double
    Dim unitPrice As Double
    Dim amount As Double
    Dim subtotal As Double
    Dim total As Double
    Dim productName As String
    Dim finishingName As String

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", vbExclamation, "MgoCorel"
        Exit Sub
    End If

    productName = Trim$(form.cmbProduk.Value)
    finishingName = Trim$(form.cmbFinishing.Value)

    If productName = "" Then
        MsgBox "Produk wajib dipilih.", vbExclamation, "MgoCorel"
        Exit Sub
    End If

    If finishingName = "" Then
        MsgBox "Finishing wajib dipilih.", vbExclamation, "MgoCorel"
        Exit Sub
    End If

    Set products = GetProducts()
    Set finishings = GetFinishings()

    For Each item In products
        If StrComp(CStr(item(0)), productName, vbTextCompare) = 0 Then
            productPrice = CDbl(item(1))
            Exit For
        End If
    Next item

    For Each item In finishings
        If StrComp(CStr(item(0)), finishingName, vbTextCompare) = 0 Then
            finishingPrice = CDbl(item(1))
            Exit For
        End If
    Next item

    pricePerM2 = productPrice + finishingPrice

    Set groups = CreateObject("Scripting.Dictionary")

    For i = 1 To sr.Count
        Set shp = sr(i)

        widthM = GetShapeWidthM(shp)
        heightM = GetShapeHeightM(shp)

        widthM = Round(widthM, 4)
        heightM = Round(heightM, 4)

        If widthM < heightM Then
            key = FormatSizeValue(widthM) & "X" & FormatSizeValue(heightM)
        Else
            key = FormatSizeValue(heightM) & "X" & FormatSizeValue(widthM)
        End If

        If groups.Exists(key) Then
            data = groups(key)
            data(2) = CDbl(data(2)) + 1
            groups(key) = data
        Else
            groups.Add key, Array(widthM, heightM, 1)
        End If
    Next i

    form.lstDetail.Clear

    keys = groups.Keys

    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        data = groups(key)

        widthM = CDbl(data(0))
        heightM = CDbl(data(1))
        quantity = CDbl(data(2))

        areaM2 = widthM * heightM
        unitPrice = areaM2 * pricePerM2
        amount = unitPrice * quantity

        subtotal = subtotal + amount

        With form.lstDetail
            .AddItem UCase$(productName)
            .List(.ListCount - 1, 1) = UCase$(finishingName)
            .List(.ListCount - 1, 2) = _
                FormatSizeValue(widthM) & "X" & FormatSizeValue(heightM)
            .List(.ListCount - 1, 3) = Format$(quantity, "0.##")
            .List(.ListCount - 1, 4) = FormatRupiah(unitPrice)
            .List(.ListCount - 1, 5) = FormatRupiah(amount)
        End With
    Next i

    If subtotal < pricePerM2 Then
        total = pricePerM2
    Else
        total = subtotal
    End If

    total = RoundDown500(total)

    form.lblHargaMeter.Caption = _
        "Harga / m² : " & FormatRupiah(pricePerM2)

    form.lblSubtotal.Caption = _
        "Subtotal : " & FormatRupiah(subtotal)

    form.lblTotal.Caption = _
        "TOTAL : " & FormatRupiah(total)
End Sub

Private Function RoundDown500(ByVal value As Double) As Double
    If value <= 0 Then
        RoundDown500 = 0
        Exit Function
    End If

    RoundDown500 = Int(value / 500#) * 500#
End Function

Private Function FormatSizeValue(ByVal value As Double) As String
    FormatSizeValue = Replace( _
        Format$(value, "0.####"), _
        ",", _
        "." _
    )
End Function

Private Function FormatRupiah(ByVal value As Double) As String
    Dim result As String

    result = Format$(Round(value, 0), "#,##0")
    result = Replace(result, ",", ".")

    FormatRupiah = "Rp " & result
End Function