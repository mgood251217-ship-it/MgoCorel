Attribute VB_Name = "modSusun"
Option Explicit

Public Sub Plugin_ShowSusunForm()
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Pilih objek terlebih dahulu.", vbExclamation, "MgoCorel"
        Exit Sub
    End If

    frmSusun.Show vbModal
End Sub

Public Sub ArrangeSelection( _
    ByVal gapMm As Double, _
    ByVal cutLine As Boolean, _
    ByVal targetPage As Boolean)

    Dim sr As ShapeRange
    Dim sourceRange As ShapeRange
    Dim targetShape As Shape
    Dim duplicateRange As ShapeRange
    Dim resultShape As Shape
    Dim cutShape As Shape
    Dim groupRange As ShapeRange

    Dim targetLeft As Double
    Dim targetBottom As Double
    Dim targetRight As Double
    Dim targetTop As Double

    Dim sourceLeft As Double
    Dim sourceTop As Double
    Dim sourceWidth As Double
    Dim sourceHeight As Double

    Dim gapDoc As Double

    Dim columns As Long
    Dim rows As Long
    Dim totalCreated As Long

    Dim usedWidth As Double
    Dim usedHeight As Double

    Dim startLeft As Double
    Dim startTop As Double

    Dim targetX As Double
    Dim targetY As Double

    Dim deltaX As Double
    Dim deltaY As Double

    Dim i As Long
    Dim j As Long

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih objek terlebih dahulu.", vbExclamation, "MgoCorel"
        Exit Sub
    End If

    gapDoc = Application.ConvertUnits( _
        gapMm, _
        cdrMillimeter, _
        ActiveDocument.Unit _
    )

    If gapDoc < 0 Then
        gapDoc = 0
    End If

    If targetPage Then
        Set sourceRange = sr

        targetLeft = ActivePage.LeftX
        targetBottom = ActivePage.BottomY
        targetRight = ActivePage.RightX
        targetTop = ActivePage.TopY
    Else
        If sr.Count < 2 Then
            MsgBox "Pilih objek sumber dan objek target.", _
                   vbExclamation, _
                   "MgoCorel"
            Exit Sub
        End If

        Set targetShape = ActiveShape

        Set sourceRange = CreateShapeRange

        For i = 1 To sr.Count
            If sr(i).Index <> targetShape.Index Then
                sourceRange.Add sr(i)
            End If
        Next i

        If sourceRange.Count = 0 Then
            MsgBox "Objek sumber tidak ditemukan.", vbExclamation, "MgoCorel"
            Exit Sub
        End If

        targetLeft = targetShape.LeftX
        targetBottom = targetShape.BottomY
        targetRight = targetShape.RightX
        targetTop = targetShape.TopY
    End If

    sourceLeft = sourceRange.LeftX
    sourceTop = sourceRange.TopY
    sourceWidth = sourceRange.SizeWidth
    sourceHeight = sourceRange.SizeHeight

    If sourceWidth <= 0 Or sourceHeight <= 0 Then
        MsgBox "Ukuran objek sumber tidak valid.", vbExclamation, "MgoCorel"
        Exit Sub
    End If

    If sourceWidth > targetRight - targetLeft Or _
       sourceHeight > targetTop - targetBottom Then

        MsgBox "Objek sumber lebih besar dari area target.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    columns = Int( _
        (targetRight - targetLeft + gapDoc) / _
        (sourceWidth + gapDoc) _
    )

    rows = Int( _
        (targetTop - targetBottom + gapDoc) / _
        (sourceHeight + gapDoc) _
    )

    If columns < 1 Then columns = 1
    If rows < 1 Then rows = 1

    usedWidth = _
        (columns * sourceWidth) + _
        ((columns - 1) * gapDoc)

    usedHeight = _
        (rows * sourceHeight) + _
        ((rows - 1) * gapDoc)

    startLeft = _
        targetLeft + _
        ((targetRight - targetLeft - usedWidth) / 2#)

    startTop = _
        targetTop - _
        ((targetTop - targetBottom - usedHeight) / 2#)

    For i = 0 To rows - 1
        For j = 0 To columns - 1

            targetX = _
                startLeft + _
                (j * (sourceWidth + gapDoc)) + _
                (sourceWidth / 2#)

            targetY = _
                startTop - _
                (i * (sourceHeight + gapDoc)) - _
                (sourceHeight / 2#)

            Set duplicateRange = sourceRange.Duplicate

            deltaX = targetX - sourceRange.CenterX
            deltaY = targetY - sourceRange.CenterY

            duplicateRange.Move deltaX, deltaY

            Set resultShape = Nothing

            If duplicateRange.Count > 1 Then
                Set resultShape = duplicateRange.Group
            Else
                Set resultShape = duplicateRange(1)
            End If

            If cutLine And IsRectangleSource(sourceRange) Then
                Set cutShape = CreateCutLine(resultShape)

                Set groupRange = CreateShapeRange
                groupRange.Add resultShape
                groupRange.Add cutShape

                Set resultShape = groupRange.Group
            End If

            totalCreated = totalCreated + 1

            If totalCreated Mod 20 = 0 Then
                DoEvents
            End If
        Next j
    Next i

    MsgBox "Susun objek selesai." & vbCrLf & vbCrLf & _
           "Kolom : " & columns & vbCrLf & _
           "Baris : " & rows & vbCrLf & _
           "Total : " & totalCreated, _
           vbInformation, _
           "MgoCorel"

    Exit Sub

ErrHandler:
    MsgBox "Gagal menyusun objek." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Private Function IsRectangleSource(ByVal sr As ShapeRange) As Boolean
    If sr.Count <> 1 Then Exit Function

    If sr(1).Type = cdrRectangleShape Then
        IsRectangleSource = True
    End If
End Function

Private Function CreateCutLine(ByVal shp As Shape) As Shape
    Dim cut As Shape

    Set cut = ActiveLayer.CreateRectangle2( _
        shp.LeftX, _
        shp.BottomY, _
        shp.SizeWidth, _
        shp.SizeHeight _
    )

    cut.Fill.ApplyNoFill
    cut.Outline.Width = ActiveDocument.ToUnits(762, cdrTenthMicron)
    cut.OrderToBack

    Set CreateCutLine = cut
End Function