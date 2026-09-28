Attribute VB_Name = "modNumbering"
Option Explicit

Public Sub Plugin_ShowNumberingForm()
    Dim sr As ShapeRange

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih satu desain terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If sr.Count <> 1 Then
        MsgBox "Numbering menggunakan satu desain/template saja.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If Not HasNumberPlaceholder(sr(1)) Then
        MsgBox "Desain tidak memiliki placeholder numbering." & vbCrLf & vbCrLf & _
               "Gunakan salah satu:" & vbCrLf & _
               "{{NO}}" & vbCrLf & _
               "{{NO1}}" & vbCrLf & _
               "{{NO2}}", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    frmNumbering.Show vbModal

    Exit Sub

ErrHandler:
    MsgBox "Gagal membuka Numbering Otomatis." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Public Sub NumberAndImpose( _
    ByVal startNumber As Long, _
    ByVal digitCount As Long, _
    ByVal quantity As Long, _
    ByVal gapMm As Double)

    Dim sourceShape As Shape
    Dim duplicateShape As Shape
    Dim allRange As ShapeRange
    Dim sr As ShapeRange
    Dim i As Long
    Dim number1 As Long
    Dim number2 As Long
    Dim formatNumber As String
    Dim hasNo As Boolean
    Dim hasNo1 As Boolean
    Dim hasNo2 As Boolean
    Dim maxNumber As Long

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count <> 1 Then
        MsgBox "Pilih satu desain/template saja.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    Set sourceShape = sr(1)

    hasNo = HasPlaceholder(sourceShape, "{{NO}}")
    hasNo1 = HasPlaceholder(sourceShape, "{{NO1}}")
    hasNo2 = HasPlaceholder(sourceShape, "{{NO2}}")

    If Not hasNo And Not hasNo1 And Not hasNo2 Then
        MsgBox "Placeholder numbering tidak ditemukan.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If startNumber < 0 Then
        MsgBox "Nomor awal tidak boleh kurang dari 0.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If digitCount < 1 Or digitCount > 9 Then
        MsgBox "Jumlah digit harus antara 1 sampai 9.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If quantity < 1 Then
        MsgBox "Quantity harus lebih dari 0.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If gapMm < 0 Then
        MsgBox "Gap tidak boleh kurang dari 0.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    maxNumber = (10 ^ digitCount) - 1

    If hasNo1 And hasNo2 Then
        If startNumber + ((quantity - 1) * 2) + 1 > maxNumber Then
            MsgBox "Nomor melebihi jumlah digit yang dipilih.", _
                   vbExclamation, _
                   "MgoCorel"
            Exit Sub
        End If
    Else
        If startNumber + quantity - 1 > maxNumber Then
            MsgBox "Nomor melebihi jumlah digit yang dipilih.", _
                   vbExclamation, _
                   "MgoCorel"
            Exit Sub
        End If
    End If

    Set allRange = CreateShapeRange

    formatNumber = String$(digitCount, "0")

    For i = 1 To quantity

        If hasNo1 And hasNo2 Then
            number1 = startNumber + ((i - 1) * 2)
            number2 = number1 + 1
        Else
            number1 = startNumber + i - 1
            number2 = number1
        End If

        Set duplicateShape = sourceShape.Duplicate

        ReplaceNumberPlaceholders _
            duplicateShape, _
            Format$(number1, formatNumber), _
            Format$(number2, formatNumber), _
            Format$(number1, formatNumber)

        allRange.Add duplicateShape

        If i Mod 20 = 0 Then
            DoEvents
        End If

    Next i

    allRange.CreateSelection

    ArrangeImposition gapMm

    Exit Sub

ErrHandler:
    MsgBox "Gagal melakukan numbering." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Private Function HasNumberPlaceholder(ByVal shp As Shape) As Boolean
    If HasPlaceholder(shp, "{{NO}}") Then
        HasNumberPlaceholder = True
        Exit Function
    End If

    If HasPlaceholder(shp, "{{NO1}}") Then
        HasNumberPlaceholder = True
        Exit Function
    End If

    If HasPlaceholder(shp, "{{NO2}}") Then
        HasNumberPlaceholder = True
        Exit Function
    End If
End Function

Private Function HasPlaceholder( _
    ByVal shp As Shape, _
    ByVal placeholder As String) As Boolean

    Dim child As Shape
    Dim textValue As String

    If shp.Type = cdrGroupShape Then

        For Each child In shp.Shapes

            If HasPlaceholder(child, placeholder) Then
                HasPlaceholder = True
                Exit Function
            End If

        Next child

    ElseIf shp.Type = cdrTextShape Then

        textValue = shp.Text.Contents

        If InStr(1, textValue, placeholder, vbBinaryCompare) > 0 Then
            HasPlaceholder = True
            Exit Function
        End If

    End If

    On Error Resume Next

    If Not shp.PowerClip Is Nothing Then

        For Each child In shp.PowerClip.Shapes

            If HasPlaceholder(child, placeholder) Then
                HasPlaceholder = True
                Exit Function
            End If

        Next child

    End If

    On Error GoTo 0
End Function

Private Sub ReplaceNumberPlaceholders( _
    ByVal shp As Shape, _
    ByVal number1 As String, _
    ByVal number2 As String, _
    ByVal sameNumber As String)

    Dim child As Shape
    Dim textValue As String
    Dim newText As String

    If shp.Type = cdrGroupShape Then

        For Each child In shp.Shapes

            ReplaceNumberPlaceholders _
                child, _
                number1, _
                number2, _
                sameNumber

        Next child

    ElseIf shp.Type = cdrTextShape Then

        textValue = shp.Text.Contents
        newText = textValue

        newText = Replace( _
            newText, _
            "{{NO1}}", _
            number1, _
            1, _
            -1, _
            vbBinaryCompare _
        )

        newText = Replace( _
            newText, _
            "{{NO2}}", _
            number2, _
            1, _
            -1, _
            vbBinaryCompare _
        )

        newText = Replace( _
            newText, _
            "{{NO}}", _
            sameNumber, _
            1, _
            -1, _
            vbBinaryCompare _
        )

        If newText <> textValue Then
            shp.Text.Contents = newText
        End If

    End If

    On Error Resume Next

    If Not shp.PowerClip Is Nothing Then

        For Each child In shp.PowerClip.Shapes

            ReplaceNumberPlaceholders _
                child, _
                number1, _
                number2, _
                sameNumber

        Next child

    End If

    On Error GoTo 0
End Sub