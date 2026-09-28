Attribute VB_Name = "modDuplicate"
Option Explicit

Public Sub Plugin_ShowDuplicateForm()
    Dim sr As ShapeRange

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    frmDuplicate.Show vbModal

    Exit Sub

ErrHandler:
    MsgBox "Gagal membuka Duplicate Quantity." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Public Sub DuplicateAndImpose( _
    ByVal quantity As Long, _
    ByVal gapMm As Double)

    Dim sourceRange As ShapeRange
    Dim allRange As ShapeRange
    Dim duplicateRange As ShapeRange
    Dim sr As ShapeRange
    Dim sourceCount As Long
    Dim i As Long
    Dim j As Long
    Dim createdCount As Long

    On Error GoTo ErrHandler

    Set sourceRange = ActiveSelectionRange
    sourceCount = sourceRange.Count

    If sourceCount = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", _
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

    Set allRange = CreateShapeRange
    allRange.AddRange sourceRange

    If quantity > 1 Then
        For i = 1 To quantity - 1

            Set duplicateRange = sourceRange.Duplicate

            allRange.AddRange duplicateRange

            createdCount = createdCount + duplicateRange.Count

            If createdCount Mod 20 = 0 Then
                DoEvents
            End If

        Next i
    End If

    allRange.CreateSelection

    ArrangeImposition gapMm

    Exit Sub

ErrHandler:
    MsgBox "Gagal Duplicate Quantity." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub