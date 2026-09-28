Attribute VB_Name = "modImposisi"
Option Explicit

Public Sub Plugin_ShowImposisiForm()
    Dim sr As ShapeRange

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    frmImposisi.Show vbModal

    Exit Sub

ErrHandler:
    MsgBox "Gagal membuka Imposisi Otomatis." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Public Sub ArrangeImposition(ByVal gapMm As Double)
    Dim sr As ShapeRange
    Dim items() As Shape
    Dim shp As Shape
    Dim firstPage As Page
    Dim currentPage As Page
    Dim firstNewPage As Page

    Dim sourceWidth As Double
    Dim sourceHeight As Double
    Dim pageWidth As Double
    Dim pageHeight As Double
    Dim gapDoc As Double

    Dim columns As Long
    Dim rows As Long
    Dim capacity As Long

    Dim totalPages As Long
    Dim pageNumber As Long
    Dim pageItemIndex As Long
    Dim itemIndex As Long
    Dim itemsOnPage As Long

    Dim pageColumns As Long
    Dim pageRows As Long

    Dim usedWidth As Double
    Dim usedHeight As Double
    Dim startLeft As Double
    Dim startTop As Double

    Dim columnIndex As Long
    Dim rowIndex As Long

    Dim targetX As Double
    Dim targetY As Double
    Dim deltaX As Double
    Dim deltaY As Double

    Dim totalCount As Long
    Dim commandStarted As Boolean
    Dim i As Long

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange
    totalCount = sr.Count

    If totalCount = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If Not ValidateSameSize( _
        sr, _
        sourceWidth, _
        sourceHeight) Then

        MsgBox "Semua objek yang dipilih harus memiliki ukuran yang sama.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    Set firstPage = ActivePage

    pageWidth = firstPage.SizeWidth
    pageHeight = firstPage.SizeHeight

    gapDoc = Application.ConvertUnits( _
        gapMm, _
        cdrMillimeter, _
        ActiveDocument.Unit _
    )

    If gapDoc < 0 Then
        gapDoc = 0
    End If

    If sourceWidth > pageWidth Or _
       sourceHeight > pageHeight Then

        MsgBox "Ukuran objek lebih besar dari ukuran page.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

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
        (totalCount + capacity - 1) \ capacity

    If totalPages < 1 Then
        totalPages = 1
    End If

    ReDim items(1 To totalCount)

    For i = 1 To totalCount
        Set items(i) = sr(i)
    Next i

    If totalPages > 1 Then
        Set firstNewPage = ActiveDocument.InsertPagesEx( _
            totalPages - 1, _
            False, _
            firstPage.Index, _
            pageWidth, _
            pageHeight _
        )
    End If

    ActiveDocument.BeginCommandGroup _
        "MgoCorel - Imposisi Otomatis"

    commandStarted = True

    itemIndex = 1

    For pageNumber = 1 To totalPages

        If pageNumber = 1 Then
            Set currentPage = firstPage
        Else
            Set currentPage = ActiveDocument.Pages( _
                firstNewPage.Index + pageNumber - 2 _
            )
        End If

        itemsOnPage = _
            totalCount - itemIndex + 1

        If itemsOnPage > capacity Then
            itemsOnPage = capacity
        End If

        pageColumns = columns

        If itemsOnPage < pageColumns Then
            pageColumns = itemsOnPage
        End If

        pageRows = _
            (itemsOnPage + pageColumns - 1) \ _
            pageColumns

        usedWidth = _
            (pageColumns * sourceWidth) + _
            ((pageColumns - 1) * gapDoc)

        usedHeight = _
            (pageRows * sourceHeight) + _
            ((pageRows - 1) * gapDoc)

        startLeft = _
            currentPage.LeftX + _
            ((pageWidth - usedWidth) / 2#)

        startTop = _
            currentPage.TopY - _
            ((pageHeight - usedHeight) / 2#)

        For pageItemIndex = 1 To itemsOnPage

            Set shp = items(itemIndex)

            If pageNumber > 1 Then
                shp.MoveToLayer currentPage.ActiveLayer
            End If

            columnIndex = _
                (pageItemIndex - 1) Mod pageColumns

            rowIndex = _
                (pageItemIndex - 1) \ pageColumns

            targetX = _
                startLeft + _
                (sourceWidth / 2#) + _
                (columnIndex * _
                (sourceWidth + gapDoc))

            targetY = _
                startTop - _
                (sourceHeight / 2#) - _
                (rowIndex * _
                (sourceHeight + gapDoc))

            deltaX = targetX - shp.CenterX
            deltaY = targetY - shp.CenterY

            shp.Move deltaX, deltaY

            itemIndex = itemIndex + 1

            If itemIndex Mod 20 = 0 Then
                DoEvents
            End If

        Next pageItemIndex

    Next pageNumber

    ActiveDocument.EndCommandGroup
    commandStarted = False

    items(totalCount).CreateSelection

    MsgBox "Imposisi selesai." & vbCrLf & vbCrLf & _
           "Total objek : " & totalCount & vbCrLf & _
           "Muat / page : " & capacity & vbCrLf & _
           "Kolom       : " & columns & vbCrLf & _
           "Baris       : " & rows & vbCrLf & _
           "Total page  : " & totalPages, _
           vbInformation, _
           "MgoCorel"

    Exit Sub

ErrHandler:
    If commandStarted Then
        On Error Resume Next
        ActiveDocument.EndCommandGroup
        On Error GoTo 0
    End If

    MsgBox "Gagal melakukan imposisi." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Private Function ValidateSameSize( _
    ByVal sr As ShapeRange, _
    ByRef widthValue As Double, _
    ByRef heightValue As Double) As Boolean

    Dim i As Long
    Dim tolerance As Double

    If sr.Count = 0 Then
        Exit Function
    End If

    tolerance = 0.000001

    widthValue = sr(1).SizeWidth
    heightValue = sr(1).SizeHeight

    For i = 2 To sr.Count

        If Abs( _
            sr(i).SizeWidth - widthValue _
        ) > tolerance Then
            Exit Function
        End If

        If Abs( _
            sr(i).SizeHeight - heightValue _
        ) > tolerance Then
            Exit Function
        End If

    Next i

    ValidateSameSize = True
End Function