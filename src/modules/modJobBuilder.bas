Attribute VB_Name = "modJobBuilder"
Option Explicit

Public Sub Plugin_ShowJobBuilderForm()
    Dim sr As ShapeRange

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count = 0 Then
        MsgBox "Pilih satu template terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If sr.Count <> 1 Then
        MsgBox "Job Builder menggunakan satu template saja.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    frmJobBuilder.Show vbModal

    Exit Sub

ErrHandler:
    MsgBox "Gagal membuka Job Builder." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Public Sub BuildJobBuilder( _
    ByVal filePath As String, _
    ByVal sheetName As String, _
    ByVal useImposition As Boolean, _
    ByVal gapMm As Double, _
    ByVal useUpperCase As Boolean)

    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object

    Dim sr As ShapeRange
    Dim sourceShape As Shape
    Dim duplicateShape As Shape
    Dim allRange As ShapeRange

    Dim headers As Collection
    Dim columns As Collection
    Dim dataRows As Collection

    Dim firstPage As Page

    Dim pageWidth As Double
    Dim pageHeight As Double
    Dim sourceWidth As Double
    Dim sourceHeight As Double
    Dim gapDoc As Double

    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim i As Long

    Dim headerName As String
    Dim variableCount As Long
    Dim dataCount As Long
    Dim outputCount As Long

    Dim columnsPerRow As Long
    Dim rowNumber As Long
    Dim columnIndex As Long

    Dim startLeft As Double
    Dim startTop As Double

    Dim targetX As Double
    Dim targetY As Double

    Dim deltaX As Double
    Dim deltaY As Double

    Dim commandStarted As Boolean

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count <> 1 Then
        MsgBox "Pilih satu template saja.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    Set sourceShape = sr(1)

    If Dir$(filePath) = "" Then
        MsgBox "File Excel tidak ditemukan.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    Set xlApp = CreateObject("Excel.Application")

    Set wb = xlApp.Workbooks.Open( _
        filePath, _
        False, _
        True _
    )

    Set ws = wb.Worksheets(sheetName)

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)

    If lastRow < 2 Then
        MsgBox "Tidak ada data pada Excel.", _
               vbExclamation, _
               "MgoCorel"
        GoTo CleanExit
    End If

    If lastCol < 1 Then
        MsgBox "Header variable tidak ditemukan.", _
               vbExclamation, _
               "MgoCorel"
        GoTo CleanExit
    End If

    Set headers = New Collection
    Set columns = New Collection
    Set dataRows = New Collection

    For colIndex = 1 To lastCol

        headerName = Trim$( _
            CStr(ws.Cells(1, colIndex).Value2) _
        )

        If headerName <> "" Then

            If CollectionContains( _
                headers, _
                headerName) Then

                MsgBox "Header Excel duplikat ditemukan:" & vbCrLf & _
                       headerName, _
                       vbExclamation, _
                       "MgoCorel"
                GoTo CleanExit
            End If

            headers.Add headerName
            columns.Add colIndex

        End If

    Next colIndex

    variableCount = headers.Count

    If variableCount = 0 Then
        MsgBox "Variable Excel tidak ditemukan.", _
               vbExclamation, _
               "MgoCorel"
        GoTo CleanExit
    End If

    For rowIndex = 2 To lastRow

        If RowHasData( _
            ws, _
            rowIndex, _
            lastCol) Then

            dataRows.Add rowIndex

        End If

    Next rowIndex

    dataCount = dataRows.Count

    If dataCount = 0 Then
        MsgBox "Tidak ada data yang bisa diproses.", _
               vbExclamation, _
               "MgoCorel"
        GoTo CleanExit
    End If

    outputCount = dataCount

    Set firstPage = ActivePage

    pageWidth = firstPage.SizeWidth
    pageHeight = firstPage.SizeHeight

    sourceWidth = sourceShape.SizeWidth
    sourceHeight = sourceShape.SizeHeight

    gapDoc = Application.ConvertUnits( _
        gapMm, _
        cdrMillimeter, _
        ActiveDocument.Unit _
    )

    If gapDoc < 0 Then
        gapDoc = 0
    End If

    ActiveDocument.BeginCommandGroup _
        "MgoCorel - Job Builder"

    commandStarted = True

    If useImposition Then

        Set allRange = CreateShapeRange

        For i = 1 To outputCount

            Set duplicateShape = sourceShape.Duplicate

            ReplaceExcelVariables _
                duplicateShape, _
                ws, _
                CLng(dataRows(i)), _
                headers, _
                columns, _
                useUpperCase

            allRange.Add duplicateShape

            If i Mod 20 = 0 Then
                DoEvents
            End If

        Next i

        allRange.CreateSelection

        ActiveDocument.EndCommandGroup
        commandStarted = False

        ArrangeImposition gapMm

    Else

        columnsPerRow = 5

        startLeft = _
            firstPage.RightX + _
            gapDoc + _
            (sourceWidth / 2#)

        startTop = _
            firstPage.TopY - _
            (sourceHeight / 2#)

        Set allRange = CreateShapeRange

        For i = 1 To outputCount

            Set duplicateShape = sourceShape.Duplicate

            ReplaceExcelVariables _
                duplicateShape, _
                ws, _
                CLng(dataRows(i)), _
                headers, _
                columns, _
                useUpperCase

            columnIndex = _
                (i - 1) Mod columnsPerRow

            rowNumber = _
                (i - 1) \ columnsPerRow

            targetX = _
                startLeft + _
                (columnIndex * _
                (sourceWidth + gapDoc))

            targetY = _
                startTop - _
                (rowNumber * _
                (sourceHeight + gapDoc))

            deltaX = _
                targetX - duplicateShape.CenterX

            deltaY = _
                targetY - duplicateShape.CenterY

            duplicateShape.Move _
                deltaX, _
                deltaY

            allRange.Add duplicateShape

            If i Mod 20 = 0 Then
                DoEvents
            End If

        Next i

        allRange.CreateSelection

        ActiveDocument.EndCommandGroup
        commandStarted = False

        firstPage.Activate

        MsgBox "Job Builder selesai." & vbCrLf & vbCrLf & _
               "Data       : " & outputCount & vbCrLf & _
               "Mode       : Satu Job / Baris" & vbCrLf & _
               "Kolom      : " & columnsPerRow & vbCrLf & _
               "Page       : Tidak digunakan", _
               vbInformation, _
               "MgoCorel"

    End If

CleanExit:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close False
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If

    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing

    On Error GoTo 0

    Exit Sub

ErrHandler:

    On Error Resume Next

    If commandStarted Then
        ActiveDocument.EndCommandGroup
    End If

    If Not wb Is Nothing Then
        wb.Close False
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If

    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing

    On Error GoTo 0

    MsgBox "Gagal menjalankan Job Builder." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Private Sub ReplaceExcelVariables( _
    ByVal shp As Shape, _
    ByVal ws As Object, _
    ByVal rowIndex As Long, _
    ByVal headers As Collection, _
    ByVal columns As Collection, _
    ByVal useUpperCase As Boolean)

    Dim child As Shape
    Dim textValue As String
    Dim newText As String
    Dim i As Long
    Dim headerName As String
    Dim placeholder As String
    Dim cellValue As String

    If shp.Type = cdrGroupShape Then
        For Each child In shp.Shapes
            ReplaceExcelVariables _
                child, _
                ws, _
                rowIndex, _
                headers, _
                columns, _
                useUpperCase
        Next child
    ElseIf shp.Type = cdrTextShape Then
        textValue = shp.Text.Contents
        newText = textValue

        For i = 1 To headers.Count
            headerName = Trim$(CStr(headers(i)))

            placeholder = "{{" & _
                          UCase$(headerName) & _
                          "}}"

            cellValue = GetExcelCellText( _
                ws, _
                rowIndex, _
                CLng(columns(i)) _
            )

            If useUpperCase Then
                cellValue = UCase$(cellValue)
            End If

            newText = Replace( _
                newText, _
                placeholder, _
                cellValue, _
                1, _
                -1, _
                vbTextCompare _
            )
        Next i

        If newText <> textValue Then
            shp.Text.Contents = newText
        End If
    End If

    On Error Resume Next

    If Not shp.PowerClip Is Nothing Then
        For Each child In shp.PowerClip.Shapes
            ReplaceExcelVariables _
                child, _
                ws, _
                rowIndex, _
                headers, _
                columns, _
                useUpperCase
        Next child
    End If

    On Error GoTo 0
End Sub

Private Function GetExcelCellText( _
    ByVal ws As Object, _
    ByVal rowIndex As Long, _
    ByVal colIndex As Long) As String

    Dim value As Variant
    Dim displayedValue As String

    On Error GoTo ErrorHandler

    displayedValue = CStr( _
        ws.Cells(rowIndex, colIndex).Text _
    )

    If displayedValue <> "" Then

        If InStr(displayedValue, "####") = 0 Then
            GetExcelCellText = displayedValue
            Exit Function
        End If

    End If

    value = ws.Cells(rowIndex, colIndex).Value2

    If IsError(value) Then
        GetExcelCellText = ""
    ElseIf IsEmpty(value) Then
        GetExcelCellText = ""
    Else
        GetExcelCellText = CStr(value)
    End If

    Exit Function

ErrorHandler:
    GetExcelCellText = ""
End Function

Private Function GetLastUsedRow(ByVal ws As Object) As Long
    Dim usedRange As Object

    Set usedRange = ws.UsedRange

    GetLastUsedRow = _
        usedRange.Row + _
        usedRange.Rows.Count - 1
End Function

Private Function GetLastUsedColumn(ByVal ws As Object) As Long
    Dim usedRange As Object

    Set usedRange = ws.UsedRange

    GetLastUsedColumn = _
        usedRange.Column + _
        usedRange.Columns.Count - 1
End Function

Private Function RowHasData( _
    ByVal ws As Object, _
    ByVal rowIndex As Long, _
    ByVal lastCol As Long) As Boolean

    Dim i As Long
    Dim value As Variant

    For i = 1 To lastCol

        value = ws.Cells(rowIndex, i).Value2

        If Not IsEmpty(value) Then

            If Trim$(CStr(value)) <> "" Then
                RowHasData = True
                Exit Function
            End If

        End If

    Next i
End Function

Private Function CollectionContains( _
    ByVal collection As Collection, _
    ByVal value As String) As Boolean

    Dim item As Variant

    For Each item In collection

        If StrComp( _
            Trim$(CStr(item)), _
            Trim$(value), _
            vbTextCompare) = 0 Then

            CollectionContains = True
            Exit Function

        End If

    Next item
End Function

Private Sub NormalizeShapeSize( _
    ByVal shp As Shape, _
    ByVal targetWidth As Double, _
    ByVal targetHeight As Double)

    Dim centerX As Double
    Dim centerY As Double

    centerX = shp.CenterX
    centerY = shp.CenterY

    shp.SizeWidth = targetWidth
    shp.SizeHeight = targetHeight

    shp.CenterX = centerX
    shp.CenterY = centerY
End Sub