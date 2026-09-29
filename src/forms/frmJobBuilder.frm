VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJobBuilder 
   Caption         =   "Job Builder"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   OleObjectBlob   =   "frmJobBuilder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJobBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private mDataCount As Long
Private mVariableCount As Long
Private mFileName As String

Private Sub UserForm_Initialize()
    txtGap.value = "2"

    optPerJob.value = True
    optImpose.value = False

    chkKapital.value = True

    With lstVariable
        .ColumnCount = 1
        .ColumnWidths = "100 pt"
        .IntegralHeight = False
        .Clear
    End With

    UpdateInfo
End Sub

Private Sub cmdBrowse_Click()
    Dim xlApp As Object
    Dim selectedFile As Variant

    On Error GoTo ErrHandler

    Set xlApp = CreateObject("Excel.Application")

    selectedFile = xlApp.GetOpenFileName( _
        "Excel Files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        1, _
        "Pilih File Excel" _
    )

    xlApp.Quit
    Set xlApp = Nothing

    If VarType(selectedFile) = vbBoolean Then
        Exit Sub
    End If

    txtFilePath.value = CStr(selectedFile)

    LoadSheets
    UpdateInfo

    Exit Sub

ErrHandler:
    On Error Resume Next

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If

    Set xlApp = Nothing

    On Error GoTo 0

    MsgBox "Gagal memilih file Excel." & vbCrLf & _
           "Pastikan Microsoft Excel terpasang.", _
           vbCritical, _
           "MgoCorel"
End Sub

Private Sub cmbSheet_Change()
    LoadSheetInfo
    UpdateInfo
End Sub

Private Sub optPerJob_Click()
    UpdateInfo
End Sub

Private Sub optImpose_Click()
    UpdateInfo
End Sub

Private Sub txtGap_Change()
    UpdateInfo
End Sub

Private Sub LoadSheets()
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object

    On Error GoTo ErrHandler

    cmbSheet.Clear
    lstVariable.Clear

    mDataCount = 0
    mVariableCount = 0

    Set xlApp = CreateObject("Excel.Application")

    Set wb = xlApp.Workbooks.Open( _
        txtFilePath.value, _
        False, _
        True _
    )

    For Each ws In wb.Worksheets
        cmbSheet.AddItem ws.Name
    Next ws

    wb.Close False
    xlApp.Quit

    Set wb = Nothing
    Set xlApp = Nothing

    If cmbSheet.ListCount > 0 Then
        cmbSheet.ListIndex = 0
    End If

    Exit Sub

ErrHandler:
    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close False
    End If

    If Not xlApp Is Nothing Then
        xlApp.Quit
    End If

    Set wb = Nothing
    Set xlApp = Nothing

    On Error GoTo 0

    MsgBox "Gagal membaca sheet Excel." & vbCrLf & _
           "Error " & Err.number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Private Sub LoadSheetInfo()
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    Dim data As Variant

    Dim lastRow As Long
    Dim lastCol As Long

    Dim i As Long
    Dim j As Long

    Dim headerName As String
    Dim columnWidth As String
    Dim maxLength As Long
    Dim cellText As String

    On Error GoTo ErrHandler

    lstVariable.Clear

    mDataCount = 0
    mVariableCount = 0

    If Trim$(txtFilePath.value) = "" Then
        UpdateInfo
        Exit Sub
    End If

    If cmbSheet.ListIndex = -1 Then
        UpdateInfo
        Exit Sub
    End If

    Set xlApp = CreateObject("Excel.Application")

    Set wb = xlApp.Workbooks.Open( _
        txtFilePath.value, _
        False, _
        True _
    )

    Set ws = wb.Worksheets(cmbSheet.value)

    lastRow = GetLastUsedRowLocal(ws)
    lastCol = GetLastUsedColumnLocal(ws)

    If lastRow < 1 Or lastCol < 1 Then
        GoTo CleanExit
    End If

    data = ws.Range( _
        ws.Cells(1, 1), _
        ws.Cells(lastRow, lastCol) _
    ).Value2

    With lstVariable
        .ColumnCount = lastCol
        .Clear
    End With

    If lastCol = 1 Then

        If IsArray(data) Then
            lstVariable.List = data
        Else
            lstVariable.AddItem CStr(data)
        End If

    Else
        lstVariable.List = data
    End If

    mVariableCount = 0

    For j = 1 To lastCol

        headerName = Trim$( _
            CStr(ws.Cells(1, j).Value2) _
        )

        If headerName <> "" Then
            mVariableCount = mVariableCount + 1
        End If

    Next j

    mDataCount = 0

    For i = 2 To lastRow

        If RowHasDataLocal( _
            ws, _
            i, _
            lastCol) Then

            mDataCount = mDataCount + 1

        End If

    Next i

    columnWidth = ""

    For j = 1 To lastCol

        maxLength = Len( _
            CStr(ws.Cells(1, j).Value2) _
        )

        For i = 2 To lastRow

            cellText = GetCellDisplayValue( _
                ws, _
                i, _
                j _
            )

            If Len(cellText) > maxLength Then
                maxLength = Len(cellText)
            End If

        Next i

        If maxLength < 8 Then
            maxLength = 8
        End If

        If maxLength > 25 Then
            maxLength = 25
        End If

        columnWidth = columnWidth & _
                      CStr(maxLength * 6) & " pt;"

    Next j

    If Len(columnWidth) > 0 Then
        columnWidth = Left$( _
            columnWidth, _
            Len(columnWidth) - 1 _
        )
    End If

    lstVariable.ColumnWidths = columnWidth

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

    UpdateInfo

    Exit Sub

ErrHandler:

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

    MsgBox "Gagal membaca data Excel." & vbCrLf & _
           "Error " & Err.number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"

    UpdateInfo
End Sub

Private Sub UpdateInfo()
    Dim modeName As String
    Dim templateCount As Long

    If optImpose.value Then
        modeName = "Job + Imposisi"
    Else
        modeName = "Satu Job / Baris"
    End If

    templateCount = ActiveSelectionRange.Count

    lblInfo.Caption = _
        "Variable   : " & mVariableCount & vbCrLf & _
        "Data       : " & mDataCount & vbCrLf & _
        "Mode       : " & modeName & vbCrLf & _
        "Template   : " & templateCount & " object" & vbCrLf & _
        "Sheet      : " & cmbSheet.value
End Sub

Private Sub cmdProses_Click()
    Dim gapMm As Double

    If Trim$(txtFilePath.value) = "" Then
        MsgBox "File Excel wajib dipilih.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If Dir$(txtFilePath.value) = "" Then
        MsgBox "File Excel tidak ditemukan.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If cmbSheet.ListIndex = -1 Then
        MsgBox "Sheet wajib dipilih.", _
               vbExclamation, _
               "MgoCorel"
        cmbSheet.SetFocus
        Exit Sub
    End If

    If ActiveSelectionRange.Count <> 1 Then
        MsgBox "Pilih satu template terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If mVariableCount = 0 Then
        MsgBox "Variable Excel tidak ditemukan.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If mDataCount = 0 Then
        MsgBox "Data Excel tidak ditemukan.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    gapMm = Val(txtGap.value)

    If gapMm < 0 Then
        MsgBox "Gap tidak boleh kurang dari 0.", _
               vbExclamation, _
               "MgoCorel"
        txtGap.SetFocus
        Exit Sub
    End If

    BuildJobBuilder _
        txtFilePath.value, _
        cmbSheet.value, _
        optImpose.value, _
        gapMm, _
        chkKapital.value

    Unload Me
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Function GetLastUsedRowLocal(ByVal ws As Object) As Long
    Dim usedRange As Object

    Set usedRange = ws.usedRange

    GetLastUsedRowLocal = _
        usedRange.Row + _
        usedRange.rows.Count - 1
End Function

Private Function GetLastUsedColumnLocal(ByVal ws As Object) As Long
    Dim usedRange As Object

    Set usedRange = ws.usedRange

    GetLastUsedColumnLocal = _
        usedRange.Column + _
        usedRange.columns.Count - 1
End Function

Private Function RowHasDataLocal( _
    ByVal ws As Object, _
    ByVal rowIndex As Long, _
    ByVal lastCol As Long) As Boolean

    Dim i As Long
    Dim value As Variant

    For i = 1 To lastCol

        value = ws.Cells(rowIndex, i).Value2

        If Not IsEmpty(value) Then

            If Trim$(CStr(value)) <> "" Then
                RowHasDataLocal = True
                Exit Function
            End If

        End If

    Next i
End Function

Private Function GetCellDisplayValue( _
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
            GetCellDisplayValue = displayedValue
            Exit Function
        End If

    End If

    value = ws.Cells(rowIndex, colIndex).Value2

    If IsError(value) Then
        GetCellDisplayValue = ""
    ElseIf IsEmpty(value) Then
        GetCellDisplayValue = ""
    Else
        GetCellDisplayValue = CStr(value)
    End If

    Exit Function

ErrorHandler:
    GetCellDisplayValue = ""
End Function

Private Function GetFileNameOnly(ByVal filePath As String) As String
    GetFileNameOnly = Mid$( _
        filePath, _
        InStrRev(filePath, "\") + 1 _
    )
End Function

