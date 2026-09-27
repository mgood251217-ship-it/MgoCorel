Attribute VB_Name = "modExport"
Option Explicit

Public Sub Plugin_ExportLabel()
    Dim sr As ShapeRange
    Dim shp As Shape
    Dim exportFolder As String
    Dim yearFolder As String
    Dim monthFolder As String
    Dim dayFolder As String
    Dim bahanFolder As String
    Dim orderFolder As String
    Dim fileName As String
    Dim filePath As String
    Dim labelText As String
    Dim parts() As String
    Dim index As Long
    Dim systemName As String
    Dim nama As String
    Dim productName As String
    Dim ukuran As String
    Dim finishing As String
    Dim quantity As String
    Dim deadline As String
    Dim operatorName As String
    Dim invoiceNumber As String
    Dim labelDate As String
    Dim documentDate As Date
    Dim monthNames As Variant
    Dim opt As StructExportOptions
    Dim widthPx As Long
    Dim heightPx As Long
    Dim resolution As Double

    On Error GoTo ErrHandler

    Set sr = ActiveSelectionRange

    If sr.Count <> 1 Then
        MsgBox "Pilih satu group label terlebih dahulu.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    Set shp = sr(1)

    If shp.Type <> cdrGroupShape Then
        MsgBox "Objek yang dipilih bukan group label MgoCorel.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    labelText = Trim$(shp.Name)

    If labelText = "" Then
        MsgBox "Group label tidak memiliki nama.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    parts = Split(labelText, "_")

    index = 0

    If UCase$(parts(0)) = "ONLINE" Then
        systemName = "ONLINE"
        index = 1
    Else
        systemName = ""
    End If

    If UBound(parts) < index + 8 Then
        MsgBox "Format nama label tidak sesuai.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    nama = parts(index)
    productName = parts(index + 1)
    ukuran = parts(index + 2)
    finishing = parts(index + 3)
    quantity = parts(index + 4)
    deadline = parts(index + 5)
    labelDate = parts(index + 6)
    operatorName = parts(index + 7)
    invoiceNumber = parts(index + 8)

    documentDate = ParseLabelDate(labelDate)

    If documentDate = 0 Then
        documentDate = Date
    End If

    monthNames = Array( _
        "", _
        "Januari", _
        "Februari", _
        "Maret", _
        "April", _
        "Mei", _
        "Juni", _
        "Juli", _
        "Agustus", _
        "September", _
        "Oktober", _
        "November", _
        "Desember" _
    )

    exportFolder = Trim$(GetExportFolder())

    If exportFolder = "" Then
        MsgBox "Export Folder belum diatur.", _
               vbExclamation, _
               "MgoCorel"
        Exit Sub
    End If

    If Right$(exportFolder, 1) <> "\" Then
        exportFolder = exportFolder & "\"
    End If

    yearFolder = exportFolder & _
                 Format$(documentDate, "yyyy")

    monthFolder = yearFolder & "\" & _
                  Format$(documentDate, "mm") & " " & _
                  monthNames(Month(documentDate))

    dayFolder = monthFolder & "\" & _
                Format$(documentDate, "dd")

    bahanFolder = dayFolder & "\" & _
                  SanitizeFileName(productName)

    If systemName <> "" Then
        orderFolder = bahanFolder & "\" & _
                      SanitizeFileName( _
                          systemName & "_" & _
                          nama & "_" & _
                          deadline & "_" & _
                          operatorName & "_" & _
                          invoiceNumber _
                      )
    Else
        orderFolder = bahanFolder & "\" & _
                      SanitizeFileName( _
                          nama & "_" & _
                          deadline & "_" & _
                          operatorName & "_" & _
                          invoiceNumber _
                      )
    End If

    CreateFolderIfNotExists yearFolder
    CreateFolderIfNotExists monthFolder
    CreateFolderIfNotExists dayFolder
    CreateFolderIfNotExists bahanFolder
    CreateFolderIfNotExists orderFolder

    fileName = SanitizeFileName( _
        productName & "_" & _
        ukuran & "_" & _
        finishing & "_" & _
        quantity & ".jpg" _
    )

    filePath = GetUniqueFilePath(orderFolder, fileName)

    shp.CreateSelection

    resolution = 100#

    widthPx = CLng(Application.ConvertUnits( _
        shp.SizeWidth, _
        ActiveDocument.Unit, _
        cdrInch _
    ) * resolution)

    heightPx = CLng(Application.ConvertUnits( _
        shp.SizeHeight, _
        ActiveDocument.Unit, _
        cdrInch _
    ) * resolution)

    Set opt = New StructExportOptions

    opt.AntiAliasingType = cdrNormalAntiAliasing
    opt.ImageType = cdrRGBColorImage
    opt.Overwrite = False
    opt.ResolutionX = resolution
    opt.ResolutionY = resolution
    opt.SizeX = widthPx
    opt.SizeY = heightPx

    ActiveDocument.Export _
        filePath, _
        cdrJPEG, _
        cdrSelection, _
        opt

    MsgBox "Export berhasil." & vbCrLf & vbCrLf & _
           filePath, _
           vbInformation, _
           "MgoCorel"

    Exit Sub

ErrHandler:
    MsgBox "Gagal export label." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Private Function GetUniqueFilePath( _
    ByVal folderPath As String, _
    ByVal fileName As String) As String

    Dim baseName As String
    Dim extension As String
    Dim dotPos As Long
    Dim counter As Long
    Dim suffix As String
    Dim candidate As String

    dotPos = InStrRev(fileName, ".")

    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
        extension = Mid$(fileName, dotPos)
    Else
        baseName = fileName
        extension = ""
    End If

    candidate = folderPath & "\" & fileName

    If Dir$(candidate) = "" Then
        GetUniqueFilePath = candidate
        Exit Function
    End If

    counter = 1

    Do
        suffix = NumberToLetters(counter)

        candidate = folderPath & "\" & _
                    baseName & "_" & _
                    suffix & _
                    extension

        If Dir$(candidate) = "" Then
            GetUniqueFilePath = candidate
            Exit Function
        End If

        counter = counter + 1
    Loop
End Function

Private Function NumberToLetters(ByVal number As Long) As String
    Dim result As String
    Dim remainder As Long

    result = ""

    Do While number > 0
        remainder = (number - 1) Mod 26
        result = Chr$(65 + remainder) & result
        number = Int((number - 1) / 26)
    Loop

    NumberToLetters = result
End Function

Private Function ParseLabelDate(ByVal value As String) As Date
    Dim parts() As String
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long

    On Error GoTo ErrorHandler

    parts = Split(value, ".")

    If UBound(parts) <> 2 Then Exit Function

    dayValue = CLng(parts(0))
    monthValue = CLng(parts(1))
    yearValue = CLng(parts(2))

    ParseLabelDate = DateSerial( _
        yearValue, _
        monthValue, _
        dayValue _
    )

    Exit Function

ErrorHandler:
    ParseLabelDate = 0
End Function

Private Sub CreateFolderIfNotExists(ByVal folderPath As String)
    If Dir$(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

Private Function SanitizeFileName(ByVal value As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    invalidChars = Array( _
        "\", _
        "/", _
        ":", _
        "*", _
        "?", _
        """", _
        "<", _
        ">", _
        "|" _
    )

    For Each item In invalidChars
        value = Replace(value, CStr(item), "_")
    Next item

    SanitizeFileName = Trim$(value)
End Function