VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Pengaturan MgoCorel"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600.001
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private selectedProductIndex As Long
Private selectedFinishingIndex As Long
Private selectedOperatorIndex As Long

Private Sub UserForm_Initialize()
    ApplyButtonTheme Me
    InitializeSettings
    LoadMonthFormat
    LoadProducts
    LoadFinishings
    LoadOperators
    txtOutputFolder.value = GetOutputFolder()
    txtExportFolder.value = GetExportFolder()
    txtTemplateFolder.value = GetTemplateFolder()
    txtCanvasExtra.value = CStr(GetCanvasExtra())
    txtLabelFont.value = GetLabelFont()
    txtLabelFontSize.value = CStr(GetLabelFontSize())
    chkLabelBold.value = GetLabelBold()
    chkProduk.value = GetUseProductFolder()
    selectedProductIndex = 0
    selectedFinishingIndex = 0
    selectedOperatorIndex = 0
    UpdatePenerapan
End Sub

Private Sub LoadMonthFormat()
    Dim currentValue As String
    currentValue = GetMonthFormat()
    cmbMonth.Clear
    cmbMonth.AddItem "10 Oktober"
    cmbMonth.AddItem "Oktober"
    cmbMonth.AddItem "10"
    Select Case currentValue
        Case "10 Oktober", "Oktober", "10"
            cmbMonth.value = currentValue
        Case Else
            cmbMonth.value = "10 Oktober"
    End Select
End Sub

Private Sub UpdatePenerapan()
    Dim basePath As String
    Dim yearFolder As String
    Dim monthFolder As String
    Dim dateFolder As String
    Dim productFolder As String
    Dim orderFolder As String
    Dim currentDate As Date

    basePath = Trim$(txtExportFolder.value)
    currentDate = Date
    yearFolder = Format$(currentDate, "yyyy")
    dateFolder = Format$(currentDate, "dd")
    orderFolder = "JALAM_13.00_VI_666666"

    Select Case cmbMonth.value
        Case "10 Oktober"
            monthFolder = Format$(currentDate, "mm") & " " & GetIndonesianMonthName(Month(currentDate))
        Case "Oktober"
            monthFolder = GetIndonesianMonthName(Month(currentDate))
        Case "10"
            monthFolder = Format$(currentDate, "mm")
        Case Else
            monthFolder = Format$(currentDate, "mm") & " " & GetIndonesianMonthName(Month(currentDate))
    End Select

    If chkProduk.value Then
        productFolder = "FX250"
    Else
        productFolder = ""
    End If

    If basePath = "" Then
        lblPenerapan.caption = yearFolder & "\" & monthFolder & "\" & dateFolder & "\"
        If chkProduk.value Then
            lblPenerapan.caption = lblPenerapan.caption & productFolder & "\"
        End If
        lblPenerapan.caption = lblPenerapan.caption & orderFolder
        Exit Sub
    End If

    If Right$(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    basePath = basePath & yearFolder & "\" & monthFolder & "\" & dateFolder & "\"

    If chkProduk.value Then
        basePath = basePath & productFolder & "\"
    End If

    basePath = basePath & orderFolder
    lblPenerapan.caption = basePath
End Sub

Private Function GetIndonesianMonthName(ByVal monthNumber As Long) As String
    Select Case monthNumber
        Case 1: GetIndonesianMonthName = "Januari"
        Case 2: GetIndonesianMonthName = "Februari"
        Case 3: GetIndonesianMonthName = "Maret"
        Case 4: GetIndonesianMonthName = "April"
        Case 5: GetIndonesianMonthName = "Mei"
        Case 6: GetIndonesianMonthName = "Juni"
        Case 7: GetIndonesianMonthName = "Juli"
        Case 8: GetIndonesianMonthName = "Agustus"
        Case 9: GetIndonesianMonthName = "September"
        Case 10: GetIndonesianMonthName = "Oktober"
        Case 11: GetIndonesianMonthName = "November"
        Case 12: GetIndonesianMonthName = "Desember"
    End Select
End Function

Private Sub cmbMonth_Change()
    UpdatePenerapan
End Sub

Private Sub chkProduk_Click()
    UpdatePenerapan
End Sub

Private Sub txtExportFolder_Change()
    UpdatePenerapan
End Sub

Private Sub LoadProducts()
    Dim item As Variant
    lstProduk.Clear
    For Each item In GetProducts()
        lstProduk.AddItem CStr(item(0))
        lstProduk.List(lstProduk.ListCount - 1, 1) = Format$(CDbl(item(1)), "0.00")
    Next item
End Sub

Private Sub LoadFinishings()
    Dim item As Variant
    lstFinishing.Clear
    For Each item In GetFinishings()
        lstFinishing.AddItem CStr(item(0))
        lstFinishing.List(lstFinishing.ListCount - 1, 1) = Format$(CDbl(item(1)), "0.00")
    Next item
End Sub

Private Sub LoadOperators()
    Dim item As Variant
    lstOperator.Clear
    For Each item In GetOperators()
        lstOperator.AddItem CStr(item)
    Next item
End Sub

Private Sub lstProduk_Click()
    If lstProduk.ListIndex = -1 Then Exit Sub
    selectedProductIndex = lstProduk.ListIndex + 1
    txtProductName.value = lstProduk.List(lstProduk.ListIndex, 0)
    txtProductPrice.value = lstProduk.List(lstProduk.ListIndex, 1)
End Sub

Private Sub lstFinishing_Click()
    If lstFinishing.ListIndex = -1 Then Exit Sub
    selectedFinishingIndex = lstFinishing.ListIndex + 1
    txtFinishingName.value = lstFinishing.List(lstFinishing.ListIndex, 0)
    txtFinishingPrice.value = lstFinishing.List(lstFinishing.ListIndex, 1)
End Sub

Private Sub lstOperator_Click()
    If lstOperator.ListIndex = -1 Then Exit Sub
    selectedOperatorIndex = lstOperator.ListIndex + 1
    txtOperatorName.value = lstOperator.List(lstOperator.ListIndex)
End Sub

Private Sub cmdProductAdd_Click()
    Dim productName As String
    Dim price As Double
    productName = Trim$(txtProductName.value)
    If productName = "" Then
        MsgBox "Nama produk wajib diisi.", vbExclamation, "MgoCorel"
        txtProductName.SetFocus
        Exit Sub
    End If
    If ProductNameExists(productName, 0) Then
        MsgBox "Produk sudah ada.", vbExclamation, "MgoCorel"
        txtProductName.SetFocus
        Exit Sub
    End If
    price = Val(txtProductPrice.value)
    AddProduct productName, price
    LoadProducts
    ClearProductInputs
End Sub

Private Sub cmdProductEdit_Click()
    Dim productName As String
    Dim price As Double
    If selectedProductIndex = 0 Then
        MsgBox "Pilih produk yang ingin diedit.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    productName = Trim$(txtProductName.value)
    If productName = "" Then
        MsgBox "Nama produk wajib diisi.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    If ProductNameExists(productName, selectedProductIndex) Then
        MsgBox "Produk sudah ada.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    price = Val(txtProductPrice.value)
    UpdateProduct selectedProductIndex, productName, price
    LoadProducts
    ClearProductInputs
End Sub

Private Sub cmdProductDelete_Click()
    If selectedProductIndex = 0 Then
        MsgBox "Pilih produk yang ingin dihapus.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    If MsgBox("Hapus produk yang dipilih?", vbQuestion Or vbYesNo, "MgoCorel") <> vbYes Then Exit Sub
    RemoveProduct selectedProductIndex
    LoadProducts
    ClearProductInputs
End Sub

Private Sub cmdProductClear_Click()
    ClearProductInputs
End Sub

Private Sub ClearProductInputs()
    selectedProductIndex = 0
    txtProductName.value = ""
    txtProductPrice.value = ""
    lstProduk.ListIndex = -1
End Sub

Private Sub cmdFinishingAdd_Click()
    Dim finishingName As String
    Dim price As Double
    finishingName = Trim$(txtFinishingName.value)
    If finishingName = "" Then
        MsgBox "Nama finishing wajib diisi.", vbExclamation, "MgoCorel"
        txtFinishingName.SetFocus
        Exit Sub
    End If
    If FinishingNameExists(finishingName, 0) Then
        MsgBox "Finishing sudah ada.", vbExclamation, "MgoCorel"
        txtFinishingName.SetFocus
        Exit Sub
    End If
    price = Val(txtFinishingPrice.value)
    AddFinishing finishingName, price
    LoadFinishings
    ClearFinishingInputs
End Sub

Private Sub cmdFinishingEdit_Click()
    Dim finishingName As String
    Dim price As Double
    If selectedFinishingIndex = 0 Then
        MsgBox "Pilih finishing yang ingin diedit.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    finishingName = Trim$(txtFinishingName.value)
    If finishingName = "" Then
        MsgBox "Nama finishing wajib diisi.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    If FinishingNameExists(finishingName, selectedFinishingIndex) Then
        MsgBox "Finishing sudah ada.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    price = Val(txtFinishingPrice.value)
    UpdateFinishing selectedFinishingIndex, finishingName, price
    LoadFinishings
    ClearFinishingInputs
End Sub

Private Sub cmdFinishingDelete_Click()
    If selectedFinishingIndex = 0 Then
        MsgBox "Pilih finishing yang ingin dihapus.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    If MsgBox("Hapus finishing yang dipilih?", vbQuestion Or vbYesNo, "MgoCorel") <> vbYes Then Exit Sub
    RemoveFinishing selectedFinishingIndex
    LoadFinishings
    ClearFinishingInputs
End Sub

Private Sub cmdFinishingClear_Click()
    ClearFinishingInputs
End Sub

Private Sub ClearFinishingInputs()
    selectedFinishingIndex = 0
    txtFinishingName.value = ""
    txtFinishingPrice.value = ""
    lstFinishing.ListIndex = -1
End Sub

Private Sub cmdOperatorAdd_Click()
    Dim operatorName As String
    operatorName = Trim$(txtOperatorName.value)
    If operatorName = "" Then
        MsgBox "Nama operator wajib diisi.", vbExclamation, "MgoCorel"
        txtOperatorName.SetFocus
        Exit Sub
    End If
    If OperatorNameExists(operatorName, 0) Then
        MsgBox "Operator sudah ada.", vbExclamation, "MgoCorel"
        txtOperatorName.SetFocus
        Exit Sub
    End If
    AddOperator operatorName
    LoadOperators
    ClearOperatorInputs
End Sub

Private Sub cmdOperatorEdit_Click()
    Dim operatorName As String
    If selectedOperatorIndex = 0 Then
        MsgBox "Pilih operator yang ingin diedit.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    operatorName = Trim$(txtOperatorName.value)
    If operatorName = "" Then
        MsgBox "Nama operator wajib diisi.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    If OperatorNameExists(operatorName, selectedOperatorIndex) Then
        MsgBox "Operator sudah ada.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    UpdateOperator selectedOperatorIndex, operatorName
    LoadOperators
    ClearOperatorInputs
End Sub

Private Sub cmdOperatorDelete_Click()
    If selectedOperatorIndex = 0 Then
        MsgBox "Pilih operator yang ingin dihapus.", vbExclamation, "MgoCorel"
        Exit Sub
    End If
    If MsgBox("Hapus operator yang dipilih?", vbQuestion Or vbYesNo, "MgoCorel") <> vbYes Then Exit Sub
    RemoveOperator selectedOperatorIndex
    LoadOperators
    ClearOperatorInputs
End Sub

Private Sub cmdOperatorClear_Click()
    ClearOperatorInputs
End Sub

Private Sub ClearOperatorInputs()
    selectedOperatorIndex = 0
    txtOperatorName.value = ""
    lstOperator.ListIndex = -1
End Sub

Private Sub cmdBrowseOutput_Click()
    Dim folderPath As String
    folderPath = PickFolder()
    If folderPath <> "" Then
        txtOutputFolder.value = folderPath
    End If
End Sub

Private Sub cmdBrowseExport_Click()
    Dim folderPath As String
    folderPath = PickFolder()
    If folderPath <> "" Then
        txtExportFolder.value = folderPath
    End If
End Sub

Private Sub cmdBrowseTemplate_Click()
    Dim folderPath As String
    folderPath = PickFolder()
    If folderPath <> "" Then
        txtTemplateFolder.value = folderPath
    End If
End Sub

Private Sub cmdSave_Click()
    Dim canvasExtra As Double
    Dim labelFont As String
    Dim labelFontSize As Double
    If cmbMonth.value = "" Then
        MsgBox "Format bulan wajib dipilih.", vbExclamation, "MgoCorel"
        cmbMonth.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtCanvasExtra.value) Then
        MsgBox "Tambahan Canvas harus berupa angka.", vbExclamation, "MgoCorel"
        txtCanvasExtra.SetFocus
        Exit Sub
    End If
    canvasExtra = CDbl(txtCanvasExtra.value)
    If canvasExtra <= 0 Then
        MsgBox "Tambahan Canvas harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtCanvasExtra.SetFocus
        Exit Sub
    End If
    labelFont = Trim$(txtLabelFont.value)
    If labelFont = "" Then
        MsgBox "Font label wajib diisi.", vbExclamation, "MgoCorel"
        txtLabelFont.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtLabelFontSize.value) Then
        MsgBox "Ukuran font harus berupa angka.", vbExclamation, "MgoCorel"
        txtLabelFontSize.SetFocus
        Exit Sub
    End If
    labelFontSize = CDbl(txtLabelFontSize.value)
    If labelFontSize <= 0 Then
        MsgBox "Ukuran font harus lebih dari 0.", vbExclamation, "MgoCorel"
        txtLabelFontSize.SetFocus
        Exit Sub
    End If
    SetOutputFolder Trim$(txtOutputFolder.value)
    SetExportFolder Trim$(txtExportFolder.value)
    SetTemplateFolder Trim$(txtTemplateFolder.value)
    SetCanvasExtra canvasExtra
    SetLabelFont labelFont
    SetLabelFontSize labelFontSize
    SetLabelBold CBool(chkLabelBold.value)
    SetMonthFormat cmbMonth.value
    SetUseProductFolder CBool(chkProduk.value)
    If SaveSettings(GetConfigPath()) Then
        UpdatePenerapan
        MsgBox "Pengaturan berhasil disimpan.", vbInformation, "MgoCorel"
    Else
        MsgBox "Pengaturan gagal disimpan.", vbCritical, "MgoCorel"
    End If
End Sub

Private Sub cmdExport_Click()
    Dim filePath As String
    filePath = PickExportFile()
    If filePath = "" Then Exit Sub
    If Not SaveCurrentFormSettings() Then Exit Sub
    If ExportSettings(filePath) Then
        MsgBox "Pengaturan berhasil di-export.", vbInformation, "MgoCorel"
    Else
        MsgBox "Pengaturan gagal di-export.", vbCritical, "MgoCorel"
    End If
End Sub

Private Sub cmdImport_Click()
    Dim filePath As String
    filePath = PickImportFile()
    If filePath = "" Then Exit Sub
    If MsgBox("Import akan mengganti pengaturan saat ini." & vbCrLf & "Lanjutkan?", vbQuestion Or vbYesNo, "MgoCorel") <> vbYes Then Exit Sub
    If ImportSettings(filePath) Then
        LoadMonthFormat
        LoadProducts
        LoadFinishings
        LoadOperators
        txtOutputFolder.value = GetOutputFolder()
        txtExportFolder.value = GetExportFolder()
        txtTemplateFolder.value = GetTemplateFolder()
        txtCanvasExtra.value = CStr(GetCanvasExtra())
        txtLabelFont.value = GetLabelFont()
        txtLabelFontSize.value = CStr(GetLabelFontSize())
        chkLabelBold.value = GetLabelBold()
        chkProduk.value = GetUseProductFolder()
        UpdatePenerapan
        MsgBox "Pengaturan berhasil di-import.", vbInformation, "MgoCorel"
    Else
        MsgBox "File pengaturan tidak valid atau gagal dibaca.", vbCritical, "MgoCorel"
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Function SaveCurrentFormSettings() As Boolean
    Dim canvasExtra As Double
    Dim labelFont As String
    Dim labelFontSize As Double
    If cmbMonth.value = "" Then
        MsgBox "Format bulan wajib dipilih.", vbExclamation, "MgoCorel"
        Exit Function
    End If
    If Not IsNumeric(txtCanvasExtra.value) Then
        MsgBox "Tambahan Canvas harus berupa angka.", vbExclamation, "MgoCorel"
        Exit Function
    End If
    canvasExtra = CDbl(txtCanvasExtra.value)
    If canvasExtra <= 0 Then
        MsgBox "Tambahan Canvas harus lebih dari 0.", vbExclamation, "MgoCorel"
        Exit Function
    End If
    labelFont = Trim$(txtLabelFont.value)
    If labelFont = "" Then
        MsgBox "Font label wajib diisi.", vbExclamation, "MgoCorel"
        Exit Function
    End If
    If Not IsNumeric(txtLabelFontSize.value) Then
        MsgBox "Ukuran font harus berupa angka.", vbExclamation, "MgoCorel"
        Exit Function
    End If
    labelFontSize = CDbl(txtLabelFontSize.value)
    If labelFontSize <= 0 Then
        MsgBox "Ukuran font harus lebih dari 0.", vbExclamation, "MgoCorel"
        Exit Function
    End If
    SetOutputFolder Trim$(txtOutputFolder.value)
    SetExportFolder Trim$(txtExportFolder.value)
    SetTemplateFolder Trim$(txtTemplateFolder.value)
    SetCanvasExtra canvasExtra
    SetLabelFont labelFont
    SetLabelFontSize labelFontSize
    SetLabelBold CBool(chkLabelBold.value)
    SetMonthFormat cmbMonth.value
    SetUseProductFolder CBool(chkProduk.value)
    SaveCurrentFormSettings = True
End Function

Private Function ProductNameExists(ByVal nameValue As String, ByVal exceptIndex As Long) As Boolean
    Dim items As collection
    Dim i As Long
    Dim item As Variant
    Set items = GetProducts()
    For i = 1 To items.count
        If i <> exceptIndex Then
            item = items(i)
            If StrComp(Trim$(CStr(item(0))), Trim$(nameValue), vbTextCompare) = 0 Then
                ProductNameExists = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FinishingNameExists(ByVal nameValue As String, ByVal exceptIndex As Long) As Boolean
    Dim items As collection
    Dim i As Long
    Dim item As Variant
    Set items = GetFinishings()
    For i = 1 To items.count
        If i <> exceptIndex Then
            item = items(i)
            If StrComp(Trim$(CStr(item(0))), Trim$(nameValue), vbTextCompare) = 0 Then
                FinishingNameExists = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function OperatorNameExists(ByVal nameValue As String, ByVal exceptIndex As Long) As Boolean
    Dim items As collection
    Dim i As Long
    Dim item As Variant
    Set items = GetOperators()
    For i = 1 To items.count
        If i <> exceptIndex Then
            item = items(i)
            If StrComp(Trim$(CStr(item)), Trim$(nameValue), vbTextCompare) = 0 Then
                OperatorNameExists = True
                Exit Function
            End If
        End If
    Next i
End Function
