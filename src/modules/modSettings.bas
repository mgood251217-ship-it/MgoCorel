Attribute VB_Name = "modSettings"
Option Explicit
Private Const CONFIG_FOLDER As String = "\MgoCorel"
Private Const CONFIG_FILE As String = "\settings.mgo"
Private mProducts As Collection
Private mFinishings As Collection
Private mOperators As Collection
Private mOutputFolder As String
Private mExportFolder As String
Private mTemplateFolder As String
Private mCanvasExtra As Double
Private mLabelFont As String
Private mLabelFontSize As Double
Private mLabelBold As Boolean
Private mInitialized As Boolean
#If VBA7 Then
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
End Type
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
#Else
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
#End If

Public Sub InitializeSettings()
    If mInitialized Then Exit Sub
    Set mProducts = New Collection
    Set mFinishings = New Collection
    Set mOperators = New Collection
    mOutputFolder = ""
    mExportFolder = ""
    mTemplateFolder = ""
    mCanvasExtra = 5
    mLabelFont = "Arial"
    mLabelFontSize = 16
    mLabelBold = True
    mInitialized = True
    If Not LoadSettings(GetConfigPath()) Then
        CreateDefaultSettings
        SaveSettings GetConfigPath()
    End If
End Sub

Public Sub Plugin_ShowSettingsForm()
    InitializeSettings
    frmSettings.Show
End Sub

Public Function GetConfigPath() As String
    GetConfigPath = Environ$("APPDATA") & CONFIG_FOLDER & CONFIG_FILE
End Function

Public Function GetProducts() As Collection
    InitializeSettings
    Set GetProducts = mProducts
End Function

Public Function GetFinishings() As Collection
    InitializeSettings
    Set GetFinishings = mFinishings
End Function

Public Function GetOperators() As Collection
    InitializeSettings
    Set GetOperators = mOperators
End Function

Public Function GetOutputFolder() As String
    InitializeSettings
    GetOutputFolder = mOutputFolder
End Function

Public Function GetExportFolder() As String
    InitializeSettings
    GetExportFolder = mExportFolder
End Function

Public Function GetTemplateFolder() As String
    InitializeSettings
    GetTemplateFolder = mTemplateFolder
End Function

Public Function GetCanvasExtra() As Double
    InitializeSettings
    GetCanvasExtra = mCanvasExtra
End Function

Public Function GetLabelFont() As String
    InitializeSettings
    GetLabelFont = mLabelFont
End Function

Public Function GetLabelFontSize() As Double
    InitializeSettings
    GetLabelFontSize = mLabelFontSize
End Function

Public Function GetLabelBold() As Boolean
    InitializeSettings
    GetLabelBold = mLabelBold
End Function

Public Sub AddProduct(ByVal productName As String, ByVal price As Double)
    InitializeSettings
    mProducts.Add Array(productName, price)
End Sub

Public Sub UpdateProduct(ByVal index As Long, ByVal productName As String, ByVal price As Double)
    InitializeSettings
    If index < 1 Or index > mProducts.Count Then Exit Sub
    mProducts.Remove index
    mProducts.Add Array(productName, price), , index
End Sub

Public Sub RemoveProduct(ByVal index As Long)
    InitializeSettings
    If index >= 1 And index <= mProducts.Count Then
        mProducts.Remove index
    End If
End Sub

Public Sub AddFinishing(ByVal finishingName As String, ByVal price As Double)
    InitializeSettings
    mFinishings.Add Array(finishingName, price)
End Sub

Public Sub UpdateFinishing(ByVal index As Long, ByVal finishingName As String, ByVal price As Double)
    InitializeSettings
    If index < 1 Or index > mFinishings.Count Then Exit Sub
    mFinishings.Remove index
    mFinishings.Add Array(finishingName, price), , index
End Sub

Public Sub RemoveFinishing(ByVal index As Long)
    InitializeSettings
    If index >= 1 And index <= mFinishings.Count Then
        mFinishings.Remove index
    End If
End Sub

Public Sub AddOperator(ByVal operatorName As String)
    InitializeSettings
    mOperators.Add operatorName
End Sub

Public Sub UpdateOperator(ByVal index As Long, ByVal operatorName As String)
    InitializeSettings
    If index < 1 Or index > mOperators.Count Then Exit Sub
    mOperators.Remove index
    mOperators.Add operatorName, , index
End Sub

Public Sub RemoveOperator(ByVal index As Long)
    InitializeSettings
    If index >= 1 And index <= mOperators.Count Then
        mOperators.Remove index
    End If
End Sub

Public Sub SetOutputFolder(ByVal folderPath As String)
    InitializeSettings
    mOutputFolder = folderPath
End Sub

Public Sub SetExportFolder(ByVal folderPath As String)
    InitializeSettings
    mExportFolder = folderPath
End Sub

Public Sub SetTemplateFolder(ByVal folderPath As String)
    InitializeSettings
    mTemplateFolder = folderPath
End Sub

Public Sub SetCanvasExtra(ByVal value As Double)
    InitializeSettings
    mCanvasExtra = value
End Sub

Public Sub SetLabelFont(ByVal value As String)
    InitializeSettings
    mLabelFont = value
End Sub

Public Sub SetLabelFontSize(ByVal value As Double)
    InitializeSettings
    mLabelFontSize = value
End Sub

Public Sub SetLabelBold(ByVal value As Boolean)
    InitializeSettings
    mLabelBold = value
End Sub

Public Function SaveSettings(ByVal filePath As String) As Boolean
    Dim fileNo As Integer
    Dim item As Variant
    Dim folderPath As String

    On Error GoTo ErrHandler

    InitializeSettings

    folderPath = Left$(filePath, InStrRev(filePath, "\") - 1)

    If Dir$(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    fileNo = FreeFile

    Open filePath For Output As #fileNo

    Print #fileNo, "[MgoCorel]"
    Print #fileNo, "Version=1"
    Print #fileNo, "OutputFolder=" & CleanValue(mOutputFolder)
    Print #fileNo, "ExportFolder=" & CleanValue(mExportFolder)
    Print #fileNo, "TemplateFolder=" & CleanValue(mTemplateFolder)
    Print #fileNo, "CanvasExtra=" & NumberToString(mCanvasExtra)
    Print #fileNo, "LabelFont=" & CleanValue(mLabelFont)
    Print #fileNo, "LabelFontSize=" & NumberToString(mLabelFontSize)
    Print #fileNo, "LabelBold=" & IIf(mLabelBold, "1", "0")

    Print #fileNo, "[Products]"

    For Each item In mProducts
        Print #fileNo, CleanValue(CStr(item(0))) & "|" & NumberToString(CDbl(item(1)))
    Next item

    Print #fileNo, "[Finishings]"

    For Each item In mFinishings
        Print #fileNo, CleanValue(CStr(item(0))) & "|" & NumberToString(CDbl(item(1)))
    Next item

    Print #fileNo, "[Operators]"

    For Each item In mOperators
        Print #fileNo, CleanValue(CStr(item))
    Next item

    Close #fileNo

    SaveSettings = True
    Exit Function

ErrHandler:
    On Error Resume Next

    If fileNo <> 0 Then
        Close #fileNo
    End If

    SaveSettings = False
End Function

Public Function LoadSettings(ByVal filePath As String) As Boolean
    Dim fileNo As Integer
    Dim lineText As String
    Dim sectionName As String
    Dim parts() As String
    Dim tempProducts As Collection
    Dim tempFinishings As Collection
    Dim tempOperators As Collection
    Dim tempOutputFolder As String
    Dim tempExportFolder As String
    Dim tempTemplateFolder As String
    Dim tempCanvasExtra As Double
    Dim tempLabelFont As String
    Dim tempLabelFontSize As Double
    Dim tempLabelBold As Boolean

    On Error GoTo ErrHandler

    If Dir$(filePath) = "" Then Exit Function

    Set tempProducts = New Collection
    Set tempFinishings = New Collection
    Set tempOperators = New Collection

    tempCanvasExtra = 5
    tempLabelFont = "Arial"
    tempLabelFontSize = 16
    tempLabelBold = True

    fileNo = FreeFile

    Open filePath For Input As #fileNo

    Do While Not EOF(fileNo)
        Line Input #fileNo, lineText
        lineText = Trim$(lineText)

        If lineText = "" Then
            GoTo ContinueLoop
        End If

        If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
            sectionName = Mid$(lineText, 2, Len(lineText) - 2)
            GoTo ContinueLoop
        End If

        Select Case sectionName
            Case "MgoCorel"
                If Left$(lineText, 13) = "OutputFolder=" Then
                    tempOutputFolder = Mid$(lineText, 14)
                ElseIf Left$(lineText, 13) = "ExportFolder=" Then
                    tempExportFolder = Mid$(lineText, 14)
                ElseIf Left$(lineText, 15) = "TemplateFolder=" Then
                    tempTemplateFolder = Mid$(lineText, 16)
                ElseIf Left$(lineText, 12) = "CanvasExtra=" Then
                    tempCanvasExtra = StringToNumber(Mid$(lineText, 13))
                ElseIf Left$(lineText, 10) = "LabelFont=" Then
                    tempLabelFont = Mid$(lineText, 11)
                ElseIf Left$(lineText, 14) = "LabelFontSize=" Then
                    tempLabelFontSize = StringToNumber(Mid$(lineText, 15))
                ElseIf Left$(lineText, 10) = "LabelBold=" Then
                    tempLabelBold = Val(Mid$(lineText, 11)) <> 0
                End If

            Case "Products"
                parts = Split(lineText, "|")

                If UBound(parts) >= 1 Then
                    tempProducts.Add Array( _
                        parts(0), _
                        StringToNumber(parts(1)) _
                    )
                End If

            Case "Finishings"
                parts = Split(lineText, "|")

                If UBound(parts) >= 1 Then
                    tempFinishings.Add Array( _
                        parts(0), _
                        StringToNumber(parts(1)) _
                    )
                End If

            Case "Operators"
                tempOperators.Add lineText
        End Select

ContinueLoop:
    Loop

    Close #fileNo

    Set mProducts = tempProducts
    Set mFinishings = tempFinishings
    Set mOperators = tempOperators
    mOutputFolder = tempOutputFolder
    mExportFolder = tempExportFolder
    mTemplateFolder = tempTemplateFolder
    mCanvasExtra = tempCanvasExtra
    mLabelFont = tempLabelFont
    mLabelFontSize = tempLabelFontSize
    mLabelBold = tempLabelBold

    LoadSettings = True
    Exit Function

ErrHandler:
    On Error Resume Next

    If fileNo <> 0 Then
        Close #fileNo
    End If

    LoadSettings = False
End Function

Public Function ExportSettings(ByVal filePath As String) As Boolean
    InitializeSettings
    ExportSettings = SaveSettings(filePath)
End Function

Public Function ImportSettings(ByVal filePath As String) As Boolean
    InitializeSettings
    ImportSettings = LoadSettings(filePath)
End Function

Public Function PickFolder() As String
    Dim shellApp As Object
    Dim folder As Object

    On Error GoTo ErrHandler

    Set shellApp = CreateObject("Shell.Application")

    Set folder = shellApp.BrowseForFolder( _
        0, _
        "Pilih Folder", _
        0 _
    )

    If Not folder Is Nothing Then
        PickFolder = folder.Self.Path
    End If

    Exit Function

ErrHandler:
    PickFolder = ""
End Function

Public Function PickImportFile() As String
    Dim ofn As OPENFILENAME
    Dim buffer As String
    Dim result As Long
    Dim nullPos As Long

    buffer = String$(260, vbNullChar)

    ofn.lStructSize = LenB(ofn)
    ofn.lpstrFilter = _
        "MgoCorel Settings (*.mgo)" & Chr$(0) & "*.mgo" & Chr$(0) & _
        "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    ofn.nFilterIndex = 1
    ofn.lpstrFile = buffer
    ofn.nMaxFile = Len(buffer)
    ofn.lpstrTitle = "Import MgoCorel Settings"
    ofn.Flags = &H80000 Or &H800 Or &H1000

    result = GetOpenFileName(ofn)

    If result <> 0 Then
        nullPos = InStr(ofn.lpstrFile, vbNullChar)

        If nullPos > 0 Then
            PickImportFile = Left$(ofn.lpstrFile, nullPos - 1)
        Else
            PickImportFile = ofn.lpstrFile
        End If
    End If
End Function

Public Function PickExportFile() As String
    Dim ofn As OPENFILENAME
    Dim buffer As String
    Dim result As Long
    Dim fileName As String
    Dim nullPos As Long

    fileName = "MgoCorel_Settings.mgo"
    buffer = fileName & String$(260 - Len(fileName), vbNullChar)

    ofn.lStructSize = LenB(ofn)
    ofn.lpstrFilter = _
        "MgoCorel Settings (*.mgo)" & Chr$(0) & "*.mgo" & Chr$(0) & _
        "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    ofn.nFilterIndex = 1
    ofn.lpstrFile = buffer
    ofn.nMaxFile = Len(buffer)
    ofn.lpstrDefExt = "mgo"
    ofn.lpstrTitle = "Export MgoCorel Settings"
    ofn.Flags = &H80000 Or &H2 Or &H4

    result = GetSaveFileName(ofn)

    If result <> 0 Then
        nullPos = InStr(ofn.lpstrFile, vbNullChar)

        If nullPos > 0 Then
            PickExportFile = Left$(ofn.lpstrFile, nullPos - 1)
        Else
            PickExportFile = ofn.lpstrFile
        End If

        If LCase$(Right$(PickExportFile, 4)) <> ".mgo" Then
            PickExportFile = PickExportFile & ".mgo"
        End If
    End If
End Function

Private Sub CreateDefaultSettings()
    Set mProducts = New Collection
    Set mFinishings = New Collection
    Set mOperators = New Collection

    mProducts.Add Array("Jersey", 0)

    mFinishings.Add Array("Tanfis", 0)
    mFinishings.Add Array("Potpass", 0)
    mFinishings.Add Array("Simming", 0)
    mFinishings.Add Array("STD", 0)
    mFinishings.Add Array("Sentand", 0)

    mOperators.Add "Operator 1"

    mOutputFolder = ""
    mExportFolder = ""
    mTemplateFolder = ""
    mCanvasExtra = 5
    mLabelFont = "Arial"
    mLabelFontSize = 16
    mLabelBold = True
End Sub

Private Function CleanValue(ByVal valueText As String) As String
    valueText = Replace(valueText, vbCr, "")
    valueText = Replace(valueText, vbLf, "")
    valueText = Replace(valueText, "|", "/")
    CleanValue = valueText
End Function

Private Function NumberToString(ByVal value As Double) As String
    NumberToString = Replace$(Trim$(Str$(value)), ",", ".")
End Function

Private Function StringToNumber(ByVal valueText As String) As Double
    valueText = Trim$(valueText)
    valueText = Replace$(valueText, ",", ".")
    StringToNumber = Val(valueText)
End Function

