Attribute VB_Name = "modToolbar"
Option Explicit

Private Const STANDARD_TOOLBAR As String = "Standard"

Private Sub SetToolbarIcon(ByVal btn As Control, ByVal iconFile As String)
    Dim iconPath As String

    On Error Resume Next

    iconPath = Application.GMSManager.UserGMSPath & "icons\" & iconFile

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    On Error GoTo 0
End Sub

Private Function GetStandardToolbar() As CommandBar
    Dim i As Long
    Dim cb As CommandBar

    On Error Resume Next

    For i = 1 To Application.CommandBars.Count
        Set cb = Nothing
        Set cb = Application.CommandBars.Item(i)

        If Not cb Is Nothing Then
            If StrComp(cb.Name, STANDARD_TOOLBAR, vbTextCompare) = 0 Then
                Set GetStandardToolbar = cb
                Exit Function
            End If
        End If
    Next i

    Set GetStandardToolbar = Nothing

    On Error GoTo 0
End Function

Private Sub RemoveMgoCorelButtons(ByVal cb As CommandBar)
    Dim i As Long
    Dim btn As Control
    Dim caption As String
    Dim tagValue As String

    On Error Resume Next

    For i = cb.Controls.Count To 1 Step -1
        Set btn = Nothing
        Set btn = cb.Controls.Item(i)

        If Not btn Is Nothing Then
            tagValue = ""
            caption = ""

            tagValue = CStr(btn.Tag)
            caption = LCase$(Trim$(btn.Caption))

            If StrComp(tagValue, "MgoCorel", vbTextCompare) = 0 Then
                cb.Controls.Remove i
            Else
                Select Case caption
                    Case "buat label spanduk", _
                         "export spanduk", _
                         "hitung harga", _
                         "susun objek", _
                         "imposisi otomatis", _
                         "duplicate quantity", _
                         "numbering otomatis", _
                         "job builder", _
                         "rectangle nesting", _
                         "nesting", _
                         "pengaturan", _
                         "detail aplikasi"

                        cb.Controls.Remove i
                End Select
            End If
        End If

        Set btn = Nothing
    Next i

    On Error GoTo 0
End Sub

Private Function AddMgoButton( _
    ByVal cb As CommandBar, _
    ByVal macroName As String, _
    ByVal caption As String, _
    ByVal iconFile As String) As Control

    Dim btn As Control
    Dim iconPath As String

    On Error GoTo ErrHandler

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        macroName, _
        cb.Controls.Count + 1, _
        False _
    )

    btn.Caption = caption
    btn.ToolTipText = caption
    btn.Tag = "MgoCorel"

    iconPath = Application.GMSManager.UserGMSPath & "icons\" & iconFile

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set AddMgoButton = btn

    Set btn = Nothing

    Exit Function

ErrHandler:
    Set AddMgoButton = Nothing
End Function

Public Sub Plugin_CreateToolbar()
    Dim cb As CommandBar
    Dim btn As Control

    On Error GoTo ErrHandler

    Set cb = GetStandardToolbar()

    If cb Is Nothing Then
        MsgBox _
            "Toolbar Standard CorelDRAW tidak ditemukan.", _
            vbCritical, _
            "MgoCorel"
        Exit Sub
    End If

    RemoveMgoCorelButtons cb

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modLabel.Plugin_ShowLabelForm", _
        "Buat Label Spanduk", _
        "label.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modExport.Plugin_ExportLabel", _
        "Export Spanduk", _
        "export.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modHitung.Plugin_ShowHitungForm", _
        "Hitung Harga", _
        "calculator.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modSusun.Plugin_ShowSusunForm", _
        "Susun Objek", _
        "susun.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modImposisi.Plugin_ShowImposisiForm", _
        "Imposisi Otomatis", _
        "imposisi.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modDuplicate.Plugin_ShowDuplicateForm", _
        "Duplicate Quantity", _
        "duplicate.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modNumbering.Plugin_ShowNumberingForm", _
        "Numbering Otomatis", _
        "numbering.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modJobBuilder.Plugin_ShowJobBuilderForm", _
        "Job Builder", _
        "excel.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modRectangleNesting.Plugin_ShowRectangleNestingForm", _
        "Rectangle Nesting", _
        "rectangleNesting.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modNesting.Plugin_ShowNestingForm", _
        "Nesting", _
        "nesting.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modSettings.Plugin_ShowSettingsForm", _
        "Pengaturan", _
        "settings.ico" _
    )

    Set btn = AddMgoButton( _
        cb, _
        "MgoCorel.modMain.Plugin_ShowMainDialog", _
        "Detail Aplikasi", _
        "main.ico" _
    )

    cb.Visible = True

    Set btn = Nothing
    Set cb = Nothing

    Exit Sub

ErrHandler:
    MsgBox _
        "Gagal memasang toolbar MgoCorel." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "MgoCorel"
End Sub

Public Sub Plugin_RemoveToolbar()
    Dim cb As CommandBar

    On Error Resume Next

    Set cb = GetStandardToolbar()

    If Not cb Is Nothing Then
        RemoveMgoCorelButtons cb
    End If

    Set cb = Nothing

    On Error GoTo 0
End Sub

Public Sub Plugin_ShowToolbar()
    On Error GoTo ErrHandler

    Plugin_CreateToolbar

    Exit Sub

ErrHandler:
    MsgBox _
        "Gagal menampilkan toolbar MgoCorel." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "MgoCorel"
End Sub

Public Sub Plugin_HideToolbar()
    Dim cb As CommandBar

    On Error GoTo ErrHandler

    Set cb = GetStandardToolbar()

    If Not cb Is Nothing Then
        RemoveMgoCorelButtons cb
    End If

    Set cb = Nothing

    Exit Sub

ErrHandler:
    MsgBox _
        "Gagal menyembunyikan toolbar MgoCorel." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "MgoCorel"
End Sub