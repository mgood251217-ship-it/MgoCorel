Attribute VB_Name = "modToolbar"
Option Explicit

Private Const TOOLBAR_NAME As String = "MgoCorel Tools"

Private Function GetMgoCorelToolbar() As CommandBar
    Dim i As Long
    Dim cb As CommandBar

    On Error Resume Next

    For i = 1 To Application.CommandBars.Count
        Set cb = Nothing
        Set cb = Application.CommandBars.Item(i)

        If Not cb Is Nothing Then
            If StrComp(cb.Name, TOOLBAR_NAME, vbTextCompare) = 0 Then
                Set GetMgoCorelToolbar = cb
                Exit Function
            End If
        End If
    Next i

    Set GetMgoCorelToolbar = Nothing

    On Error GoTo 0
End Function

Public Sub Plugin_CreateToolbar()

    Dim cb As CommandBar
    Dim btn As Control
    Dim iconPath As String

    On Error GoTo ErrHandler
    Set cb = GetMgoCorelToolbar()

    If Not cb Is Nothing Then
        cb.Visible = True
        Set cb = Nothing
        Exit Sub
    End If

    Set cb = Application.CommandBars.Add( _
        TOOLBAR_NAME, _
        cuiBarFloating, _
        True _
    )

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modLabel.Plugin_ShowLabelForm" _
    )

    btn.Caption = "Buat Label Spanduk"
    btn.ToolTipText = "Buat Label Spanduk"

    iconPath = Application.GMSManager.UserGMSPath & "icons\label.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modExport.Plugin_ExportLabel" _
    )

    btn.Caption = "Export Spanduk"
    btn.ToolTipText = "Export Spanduk"

    iconPath = Application.GMSManager.UserGMSPath & "icons\export.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modHitung.Plugin_ShowHitungForm" _
    )

    btn.Caption = "Hitung Harga"
    btn.ToolTipText = "Hitung Harga"

    iconPath = Application.GMSManager.UserGMSPath & "icons\calculator.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modSusun.Plugin_ShowSusunForm" _
    )

    btn.Caption = "Susun Objek"
    btn.ToolTipText = "Susun Objek"

    iconPath = Application.GMSManager.UserGMSPath & "icons\susun.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modImposisi.Plugin_ShowImposisiForm" _
    )

    btn.Caption = "Imposisi Otomatis"
    btn.ToolTipText = "Imposisi Otomatis"

    iconPath = Application.GMSManager.UserGMSPath & "icons\imposisi.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modDuplicate.Plugin_ShowDuplicateForm" _
    )

    btn.Caption = "Duplicate Quantity"
    btn.ToolTipText = "Duplicate Quantity"

    iconPath = Application.GMSManager.UserGMSPath & "icons\duplicate.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modNumbering.Plugin_ShowNumberingForm" _
    )

    btn.Caption = "Numbering Otomatis"
    btn.ToolTipText = "Numbering Otomatis"

    iconPath = Application.GMSManager.UserGMSPath & "icons\numbering.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modJobBuilder.Plugin_ShowJobBuilderForm" _
    )

    btn.Caption = "Job Builder"
    btn.ToolTipText = "Job Builder"

    iconPath = Application.GMSManager.UserGMSPath & "icons\excel.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If


    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modRectangleNesting.Plugin_ShowRectangleNestingForm" _
    )

    btn.Caption = "Rectangle Nesting"
    btn.ToolTipText = "Rectangle Nesting"

    iconPath = Application.GMSManager.UserGMSPath & "icons\rectangleNesting.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If


    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modNesting.Plugin_ShowNestingForm" _
    )

    btn.Caption = "Nesting"
    btn.ToolTipText = "Nesting"

    iconPath = Application.GMSManager.UserGMSPath & "icons\nesting.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modSettings.Plugin_ShowSettingsForm" _
    )

    btn.Caption = "Pengaturan"
    btn.ToolTipText = "Pengaturan"

    iconPath = Application.GMSManager.UserGMSPath & "icons\settings.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modMain.Plugin_ShowMainDialog" _
    )

    btn.Caption = "Detail Aplikasi"
    btn.ToolTipText = "Detail Aplikasi"

    iconPath = Application.GMSManager.UserGMSPath & "icons\main.ico"

    If Dir$(iconPath) <> "" Then
        btn.SetIcon2 iconPath
    End If

    cb.Visible = True

    Set btn = Nothing
    Set cb = Nothing

    Exit Sub


ErrHandler:

    MsgBox _
        "Gagal membuat toolbar MgoCorel." & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "MgoCorel"

End Sub


Public Sub Plugin_RemoveToolbar()

    Dim i As Long
    Dim cb As CommandBar

    On Error Resume Next

    For i = Application.CommandBars.Count To 1 Step -1

        Set cb = Nothing
        Set cb = Application.CommandBars.Item(i)

        If Not cb Is Nothing Then

            If StrComp(cb.Name, TOOLBAR_NAME, vbTextCompare) = 0 Then
                cb.Delete
            End If

        End If

    Next i

    Set cb = Nothing

    On Error GoTo 0

End Sub

Public Sub Plugin_ShowToolbar()

    Dim cb As CommandBar

    On Error GoTo ErrHandler

    Set cb = GetMgoCorelToolbar()

    If cb Is Nothing Then
        Plugin_CreateToolbar
        Exit Sub
    End If
    cb.Visible = True

    Set cb = Nothing

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

    Set cb = GetMgoCorelToolbar()

    If Not cb Is Nothing Then
        cb.Visible = False
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