Attribute VB_Name = "modToolbar"
Option Explicit

Private Const TOOLBAR_NAME As String = "MgoCorel Tools"

Public Sub Plugin_CreateToolbar()
    Dim cb As CommandBar
    Dim btn As Control

    On Error GoTo ErrHandler

    Plugin_RemoveToolbar

    Set cb = Application.CommandBars.Add( _
        TOOLBAR_NAME, _
        cuiBarFloating, _
        False _
    )

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modMain.Plugin_ShowMainDialog" _
    )
    btn.Caption = "Buka Form"
    btn.ToolTipText = "Buka Main Dialog"
    btn.DescriptionText = "Buka Main Dialog MgoCorel"

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modLabel.Plugin_ShowLabelForm" _
    )
    btn.Caption = "Buat Label"
    btn.ToolTipText = "Buat Label"
    btn.DescriptionText = "Membuka form untuk membuat label"

    Set btn = cb.Controls.AddCustomButton( _
        cdrCmdCategoryMacros, _
        "MgoCorel.modSettings.Plugin_ShowSettingsForm" _
    )
    btn.Caption = "Pengaturan"
    btn.ToolTipText = "Pengaturan"
    btn.DescriptionText = "Membuka pengaturan MgoCorel"

    cb.Visible = True

    Set btn = Nothing
    Set cb = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Gagal membuat toolbar MgoCorel." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Public Sub Plugin_RemoveToolbar()
    Dim cb As CommandBar

    On Error Resume Next

    Set cb = Application.CommandBars(TOOLBAR_NAME)

    If Not cb Is Nothing Then
        cb.Delete
    End If

    Set cb = Nothing

    On Error GoTo 0
End Sub

Public Sub Plugin_ShowToolbar()
    Dim cb As CommandBar

    On Error GoTo ErrHandler

    Set cb = Application.CommandBars(TOOLBAR_NAME)

    If cb Is Nothing Then
        Plugin_CreateToolbar
        Exit Sub
    End If

    cb.Visible = True

    Set cb = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Gagal menampilkan toolbar MgoCorel." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

Public Sub Plugin_HideToolbar()
    Dim cb As CommandBar

    On Error GoTo ErrHandler

    Set cb = Application.CommandBars(TOOLBAR_NAME)

    If Not cb Is Nothing Then
        cb.Visible = False
    End If

    Set cb = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Gagal menyembunyikan toolbar MgoCorel." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "MgoCorel"
End Sub

