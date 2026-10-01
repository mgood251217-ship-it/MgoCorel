Attribute VB_Name = "modToolbar"
Option Explicit

Private Const STANDARD_TOOLBAR As String = "Standard"
Private Const LEGACY_TOOLBAR As String = "MgoCorel Tools"

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
                    Case "modlabel.plugin_showlabelform", _
                         "modexport.plugin_exportlabel", _
                         "modhitung.plugin_showhitungform", _
                         "modsusun.plugin_showsusunform", _
                         "modimposisi.plugin_showimposisiform", _
                         "modduplicate.plugin_showduplicateform", _
                         "modnumbering.plugin_shownumberingform", _
                         "modjobbuilder.plugin_showjobbuilderform", _
                         "modrectanglenesting.plugin_showrectanglenestingform", _
                         "modnesting.plugin_shownestingform", _
                         "modsettings.plugin_showsettingsform", _
                         "modmain.plugin_showmaindialog", _
                        "buat label spanduk", _
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

Public Sub Plugin_RemoveToolbar()
    Dim cb As CommandBar
    Dim i As Long

    On Error Resume Next

    Set cb = GetStandardToolbar()

    If Not cb Is Nothing Then
        RemoveMgoCorelButtons cb
    End If

    Set cb = Nothing

    For i = Application.CommandBars.Count To 1 Step -1
        Set cb = Nothing
        Set cb = Application.CommandBars.Item(i)

        If Not cb Is Nothing Then
            If StrComp(cb.Name, LEGACY_TOOLBAR, vbTextCompare) = 0 Then
                cb.Delete
            End If
        End If
    Next i

    On Error GoTo 0
End Sub