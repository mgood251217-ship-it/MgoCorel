Attribute VB_Name = "modTheme"
Option Explicit

Public Function ThemeColorPrimary() As Long
    ThemeColorPrimary = RGB(43, 108, 176)
End Function

Public Function ThemeColorSuccess() As Long
    ThemeColorSuccess = RGB(22, 135, 95)
End Function

Public Function ThemeColorDanger() As Long
    ThemeColorDanger = RGB(196, 62, 62)
End Function

Public Function ThemeColorSecondary() As Long
    ThemeColorSecondary = RGB(226, 232, 240)
End Function

Public Sub ApplyButtonTheme(ByVal targetForm As Object)
    Dim control As Object
    Dim controlName As String
    Dim buttonColor As Long
    Dim textColor As Long

    On Error Resume Next

    For Each control In targetForm.Controls
        If TypeName(control) = "CommandButton" Then
            controlName = LCase$(control.Name)
            buttonColor = ThemeColorPrimary()
            textColor = RGB(255, 255, 255)

            If InStr(controlName, "batal") > 0 _
                Or InStr(controlName, "tutup") > 0 _
                Or InStr(controlName, "close") > 0 _
                Or InStr(controlName, "browse") > 0 Then
                buttonColor = ThemeColorSecondary()
                textColor = RGB(51, 65, 85)
            ElseIf InStr(controlName, "delete") > 0 _
                Or InStr(controlName, "clear") > 0 Then
                buttonColor = ThemeColorDanger()
            ElseIf InStr(controlName, "simpan") > 0 _
                Or InStr(controlName, "save") > 0 _
                Or InStr(controlName, "proses") > 0 _
                Or InStr(controlName, "hitung") > 0 Then
                buttonColor = ThemeColorSuccess()
            End If

            With control
                .Style = 1
                .BackColor = buttonColor
                .ForeColor = textColor
                .Font.Bold = True
            End With

            Err.Clear
        End If
    Next control

    On Error GoTo 0
End Sub
