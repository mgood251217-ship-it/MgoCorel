VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private ToolbarStartupPending As Boolean

Private Sub GlobalMacroStorage_Start()
    On Error Resume Next

    ToolbarStartupPending = True

    On Error GoTo 0
End Sub

Private Sub GlobalMacroStorage_WindowActivate(ByVal Doc As Document, ByVal Window As Window)
    On Error Resume Next

    If Not ToolbarStartupPending Then Exit Sub

    ToolbarStartupPending = False

    modToolbar.Plugin_RemoveToolbar

    On Error GoTo 0
End Sub
