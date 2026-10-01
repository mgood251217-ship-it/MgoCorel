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
Private ToolbarInitializing As Boolean

Private Sub GlobalMacroStorage_Start()
    On Error Resume Next

    ToolbarStartupPending = True
    ToolbarInitializing = False

    On Error GoTo 0
End Sub

Private Sub GlobalMacroStorage_WindowActivate(ByVal Doc As Document, ByVal Window As Window)
    On Error Resume Next

    If Not ToolbarStartupPending Then Exit Sub
    If ToolbarInitializing Then Exit Sub

    ToolbarInitializing = True
    ToolbarStartupPending = False

    modToolbar.Plugin_CreateToolbar
    modToolbar.Plugin_RemoveToolbar
    modToolbar.Plugin_CreateToolbar

    ToolbarInitializing = False

    On Error GoTo 0
End Sub
