Attribute VB_Name = "modDevTest"
Option Explicit

Public Sub Plugin_DevTest()
    MsgBox "MgoCorel Build OK" & vbCrLf & _
           "Project: MgoCorel" & vbCrLf & _
           "Version: 1.0.0", _
           vbInformation, _
           "MgoCorel"
End Sub