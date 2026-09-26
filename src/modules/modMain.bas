Attribute VB_Name = "modMain"
Option Explicit

Private Const MGOCoreL_VERSION As String = "1.0.0"

Public Sub Plugin_ShowMainDialog()
    Dim info As String
    Dim docInfo As String
    Dim pageWidth As Double
    Dim pageHeight As Double

    info = "MgoCorel" & vbCrLf & _
           String(32, "-") & vbCrLf & _
           "Version       : " & MGOCoreL_VERSION & vbCrLf & _
           "CorelDRAW     : " & Application.Version & vbCrLf & _
           "GMS Project   : MgoCorel" & vbCrLf & _
           "Environment   : " & Environ$("PROCESSOR_ARCHITECTURE")

    If Not ActiveDocument Is Nothing Then
        pageWidth = Application.ConvertUnits( _
            ActivePage.SizeWidth, _
            ActiveDocument.Unit, _
            cdrCentimeter _
        )

        pageHeight = Application.ConvertUnits( _
            ActivePage.SizeHeight, _
            ActiveDocument.Unit, _
            cdrCentimeter _
        )

        docInfo = vbCrLf & vbCrLf & _
                  "Dokumen Aktif" & vbCrLf & _
                  String(32, "-") & vbCrLf & _
                  "Unit Dokumen  : " & ActiveDocument.Unit & vbCrLf & _
                  "Ukuran Halaman: " & _
                  Format$(pageWidth, "0.##") & " x " & _
                  Format$(pageHeight, "0.##") & " cm"
    Else
        docInfo = vbCrLf & vbCrLf & _
                  "Dokumen Aktif : Tidak ada"
    End If

    MsgBox info & docInfo, _
           vbInformation, _
           "MgoCorel - Informasi"
End Sub