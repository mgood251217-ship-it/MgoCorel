Attribute VB_Name = "modLabel"
Option Explicit

Public Sub Plugin_ShowLabelForm()
    frmLabel.Show
End Sub

Public Function GetSelectedWidthCm() As Double
    Dim sr As ShapeRange

    If ActiveSelectionRange.Count = 0 Then Exit Function

    Set sr = ActiveSelectionRange

    GetSelectedWidthCm = Application.ConvertUnits( _
        sr.SizeWidth, _
        ActiveDocument.Unit, _
        cdrCentimeter _
    )
End Function

Public Function GetSelectedHeightCm() As Double
    Dim sr As ShapeRange

    If ActiveSelectionRange.Count = 0 Then Exit Function

    Set sr = ActiveSelectionRange

    GetSelectedHeightCm = Application.ConvertUnits( _
        sr.SizeHeight, _
        ActiveDocument.Unit, _
        cdrCentimeter _
    )
End Function

Public Function CreateLabelCanvas(ByVal sr As ShapeRange) As Shape
    Dim extraCm As Double
    Dim extraDoc As Double
    Dim canvasWidth As Double
    Dim canvasHeight As Double
    Dim canvas As Shape

    extraCm = GetCanvasExtra()

    extraDoc = Application.ConvertUnits( _
        extraCm, _
        cdrCentimeter, _
        ActiveDocument.Unit _
    )

    canvasWidth = sr.SizeWidth + extraDoc
    canvasHeight = sr.SizeHeight + extraDoc

    Set canvas = ActiveLayer.CreateRectangle2(0, 0, 1, 1)

    canvas.SizeWidth = canvasWidth
    canvas.SizeHeight = canvasHeight
    canvas.centerX = sr.centerX
    canvas.centerY = sr.centerY

    canvas.Fill.ApplyNoFill
    canvas.OrderToBack

    Set CreateLabelCanvas = canvas
End Function

Private Function BuildLabelText( _
    ByVal nama As String, _
    ByVal productName As String, _
    ByVal widthCm As Double, _
    ByVal heightCm As Double, _
    ByVal finishing As String, _
    ByVal quantity As String, _
    ByVal deadline As String, _
    ByVal operatorName As String, _
    ByVal invoiceNumber As String) As String

    BuildLabelText = UCase$( _
        Trim$(nama) & "_" & _
        Trim$(productName) & "_" & _
        Format$(widthCm, "0.##") & "X" & _
        Format$(heightCm, "0.##") & "_" & _
        Trim$(finishing) & "_" & _
        Trim$(quantity) & "_" & _
        Trim$(deadline) & "_" & _
        Trim$(operatorName) & "_" & _
        Trim$(invoiceNumber) _
    )
End Function

Private Function CreateOneLabelText( _
    ByVal labelText As String, _
    ByVal centerX As Double, _
    ByVal centerY As Double, _
    ByVal rotation As Double) As Shape

    Dim label As Shape
    Dim fontName As String
    Dim fontSize As Double
    Dim boldValue As Long

    fontName = GetLabelFont()
    fontSize = GetLabelFontSize()

    If GetLabelBold() Then
        boldValue = cdrTrue
    Else
        boldValue = cdrFalse
    End If

    Set label = ActiveLayer.CreateArtisticText( _
        0, _
        0, _
        labelText, _
        cdrLanguageNone, _
        cdrCharSetMixed, _
        fontName, _
        fontSize, _
        boldValue, _
        cdrFalse, _
        cdrMixedFontLine, _
        cdrCenterAlignment _
    )

    label.centerX = centerX
    label.centerY = centerY

    If rotation <> 0 Then
        label.Rotate rotation
        label.centerX = centerX
        label.centerY = centerY
    End If

    Set CreateOneLabelText = label
End Function

Public Sub CreateLabelTexts( _
    ByVal sr As ShapeRange, _
    ByVal nama As String, _
    ByVal productName As String, _
    ByVal widthCm As Double, _
    ByVal heightCm As Double, _
    ByVal finishing As String, _
    ByVal quantity As String, _
    ByVal deadline As String, _
    ByVal operatorName As String, _
    ByVal invoiceNumber As String)

    Dim extraDoc As Double
    Dim stripDoc As Double
    Dim labelCenter As Double
    Dim labelText As String
    Dim labelTop As Shape
    Dim labelBottom As Shape
    Dim labelLeft As Shape
    Dim labelRight As Shape

    extraDoc = Application.ConvertUnits( _
        GetCanvasExtra(), _
        cdrCentimeter, _
        ActiveDocument.Unit _
    )

    stripDoc = extraDoc / 2
    labelCenter = stripDoc / 2

    labelText = BuildLabelText( _
        nama, _
        productName, _
        widthCm, _
        heightCm, _
        finishing, _
        quantity, _
        deadline, _
        operatorName, _
        invoiceNumber _
    )

    If sr.SizeHeight >= sr.SizeWidth Then
        Set labelTop = CreateOneLabelText( _
            labelText, _
            sr.centerX, _
            sr.TopY + labelCenter, _
            0 _
        )

        Set labelBottom = CreateOneLabelText( _
            labelText, _
            sr.centerX, _
            sr.BottomY - labelCenter, _
            0 _
        )
    Else
        Set labelLeft = CreateOneLabelText( _
            labelText, _
            sr.LeftX - labelCenter, _
            sr.centerY, _
            90 _
        )

        Set labelRight = CreateOneLabelText( _
            labelText, _
            sr.RightX + labelCenter, _
            sr.centerY, _
            270 _
        )
    End If
End Sub

