Attribute VB_Name = "modLabel"
Option Explicit

Private mLabelStateLoaded As Boolean
Private mLastSystem As String
Private mLastNoInv As String
Private mLastNama As String
Private mLastProduk As String
Private mLastQuantity As String
Private mLastFinishing As String
Private mLastDeadline As String
Private mLastOperator As String

Public Sub Plugin_ShowLabelForm()
    frmLabel.Show
End Sub

Public Sub LoadLabelState(ByVal form As Object)
    If Not mLabelStateLoaded Then
        mLastSystem = ""
        mLastNoInv = ""
        mLastNama = ""
        mLastProduk = ""
        mLastQuantity = "1"
        mLastFinishing = ""
        mLastDeadline = GetDefaultDeadline()
        mLastOperator = ""
        mLabelStateLoaded = True
    End If

    If mLastSystem <> "" Then
        form.cmbSystem.Value = mLastSystem
    End If

    If form.cmbSystem.ListIndex = -1 Then
        If form.cmbSystem.ListCount > 0 Then
            form.cmbSystem.ListIndex = 0
        End If
    End If

    form.txtNoInv.Value = mLastNoInv
    form.txtNama.Value = mLastNama
    form.txtQuantity.Value = mLastQuantity
    form.txtDeadline.Value = mLastDeadline

    If form.cmbProduk.ListCount > 0 Then
        If mLastProduk <> "" Then
            form.cmbProduk.Value = mLastProduk
        End If

        If form.cmbProduk.ListIndex = -1 Then
            form.cmbProduk.ListIndex = 0
        End If
    End If

    If form.cmbFinishing.ListCount > 0 Then
        If mLastFinishing <> "" Then
            form.cmbFinishing.Value = mLastFinishing
        End If

        If form.cmbFinishing.ListIndex = -1 Then
            form.cmbFinishing.ListIndex = 0
        End If
    End If

    If form.cmbOperator.ListCount > 0 Then
        If mLastOperator <> "" Then
            form.cmbOperator.Value = mLastOperator
        End If

        If form.cmbOperator.ListIndex = -1 Then
            form.cmbOperator.ListIndex = 0
        End If
    End If
End Sub

Public Sub SaveLabelState(ByVal form As Object)
    mLastSystem = Trim$(form.cmbSystem.Value)
    mLastNoInv = Trim$(form.txtNoInv.Value)
    mLastNama = Trim$(form.txtNama.Value)
    mLastProduk = Trim$(form.cmbProduk.Value)
    mLastQuantity = Trim$(form.txtQuantity.Value)
    mLastFinishing = Trim$(form.cmbFinishing.Value)
    mLastDeadline = Trim$(form.txtDeadline.Value)
    mLastOperator = Trim$(form.cmbOperator.Value)

    If mLastSystem = "" Then
        mLastSystem = "Offline"
    End If

    If mLastQuantity = "" Then
        mLastQuantity = "1"
    End If

    mLabelStateLoaded = True
End Sub

Public Sub ClearLabelState(ByVal form As Object)
    mLastSystem = ""
    mLastNoInv = ""
    mLastNama = ""
    mLastProduk = ""
    mLastQuantity = "1"
    mLastFinishing = ""
    mLastDeadline = GetDefaultDeadline()
    mLastOperator = ""

    If form.cmbSystem.ListCount > 0 Then
        form.cmbSystem.ListIndex = 0
    End If

    form.txtNoInv.Value = ""
    form.txtNama.Value = ""
    form.txtQuantity.Value = "1"
    form.txtDeadline.Value = GetDefaultDeadline()

    If form.cmbProduk.ListCount > 0 Then
        form.cmbProduk.ListIndex = 0
    End If

    If form.cmbFinishing.ListCount > 0 Then
        form.cmbFinishing.ListIndex = 0
    End If

    If form.cmbOperator.ListCount > 0 Then
        form.cmbOperator.ListIndex = 0
    End If

    mLabelStateLoaded = True
End Sub

Public Function GetShapeWidthM(ByVal shp As Shape) As Double
    GetShapeWidthM = Application.ConvertUnits( _
        shp.SizeWidth, _
        ActiveDocument.Unit, _
        cdrMeter _
    )
End Function

Public Function GetShapeHeightM(ByVal shp As Shape) As Double
    GetShapeHeightM = Application.ConvertUnits( _
        shp.SizeHeight, _
        ActiveDocument.Unit, _
        cdrMeter _
    )
End Function

Public Function GetDefaultDeadline() As String
    Dim deadlineTime As Date

    deadlineTime = DateAdd("h", 1, Now)

    If Minute(deadlineTime) > 0 Then
        deadlineTime = DateAdd("h", 1, deadlineTime)
    End If

    deadlineTime = DateValue(deadlineTime) + _
                   TimeSerial(Hour(deadlineTime), 0, 0)

    GetDefaultDeadline = Format$(deadlineTime, "hh.mm")
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
    canvas.CenterX = sr.CenterX
    canvas.CenterY = sr.CenterY

    canvas.Fill.ApplyNoFill
    canvas.OrderToBack

    Set CreateLabelCanvas = canvas
End Function

Private Sub ApplyLabelTextColor( _
    ByVal label As Shape, _
    ByVal systemName As String, _
    ByVal nama As String, _
    ByVal invoiceNumber As String)

    Dim labelText As String
    Dim labelLength As Long
    Dim namaLength As Long
    Dim invoiceLength As Long
    Dim invoiceStart As Long

    labelText = label.Text.Story.Text
    labelLength = Len(labelText)
    namaLength = Len(Trim$(nama))
    invoiceLength = Len(Trim$(invoiceNumber))

    If labelLength <= 0 Then Exit Sub

    If StrComp(Trim$(systemName), "Online", vbTextCompare) = 0 Then
        label.Text.Range(0, labelLength).Fill.UniformColor.RGBAssign 255, 0, 0
        Exit Sub
    End If

    If namaLength > 0 Then
        label.Text.Range(0, namaLength).Fill.UniformColor.RGBAssign 255, 0, 0
    End If

    If invoiceLength > 0 Then
        invoiceStart = labelLength - invoiceLength

        label.Text.Range( _
            invoiceStart, _
            labelLength _
        ).Fill.UniformColor.RGBAssign 255, 0, 0
    End If
End Sub

Private Function BuildLabelText( _
    ByVal systemName As String, _
    ByVal nama As String, _
    ByVal productName As String, _
    ByVal widthM As Double, _
    ByVal heightM As Double, _
    ByVal finishing As String, _
    ByVal quantity As String, _
    ByVal deadline As String, _
    ByVal operatorName As String, _
    ByVal invoiceNumber As String) As String

    Dim labelDate As String
    Dim prefix As String

    labelDate = Format$(Date, "dd.mm.yyyy")

    If StrComp(Trim$(systemName), "Online", vbTextCompare) = 0 Then
        prefix = "ONLINE_"
    Else
        prefix = ""
    End If

    BuildLabelText = UCase$( _
        prefix & _
        Trim$(nama) & "_" & _
        Trim$(productName) & "_" & _
        Format$(widthM, "0.##") & "X" & _
        Format$(heightM, "0.##") & "_" & _
        Trim$(finishing) & "_" & _
        Trim$(quantity) & "PCS_" & _
        Trim$(deadline) & "_" & _
        labelDate & "_" & _
        Trim$(operatorName) & "_" & _
        Trim$(invoiceNumber) _
    )
End Function

Private Function CreateOneLabelText( _
    ByVal labelText As String, _
    ByVal systemName As String, _
    ByVal nama As String, _
    ByVal invoiceNumber As String, _
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

    ApplyLabelTextColor _
        label, _
        systemName, _
        nama, _
        invoiceNumber

    label.CenterX = centerX
    label.CenterY = centerY

    If rotation <> 0 Then
        label.Rotate rotation
        label.CenterX = centerX
        label.CenterY = centerY
    End If

    Set CreateOneLabelText = label
End Function

Public Function CreateLabelTexts( _
    ByVal sr As ShapeRange, _
    ByVal canvas As Shape, _
    ByVal systemName As String, _
    ByVal nama As String, _
    ByVal productName As String, _
    ByVal widthM As Double, _
    ByVal heightM As Double, _
    ByVal finishing As String, _
    ByVal quantity As String, _
    ByVal deadline As String, _
    ByVal operatorName As String, _
    ByVal invoiceNumber As String) As Shape

    Dim extraDoc As Double
    Dim stripDoc As Double
    Dim labelCenter As Double
    Dim labelText As String
    Dim labelTop As Shape
    Dim labelBottom As Shape
    Dim labelLeft As Shape
    Dim labelRight As Shape
    Dim groupRange As ShapeRange
    Dim groupShape As Shape

    extraDoc = Application.ConvertUnits( _
        GetCanvasExtra(), _
        cdrCentimeter, _
        ActiveDocument.Unit _
    )

    stripDoc = extraDoc / 2#
    labelCenter = stripDoc / 2#

    labelText = BuildLabelText( _
        systemName, _
        nama, _
        productName, _
        widthM, _
        heightM, _
        finishing, _
        quantity, _
        deadline, _
        operatorName, _
        invoiceNumber _
    )

    If sr.SizeHeight >= sr.SizeWidth Then
        Set labelTop = CreateOneLabelText( _
            labelText, _
            systemName, _
            nama, _
            invoiceNumber, _
            sr.CenterX, _
            sr.TopY + labelCenter, _
            0 _
        )

        Set labelBottom = CreateOneLabelText( _
            labelText, _
            systemName, _
            nama, _
            invoiceNumber, _
            sr.CenterX, _
            sr.BottomY - labelCenter, _
            0 _
        )
    Else
        Set labelLeft = CreateOneLabelText( _
            labelText, _
            systemName, _
            nama, _
            invoiceNumber, _
            sr.LeftX - labelCenter, _
            sr.CenterY, _
            90 _
        )

        Set labelRight = CreateOneLabelText( _
            labelText, _
            systemName, _
            nama, _
            invoiceNumber, _
            sr.RightX + labelCenter, _
            sr.CenterY, _
            270 _
        )
    End If

    Set groupRange = CreateShapeRange
    groupRange.AddRange sr
    groupRange.Add canvas

    If Not labelTop Is Nothing Then
        groupRange.Add labelTop
    End If

    If Not labelBottom Is Nothing Then
        groupRange.Add labelBottom
    End If

    If Not labelLeft Is Nothing Then
        groupRange.Add labelLeft
    End If

    If Not labelRight Is Nothing Then
        groupRange.Add labelRight
    End If

    If StrComp(Trim$(finishing), "SESTAND", vbTextCompare) = 0 Or _
       StrComp(Trim$(finishing), "MATIK", vbTextCompare) = 0 Then

        AddFinishingDots groupRange, sr, canvas
    End If

    Set groupShape = groupRange.Group

    groupShape.Name = labelText

    Set CreateLabelTexts = groupShape
End Function

Private Sub AddFinishingDots( _
    ByVal groupRange As ShapeRange, _
    ByVal sr As ShapeRange, _
    ByVal canvas As Shape)

    Dim dotDiameterDoc As Double
    Dim dotRadiusDoc As Double
    Dim insetDoc As Double
    Dim widthM As Double
    Dim heightM As Double
    Dim widthDoc As Double
    Dim heightDoc As Double
    Dim horizontalSegments As Long
    Dim verticalSegments As Long
    Dim i As Long
    Dim x As Double
    Dim y As Double
    Dim stepX As Double
    Dim stepY As Double
    Dim dot As Shape

    dotDiameterDoc = Application.ConvertUnits( _
        1#, _
        cdrCentimeter, _
        ActiveDocument.Unit _
    )

    dotRadiusDoc = dotDiameterDoc / 2#

    insetDoc = Application.ConvertUnits( _
        1.25, _
        cdrCentimeter, _
        ActiveDocument.Unit _
    )

    widthDoc = sr.SizeWidth
    heightDoc = sr.SizeHeight

    widthM = Application.ConvertUnits( _
        widthDoc, _
        ActiveDocument.Unit, _
        cdrMeter _
    )

    heightM = Application.ConvertUnits( _
        heightDoc, _
        ActiveDocument.Unit, _
        cdrMeter _
    )

    horizontalSegments = Int(widthM)

    If widthM > horizontalSegments Then
        horizontalSegments = horizontalSegments + 1
    End If

    verticalSegments = Int(heightM)

    If heightM > verticalSegments Then
        verticalSegments = verticalSegments + 1
    End If

    If horizontalSegments < 1 Then
        horizontalSegments = 1
    End If

    If verticalSegments < 1 Then
        verticalSegments = 1
    End If

    stepX = widthDoc / horizontalSegments
    stepY = heightDoc / verticalSegments

    Set dot = ActiveLayer.CreateEllipse2( _
        canvas.LeftX + insetDoc, _
        canvas.TopY - insetDoc, _
        dotRadiusDoc, _
        dotRadiusDoc _
    )

    dot.Fill.UniformColor.RGBAssign 255, 0, 0
    dot.Outline.SetNoOutline
    groupRange.Add dot

    Set dot = ActiveLayer.CreateEllipse2( _
        canvas.RightX - insetDoc, _
        canvas.TopY - insetDoc, _
        dotRadiusDoc, _
        dotRadiusDoc _
    )

    dot.Fill.UniformColor.RGBAssign 255, 0, 0
    dot.Outline.SetNoOutline
    groupRange.Add dot

    Set dot = ActiveLayer.CreateEllipse2( _
        canvas.LeftX + insetDoc, _
        canvas.BottomY + insetDoc, _
        dotRadiusDoc, _
        dotRadiusDoc _
    )

    dot.Fill.UniformColor.RGBAssign 255, 0, 0
    dot.Outline.SetNoOutline
    groupRange.Add dot

    Set dot = ActiveLayer.CreateEllipse2( _
        canvas.RightX - insetDoc, _
        canvas.BottomY + insetDoc, _
        dotRadiusDoc, _
        dotRadiusDoc _
    )

    dot.Fill.UniformColor.RGBAssign 255, 0, 0
    dot.Outline.SetNoOutline
    groupRange.Add dot

    If horizontalSegments > 1 Then
        For i = 1 To horizontalSegments - 1
            x = sr.LeftX + (stepX * i)

            Set dot = ActiveLayer.CreateEllipse2( _
                x, _
                canvas.TopY - insetDoc, _
                dotRadiusDoc, _
                dotRadiusDoc _
            )

            dot.Fill.UniformColor.RGBAssign 255, 0, 0
            dot.Outline.SetNoOutline
            groupRange.Add dot

            Set dot = ActiveLayer.CreateEllipse2( _
                x, _
                canvas.BottomY + insetDoc, _
                dotRadiusDoc, _
                dotRadiusDoc _
            )

            dot.Fill.UniformColor.RGBAssign 255, 0, 0
            dot.Outline.SetNoOutline
            groupRange.Add dot
        Next i
    End If

    If verticalSegments > 1 Then
        For i = 1 To verticalSegments - 1
            y = sr.TopY - (stepY * i)

            Set dot = ActiveLayer.CreateEllipse2( _
                canvas.LeftX + insetDoc, _
                y, _
                dotRadiusDoc, _
                dotRadiusDoc _
            )

            dot.Fill.UniformColor.RGBAssign 255, 0, 0
            dot.Outline.SetNoOutline
            groupRange.Add dot

            Set dot = ActiveLayer.CreateEllipse2( _
                canvas.RightX - insetDoc, _
                y, _
                dotRadiusDoc, _
                dotRadiusDoc _
            )

            dot.Fill.UniformColor.RGBAssign 255, 0, 0
            dot.Outline.SetNoOutline
            groupRange.Add dot
        Next i
    End If
End Sub