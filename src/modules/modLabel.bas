Attribute VB_Name = "modLabel"
Option Explicit

Private mLabelStateLoaded As Boolean
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
        mLastNoInv = ""
        mLastNama = ""
        mLastProduk = ""
        mLastQuantity = "1"
        mLastFinishing = ""
        mLastDeadline = Format$(Date, "dd/mm/yyyy")
        mLastOperator = ""
        mLabelStateLoaded = True
    End If

    form.txtNoInv.value = mLastNoInv
    form.txtNama.value = mLastNama
    form.txtQuantity.value = mLastQuantity
    form.txtDeadline.value = mLastDeadline

    If form.cmbProduk.ListCount > 0 Then
        If mLastProduk <> "" Then
            form.cmbProduk.value = mLastProduk
            If form.cmbProduk.ListIndex = -1 Then
                form.cmbProduk.ListIndex = 0
            End If
        Else
            form.cmbProduk.ListIndex = 0
        End If
    End If

    If form.cmbFinishing.ListCount > 0 Then
        If mLastFinishing <> "" Then
            form.cmbFinishing.value = mLastFinishing
            If form.cmbFinishing.ListIndex = -1 Then
                form.cmbFinishing.ListIndex = 0
            End If
        Else
            form.cmbFinishing.ListIndex = 0
        End If
    End If

    If form.cmbOperator.ListCount > 0 Then
        If mLastOperator <> "" Then
            form.cmbOperator.value = mLastOperator
            If form.cmbOperator.ListIndex = -1 Then
                form.cmbOperator.ListIndex = 0
            End If
        Else
            form.cmbOperator.ListIndex = 0
        End If
    End If
End Sub

Public Sub SaveLabelState(ByVal form As Object)
    mLastNoInv = Trim$(form.txtNoInv.value)
    mLastNama = Trim$(form.txtNama.value)
    mLastProduk = Trim$(form.cmbProduk.value)
    mLastQuantity = Trim$(form.txtQuantity.value)
    mLastFinishing = Trim$(form.cmbFinishing.value)
    mLastDeadline = Trim$(form.txtDeadline.value)
    mLastOperator = Trim$(form.cmbOperator.value)

    If mLastQuantity = "" Then
        mLastQuantity = "1"
    End If

    mLabelStateLoaded = True
End Sub

Public Sub ClearLabelState(ByVal form As Object)
    mLastNoInv = ""
    mLastNama = ""
    mLastProduk = ""
    mLastQuantity = "1"
    mLastFinishing = ""
    mLastDeadline = Format$(Date, "dd/mm/yyyy")
    mLastOperator = ""

    form.txtNoInv.value = ""
    form.txtNama.value = ""
    form.txtQuantity.value = "1"
    form.txtDeadline.value = Format$(Date, "dd/mm/yyyy")

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

    label.CenterX = centerX
    label.CenterY = centerY

    If rotation <> 0 Then
        label.Rotate rotation
        label.CenterX = centerX
        label.CenterY = centerY
    End If

    Set CreateOneLabelText = label
End Function

Public Sub CreateLabelTexts( _
    ByVal sr As ShapeRange, _
    ByVal nama As String, _
    ByVal productName As String, _
    ByVal widthM As Double, _
    ByVal heightM As Double, _
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

    stripDoc = extraDoc / 2#
    labelCenter = stripDoc / 2#

    labelText = BuildLabelText( _
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
            sr.CenterX, _
            sr.TopY + labelCenter, _
            0 _
        )

        Set labelBottom = CreateOneLabelText( _
            labelText, _
            sr.CenterX, _
            sr.BottomY - labelCenter, _
            0 _
        )
    Else
        Set labelLeft = CreateOneLabelText( _
            labelText, _
            sr.LeftX - labelCenter, _
            sr.CenterY, _
            90 _
        )

        Set labelRight = CreateOneLabelText( _
            labelText, _
            sr.RightX + labelCenter, _
            sr.CenterY, _
            270 _
        )
    End If
End Sub