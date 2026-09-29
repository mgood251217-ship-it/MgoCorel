Attribute VB_Name = "modNesting"
Option Explicit

Public NestingCancelRequested As Boolean

Private Const EPSILON As Double = 0.000001

Public Sub Plugin_ShowNestingForm()
    frmNesting.Show vbModal
End Sub

Public Function ArrangeNesting( _
    ByVal boxWidthM As Double, _
    ByVal boxHeightM As Double, _
    ByVal gapMm As Double, _
    ByVal rotationMode As Long) As Boolean

    Dim sr As ShapeRange
    Dim items() As Shape
    Dim itemWidth() As Double
    Dim itemHeight() As Double
    Dim itemArea() As Double
    Dim itemOrder() As Long
    Dim originalAngle() As Double

    Dim bestLeft() As Double
    Dim bestBottom() As Double
    Dim bestAngle() As Double

    Dim testLeft() As Double
    Dim testBottom() As Double
    Dim testAngle() As Double

    Dim boxWidthDoc As Double
    Dim boxHeightDoc As Double
    Dim gapDoc As Double

    Dim boxLeft As Double
    Dim boxBottom As Double

    Dim itemCount As Long
    Dim attemptCount As Long
    Dim attempt As Long
    Dim i As Long

    Dim testUsedWidth As Double
    Dim testUsedHeight As Double
    Dim bestUsedWidth As Double
    Dim bestUsedHeight As Double

    Dim totalArea As Double
    Dim boxArea As Double

    Dim hasBest As Boolean
    Dim success As Boolean

    Dim shp As Shape
    Dim boxShape As Shape
    Dim resultRange As ShapeRange

    Dim commandStarted As Boolean

    On Error GoTo ErrHandler

    ArrangeNesting = False
    NestingCancelRequested = False

    Set sr = ActiveSelectionRange
    itemCount = sr.Count

    If itemCount = 0 Then
        MsgBox "Pilih minimal satu objek terlebih dahulu.", vbExclamation, "MgoCorel"
        Exit Function
    End If

    If boxWidthM <= 0 Or boxHeightM <= 0 Then
        MsgBox "Ukuran nesting harus lebih dari 0.", vbExclamation, "MgoCorel"
        Exit Function
    End If

    If gapMm < 0 Then
        MsgBox "Gap tidak boleh kurang dari 0.", vbExclamation, "MgoCorel"
        Exit Function
    End If

    If rotationMode < 0 Or rotationMode > 3 Then
        rotationMode = 0
    End If

    boxWidthDoc = Application.ConvertUnits( _
        boxWidthM, _
        cdrMeter, _
        ActiveDocument.Unit)

    boxHeightDoc = Application.ConvertUnits( _
        boxHeightM, _
        cdrMeter, _
        ActiveDocument.Unit)

    gapDoc = Application.ConvertUnits( _
        gapMm, _
        cdrMillimeter, _
        ActiveDocument.Unit)

    ReDim items(1 To itemCount)
    ReDim itemWidth(1 To itemCount)
    ReDim itemHeight(1 To itemCount)
    ReDim itemArea(1 To itemCount)
    ReDim itemOrder(1 To itemCount)
    ReDim originalAngle(1 To itemCount)

    ReDim bestLeft(1 To itemCount)
    ReDim bestBottom(1 To itemCount)
    ReDim bestAngle(1 To itemCount)

    ReDim testLeft(1 To itemCount)
    ReDim testBottom(1 To itemCount)
    ReDim testAngle(1 To itemCount)

    totalArea = 0

    For i = 1 To itemCount

        Set items(i) = sr(i)

        itemWidth(i) = items(i).SizeWidth
        itemHeight(i) = items(i).SizeHeight
        itemArea(i) = itemWidth(i) * itemHeight(i)
        originalAngle(i) = items(i).RotationAngle

        totalArea = totalArea + itemArea(i)

    Next i

    boxArea = boxWidthDoc * boxHeightDoc

    If totalArea > boxArea + EPSILON Then

        MsgBox "Luas total objek lebih besar dari area nesting." & vbCrLf & vbCrLf & _
               "Tidak ada objek yang dipindahkan.", _
               vbExclamation, "MgoCorel"

        Exit Function
    End If

    Select Case rotationMode
        Case 0
            attemptCount = 2
        Case 1
            attemptCount = 3
        Case 2
            attemptCount = 3
        Case 3
            attemptCount = 3
    End Select

    StartNestingLoading itemCount, attemptCount

    boxLeft = sr.RightX + gapDoc + _
              Application.ConvertUnits( _
                  20#, _
                  cdrMillimeter, _
                  ActiveDocument.Unit)

    boxBottom = sr.CenterY - boxHeightDoc / 2#

    bestUsedWidth = 1E+30
    bestUsedHeight = 1E+30

    For attempt = 1 To attemptCount

        If NestingCancelRequested Then
            GoTo Cancelled
        End If

        UpdateNestingLoading _
            "Mencari susunan " & attempt & "/" & attemptCount, _
            attempt - 1, _
            attemptCount

        BuildItemOrder _
            itemOrder, _
            itemWidth, _
            itemHeight, _
            itemArea, _
            itemCount, _
            attempt

        success = BuildRowNestingPlan( _
            items, _
            itemOrder, _
            itemCount, _
            originalAngle, _
            boxLeft, _
            boxBottom, _
            boxWidthDoc, _
            boxHeightDoc, _
            gapDoc, _
            rotationMode, _
            testLeft, _
            testBottom, _
            testAngle, _
            testUsedWidth, _
            testUsedHeight)

        If NestingCancelRequested Then
            GoTo Cancelled
        End If

        If success Then

            If Not hasBest Then

                hasBest = True
                bestUsedWidth = testUsedWidth
                bestUsedHeight = testUsedHeight

                CopyDoubleArray testLeft, bestLeft, itemCount
                CopyDoubleArray testBottom, bestBottom, itemCount
                CopyDoubleArray testAngle, bestAngle, itemCount

            ElseIf IsBetterPlan( _
                testUsedWidth, _
                testUsedHeight, _
                bestUsedWidth, _
                bestUsedHeight) Then

                bestUsedWidth = testUsedWidth
                bestUsedHeight = testUsedHeight

                CopyDoubleArray testLeft, bestLeft, itemCount
                CopyDoubleArray testBottom, bestBottom, itemCount
                CopyDoubleArray testAngle, bestAngle, itemCount

            End If
        End If

    Next attempt

    If NestingCancelRequested Then
        GoTo Cancelled
    End If

    If Not hasBest Then

        StopNestingLoading

        MsgBox "Tidak ditemukan susunan yang dapat memuat seluruh objek." & vbCrLf & vbCrLf & _
               "Total objek : " & itemCount & vbCrLf & _
               "Area nesting : " & _
               Format$(boxWidthM, "0.###") & " x " & _
               Format$(boxHeightM, "0.###") & " m" & vbCrLf & vbCrLf & _
               "Tidak ada objek yang dipindahkan.", _
               vbExclamation, "MgoCorel"

        Exit Function
    End If

    UpdateNestingLoading _
        "Menerapkan hasil nesting...", _
        100, _
        100

    On Error Resume Next
    frmNestingLoading.cmdBatal.Enabled = False
    On Error GoTo ErrHandler

    ActiveDocument.BeginCommandGroup "MgoCorel - Nesting"
    commandStarted = True

    Set boxShape = ActiveLayer.CreateRectangle2(0, 0, 1, 1)

    boxShape.SizeWidth = boxWidthDoc
    boxShape.SizeHeight = boxHeightDoc
    boxShape.LeftX = boxLeft
    boxShape.BottomY = boxBottom
    boxShape.Fill.ApplyNoFill
    boxShape.Name = "NESTING AREA"

    Set resultRange = CreateShapeRange

    For i = 1 To itemCount

        Set shp = items(i)

        shp.RotationAngle = _
            originalAngle(i) + bestAngle(i)

        shp.LeftX = bestLeft(i)
        shp.BottomY = bestBottom(i)

        resultRange.Add shp

        UpdateNestingLoading _
            "Memindahkan objek " & i & "/" & itemCount, _
            i, _
            itemCount

    Next i

    resultRange.Add boxShape

    ActiveDocument.EndCommandGroup
    commandStarted = False

    StopNestingLoading

    resultRange.CreateSelection

    ArrangeNesting = True

    MsgBox "Nesting selesai." & vbCrLf & vbCrLf & _
           "Total objek : " & itemCount & vbCrLf & _
           "Area : " & _
           Format$(boxWidthM, "0.###") & " x " & _
           Format$(boxHeightM, "0.###") & " m" & vbCrLf & _
           "Pemakaian lebar : " & _
           Format$(Application.ConvertUnits( _
               bestUsedWidth, _
               ActiveDocument.Unit, _
               cdrMeter), "0.###") & " m" & vbCrLf & _
           "Pemakaian tinggi : " & _
           Format$(Application.ConvertUnits( _
               bestUsedHeight, _
               ActiveDocument.Unit, _
               cdrMeter), "0.###") & " m" & vbCrLf & _
           "Rotasi : " & GetRotationModeText(rotationMode), _
           vbInformation, "MgoCorel"

    Exit Function

Cancelled:

    If commandStarted Then
        On Error Resume Next
        ActiveDocument.EndCommandGroup
        On Error GoTo 0
    End If

    StopNestingLoading

    MsgBox "Nesting dibatalkan." & vbCrLf & vbCrLf & _
           "Tidak ada objek yang dipindahkan.", _
           vbInformation, "MgoCorel"

    Exit Function

ErrHandler:

    If commandStarted Then
        On Error Resume Next
        ActiveDocument.EndCommandGroup
        On Error GoTo 0
    End If

    StopNestingLoading

    MsgBox "Gagal melakukan Nesting." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "MgoCorel"
End Function

Private Function BuildRowNestingPlan( _
    ByRef items() As Shape, _
    ByRef itemOrder() As Long, _
    ByVal itemCount As Long, _
    ByRef originalAngle() As Double, _
    ByVal boxLeft As Double, _
    ByVal boxBottom As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double, _
    ByVal rotationMode As Long, _
    ByRef placeLeft() As Double, _
    ByRef placeBottom() As Double, _
    ByRef placeAngle() As Double, _
    ByRef usedWidth As Double, _
    ByRef usedHeight As Double) As Boolean

    Dim tempLayer As Layer
    Dim placed() As Shape
    Dim used() As Boolean

    Dim placedCount As Long
    Dim remainingCount As Long

    Dim bestItemIndex As Long
    Dim bestCandidate As Shape

    Dim bestLeft As Double
    Dim bestBottom As Double
    Dim bestAngle As Double

    Dim minX As Double
    Dim minY As Double
    Dim maxX As Double
    Dim maxY As Double

    Dim success As Boolean
    Dim i As Long

    On Error GoTo ErrHandler

    Set tempLayer = ActivePage.CreateLayer("__MGO_NESTING_TEMP")
    tempLayer.Visible = False
    tempLayer.Editable = True

    ReDim placed(1 To itemCount)
    ReDim used(1 To itemCount)

    placedCount = 0
    remainingCount = itemCount

    minX = 1E+30
    minY = 1E+30
    maxX = -1E+30
    maxY = -1E+30

    Do While remainingCount > 0

        If NestingCancelRequested Then
            DeleteTemporaryShapes placed, placedCount
            tempLayer.Delete
            Exit Function
        End If

        Set bestCandidate = Nothing
        bestItemIndex = 0
        bestLeft = 0#
        bestBottom = 0#
        bestAngle = 0#

        If placedCount = 0 Then

            success = FindBestRowSeed( _
                items, _
                used, _
                itemOrder, _
                itemCount, _
                originalAngle, _
                boxLeft, _
                boxBottom, _
                boxWidth, _
                boxHeight, _
                gapDoc, _
                rotationMode, _
                tempLayer, _
                bestCandidate, _
                bestItemIndex, _
                bestLeft, _
                bestBottom, _
                bestAngle)

        Else

            success = FindBestRowItem( _
                items, _
                used, _
                itemOrder, _
                itemCount, _
                originalAngle, _
                placed, _
                placedCount, _
                boxLeft, _
                boxBottom, _
                boxWidth, _
                boxHeight, _
                gapDoc, _
                rotationMode, _
                tempLayer, _
                bestCandidate, _
                bestItemIndex, _
                bestLeft, _
                bestBottom, _
                bestAngle)

        End If

        If Not success Then
            DeleteTemporaryShapes placed, placedCount
            tempLayer.Delete
            Exit Function
        End If

        Set placed(placedCount + 1) = bestCandidate
        placedCount = placedCount + 1

        used(bestItemIndex) = True
        remainingCount = remainingCount - 1

        placeLeft(bestItemIndex) = bestLeft
        placeBottom(bestItemIndex) = bestBottom
        placeAngle(bestItemIndex) = bestAngle

        If bestCandidate.LeftX < minX Then
            minX = bestCandidate.LeftX
        End If

        If bestCandidate.BottomY < minY Then
            minY = bestCandidate.BottomY
        End If

        If bestCandidate.RightX > maxX Then
            maxX = bestCandidate.RightX
        End If

        If bestCandidate.TopY > maxY Then
            maxY = bestCandidate.TopY
        End If

        UpdateNestingLoading _
            "Menyusun objek " & placedCount & "/" & itemCount, _
            placedCount, _
            itemCount

        DoEvents

    Loop

    usedWidth = maxX - minX
    usedHeight = maxY - minY

    DeleteTemporaryShapes placed, placedCount
    tempLayer.Delete

    BuildRowNestingPlan = True

    Exit Function

ErrHandler:

    DeleteTemporaryShapes placed, placedCount

    On Error Resume Next
    tempLayer.Delete
    On Error GoTo 0
End Function

Private Function FindBestRowSeed( _
    ByRef items() As Shape, _
    ByRef used() As Boolean, _
    ByRef itemOrder() As Long, _
    ByVal itemCount As Long, _
    ByRef originalAngle() As Double, _
    ByVal boxLeft As Double, _
    ByVal boxBottom As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double, _
    ByVal rotationMode As Long, _
    ByVal tempLayer As Layer, _
    ByRef bestCandidate As Shape, _
    ByRef bestItemIndex As Long, _
    ByRef bestLeft As Double, _
    ByRef bestBottom As Double, _
    ByRef bestAngle As Double) As Boolean

    Dim angles() As Double
    Dim angleCount As Long

    Dim i As Long
    Dim a As Long
    Dim index As Long

    Dim candidate As Shape

    Dim candidateWidth As Double
    Dim candidateHeight As Double
    Dim candidateArea As Double

    Dim bestArea As Double
    Dim bestWidth As Double
    Dim bestHeight As Double

    Dim emptyPlaced() As Shape

    GetRotationAngles rotationMode, angles, angleCount

    bestItemIndex = 0
    bestArea = -1E+30
    bestWidth = -1E+30
    bestHeight = -1E+30

    For i = 1 To itemCount

        index = itemOrder(i)

        If Not used(index) Then

            For a = 1 To angleCount

                If NestingCancelRequested Then
                    Exit Function
                End If

                Set candidate = items(index).Duplicate(0, 0)
                candidate.MoveToLayer tempLayer
                candidate.RotationAngle = _
                    originalAngle(index) + angles(a)

                candidateWidth = candidate.SizeWidth
                candidateHeight = candidate.SizeHeight
                candidateArea = candidateWidth * candidateHeight

                candidate.LeftX = boxLeft + gapDoc
                candidate.BottomY = _
                    boxBottom + boxHeight - gapDoc - candidateHeight

                If IsPlacementValid( _
                    candidate, _
                    emptyPlaced, _
                    0, _
                    boxLeft, _
                    boxBottom, _
                    boxWidth, _
                    boxHeight, _
                    gapDoc) Then

                    If candidateArea > bestArea + EPSILON Then

                        If Not bestCandidate Is Nothing Then
                            On Error Resume Next
                            bestCandidate.Delete
                            On Error GoTo 0
                        End If

                        bestArea = candidateArea
                        bestWidth = candidateWidth
                        bestHeight = candidateHeight

                        bestItemIndex = index
                        bestLeft = candidate.LeftX
                        bestBottom = candidate.BottomY
                        bestAngle = angles(a)

                        Set bestCandidate = candidate
                        Set candidate = Nothing

                        FindBestRowSeed = True

                    ElseIf Abs(candidateArea - bestArea) <= EPSILON Then

                        If candidateWidth > bestWidth + EPSILON Then

                            If Not bestCandidate Is Nothing Then
                                On Error Resume Next
                                bestCandidate.Delete
                                On Error GoTo 0
                            End If

                            bestWidth = candidateWidth
                            bestHeight = candidateHeight

                            bestItemIndex = index
                            bestLeft = candidate.LeftX
                            bestBottom = candidate.BottomY
                            bestAngle = angles(a)

                            Set bestCandidate = candidate
                            Set candidate = Nothing

                            FindBestRowSeed = True

                        ElseIf Abs(candidateWidth - bestWidth) <= EPSILON Then

                            If candidateHeight > bestHeight + EPSILON Then

                                If Not bestCandidate Is Nothing Then
                                    On Error Resume Next
                                    bestCandidate.Delete
                                    On Error GoTo 0
                                End If

                                bestHeight = candidateHeight

                                bestItemIndex = index
                                bestLeft = candidate.LeftX
                                bestBottom = candidate.BottomY
                                bestAngle = angles(a)

                                Set bestCandidate = candidate
                                Set candidate = Nothing

                                FindBestRowSeed = True

                            End If
                        End If
                    End If
                End If

                If Not candidate Is Nothing Then
                    On Error Resume Next
                    candidate.Delete
                    On Error GoTo 0
                End If

            Next a
        End If
    Next i

    If Not FindBestRowSeed Then
        Set bestCandidate = Nothing
        bestItemIndex = 0
    End If
End Function

Private Function FindBestRowItem( _
    ByRef items() As Shape, _
    ByRef used() As Boolean, _
    ByRef itemOrder() As Long, _
    ByVal itemCount As Long, _
    ByRef originalAngle() As Double, _
    ByRef placed() As Shape, _
    ByVal placedCount As Long, _
    ByVal boxLeft As Double, _
    ByVal boxBottom As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double, _
    ByVal rotationMode As Long, _
    ByVal tempLayer As Layer, _
    ByRef bestCandidate As Shape, _
    ByRef bestItemIndex As Long, _
    ByRef bestLeft As Double, _
    ByRef bestBottom As Double, _
    ByRef bestAngle As Double) As Boolean

    Dim angles() As Double
    Dim angleCount As Long

    Dim i As Long
    Dim a As Long
    Dim index As Long

    Dim candidate As Shape

    Dim candidateLeft As Double
    Dim candidateBottom As Double
    Dim candidateTop As Double
    Dim candidateRight As Double
    Dim candidateArea As Double

    Dim bestTop As Double
    Dim bestX As Double
    Dim bestRight As Double
    Dim bestArea As Double

    Dim foundPosition As Boolean

    GetRotationAngles rotationMode, angles, angleCount

    bestItemIndex = 0
    bestTop = -1E+30
    bestX = 1E+30
    bestRight = -1E+30
    bestArea = -1E+30

    For i = 1 To itemCount

        index = itemOrder(i)

        If Not used(index) Then

            For a = 1 To angleCount

                If NestingCancelRequested Then
                    Exit Function
                End If

                Set candidate = items(index).Duplicate(0, 0)
                candidate.MoveToLayer tempLayer

                candidate.RotationAngle = _
                    originalAngle(index) + angles(a)

                foundPosition = FindBestNestingPosition( _
                    candidate, _
                    placed, _
                    placedCount, _
                    boxLeft, _
                    boxBottom, _
                    boxWidth, _
                    boxHeight, _
                    gapDoc, _
                    candidateLeft, _
                    candidateBottom, _
                    candidateTop, _
                    candidateRight)

                If foundPosition Then

                    candidateArea = _
                        candidate.SizeWidth * candidate.SizeHeight

                    If candidateTop > bestTop + EPSILON Then

                        If Not bestCandidate Is Nothing Then
                            On Error Resume Next
                            bestCandidate.Delete
                            On Error GoTo 0
                        End If

                        bestTop = candidateTop
                        bestX = candidateLeft
                        bestRight = candidateRight
                        bestArea = candidateArea

                        bestItemIndex = index
                        bestLeft = candidateLeft
                        bestBottom = candidateBottom
                        bestAngle = angles(a)

                        candidate.LeftX = candidateLeft
                        candidate.BottomY = candidateBottom

                        Set bestCandidate = candidate
                        Set candidate = Nothing

                        FindBestRowItem = True

                    ElseIf Abs(candidateTop - bestTop) <= EPSILON Then

                        If candidateLeft < bestX - EPSILON Then

                            If Not bestCandidate Is Nothing Then
                                On Error Resume Next
                                bestCandidate.Delete
                                On Error GoTo 0
                            End If

                            bestX = candidateLeft
                            bestRight = candidateRight
                            bestArea = candidateArea

                            bestItemIndex = index
                            bestLeft = candidateLeft
                            bestBottom = candidateBottom
                            bestAngle = angles(a)

                            candidate.LeftX = candidateLeft
                            candidate.BottomY = candidateBottom

                            Set bestCandidate = candidate
                            Set candidate = Nothing

                            FindBestRowItem = True

                        ElseIf Abs(candidateLeft - bestX) <= EPSILON Then

                            If candidateRight > bestRight + EPSILON Then

                                If Not bestCandidate Is Nothing Then
                                    On Error Resume Next
                                    bestCandidate.Delete
                                    On Error GoTo 0
                                End If

                                bestRight = candidateRight
                                bestArea = candidateArea

                                bestItemIndex = index
                                bestLeft = candidateLeft
                                bestBottom = candidateBottom
                                bestAngle = angles(a)

                                candidate.LeftX = candidateLeft
                                candidate.BottomY = candidateBottom

                                Set bestCandidate = candidate
                                Set candidate = Nothing

                                FindBestRowItem = True

                            ElseIf Abs(candidateRight - bestRight) <= EPSILON Then

                                If candidateArea > bestArea + EPSILON Then

                                    If Not bestCandidate Is Nothing Then
                                        On Error Resume Next
                                        bestCandidate.Delete
                                        On Error GoTo 0
                                    End If

                                    bestArea = candidateArea

                                    bestItemIndex = index
                                    bestLeft = candidateLeft
                                    bestBottom = candidateBottom
                                    bestAngle = angles(a)

                                    candidate.LeftX = candidateLeft
                                    candidate.BottomY = candidateBottom

                                    Set bestCandidate = candidate
                                    Set candidate = Nothing

                                    FindBestRowItem = True

                                End If
                            End If
                        End If
                    End If
                End If

                If Not candidate Is Nothing Then
                    On Error Resume Next
                    candidate.Delete
                    On Error GoTo 0
                End If

            Next a
        End If
    Next i

    If Not FindBestRowItem Then
        Set bestCandidate = Nothing
        bestItemIndex = 0
    End If
End Function

Private Function FindBestNestingPosition( _
    ByVal candidate As Shape, _
    ByRef placed() As Shape, _
    ByVal placedCount As Long, _
    ByVal boxLeft As Double, _
    ByVal boxBottom As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double, _
    ByRef bestLeft As Double, _
    ByRef bestBottom As Double, _
    ByRef bestTop As Double, _
    ByRef bestRight As Double) As Boolean

    Dim pointX() As Double
    Dim pointY() As Double

    Dim pointCountX As Long
    Dim pointCountY As Long

    Dim candidateWidth As Double
    Dim candidateHeight As Double

    Dim i As Long
    Dim j As Long

    Dim x As Double
    Dim y As Double

    Dim topValue As Double
    Dim rightValue As Double

    Dim found As Boolean

    candidateWidth = candidate.SizeWidth
    candidateHeight = candidate.SizeHeight

    pointCountX = BuildNestingCandidateX( _
        placed, _
        placedCount, _
        boxLeft, _
        gapDoc, _
        candidateWidth, _
        pointX)

    pointCountY = BuildNestingCandidateY( _
        placed, _
        placedCount, _
        boxBottom, _
        boxHeight, _
        gapDoc, _
        candidateHeight, _
        pointY)

    bestTop = -1E+30
    bestRight = -1E+30
    bestLeft = 1E+30
    bestBottom = 0#

    For i = 1 To pointCountY

        y = pointY(i)

        For j = 1 To pointCountX

            If NestingCancelRequested Then
                Exit Function
            End If

            x = pointX(j)

            candidate.LeftX = x
            candidate.BottomY = y

            If IsPlacementValid( _
                candidate, _
                placed, _
                placedCount, _
                boxLeft, _
                boxBottom, _
                boxWidth, _
                boxHeight, _
                gapDoc) Then

                topValue = candidate.TopY
                rightValue = candidate.RightX

                If Not found Then

                    found = True
                    bestTop = topValue
                    bestLeft = candidate.LeftX
                    bestBottom = candidate.BottomY
                    bestRight = rightValue

                ElseIf topValue > bestTop + EPSILON Then

                    bestTop = topValue
                    bestLeft = candidate.LeftX
                    bestBottom = candidate.BottomY
                    bestRight = rightValue

                ElseIf Abs(topValue - bestTop) <= EPSILON Then

                    If candidate.LeftX < bestLeft - EPSILON Then

                        bestLeft = candidate.LeftX
                        bestBottom = candidate.BottomY
                        bestRight = rightValue

                    ElseIf Abs(candidate.LeftX - bestLeft) <= EPSILON Then

                        If rightValue > bestRight + EPSILON Then
                            bestBottom = candidate.BottomY
                            bestRight = rightValue
                        End If
                    End If
                End If
            End If
        Next j
    Next i

    FindBestNestingPosition = found
End Function

Private Function BuildNestingCandidateX( _
    ByRef placed() As Shape, _
    ByVal placedCount As Long, _
    ByVal boxLeft As Double, _
    ByVal gapDoc As Double, _
    ByVal candidateWidth As Double, _
    ByRef pointX() As Double) As Long

    Dim capacity As Long
    Dim count As Long
    Dim i As Long

    capacity = placedCount * 4 + 4

    If capacity < 4 Then
        capacity = 4
    End If

    ReDim pointX(1 To capacity)

    AddCandidatePoint _
        pointX, _
        count, _
        boxLeft + gapDoc

    For i = 1 To placedCount

        AddCandidatePoint _
            pointX, _
            count, _
            placed(i).RightX + gapDoc

        AddCandidatePoint _
            pointX, _
            count, _
            placed(i).LeftX - candidateWidth - gapDoc

        AddCandidatePoint _
            pointX, _
            count, _
            placed(i).RightX - candidateWidth

        AddCandidatePoint _
            pointX, _
            count, _
            placed(i).LeftX

    Next i

    BuildNestingCandidateX = count
End Function

Private Function BuildNestingCandidateY( _
    ByRef placed() As Shape, _
    ByVal placedCount As Long, _
    ByVal boxBottom As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double, _
    ByVal candidateHeight As Double, _
    ByRef pointY() As Double) As Long

    Dim capacity As Long
    Dim count As Long
    Dim i As Long

    capacity = placedCount * 5 + 5

    If capacity < 5 Then
        capacity = 5
    End If

    ReDim pointY(1 To capacity)

    AddCandidatePoint _
        pointY, _
        count, _
        boxBottom + boxHeight - gapDoc - candidateHeight

    AddCandidatePoint _
        pointY, _
        count, _
        boxBottom + gapDoc

    For i = 1 To placedCount

        AddCandidatePoint _
            pointY, _
            count, _
            placed(i).TopY - candidateHeight - gapDoc

        AddCandidatePoint _
            pointY, _
            count, _
            placed(i).TopY - candidateHeight

        AddCandidatePoint _
            pointY, _
            count, _
            placed(i).BottomY - candidateHeight - gapDoc

        AddCandidatePoint _
            pointY, _
            count, _
            placed(i).BottomY - candidateHeight

        AddCandidatePoint _
            pointY, _
            count, _
            placed(i).TopY + gapDoc

    Next i

    BuildNestingCandidateY = count
End Function

Private Sub AddCandidatePoint( _
    ByRef points() As Double, _
    ByRef count As Long, _
    ByVal value As Double)

    Dim i As Long
    Dim capacity As Long

    For i = 1 To count

        If Abs(points(i) - value) <= EPSILON Then
            Exit Sub
        End If

    Next i

    capacity = UBound(points)

    If count >= capacity Then
        ReDim Preserve points(1 To capacity * 2)
    End If

    count = count + 1
    points(count) = value
End Sub

Private Function IsPlacementValid( _
    ByVal candidate As Shape, _
    ByRef placed() As Shape, _
    ByVal placedCount As Long, _
    ByVal boxLeft As Double, _
    ByVal boxBottom As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double) As Boolean

    Dim i As Long

    If candidate.LeftX < _
       boxLeft + gapDoc - EPSILON Then
        Exit Function
    End If

    If candidate.BottomY < _
       boxBottom + gapDoc - EPSILON Then
        Exit Function
    End If

    If candidate.RightX > _
       boxLeft + boxWidth - gapDoc + EPSILON Then
        Exit Function
    End If

    If candidate.TopY > _
       boxBottom + boxHeight - gapDoc + EPSILON Then
        Exit Function
    End If

    For i = 1 To placedCount

        If NestingCancelRequested Then
            Exit Function
        End If

        If Not BoundingSeparated( _
            candidate, _
            placed(i), _
            gapDoc) Then

            If ShapeCollision( _
                candidate, _
                placed(i)) Then

                Exit Function
            End If
        End If
    Next i

    IsPlacementValid = True
End Function

Private Function BoundingSeparated( _
    ByVal shapeA As Shape, _
    ByVal shapeB As Shape, _
    ByVal margin As Double) As Boolean

    If shapeA.RightX + margin <= _
       shapeB.LeftX + EPSILON Then

        BoundingSeparated = True
        Exit Function
    End If

    If shapeA.LeftX - margin >= _
       shapeB.RightX - EPSILON Then

        BoundingSeparated = True
        Exit Function
    End If

    If shapeA.TopY + margin <= _
       shapeB.BottomY + EPSILON Then

        BoundingSeparated = True
        Exit Function
    End If

    If shapeA.BottomY - margin >= _
       shapeB.TopY - EPSILON Then

        BoundingSeparated = True
        Exit Function
    End If
End Function

Private Function ShapeCollision( _
    ByVal shapeA As Shape, _
    ByVal shapeB As Shape) As Boolean

    On Error GoTo CollisionError

    If shapeA.DisplayCurve.IntersectsWith( _
        shapeB.DisplayCurve) Then

        ShapeCollision = True
        Exit Function
    End If

    If PointInsideShape( _
        shapeA.CenterX, _
        shapeA.CenterY, _
        shapeB) Then

        ShapeCollision = True
        Exit Function
    End If

    If PointInsideShape( _
        shapeB.CenterX, _
        shapeB.CenterY, _
        shapeA) Then

        ShapeCollision = True
        Exit Function
    End If

    ShapeCollision = False

    Exit Function

CollisionError:
    ShapeCollision = True
End Function

Private Function PointInsideShape( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal shp As Shape) As Boolean

    Dim position As cdrPositionOfPointOverShape

    On Error GoTo PointError

    position = shp.IsOnShape(x, y, 0#)

    If position = cdrInsideShape Or _
       position = cdrOnMarginOfShape Then

        PointInsideShape = True
    End If

    Exit Function

PointError:
    PointInsideShape = False
End Function

Private Sub GetRotationAngles( _
    ByVal rotationMode As Long, _
    ByRef angles() As Double, _
    ByRef angleCount As Long)

    Dim i As Long

    Select Case rotationMode

        Case 0

            angleCount = 1
            ReDim angles(1 To 1)
            angles(1) = 0#

        Case 1

            angleCount = 1
            ReDim angles(1 To 1)
            angles(1) = 90#

        Case 2

            angleCount = 1
            ReDim angles(1 To 1)
            angles(1) = 180#

        Case 3

            angleCount = 24
            ReDim angles(1 To 24)

            For i = 1 To 24
                angles(i) = (i - 1) * 15#
            Next i

        Case Else

            angleCount = 1
            ReDim angles(1 To 1)
            angles(1) = 0#

    End Select
End Sub

Private Function IsBetterPlan( _
    ByVal usedWidth As Double, _
    ByVal usedHeight As Double, _
    ByVal bestWidth As Double, _
    ByVal bestHeight As Double) As Boolean

    Dim areaValue As Double
    Dim bestArea As Double

    areaValue = usedWidth * usedHeight
    bestArea = bestWidth * bestHeight

    If areaValue < bestArea - EPSILON Then
        IsBetterPlan = True
        Exit Function
    End If

    If Abs(areaValue - bestArea) <= EPSILON Then

        If usedHeight < bestHeight - EPSILON Then
            IsBetterPlan = True
            Exit Function
        End If

        If Abs(usedHeight - bestHeight) <= EPSILON Then

            If usedWidth < bestWidth - EPSILON Then
                IsBetterPlan = True
            End If

        End If
    End If
End Function

Private Sub BuildItemOrder( _
    ByRef itemOrder() As Long, _
    ByRef itemWidth() As Double, _
    ByRef itemHeight() As Double, _
    ByRef itemArea() As Double, _
    ByVal itemCount As Long, _
    ByVal sortMode As Long)

    Dim i As Long
    Dim j As Long
    Dim temp As Long

    For i = 1 To itemCount
        itemOrder(i) = i
    Next i

    For i = 1 To itemCount - 1

        For j = i + 1 To itemCount

            If ShouldComeBefore( _
                itemOrder(j), _
                itemOrder(i), _
                itemWidth, _
                itemHeight, _
                itemArea, _
                sortMode) Then

                temp = itemOrder(i)
                itemOrder(i) = itemOrder(j)
                itemOrder(j) = temp

            End If
        Next j
    Next i
End Sub

Private Function ShouldComeBefore( _
    ByVal indexA As Long, _
    ByVal indexB As Long, _
    ByRef itemWidth() As Double, _
    ByRef itemHeight() As Double, _
    ByRef itemArea() As Double, _
    ByVal sortMode As Long) As Boolean

    Dim valueA As Double
    Dim valueB As Double

    Select Case sortMode

        Case 1

            valueA = itemArea(indexA)
            valueB = itemArea(indexB)

        Case 2

            valueA = MaxValue( _
                itemWidth(indexA), _
                itemHeight(indexA))

            valueB = MaxValue( _
                itemWidth(indexB), _
                itemHeight(indexB))

        Case 3

            valueA = itemHeight(indexA)
            valueB = itemHeight(indexB)

        Case Else

            valueA = itemWidth(indexA) + itemHeight(indexA)
            valueB = itemWidth(indexB) + itemHeight(indexB)

    End Select

    If valueA > valueB + EPSILON Then

        ShouldComeBefore = True

    ElseIf Abs(valueA - valueB) <= EPSILON Then

        ShouldComeBefore = indexA < indexB

    End If
End Function

Private Function MaxValue( _
    ByVal valueA As Double, _
    ByVal valueB As Double) As Double

    If valueA > valueB Then
        MaxValue = valueA
    Else
        MaxValue = valueB
    End If
End Function

Private Sub CopyDoubleArray( _
    ByRef source() As Double, _
    ByRef target() As Double, _
    ByVal itemCount As Long)

    Dim i As Long

    For i = 1 To itemCount
        target(i) = source(i)
    Next i
End Sub

Private Sub DeleteTemporaryShapes( _
    ByRef placed() As Shape, _
    ByVal placedCount As Long)

    Dim i As Long

    On Error Resume Next

    For i = 1 To placedCount

        If Not placed(i) Is Nothing Then
            placed(i).Delete
        End If

    Next i

    On Error GoTo 0
End Sub

Private Function GetRotationModeText( _
    ByVal rotationMode As Long) As String

    Select Case rotationMode

        Case 0
            GetRotationModeText = "Tanpa Rotasi"

        Case 1
            GetRotationModeText = "90°"

        Case 2
            GetRotationModeText = "180°"

        Case 3
            GetRotationModeText = "Bebas 15°"

        Case Else
            GetRotationModeText = "Tanpa Rotasi"

    End Select
End Function

Private Sub StartNestingLoading( _
    ByVal itemCount As Long, _
    ByVal attemptCount As Long)

    On Error Resume Next

    NestingCancelRequested = False

    Load frmNestingLoading

    frmNestingLoading.lblStatus.Caption = _
        "Menyiapkan nesting..."

    frmNestingLoading.lblPercent.Caption = "0%"

    frmNestingLoading.cmdBatal.Caption = "Batalkan"
    frmNestingLoading.cmdBatal.Enabled = True

    frmNestingLoading.Show vbModeless
    frmNestingLoading.Repaint

    DoEvents

    On Error GoTo 0
End Sub

Private Sub UpdateNestingLoading( _
    ByVal statusText As String, _
    ByVal currentValue As Long, _
    ByVal totalValue As Long)

    Dim percentValue As Long

    On Error Resume Next

    If totalValue <= 0 Then
        percentValue = 0
    Else
        percentValue = CLng( _
            (currentValue / totalValue) * 100#)
    End If

    If percentValue < 0 Then
        percentValue = 0
    End If

    If percentValue > 100 Then
        percentValue = 100
    End If

    frmNestingLoading.lblStatus.Caption = statusText
    frmNestingLoading.lblPercent.Caption = _
        percentValue & "%"

    frmNestingLoading.Repaint

    DoEvents

    On Error GoTo 0
End Sub

Private Sub StopNestingLoading()

    On Error Resume Next

    Unload frmNestingLoading

    On Error GoTo 0
End Sub