Attribute VB_Name = "modRectangleNesting"
Option Explicit

Private Const EPSILON As Double = 0.000001

Public Sub Plugin_ShowRectangleNestingForm()
    frmRectangleNesting.Show vbModal
End Sub

Public Function ArrangeRectangleNesting( _
    ByVal boxWidthM As Double, _
    ByVal boxHeightM As Double, _
    ByVal gapMm As Double, _
    ByVal autoRotate As Boolean) As Boolean

    Dim sr As ShapeRange
    Dim items() As Shape
    Dim itemWidth() As Double
    Dim itemHeight() As Double
    Dim itemArea() As Double
    Dim itemOrder() As Long

    Dim bestPlaceX() As Double
    Dim bestPlaceY() As Double
    Dim bestPlaceWidth() As Double
    Dim bestPlaceHeight() As Double
    Dim bestPlaceRotate() As Boolean

    Dim testPlaceX() As Double
    Dim testPlaceY() As Double
    Dim testPlaceWidth() As Double
    Dim testPlaceHeight() As Double
    Dim testPlaceRotate() As Boolean

    Dim boxWidthDoc As Double
    Dim boxHeightDoc As Double
    Dim gapDoc As Double

    Dim boxLeft As Double
    Dim boxBottom As Double

    Dim itemCount As Long
    Dim i As Long
    Dim attempt As Long
    Dim orderMode As Long
    Dim scoreMode As Long

    Dim success As Boolean
    Dim attemptUsedWidth As Double
    Dim attemptUsedHeight As Double
    Dim attemptFragmentation As Long

    Dim bestUsedWidth As Double
    Dim bestUsedHeight As Double
    Dim bestFragmentation As Long

    Dim hasBest As Boolean

    Dim shp As Shape
    Dim boxShape As Shape
    Dim resultRange As ShapeRange

    Dim offsetX As Double
    Dim offsetY As Double

    Dim commandStarted As Boolean

    On Error GoTo ErrHandler

    ArrangeRectangleNesting = False

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

    boxWidthDoc = Application.ConvertUnits(boxWidthM, cdrMeter, ActiveDocument.Unit)
    boxHeightDoc = Application.ConvertUnits(boxHeightM, cdrMeter, ActiveDocument.Unit)
    gapDoc = Application.ConvertUnits(gapMm, cdrMillimeter, ActiveDocument.Unit)

    ReDim items(1 To itemCount)
    ReDim itemWidth(1 To itemCount)
    ReDim itemHeight(1 To itemCount)
    ReDim itemArea(1 To itemCount)
    ReDim itemOrder(1 To itemCount)

    ReDim bestPlaceX(1 To itemCount)
    ReDim bestPlaceY(1 To itemCount)
    ReDim bestPlaceWidth(1 To itemCount)
    ReDim bestPlaceHeight(1 To itemCount)
    ReDim bestPlaceRotate(1 To itemCount)

    ReDim testPlaceX(1 To itemCount)
    ReDim testPlaceY(1 To itemCount)
    ReDim testPlaceWidth(1 To itemCount)
    ReDim testPlaceHeight(1 To itemCount)
    ReDim testPlaceRotate(1 To itemCount)

    For i = 1 To itemCount
        Set items(i) = sr(i)

        itemWidth(i) = items(i).SizeWidth
        itemHeight(i) = items(i).SizeHeight
        itemArea(i) = itemWidth(i) * itemHeight(i)

        If Not CanFitItem( _
            itemWidth(i), _
            itemHeight(i), _
            boxWidthDoc, _
            boxHeightDoc, _
            gapDoc, _
            autoRotate) Then

            MsgBox "Area nesting terlalu kecil." & vbCrLf & vbCrLf & _
                   "Objek ke-" & i & " tidak dapat masuk." & vbCrLf & _
                   "Ukuran objek: " & FormatObjectSize(itemWidth(i), itemHeight(i)) & vbCrLf & _
                   "Area nesting: " & FormatObjectSize(boxWidthDoc, boxHeightDoc) & vbCrLf & vbCrLf & _
                   "Tidak ada objek yang dipindahkan.", _
                   vbExclamation, "MgoCorel"

            Exit Function
        End If
    Next i

    bestUsedWidth = 1E+30
    bestUsedHeight = 1E+30
    bestFragmentation = 2147483647

    For attempt = 1 To 24

        orderMode = ((attempt - 1) Mod 12) + 1
        scoreMode = ((attempt - 1) Mod 3) + 1

        BuildItemOrder _
            itemOrder, _
            itemWidth, _
            itemHeight, _
            itemArea, _
            itemCount, _
            orderMode

        success = BuildMaxRectsPlan( _
            itemOrder, _
            itemWidth, _
            itemHeight, _
            itemCount, _
            boxWidthDoc, _
            boxHeightDoc, _
            gapDoc, _
            autoRotate, _
            scoreMode, _
            testPlaceX, _
            testPlaceY, _
            testPlaceWidth, _
            testPlaceHeight, _
            testPlaceRotate, _
            attemptUsedWidth, _
            attemptUsedHeight, _
            attemptFragmentation)

        If success Then

            If Not hasBest Then

                hasBest = True
                bestUsedWidth = attemptUsedWidth
                bestUsedHeight = attemptUsedHeight
                bestFragmentation = attemptFragmentation

                CopyPlacement testPlaceX, bestPlaceX, itemCount
                CopyPlacement testPlaceY, bestPlaceY, itemCount
                CopyPlacement testPlaceWidth, bestPlaceWidth, itemCount
                CopyPlacement testPlaceHeight, bestPlaceHeight, itemCount
                CopyPlacementBoolean testPlaceRotate, bestPlaceRotate, itemCount

            ElseIf IsBetterPlan( _
                attemptUsedWidth, _
                attemptUsedHeight, _
                attemptFragmentation, _
                bestUsedWidth, _
                bestUsedHeight, _
                bestFragmentation) Then

                bestUsedWidth = attemptUsedWidth
                bestUsedHeight = attemptUsedHeight
                bestFragmentation = attemptFragmentation

                CopyPlacement testPlaceX, bestPlaceX, itemCount
                CopyPlacement testPlaceY, bestPlaceY, itemCount
                CopyPlacement testPlaceWidth, bestPlaceWidth, itemCount
                CopyPlacement testPlaceHeight, bestPlaceHeight, itemCount
                CopyPlacementBoolean testPlaceRotate, bestPlaceRotate, itemCount
            End If
        End If

        If attempt Mod 3 = 0 Then DoEvents
    Next attempt

    If Not hasBest Then
        MsgBox "Area nesting tidak cukup untuk seluruh objek." & vbCrLf & vbCrLf & _
               "Total objek : " & itemCount & vbCrLf & vbCrLf & _
               "Tidak ada objek yang dipindahkan.", _
               vbExclamation, "MgoCorel"

        Exit Function
    End If

    boxLeft = sr.RightX + gapDoc + _
              Application.ConvertUnits(20#, cdrMillimeter, ActiveDocument.Unit)

    boxBottom = sr.CenterY - boxHeightDoc / 2#

    offsetX = (boxWidthDoc - bestUsedWidth) / 2#
    offsetY = (boxHeightDoc - bestUsedHeight) / 2#

    If offsetX < 0 Then offsetX = 0
    If offsetY < 0 Then offsetY = 0

    ActiveDocument.BeginCommandGroup "MgoCorel - Rectangle Nesting"
    commandStarted = True

    Set boxShape = ActiveLayer.CreateRectangle2(0, 0, 1, 1)
    boxShape.SizeWidth = boxWidthDoc
    boxShape.SizeHeight = boxHeightDoc
    boxShape.LeftX = boxLeft
    boxShape.BottomY = boxBottom
    boxShape.Fill.ApplyNoFill
    boxShape.Name = "RECTANGLE NESTING AREA"

    Set resultRange = CreateShapeRange

    For i = 1 To itemCount

        Set shp = items(i)

        If bestPlaceRotate(i) Then
            shp.Rotate 90#
        End If

        shp.CenterX = _
            boxLeft + _
            offsetX + _
            bestPlaceX(i) + _
            bestPlaceWidth(i) / 2#

        shp.CenterY = _
            boxBottom + _
            offsetY + _
            bestPlaceY(i) + _
            bestPlaceHeight(i) / 2#

        resultRange.Add shp
    Next i

    resultRange.Add boxShape

    ActiveDocument.EndCommandGroup
    commandStarted = False

    resultRange.CreateSelection

    ArrangeRectangleNesting = True

    MsgBox "Rectangle Nesting selesai." & vbCrLf & vbCrLf & _
           "Total objek : " & itemCount & vbCrLf & _
           "Kotak : " & Format$(boxWidthM, "0.###") & " x " & _
                       Format$(boxHeightM, "0.###") & " m" & vbCrLf & _
           "Pemakaian lebar : " & _
                       Format$(Application.ConvertUnits(bestUsedWidth, ActiveDocument.Unit, cdrMeter), "0.###") & " m" & vbCrLf & _
           "Pemakaian tinggi : " & _
                       Format$(Application.ConvertUnits(bestUsedHeight, ActiveDocument.Unit, cdrMeter), "0.###") & " m" & vbCrLf & _
           "Auto Rotate : " & IIf(autoRotate, "Ya", "Tidak"), _
           vbInformation, "MgoCorel"

    Exit Function

ErrHandler:
    If commandStarted Then
        On Error Resume Next
        ActiveDocument.EndCommandGroup
        On Error GoTo 0
    End If

    MsgBox "Gagal melakukan Rectangle Nesting." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "MgoCorel"
End Function

Private Function BuildMaxRectsPlan( _
    ByRef itemOrder() As Long, _
    ByRef itemWidth() As Double, _
    ByRef itemHeight() As Double, _
    ByVal itemCount As Long, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double, _
    ByVal autoRotate As Boolean, _
    ByVal scoreMode As Long, _
    ByRef placeX() As Double, _
    ByRef placeY() As Double, _
    ByRef placeWidth() As Double, _
    ByRef placeHeight() As Double, _
    ByRef placeRotate() As Boolean, _
    ByRef usedWidth As Double, _
    ByRef usedHeight As Double, _
    ByRef fragmentation As Long) As Boolean

    Dim freeLeft() As Double
    Dim freeBottom() As Double
    Dim freeWidth() As Double
    Dim freeHeight() As Double

    Dim freeCount As Long

    Dim i As Long
    Dim index As Long

    Dim bestX As Double
    Dim bestY As Double
    Dim bestObjectWidth As Double
    Dim bestObjectHeight As Double
    Dim bestPackedWidth As Double
    Dim bestPackedHeight As Double
    Dim bestRotate As Boolean

    Dim rightValue As Double
    Dim topValue As Double

    ReDim freeLeft(1 To 10)
    ReDim freeBottom(1 To 10)
    ReDim freeWidth(1 To 10)
    ReDim freeHeight(1 To 10)

    freeCount = 1

    freeLeft(1) = 0
    freeBottom(1) = 0
    freeWidth(1) = boxWidth
    freeHeight(1) = boxHeight

    For i = 1 To itemCount

        index = itemOrder(i)

        If Not FindBestMaxRectsPlacement( _
            itemWidth(index), _
            itemHeight(index), _
            autoRotate, _
            gapDoc, _
            scoreMode, _
            freeLeft, _
            freeBottom, _
            freeWidth, _
            freeHeight, _
            freeCount, _
            bestX, _
            bestY, _
            bestObjectWidth, _
            bestObjectHeight, _
            bestPackedWidth, _
            bestPackedHeight, _
            bestRotate) Then

            Exit Function
        End If

        placeX(index) = bestX
        placeY(index) = bestY
        placeWidth(index) = bestObjectWidth
        placeHeight(index) = bestObjectHeight
        placeRotate(index) = bestRotate

        SplitMaxRects _
            bestX, _
            bestY, _
            bestPackedWidth, _
            bestPackedHeight, _
            freeLeft, _
            freeBottom, _
            freeWidth, _
            freeHeight, _
            freeCount

        If freeCount <= 0 Then
            Exit Function
        End If

        rightValue = bestX + bestObjectWidth
        topValue = bestY + bestObjectHeight

        If rightValue > usedWidth Then
            usedWidth = rightValue
        End If

        If topValue > usedHeight Then
            usedHeight = topValue
        End If

        If i = 1 Then
            usedWidth = rightValue
            usedHeight = topValue
        End If

        If i Mod 10 = 0 Then DoEvents
    Next i

    fragmentation = freeCount

    BuildMaxRectsPlan = True
End Function

Private Function FindBestMaxRectsPlacement( _
    ByVal objectWidth As Double, _
    ByVal objectHeight As Double, _
    ByVal autoRotate As Boolean, _
    ByVal gapDoc As Double, _
    ByVal scoreMode As Long, _
    ByRef freeLeft() As Double, _
    ByRef freeBottom() As Double, _
    ByRef freeWidth() As Double, _
    ByRef freeHeight() As Double, _
    ByVal freeCount As Long, _
    ByRef bestX As Double, _
    ByRef bestY As Double, _
    ByRef bestObjectWidth As Double, _
    ByRef bestObjectHeight As Double, _
    ByRef bestPackedWidth As Double, _
    ByRef bestPackedHeight As Double, _
    ByRef bestRotate As Boolean) As Boolean

    Dim orientation As Long
    Dim i As Long

    Dim testWidth As Double
    Dim testHeight As Double

    Dim packedWidth As Double
    Dim packedHeight As Double

    Dim candidateX As Double
    Dim candidateY As Double

    Dim candidateTop As Double
    Dim candidateRight As Double

    Dim remainingWidth As Double
    Dim remainingHeight As Double

    Dim shortFit As Double
    Dim longFit As Double
    Dim areaFit As Double

    Dim bestShortFit As Double
    Dim bestLongFit As Double
    Dim bestAreaFit As Double
    Dim bestCandidateY As Double
    Dim bestCandidateX As Double

    bestShortFit = 1E+30
    bestLongFit = 1E+30
    bestAreaFit = 1E+30
    bestCandidateY = 1E+30
    bestCandidateX = 1E+30

    For orientation = 0 To 1

        If orientation = 0 Then
            testWidth = objectWidth
            testHeight = objectHeight
        Else
            If Not autoRotate Then Exit For

            testWidth = objectHeight
            testHeight = objectWidth
        End If

        If testWidth > 0 And testHeight > 0 Then

            For i = 1 To freeCount

                If testWidth <= freeWidth(i) + EPSILON And _
                   testHeight <= freeHeight(i) + EPSILON Then

                    packedWidth = testWidth + gapDoc
                    packedHeight = testHeight + gapDoc

                    If packedWidth > freeWidth(i) + EPSILON Then
                        packedWidth = freeWidth(i)
                    End If

                    If packedHeight > freeHeight(i) + EPSILON Then
                        packedHeight = freeHeight(i)
                    End If

                    candidateX = freeLeft(i)
                    candidateY = freeBottom(i)

                    candidateRight = candidateX + packedWidth
                    candidateTop = candidateY + packedHeight

                    If candidateRight <= MaxValue( _
                        freeLeft(i) + freeWidth(i), _
                        candidateRight) + EPSILON Then

                        If candidateTop <= MaxValue( _
                            freeBottom(i) + freeHeight(i), _
                            candidateTop) + EPSILON Then

                            remainingWidth = _
                                freeWidth(i) - packedWidth

                            remainingHeight = _
                                freeHeight(i) - packedHeight

                            If remainingWidth < 0 Then
                                remainingWidth = 0
                            End If

                            If remainingHeight < 0 Then
                                remainingHeight = 0
                            End If

                            shortFit = _
                                MinValue(remainingWidth, remainingHeight)

                            longFit = _
                                MaxValue(remainingWidth, remainingHeight)

                            areaFit = _
                                (freeWidth(i) * freeHeight(i)) - _
                                (packedWidth * packedHeight)

                            Select Case scoreMode

                                Case 1

                                    If IsBetterShortSideFit( _
                                        shortFit, _
                                        longFit, _
                                        candidateY, _
                                        candidateX, _
                                        bestShortFit, _
                                        bestLongFit, _
                                        bestCandidateY, _
                                        bestCandidateX) Then

                                        bestShortFit = shortFit
                                        bestLongFit = longFit
                                        bestCandidateY = candidateY
                                        bestCandidateX = candidateX

                                        bestX = candidateX
                                        bestY = candidateY
                                        bestObjectWidth = testWidth
                                        bestObjectHeight = testHeight
                                        bestPackedWidth = packedWidth
                                        bestPackedHeight = packedHeight
                                        bestRotate = (orientation = 1)

                                        FindBestMaxRectsPlacement = True
                                    End If

                                Case 2

                                    If IsBetterAreaFit( _
                                        areaFit, _
                                        shortFit, _
                                        candidateY, _
                                        candidateX, _
                                        bestAreaFit, _
                                        bestShortFit, _
                                        bestCandidateY, _
                                        bestCandidateX) Then

                                        bestAreaFit = areaFit
                                        bestShortFit = shortFit
                                        bestCandidateY = candidateY
                                        bestCandidateX = candidateX

                                        bestX = candidateX
                                        bestY = candidateY
                                        bestObjectWidth = testWidth
                                        bestObjectHeight = testHeight
                                        bestPackedWidth = packedWidth
                                        bestPackedHeight = packedHeight
                                        bestRotate = (orientation = 1)

                                        FindBestMaxRectsPlacement = True
                                    End If

                                Case 3

                                    If IsBetterBottomLeft( _
                                        candidateY, _
                                        candidateX, _
                                        areaFit, _
                                        bestCandidateY, _
                                        bestCandidateX, _
                                        bestAreaFit) Then

                                        bestCandidateY = candidateY
                                        bestCandidateX = candidateX
                                        bestAreaFit = areaFit

                                        bestX = candidateX
                                        bestY = candidateY
                                        bestObjectWidth = testWidth
                                        bestObjectHeight = testHeight
                                        bestPackedWidth = packedWidth
                                        bestPackedHeight = packedHeight
                                        bestRotate = (orientation = 1)

                                        FindBestMaxRectsPlacement = True
                                    End If

                            End Select
                        End If
                    End If
                End If
            Next i
        End If
    Next orientation
End Function

Private Sub SplitMaxRects( _
    ByVal usedLeft As Double, _
    ByVal usedBottom As Double, _
    ByVal usedWidth As Double, _
    ByVal usedHeight As Double, _
    ByRef freeLeft() As Double, _
    ByRef freeBottom() As Double, _
    ByRef freeWidth() As Double, _
    ByRef freeHeight() As Double, _
    ByRef freeCount As Long)

    Dim newLeft() As Double
    Dim newBottom() As Double
    Dim newWidth() As Double
    Dim newHeight() As Double

    Dim newCount As Long
    Dim i As Long

    Dim freeRight As Double
    Dim freeTop As Double

    Dim usedRight As Double
    Dim usedTop As Double

    ReDim newLeft(1 To 10)
    ReDim newBottom(1 To 10)
    ReDim newWidth(1 To 10)
    ReDim newHeight(1 To 10)

    usedRight = usedLeft + usedWidth
    usedTop = usedBottom + usedHeight

    For i = 1 To freeCount

        freeRight = freeLeft(i) + freeWidth(i)
        freeTop = freeBottom(i) + freeHeight(i)

        If Not RectanglesIntersect( _
            freeLeft(i), _
            freeBottom(i), _
            freeWidth(i), _
            freeHeight(i), _
            usedLeft, _
            usedBottom, _
            usedWidth, _
            usedHeight) Then

            AddFreeRectangle _
                newLeft, _
                newBottom, _
                newWidth, _
                newHeight, _
                newCount, _
                freeLeft(i), _
                freeBottom(i), _
                freeWidth(i), _
                freeHeight(i)

        Else

            If usedLeft > freeLeft(i) + EPSILON Then

                AddFreeRectangle _
                    newLeft, _
                    newBottom, _
                    newWidth, _
                    newHeight, _
                    newCount, _
                    freeLeft(i), _
                    freeBottom(i), _
                    usedLeft - freeLeft(i), _
                    freeHeight(i)
            End If

            If usedRight < freeRight - EPSILON Then

                AddFreeRectangle _
                    newLeft, _
                    newBottom, _
                    newWidth, _
                    newHeight, _
                    newCount, _
                    usedRight, _
                    freeBottom(i), _
                    freeRight - usedRight, _
                    freeHeight(i)
            End If

            If usedBottom > freeBottom(i) + EPSILON Then

                AddFreeRectangle _
                    newLeft, _
                    newBottom, _
                    newWidth, _
                    newHeight, _
                    newCount, _
                    freeLeft(i), _
                    freeBottom(i), _
                    freeWidth(i), _
                    usedBottom - freeBottom(i)
            End If

            If usedTop < freeTop - EPSILON Then

                AddFreeRectangle _
                    newLeft, _
                    newBottom, _
                    newWidth, _
                    newHeight, _
                    newCount, _
                    freeLeft(i), _
                    usedTop, _
                    freeWidth(i), _
                    freeTop - usedTop
            End If
        End If
    Next i

    freeCount = newCount

    ReDim freeLeft(1 To MaxValueLong(newCount, 1))
    ReDim freeBottom(1 To MaxValueLong(newCount, 1))
    ReDim freeWidth(1 To MaxValueLong(newCount, 1))
    ReDim freeHeight(1 To MaxValueLong(newCount, 1))

    For i = 1 To newCount
        freeLeft(i) = newLeft(i)
        freeBottom(i) = newBottom(i)
        freeWidth(i) = newWidth(i)
        freeHeight(i) = newHeight(i)
    Next i

    PruneFreeRectangles _
        freeLeft, _
        freeBottom, _
        freeWidth, _
        freeHeight, _
        freeCount
End Sub

Private Sub AddFreeRectangle( _
    ByRef rectLeft() As Double, _
    ByRef rectBottom() As Double, _
    ByRef rectWidth() As Double, _
    ByRef rectHeight() As Double, _
    ByRef rectCount As Long, _
    ByVal leftValue As Double, _
    ByVal bottomValue As Double, _
    ByVal widthValue As Double, _
    ByVal heightValue As Double)

    Dim capacity As Long

    If widthValue <= EPSILON Or heightValue <= EPSILON Then
        Exit Sub
    End If

    capacity = UBound(rectLeft)

    If rectCount >= capacity Then

        capacity = capacity * 2

        If capacity < 10 Then
            capacity = 10
        End If

        ReDim Preserve rectLeft(1 To capacity)
        ReDim Preserve rectBottom(1 To capacity)
        ReDim Preserve rectWidth(1 To capacity)
        ReDim Preserve rectHeight(1 To capacity)
    End If

    rectCount = rectCount + 1

    rectLeft(rectCount) = leftValue
    rectBottom(rectCount) = bottomValue
    rectWidth(rectCount) = widthValue
    rectHeight(rectCount) = heightValue
End Sub

Private Sub PruneFreeRectangles( _
    ByRef freeLeft() As Double, _
    ByRef freeBottom() As Double, _
    ByRef freeWidth() As Double, _
    ByRef freeHeight() As Double, _
    ByRef freeCount As Long)

    Dim remove() As Boolean
    Dim i As Long
    Dim j As Long
    Dim newCount As Long

    ReDim remove(1 To MaxValueLong(freeCount, 1))

    For i = 1 To freeCount

        If Not remove(i) Then

            For j = 1 To freeCount

                If i <> j Then

                    If RectContains( _
                        freeLeft(j), _
                        freeBottom(j), _
                        freeWidth(j), _
                        freeHeight(j), _
                        freeLeft(i), _
                        freeBottom(i), _
                        freeWidth(i), _
                        freeHeight(i)) Then

                        remove(i) = True
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i

    newCount = 0

    For i = 1 To freeCount

        If Not remove(i) Then

            newCount = newCount + 1

            freeLeft(newCount) = freeLeft(i)
            freeBottom(newCount) = freeBottom(i)
            freeWidth(newCount) = freeWidth(i)
            freeHeight(newCount) = freeHeight(i)
        End If
    Next i

    freeCount = newCount

    If freeCount = 0 Then
        ReDim freeLeft(1 To 1)
        ReDim freeBottom(1 To 1)
        ReDim freeWidth(1 To 1)
        ReDim freeHeight(1 To 1)
    End If
End Sub

Private Function RectanglesIntersect( _
    ByVal leftA As Double, _
    ByVal bottomA As Double, _
    ByVal widthA As Double, _
    ByVal heightA As Double, _
    ByVal leftB As Double, _
    ByVal bottomB As Double, _
    ByVal widthB As Double, _
    ByVal heightB As Double) As Boolean

    Dim rightA As Double
    Dim topA As Double
    Dim rightB As Double
    Dim topB As Double

    rightA = leftA + widthA
    topA = bottomA + heightA

    rightB = leftB + widthB
    topB = bottomB + heightB

    If rightA <= leftB + EPSILON Then Exit Function
    If leftA >= rightB - EPSILON Then Exit Function
    If topA <= bottomB + EPSILON Then Exit Function
    If bottomA >= topB - EPSILON Then Exit Function

    RectanglesIntersect = True
End Function

Private Function RectContains( _
    ByVal outerLeft As Double, _
    ByVal outerBottom As Double, _
    ByVal outerWidth As Double, _
    ByVal outerHeight As Double, _
    ByVal innerLeft As Double, _
    ByVal innerBottom As Double, _
    ByVal innerWidth As Double, _
    ByVal innerHeight As Double) As Boolean

    Dim outerRight As Double
    Dim outerTop As Double

    Dim innerRight As Double
    Dim innerTop As Double

    outerRight = outerLeft + outerWidth
    outerTop = outerBottom + outerHeight

    innerRight = innerLeft + innerWidth
    innerTop = innerBottom + innerHeight

    If outerLeft > innerLeft + EPSILON Then Exit Function
    If outerBottom > innerBottom + EPSILON Then Exit Function
    If outerRight < innerRight - EPSILON Then Exit Function
    If outerTop < innerTop - EPSILON Then Exit Function

    RectContains = True
End Function

Private Function IsBetterShortSideFit( _
    ByVal shortFit As Double, _
    ByVal longFit As Double, _
    ByVal candidateY As Double, _
    ByVal candidateX As Double, _
    ByVal bestShortFit As Double, _
    ByVal bestLongFit As Double, _
    ByVal bestY As Double, _
    ByVal bestX As Double) As Boolean

    If shortFit < bestShortFit - EPSILON Then
        IsBetterShortSideFit = True
        Exit Function
    End If

    If Abs(shortFit - bestShortFit) <= EPSILON Then

        If longFit < bestLongFit - EPSILON Then
            IsBetterShortSideFit = True
            Exit Function
        End If

        If Abs(longFit - bestLongFit) <= EPSILON Then

            If candidateY < bestY - EPSILON Then
                IsBetterShortSideFit = True
                Exit Function
            End If

            If Abs(candidateY - bestY) <= EPSILON Then
                If candidateX < bestX - EPSILON Then
                    IsBetterShortSideFit = True
                End If
            End If
        End If
    End If
End Function

Private Function IsBetterAreaFit( _
    ByVal areaFit As Double, _
    ByVal shortFit As Double, _
    ByVal candidateY As Double, _
    ByVal candidateX As Double, _
    ByVal bestAreaFit As Double, _
    ByVal bestShortFit As Double, _
    ByVal bestY As Double, _
    ByVal bestX As Double) As Boolean

    If areaFit < bestAreaFit - EPSILON Then
        IsBetterAreaFit = True
        Exit Function
    End If

    If Abs(areaFit - bestAreaFit) <= EPSILON Then

        If shortFit < bestShortFit - EPSILON Then
            IsBetterAreaFit = True
            Exit Function
        End If

        If Abs(shortFit - bestShortFit) <= EPSILON Then

            If candidateY < bestY - EPSILON Then
                IsBetterAreaFit = True
                Exit Function
            End If

            If Abs(candidateY - bestY) <= EPSILON Then
                If candidateX < bestX - EPSILON Then
                    IsBetterAreaFit = True
                End If
            End If
        End If
    End If
End Function

Private Function IsBetterBottomLeft( _
    ByVal candidateY As Double, _
    ByVal candidateX As Double, _
    ByVal areaFit As Double, _
    ByVal bestY As Double, _
    ByVal bestX As Double, _
    ByVal bestAreaFit As Double) As Boolean

    If candidateY < bestY - EPSILON Then
        IsBetterBottomLeft = True
        Exit Function
    End If

    If Abs(candidateY - bestY) <= EPSILON Then

        If candidateX < bestX - EPSILON Then
            IsBetterBottomLeft = True
            Exit Function
        End If

        If Abs(candidateX - bestX) <= EPSILON Then
            If areaFit < bestAreaFit - EPSILON Then
                IsBetterBottomLeft = True
            End If
        End If
    End If
End Function

Private Function CanFitItem( _
    ByVal objectWidth As Double, _
    ByVal objectHeight As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double, _
    ByVal gapDoc As Double, _
    ByVal autoRotate As Boolean) As Boolean

    If objectWidth <= boxWidth + EPSILON And _
       objectHeight <= boxHeight + EPSILON Then

        CanFitItem = True
        Exit Function
    End If

    If autoRotate Then

        If objectHeight <= boxWidth + EPSILON And _
           objectWidth <= boxHeight + EPSILON Then

            CanFitItem = True
            Exit Function
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
            valueA = MaxValue(itemWidth(indexA), itemHeight(indexA))
            valueB = MaxValue(itemWidth(indexB), itemHeight(indexB))

        Case 3
            valueA = itemWidth(indexA)
            valueB = itemWidth(indexB)

        Case 4
            valueA = itemHeight(indexA)
            valueB = itemHeight(indexB)

        Case 5
            valueA = itemWidth(indexA) + itemHeight(indexA)
            valueB = itemWidth(indexB) + itemHeight(indexB)

        Case 6
            valueA = _
                MaxValue(itemWidth(indexA), itemHeight(indexA)) / _
                MinValue(itemWidth(indexA), itemHeight(indexA))

            valueB = _
                MaxValue(itemWidth(indexB), itemHeight(indexB)) / _
                MinValue(itemWidth(indexB), itemHeight(indexB))

        Case 7
            valueA = MinValue(itemWidth(indexA), itemHeight(indexA))
            valueB = MinValue(itemWidth(indexB), itemHeight(indexB))

        Case 8
            valueA = itemArea(indexA)
            valueB = itemArea(indexB)

        Case 9
            valueA = MaxValue(itemWidth(indexA), itemHeight(indexA))
            valueB = MaxValue(itemWidth(indexB), itemHeight(indexB))

        Case 10
            valueA = itemWidth(indexA)
            valueB = itemWidth(indexB)

        Case 11
            valueA = itemHeight(indexA)
            valueB = itemHeight(indexB)

        Case 12
            valueA = MinValue(itemWidth(indexA), itemHeight(indexA))
            valueB = MinValue(itemWidth(indexB), itemHeight(indexB))

        Case Else
            valueA = itemArea(indexA)
            valueB = itemArea(indexB)
    End Select

    Select Case sortMode

        Case 8, 9, 10, 11, 12

            If valueA < valueB - EPSILON Then
                ShouldComeBefore = True
            ElseIf Abs(valueA - valueB) <= EPSILON Then
                ShouldComeBefore = indexA < indexB
            End If

        Case Else

            If valueA > valueB + EPSILON Then
                ShouldComeBefore = True
            ElseIf Abs(valueA - valueB) <= EPSILON Then
                ShouldComeBefore = indexA < indexB
            End If
    End Select
End Function

Private Function IsBetterPlan( _
    ByVal usedWidth As Double, _
    ByVal usedHeight As Double, _
    ByVal fragmentation As Long, _
    ByVal bestWidth As Double, _
    ByVal bestHeight As Double, _
    ByVal bestFragmentation As Long) As Boolean

    Dim usedArea As Double
    Dim bestArea As Double

    usedArea = usedWidth * usedHeight
    bestArea = bestWidth * bestHeight

    If usedArea < bestArea - EPSILON Then
        IsBetterPlan = True
        Exit Function
    End If

    If Abs(usedArea - bestArea) <= EPSILON Then

        If usedHeight < bestHeight - EPSILON Then
            IsBetterPlan = True
            Exit Function
        End If

        If Abs(usedHeight - bestHeight) <= EPSILON Then

            If usedWidth < bestWidth - EPSILON Then
                IsBetterPlan = True
                Exit Function
            End If

            If Abs(usedWidth - bestWidth) <= EPSILON Then

                If fragmentation < bestFragmentation Then
                    IsBetterPlan = True
                End If
            End If
        End If
    End If
End Function

Private Sub CopyPlacement( _
    ByRef source() As Double, _
    ByRef target() As Double, _
    ByVal itemCount As Long)

    Dim i As Long

    For i = 1 To itemCount
        target(i) = source(i)
    Next i
End Sub

Private Sub CopyPlacementBoolean( _
    ByRef source() As Boolean, _
    ByRef target() As Boolean, _
    ByVal itemCount As Long)

    Dim i As Long

    For i = 1 To itemCount
        target(i) = source(i)
    Next i
End Sub

Private Function MinValue( _
    ByVal valueA As Double, _
    ByVal valueB As Double) As Double

    If valueA < valueB Then
        MinValue = valueA
    Else
        MinValue = valueB
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

Private Function MaxValueLong( _
    ByVal valueA As Long, _
    ByVal valueB As Long) As Long

    If valueA > valueB Then
        MaxValueLong = valueA
    Else
        MaxValueLong = valueB
    End If
End Function

Private Function FormatObjectSize( _
    ByVal widthValue As Double, _
    ByVal heightValue As Double) As String

    Dim widthM As Double
    Dim heightM As Double

    widthM = Application.ConvertUnits( _
        widthValue, _
        ActiveDocument.Unit, _
        cdrMeter)

    heightM = Application.ConvertUnits( _
        heightValue, _
        ActiveDocument.Unit, _
        cdrMeter)

    FormatObjectSize = _
        Format$(widthM, "0.###") & " x " & _
        Format$(heightM, "0.###") & " m"
End Function