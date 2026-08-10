Public Sub GenerateMorningSummaryMini()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    ' =========================
    ' FIXED / DAILY VALUES
    ' =========================
    Const DEFAULT_BHA_NUMBER As String = "1"

    ' =========================
    ' BASIC DRILLING VALUES
    ' =========================
    Dim bhaNo As Variant
    Dim depth00 As Variant
    Dim depth24 As Variant
    Dim metersDrilledRaw As Variant

    bhaNo = GetByLabelAny(ws, Array("BHA #", "BHA Number", "BHA"))
    If Len(CleanText(bhaNo)) = 0 Then bhaNo = DEFAULT_BHA_NUMBER

    ' Prefer midnight-totals row on Data sheet after layout swap:
    ' I23 = Start Depth, J23 = End Depth, K23 = Total, L23 = Meters Sliding
    depth00 = ws.Range("I23").Value
    depth24 = ws.Range("J23").Value
    metersDrilledRaw = ws.Range("K23").Value

    ' =========================
    ' REQUIRED CELL ADDRESS VALUES
    ' =========================
    Dim slidMetersRaw As Variant
    Dim bitHoursRaw As Variant
    Dim totalBitHoursRaw As Variant
    Dim motorHoursRaw As Variant
    Dim totalMotorHoursRaw As Variant

    Dim lastSurveyRaw As Variant
    Dim incRaw As Variant
    Dim azRaw As Variant
    Dim ewPositionRaw As Variant
    Dim nsPositionRaw As Variant

    slidMetersRaw = ws.Range("L23").Value
    bitHoursRaw = ws.Range("P23").Value
    totalBitHoursRaw = ws.Range("P24").Value
    motorHoursRaw = ws.Range("Q23").Value
    totalMotorHoursRaw = ws.Range("Q24").Value

    lastSurveyRaw = ws.Range("E7").Value
    incRaw = ws.Range("E8").Value
    azRaw = ws.Range("E9").Value
    ewPositionRaw = ws.Range("E13").Value
    nsPositionRaw = ws.Range("E14").Value

    ' =========================
    ' DEPTH MATH (I23 / J23 / K23)
    ' =========================
    Dim n00 As Double
    Dim n24 As Double
    Dim totalDrilled As Double
    Dim nTotalCell As Double

    Dim haveDepth00 As Boolean
    Dim haveDepth24 As Boolean
    Dim haveTotalDrilled As Boolean

    haveDepth00 = TryFirstNumber(depth00, n00)
    haveDepth24 = TryFirstNumber(depth24, n24)

    If haveDepth00 Then depth00 = n00
    If haveDepth24 Then depth24 = n24

    If TryFirstNumber(ws.Range("K23").Value, nTotalCell) Then
        totalDrilled = Abs(nTotalCell)
        haveTotalDrilled = True
    ElseIf haveDepth00 And haveDepth24 Then
        totalDrilled = Abs(n24 - n00)
        haveTotalDrilled = True
    Else
        haveTotalDrilled = TryFirstNumber(metersDrilledRaw, totalDrilled)
        If haveTotalDrilled Then totalDrilled = Abs(totalDrilled)
    End If

    ' =========================
    ' ROTATE / SLIDE MATH
    ' Rotating meters = K23 - L23
    ' Slide% = L23 / K23 ; Rotate% = rotating / K23
    ' =========================
    Dim slidePct As Double
    Dim rotatePct As Double
    Dim slidMeters As Double
    Dim rotateMeters As Double
    Dim nSlidCell As Double

    Dim haveSlidePct As Boolean
    Dim haveRotatePct As Boolean
    Dim haveSlidMeters As Boolean
    Dim haveRotateMeters As Boolean

    haveSlidMeters = TryFirstNumber(slidMetersRaw, nSlidCell)
    If haveSlidMeters Then slidMeters = Abs(nSlidCell)

    If haveTotalDrilled And haveSlidMeters Then
        rotateMeters = totalDrilled - slidMeters
        If rotateMeters < 0 Then rotateMeters = 0
        haveRotateMeters = True
    End If

    If haveTotalDrilled And totalDrilled > 0 Then
        If haveSlidMeters Then
            slidePct = slidMeters / totalDrilled * 100#
            haveSlidePct = True
        End If
        If haveRotateMeters Then
            rotatePct = rotateMeters / totalDrilled * 100#
            haveRotatePct = True
        End If
    End If

    ' =========================
    ' WELL POSITION DIRECTIONS
    ' =========================
    Dim nsValue As Double
    Dim ewValue As Double
    Dim haveNS As Boolean
    Dim haveEW As Boolean
    Dim nsText As String
    Dim ewText As String

    haveNS = TryFirstNumber(nsPositionRaw, nsValue)
    haveEW = TryFirstNumber(ewPositionRaw, ewValue)

    If haveNS Then
        If nsValue < 0 Then
            nsText = Format$(Abs(nsValue), "0.00") & "m South"
        Else
            nsText = Format$(nsValue, "0.00") & "m North"
        End If
    Else
        nsText = "[E14]m North"
    End If

    If haveEW Then
        If ewValue < 0 Then
            ewText = Format$(Abs(ewValue), "0.00") & "m West"
        Else
            ewText = Format$(ewValue, "0.00") & "m East"
        End If
    Else
        ewText = "[E13]m East"
    End If

    ' =========================
    ' BUILD MINI SUMMARY
    ' =========================
    Dim report As String

    report = ""
    report = report & "BHA #" & MiniBhaNumber(bhaNo) & vbCrLf
    report = report & "Last 24Hrs" & vbCrLf
    report = report & MiniDepth0(depth00) & " - " & MiniDepth0(depth24) & " =  Total " & MiniMeters0(totalDrilled, haveTotalDrilled) & vbCrLf
    report = report & "Rotating " & MiniMeters0(rotateMeters, haveRotateMeters) & "  Rotate " & MiniPct0(rotatePct, haveRotatePct) & vbCrLf
    report = report & "Slid " & MiniNumber0(slidMeters, haveSlidMeters) & "M Slid " & MiniPct0(slidePct, haveSlidePct) & vbCrLf
    report = report & "Bit Hours = " & MiniHoursFromCell(bitHoursRaw, "P23") & " Total Bit Hours = " & MiniHoursFromCell(totalBitHoursRaw, "P24") & vbCrLf
    report = report & "Motor Hours = " & MiniHoursFromCell(motorHoursRaw, "Q23") & " Total Motor Hours = " & MiniHoursFromCell(totalMotorHoursRaw, "Q24") & vbCrLf
    report = report & "Last Survey = " & MiniMeters2FromCell(lastSurveyRaw, "E7") & " Inc = " & MiniDegrees2FromCell(incRaw, "E8") & ", Az = " & MiniDegrees2FromCell(azRaw, "E9") & vbCrLf
    report = report & "Well Position @ Survey = " & nsText & ", " & ewText

    PutTextOnClipboard report

End Sub
