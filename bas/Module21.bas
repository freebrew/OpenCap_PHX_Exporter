Attribute VB_Name = "Module21"
Option Explicit

Public Sub GenerateMorningReport()
    Dim ws As Worksheet
    Dim WELL_NUMBER As String
    Set ws = ThisWorkbook.Worksheets("Data")
    
    ' =========================
    ' EDIT THESE DAILY/AS NEEDED
    ' =========================
    
    WELL_NUMBER = CStr(ThisWorkbook.Worksheets("Setup").Range("D4").Value)
    Const HOLE_SIZE As String = ""
    Const NEXT_24_TEXT As String = ""
    Const CURRENTLY_TEXT As String = "Drilling ahead no issues"
    Const ISSUE_TEXT As String = "No issues to report."
    
    ' If you run this after midnight but want yesterday's report date, use Date - 1.
    Dim reportDate As Date
    reportDate = Date
    
    Dim dash As String
    dash = " " & ChrW(8211) & " "
    
    ' =========================
    ' DEPTH / DRILLING VALUES
    ' =========================
    Dim depth00 As Variant
    Dim depth24 As Variant
    Dim metersDrilled As Variant
    Dim slidingPct As Variant
    Dim rotatingPct As Variant
    
    depth00 = GetByLabelAny(ws, Array("00:00 Depth", "Start Depth"))
    depth24 = GetByLabelAny(ws, Array("24:00 Depth", "Midnight Depth"))
    metersDrilled = GetByLabelAny(ws, Array("Meters Drilled"))
    slidingPct = GetByLabelAny(ws, Array("Sliding Percentage", "Sliding Values"))
    rotatingPct = GetByLabelAny(ws, Array("Rotating Percentage", "Rotating Values"))
    
    ' If 00:00 depth is missing, calculate it from midnight depth minus meters drilled.
    If Len(CleanText(depth00)) = 0 Then
        Dim n24 As Double, nDrilled As Double
        If TryFirstNumber(depth24, n24) And TryFirstNumber(metersDrilled, nDrilled) Then
            depth00 = n24 - nDrilled
        End If
    End If
    
    ' =========================
    ' MOTOR INFORMATION
    ' =========================
    Dim motorSN As Variant
    Dim config As Variant
    Dim revL As Variant
    Dim pumpRate As Variant
    Dim rotary As Variant
    
    motorSN = GetByLabelAny(ws, Array("Motor S/N", "Motor SN"))
    config = GetByLabelAny(ws, Array("Configuration", "Config"))
    revL = GetByLabelAny(ws, Array("Rev/L", "Rev / L", "RevL"))
    pumpRate = GetByLabelAny(ws, Array("Pump Rate"))
    rotary = GetByLabelAny(ws, Array("Rotary"))
    
    Dim motorRevsText As String
    Dim totalRPMText As String
    
    motorRevsText = BuildMotorRevsText(pumpRate, revL)
    totalRPMText = BuildTotalRPMText(motorRevsText, rotary, GetByLabelAny(ws, Array("Total RPM")))
    
    Dim todayBitHours As Variant
    Dim totalBitHours As Variant
    Dim todayCircHours As Variant
    Dim totalCircHours As Variant
    Dim totalMotorHours As Variant
    
    todayBitHours = GetByLabelAny(ws, Array("Today Bit Hours", "Daily Bit Hours"))
    totalBitHours = GetByLabelAny(ws, Array("Total Bit Hours"))
    todayCircHours = GetByLabelAny(ws, Array("Today Circ Hours", "Daily Circ Hours"))
    totalCircHours = GetByLabelAny(ws, Array("Total Circ Hours"))
    totalMotorHours = GetByLabelAny(ws, Array("Total Motor Hours"))
    
    ' =========================
    ' WELLBORE POSITION
    ' Same Data cells as Mini Report (E13/E14/E15 + F13/F14 directions)
    ' =========================
    Dim lateralDist As Variant, lateralDir As Variant
    Dim vertDist As Variant, vertDir As Variant
    Dim totalDist As Variant
    Dim planRef As String
    
    lateralDist = ws.Range("E13").Value
    lateralDir = ws.Range("F13").Value
    vertDist = ws.Range("E14").Value
    vertDir = ws.Range("F14").Value
    totalDist = ws.Range("E15").Value
    planRef = "P1"
    
    ' Label fallbacks if direct cells are blank
    If Len(CleanText(lateralDist)) = 0 Then
        lateralDist = GetByLabelAny(ws, Array("R/L from Plan P1", "R/L from Plan P3", "E/W Direction from Plan P3", "E/W Direction from Plan", "E/W Direction"), 1)
    End If
    If Len(CleanText(lateralDir)) = 0 Then
        lateralDir = GetByLabelAny(ws, Array("R/L from Plan P1", "R/L from Plan P3", "E/W Direction from Plan P3", "E/W Direction from Plan", "E/W Direction"), 2)
    End If
    If Len(CleanText(vertDist)) = 0 Then
        vertDist = GetByLabelAny(ws, Array("Above/Below from Plan P1", "Above/Below from Plan P3", "N/S Direction from Plan P3", "N/S Direction from Plan", "N/S Direction"), 1)
    End If
    If Len(CleanText(vertDir)) = 0 Then
        vertDir = GetByLabelAny(ws, Array("Above/Below from Plan P1", "Above/Below from Plan P3", "N/S Direction from Plan P3", "N/S Direction from Plan", "N/S Direction"), 2)
    End If
    If Len(CleanText(totalDist)) = 0 Then
        totalDist = GetByLabelAny(ws, Array("Distance from Current Geo Target", "Distance from Current GeoTarget", "Distance from Target"), 1)
    End If
    
    If Len(CleanText(lateralDir)) = 0 Then lateralDir = DirectionFromText(lateralDist)
    If Len(CleanText(vertDir)) = 0 Then vertDir = DirectionFromText(vertDist)
    
    ' =========================
    ' MOTOR PERFORMANCE
    ' =========================
    Dim motorOutput As Variant
    Dim doglegNeeded As Variant
    Dim downHoleTemp As Variant
    Dim tvdMidnight As Variant
    
    motorOutput = GetByLabelAny(ws, Array("Average Motor Output", "Motor Output"))
    doglegNeeded = GetByLabelAny(ws, Array("Dogleg Needed"))
    downHoleTemp = GetByLabelAny(ws, Array("Down Hole Temp", "Downhole Temp"))
    tvdMidnight = GetByLabelAny(ws, Array("TVD @ Midnight", "TVD at Midnight", "TVD:"))
    
    ' =========================
    ' BUILD REPORT
    ' =========================
    Dim report As String
    
    report = ""
    report = report & "Daily Drilling Summary " & ChrW(8211) & " " & WELL_NUMBER & " " & Format$(reportDate, "mmmm dd, yyyy") & vbCrLf
    report = report & "Drill " & HOLE_SIZE & " " & FormatDepthNumber(depth00) & "m" & ChrW(8211) & " " & FormatDepthNumber(depth24) & "m" & vbCrLf
    report = report & "Next 24 hrs. " & ChrW(8211) & " " & NEXT_24_TEXT & vbCrLf
    report = report & "Currently " & ChrW(8211) & " " & CURRENTLY_TEXT & vbCrLf
    report = report & ISSUE_TEXT & vbCrLf & vbCrLf
    
    report = report & "00:00 Depth: " & FormatDepthNumber(depth00) & "m" & vbCrLf
    report = report & "24:00 Depth: " & FormatDepthNumber(depth24) & "m" & vbCrLf
    report = report & "Meters Drilled: " & FormatDepthNumber(metersDrilled) & "m" & vbCrLf
    report = report & "Sliding Percentage: " & FormatPercentSmart(slidingPct) & vbCrLf
    report = report & "Rotating Percentage: " & FormatPercentSmart(rotatingPct) & vbCrLf & vbCrLf
    
    report = report & "Motor Information" & vbCrLf
    report = report & "Motor S/N: " & CleanText(motorSN) & vbCrLf
    report = report & "Configuration " & ChrW(8211) & " " & CleanText(config) & vbCrLf
    report = report & "Rev/L: " & FormatRevL(revL) & vbCrLf
    report = report & "Pump Rate: " & FormatPumpRate(pumpRate) & vbCrLf
    report = report & "Motor Revs: " & motorRevsText & vbCrLf
    report = report & "Rotary: " & FormatRPMValue(rotary) & vbCrLf
    report = report & "Total RPM: " & totalRPMText & vbCrLf & vbCrLf
    
    report = report & "Today Bit Hours: " & FormatHoursValue(todayBitHours) & vbCrLf
    report = report & "Total Bit Hours: " & FormatHoursValue(totalBitHours) & vbCrLf
    report = report & "Today Circ Hours: " & FormatHoursValue(todayCircHours) & vbCrLf
    report = report & "Total Circ Hours: " & FormatHoursValue(totalCircHours) & vbCrLf
    report = report & "Total Motor Hours:" & FormatHoursValue(totalMotorHours) & vbCrLf & vbCrLf
    
    report = report & "Position of Wellbore (as of last survey)" & vbCrLf
    report = report & FormatPositionDistance(lateralDist) & "m " & ProperCaseDirection(lateralDir) & " of " & planRef & vbCrLf
    report = report & FormatPositionDistance(vertDist) & "m " & ProperCaseDirection(vertDir) & " of " & planRef & vbCrLf
    report = report & FormatPositionDistance(totalDist) & "m Total" & vbCrLf & vbCrLf
    
    
    report = report & "Motor Performance" & vbCrLf
    report = report & "Motor Output: " & FormatDegreePer30m(motorOutput, 2, False) & vbCrLf
    report = report & "Dogleg Needed: " & FormatDegreePer30m(doglegNeeded, 0, True) & vbCrLf & vbCrLf
    
    report = report & "Down Hole Temp " & FormatTemperature(downHoleTemp) & vbCrLf
    report = report & "TVD @ Midnight: " & FormatTVD(tvdMidnight)
    
    PutTextOnClipboard report
    
End Sub
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

    bhaNo = ws.Range("H3").Value
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


Private Function MorningSummaryClean(ByVal v As Variant) As String
    If isError(v) Or IsEmpty(v) Then
        MorningSummaryClean = ""
    Else
        MorningSummaryClean = Trim$(Replace(CStr(v), Chr$(160), " "))
    End If
End Function
Private Function MiniTryPercent(ByVal v As Variant, ByRef pct As Double) As Boolean
    Dim n As Double
    Dim s As String

    s = CleanText(v)

    If TryFirstNumber(v, n) Then
        If InStr(1, s, "%", vbTextCompare) > 0 Then
            pct = n
        ElseIf Abs(n) <= 1 Then
            pct = n * 100
        Else
            pct = n
        End If

        MiniTryPercent = True
    Else
        MiniTryPercent = False
    End If
End Function

Private Function MiniBhaText(ByVal v As Variant) As String
    Dim s As String

    s = CleanText(v)
    s = Replace(s, "BHA", "", , , vbTextCompare)
    s = Replace(s, "#", "")
    s = Trim$(s)

    If Len(s) = 0 Then s = "1"

    MiniBhaText = s
End Function

Private Function MiniFmtDepth0(ByVal v As Variant) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        MiniFmtDepth0 = Format$(n, "0") & "m"
    Else
        MiniFmtDepth0 = "XXm"
    End If
End Function

Private Function MiniFmtMeters2(ByVal n As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniFmtMeters2 = Format$(n, "0.00") & "m"
    Else
        MiniFmtMeters2 = "XXm"
    End If
End Function

Private Function MiniFmtMeters2FromVariant(ByVal v As Variant) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        MiniFmtMeters2FromVariant = Format$(n, "0.00") & "m"
    Else
        MiniFmtMeters2FromVariant = "XXm"
    End If
End Function

Private Function MiniFmtMeters0(ByVal n As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniFmtMeters0 = Format$(n, "0") & "m"
    Else
        MiniFmtMeters0 = "XXm"
    End If
End Function

Private Function MiniFmtNumber0(ByVal n As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniFmtNumber0 = Format$(n, "0")
    Else
        MiniFmtNumber0 = "XX"
    End If
End Function

Private Function MiniFmtPercent0(ByVal pct As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniFmtPercent0 = Format$(pct, "0") & "%"
    Else
        MiniFmtPercent0 = "XX%"
    End If
End Function

Private Function MiniFmtHours(ByVal v As Variant) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        If Abs(n - Round(n, 0)) < 0.005 Then
            MiniFmtHours = Format$(n, "0") & "Hrs"
        Else
            MiniFmtHours = Format$(n, "0.0") & "Hrs"
        End If
    Else
        MiniFmtHours = "XXHrs"
    End If
End Function

Private Function MiniFmtDegrees2(ByVal v As Variant) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        MiniFmtDegrees2 = Format$(n, "0.00") & ChrW(176)
    Else
        MiniFmtDegrees2 = "XX" & ChrW(176)
    End If
End Function

Private Function MiniPositionText(ByVal dist As Variant, ByVal dirValue As Variant, ByVal axis As String) As String
    Dim n As Double
    Dim directionText As String
    Dim signedDistance As Double

    directionText = MiniDirectionText(dirValue)

    If Len(directionText) = 0 Then
        directionText = MiniDirectionText(dist)
    End If

    If TryFirstNumber(dist, n) Then

        If Len(directionText) = 0 Then
            signedDistance = n

            If UCase$(axis) = "NS" Then
                If n < 0 Then
                    directionText = "South"
                Else
                    directionText = "North"
                End If
            Else
                If n < 0 Then
                    directionText = "West"
                Else
                    directionText = "East"
                End If
            End If
        Else
            signedDistance = Abs(n)

            If directionText = "South" Or directionText = "West" Then
                signedDistance = signedDistance * -1
            End If
        End If

        MiniPositionText = Format$(signedDistance, "0.00") & "m " & directionText

    Else
        If UCase$(axis) = "NS" Then
            MiniPositionText = "XXm North"
        Else
            MiniPositionText = "XXm East"
        End If
    End If
End Function

Private Function MiniDirectionText(ByVal v As Variant) As String
    Dim s As String

    s = UCase$(CleanText(v))

    If Len(s) = 0 Then
        MiniDirectionText = ""
    ElseIf InStr(1, s, "NORTH", vbTextCompare) > 0 Or s = "N" Then
        MiniDirectionText = "North"
    ElseIf InStr(1, s, "SOUTH", vbTextCompare) > 0 Or s = "S" Then
        MiniDirectionText = "South"
    ElseIf InStr(1, s, "EAST", vbTextCompare) > 0 Or s = "E" Then
        MiniDirectionText = "East"
    ElseIf InStr(1, s, "WEST", vbTextCompare) > 0 Or s = "W" Then
        MiniDirectionText = "West"
    Else
        MiniDirectionText = ""
    End If
End Function
Private Function MorningSummaryTryNumber(ByVal v As Variant, ByRef n As Double) As Boolean
    MorningSummaryTryNumber = TryFirstNumber(v, n)
End Function

Private Function MorningSummaryPctNumber(ByVal v As Variant) As Variant
    Dim n As Double
    Dim s As String
    
    s = CleanText(v)
    
    If TryFirstNumber(v, n) Then
        If InStr(1, s, "%", vbTextCompare) > 0 Then
            MorningSummaryPctNumber = n
        ElseIf Abs(n) <= 1 Then
            MorningSummaryPctNumber = n * 100
        Else
            MorningSummaryPctNumber = n
        End If
    Else
        MorningSummaryPctNumber = Empty
    End If
End Function

Private Function MorningSummaryBhaText(ByVal v As Variant) As String
    Dim s As String
    
    s = MorningSummaryClean(v)
    s = Replace(s, "BHA", "", , , vbTextCompare)
    s = Replace(s, "#", "")
    s = Trim$(s)
    
    If Len(s) = 0 Then s = "1"
    
    MorningSummaryBhaText = s
End Function

Private Function MorningSummaryFmtDepth0(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        MorningSummaryFmtDepth0 = Format$(n, "0") & "m"
    Else
        MorningSummaryFmtDepth0 = "XXm"
    End If
End Function

Private Function MorningSummaryFmtNumber(ByVal v As Variant, ByVal numberFormat As String) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        MorningSummaryFmtNumber = Format$(n, numberFormat)
    Else
        MorningSummaryFmtNumber = "XX"
    End If
End Function

Private Function MorningSummaryFmtPct0(ByVal v As Variant) As String
    If IsEmpty(v) Then
        MorningSummaryFmtPct0 = "XX%"
    Else
        MorningSummaryFmtPct0 = Format$(CDbl(v), "0") & "%"
    End If
End Function

Private Function MorningSummaryFmtHours(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        If Abs(n - Round(n, 0)) < 0.005 Then
            MorningSummaryFmtHours = Format$(n, "0") & "Hrs"
        Else
            MorningSummaryFmtHours = Format$(n, "0.0") & "Hrs"
        End If
    Else
        MorningSummaryFmtHours = "XXHrs"
    End If
End Function

Private Function MorningSummaryFmtMeters2(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        MorningSummaryFmtMeters2 = Format$(n, "0.00") & "m"
    Else
        MorningSummaryFmtMeters2 = "XXm"
    End If
End Function

Private Function MorningSummaryFmtDegrees2(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        MorningSummaryFmtDegrees2 = Format$(n, "0.00") & ChrW(176)
    Else
        MorningSummaryFmtDegrees2 = "XX" & ChrW(176)
    End If
End Function

Private Function MorningSummaryPositionText(ByVal dist As Variant, ByVal dirValue As Variant, ByVal axis As String) As String
    Dim n As Double
    Dim directionText As String
    Dim signedDistance As Double
    
    directionText = MorningSummaryDirectionText(dirValue, axis)
    
    If Len(directionText) = 0 Then
        directionText = MorningSummaryDirectionText(dist, axis)
    End If
    
    If TryFirstNumber(dist, n) Then
        
        If Len(directionText) = 0 Then
            If UCase$(axis) = "NS" Then
                If n < 0 Then
                    directionText = "South"
                Else
                    directionText = "North"
                End If
            Else
                If n < 0 Then
                    directionText = "West"
                Else
                    directionText = "East"
                End If
            End If
            
            signedDistance = n
        Else
            signedDistance = Abs(n)
            
            If directionText = "South" Or directionText = "West" Then
                signedDistance = signedDistance * -1
            End If
        End If
        
        MorningSummaryPositionText = Format$(signedDistance, "0.00") & "m " & directionText
    Else
        If UCase$(axis) = "NS" Then
            MorningSummaryPositionText = "XXm North"
        Else
            MorningSummaryPositionText = "XXm East"
        End If
    End If
End Function

Private Function MorningSummaryDirectionText(ByVal v As Variant, ByVal axis As String) As String
    Dim s As String
    
    s = UCase$(MorningSummaryClean(v))
    
    If Len(s) = 0 Then
        MorningSummaryDirectionText = ""
        Exit Function
    End If
    
    If InStr(1, s, "NORTH", vbTextCompare) > 0 Or s = "N" Then
        MorningSummaryDirectionText = "North"
    ElseIf InStr(1, s, "SOUTH", vbTextCompare) > 0 Or s = "S" Then
        MorningSummaryDirectionText = "South"
    ElseIf InStr(1, s, "EAST", vbTextCompare) > 0 Or s = "E" Then
        MorningSummaryDirectionText = "East"
    ElseIf InStr(1, s, "WEST", vbTextCompare) > 0 Or s = "W" Then
        MorningSummaryDirectionText = "West"
    Else
        MorningSummaryDirectionText = ""
    End If
End Function

' ============================================================
' LABEL LOOKUP HELPERS
' ============================================================

Private Function GetByLabelAny(ByVal ws As Worksheet, ByVal labels As Variant, Optional ByVal valueOffset As Long = 1, Optional ByVal defaultValue As String = "") As Variant
    Dim i As Long
    Dim v As Variant
    
    For i = LBound(labels) To UBound(labels)
        v = FindLabelValue(ws, CStr(labels(i)), valueOffset)
        If Len(CleanText(v)) > 0 Then
            GetByLabelAny = v
            Exit Function
        End If
    Next i
    
    GetByLabelAny = defaultValue
End Function

Private Function FindLabelValue(ByVal ws As Worksheet, ByVal labelText As String, Optional ByVal valueOffset As Long = 1) As Variant
    Dim ur As Range
    Dim c As Range
    Dim cellText As String
    Dim inlineValue As String
    
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    
    If ur Is Nothing Then Exit Function
    
    For Each c In ur.Cells
        cellText = CleanText(c.text)
        
        If IsLabelMatch(cellText, labelText) Then
            inlineValue = InlineAfterLabel(cellText, labelText)
            
            If valueOffset = 1 And Len(inlineValue) > 0 Then
                FindLabelValue = inlineValue
                Exit Function
            End If
            
            FindLabelValue = NthNonBlankToRight(ws, c, valueOffset)
            If Len(CleanText(FindLabelValue)) > 0 Then Exit Function
        End If
    Next c
End Function

Private Function NthNonBlankToRight(ByVal ws As Worksheet, ByVal startCell As Range, ByVal n As Long) As Variant
    Dim col As Long
    Dim found As Long
    Dim lastCol As Long
    Dim txt As String
    
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    
    For col = startCell.Column + 1 To Application.Min(lastCol, startCell.Column + 10)
        txt = CleanText(ws.Cells(startCell.Row, col).text)
        If Len(txt) > 0 Then
            found = found + 1
            If found = n Then
                NthNonBlankToRight = txt
                Exit Function
            End If
        End If
    Next col
End Function

Private Function IsLabelMatch(ByVal cellText As String, ByVal labelText As String) As Boolean
    Dim a As String
    Dim b As String
    
    a = NormalizeForMatch(cellText)
    b = NormalizeForMatch(labelText)
    
    IsLabelMatch = InStr(1, a, b, vbTextCompare) > 0
End Function

Private Function NormalizeForMatch(ByVal s As String) As String
    s = UCase$(CleanText(s))
    s = Replace(s, ":", "")
    s = Replace(s, ChrW(8211), "-")
    s = Replace(s, ChrW(8212), "-")
    NormalizeForMatch = s
End Function

Private Function InlineAfterLabel(ByVal cellText As String, ByVal labelText As String) As String
    Dim p As Long
    Dim s As String
    
    p = InStr(1, cellText, labelText, vbTextCompare)
    If p = 0 Then Exit Function
    
    s = mid$(cellText, p + Len(labelText))
    s = Trim$(s)
    
    If Left$(s, 1) = ":" Then s = Trim$(mid$(s, 2))
    If Left$(s, 1) = "-" Then s = Trim$(mid$(s, 2))
    If Left$(s, 1) = ChrW(8211) Then s = Trim$(mid$(s, 2))
    
    InlineAfterLabel = s
End Function

' ============================================================
' FORMAT HELPERS
' ============================================================

Private Function CleanText(ByVal v As Variant) As String
    If isError(v) Then
        CleanText = ""
    Else
        CleanText = CStr(v)
        CleanText = Replace(CleanText, Chr$(160), " ")
        CleanText = Application.WorksheetFunction.Trim(CleanText)
    End If
End Function

Private Function TryFirstNumber(ByVal v As Variant, ByRef result As Double) As Boolean
    On Error GoTo Fail
    
    If isError(v) Or Len(CleanText(v)) = 0 Then Exit Function
    
    If IsNumeric(v) Then
        result = CDbl(v)
        TryFirstNumber = True
        Exit Function
    End If
    
    Dim re As Object
    Dim matches As Object
    Dim s As String
    
    s = Replace(CleanText(v), ",", "")
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[-+]?\d*\.?\d+"
    re.Global = False
    
    If re.Test(s) Then
        Set matches = re.Execute(s)
        result = CDbl(matches(0).Value)
        TryFirstNumber = True
    End If
    
    Exit Function
    
Fail:
    TryFirstNumber = False
End Function

Private Function CountNumbers(ByVal v As Variant) As Long
    On Error GoTo Fail
    
    Dim re As Object
    Dim matches As Object
    Dim s As String
    
    s = Replace(CleanText(v), ",", "")
    
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "[-+]?\d*\.?\d+"
    re.Global = True
    
    If re.Test(s) Then
        Set matches = re.Execute(s)
        CountNumbers = matches.Count
    End If
    
    Exit Function
    
Fail:
    CountNumbers = 0
End Function

Private Function FormatDepthNumber(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        If Abs(n - Round(n, 0)) < 0.0001 Then
            FormatDepthNumber = Format$(n, "0")
        Else
            FormatDepthNumber = Format$(n, "0.###")
        End If
    Else
        FormatDepthNumber = CleanText(v)
        FormatDepthNumber = Replace(FormatDepthNumber, "m", "", , , vbTextCompare)
    End If
End Function

Private Function FormatPercentSmart(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        If InStr(1, CleanText(v), "%", vbTextCompare) > 0 Then
            FormatPercentSmart = Format$(n / 100, "0.00%")
        ElseIf Abs(n) <= 1 Then
            FormatPercentSmart = Format$(n, "0.00%")
        Else
            FormatPercentSmart = Format$(n / 100, "0.00%")
        End If
    Else
        FormatPercentSmart = CleanText(v)
    End If
End Function

Private Function FormatRevL(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        FormatRevL = Format$(n, "0.000")
    Else
        FormatRevL = CleanText(v)
    End If
End Function

Private Function FormatPumpRate(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        FormatPumpRate = Format$(n, "0.00") & "m" & ChrW(179) & "/min"
    Else
        FormatPumpRate = CleanText(v)
    End If
End Function

Private Function BuildMotorRevsText(ByVal pumpRate As Variant, ByVal revL As Variant) As String
    Dim pumpM3Min As Double
    Dim revPerL As Double
    Dim motorRPM As Double
    
    If TryFirstNumber(pumpRate, pumpM3Min) And TryFirstNumber(revL, revPerL) Then
        ' Motor RPM = pump rate in m3/min * 1000 L/m3 * rev/L
        motorRPM = pumpM3Min * 1000 * revPerL
        BuildMotorRevsText = Format$(motorRPM, "0") & " rpm"
    Else
        BuildMotorRevsText = ""
    End If
End Function

Private Function FormatRPMValue(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        FormatRPMValue = Format$(n, "0") & " rpm"
    Else
        FormatRPMValue = CleanText(v)
    End If
End Function

Private Function BuildTotalRPMText(ByVal motorRevsText As String, ByVal rotary As Variant, ByVal totalRPMFallback As Variant) As String
    Dim motorRPM As Double
    Dim rotaryRPM As Double
    
    If TryFirstNumber(motorRevsText, motorRPM) And TryFirstNumber(rotary, rotaryRPM) Then
        BuildTotalRPMText = Format$(motorRPM + rotaryRPM, "0") & " rpm"
    ElseIf Len(CleanText(totalRPMFallback)) > 0 Then
        BuildTotalRPMText = FormatRPMValue(totalRPMFallback)
    Else
        BuildTotalRPMText = ""
    End If
End Function

Private Function FormatHoursValue(ByVal v As Variant) As String
    Dim n As Double
    Dim s As String
    
    s = CleanText(v)
    
    If Len(s) = 0 Then
        FormatHoursValue = ""
        Exit Function
    End If
    
    ' Preserve odd manually typed values like "1hr8s."
    If CountNumbers(s) > 1 Then
        If right$(s, 1) <> "." Then s = s & "."
        FormatHoursValue = s
        Exit Function
    End If
    
    If TryFirstNumber(v, n) Then
        If Abs(n - Round(n, 0)) < 0.0001 Then
            FormatHoursValue = Format$(n, "0") & "hrs."
        Else
            FormatHoursValue = Format$(n, "0.0") & "hrs."
        End If
    Else
        s = Replace(s, "Hrs", "hrs", , , vbTextCompare)
        s = Replace(s, "Hr", "hr", , , vbTextCompare)
        If right$(s, 1) <> "." Then s = s & "."
        FormatHoursValue = s
    End If
End Function

Private Function FormatPositionDistance(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        FormatPositionDistance = Format$(n, "0.00")
    Else
        FormatPositionDistance = CleanText(v)
        FormatPositionDistance = Replace(FormatPositionDistance, "m", "", , , vbTextCompare)
    End If
End Function

Private Function DirectionFromText(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(CleanText(v))
    
    If InStr(s, "LEFT") > 0 Then DirectionFromText = "LEFT": Exit Function
    If InStr(s, "RIGHT") > 0 Then DirectionFromText = "RIGHT": Exit Function
    If InStr(s, "ABOVE") > 0 Then DirectionFromText = "ABOVE": Exit Function
    If InStr(s, "BELOW") > 0 Then DirectionFromText = "BELOW": Exit Function
    If InStr(s, "NORTH") > 0 Then DirectionFromText = "NORTH": Exit Function
    If InStr(s, "SOUTH") > 0 Then DirectionFromText = "SOUTH": Exit Function
    If InStr(s, "EAST") > 0 Then DirectionFromText = "EAST": Exit Function
    If InStr(s, "WEST") > 0 Then DirectionFromText = "WEST": Exit Function
    
    DirectionFromText = ""
End Function

Private Function FormatDegreePer30m(ByVal v As Variant, ByVal decimals As Long, ByVal spaceBeforeDegree As Boolean) As String
    Dim n As Double
    Dim spacer As String
    
    spacer = IIf(spaceBeforeDegree, " ", "")
    
    If TryFirstNumber(v, n) Then
        If decimals = 0 Then
            FormatDegreePer30m = Format$(n, "0") & spacer & ChrW(176) & "/30m"
        Else
            FormatDegreePer30m = Format$(n, "0." & String$(decimals, "0")) & spacer & ChrW(176) & "/30m"
        End If
    Else
        FormatDegreePer30m = CleanText(v)
    End If
End Function

Private Function FormatTemperature(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        FormatTemperature = Format$(n, "0") & ChrW(176) & "C"
    Else
        FormatTemperature = CleanText(v)
    End If
End Function

Private Function FormatTVD(ByVal v As Variant) As String
    Dim n As Double
    
    If TryFirstNumber(v, n) Then
        FormatTVD = Format$(n, "0.###") & "mTVD"
    Else
        FormatTVD = CleanText(v)
    End If
End Function
Private Function ProperCaseDirection(ByVal v As Variant) As String
    Dim s As String
    
    s = LCase$(CleanText(v))
    
    Select Case s
        Case "north"
            ProperCaseDirection = "North"
        Case "south"
            ProperCaseDirection = "South"
        Case "east"
            ProperCaseDirection = "East"
        Case "west"
            ProperCaseDirection = "West"
        Case "left"
            ProperCaseDirection = "Left"
        Case "right"
            ProperCaseDirection = "Right"
        Case "above"
            ProperCaseDirection = "Above"
        Case "below"
            ProperCaseDirection = "Below"
        Case Else
            ProperCaseDirection = StrConv(s, vbProperCase)
    End Select
End Function

' ============================================================
' CLIPBOARD
' ============================================================

Private Sub PutTextOnClipboard(ByVal txt As String)
    Dim dataObj As Object
    
    ' Late-bound MSForms DataObject.
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    dataObj.SetText txt
    dataObj.PutInClipboard
End Sub
Private Sub PickAttachmentFileRow(ByVal r As Long)
    Dim ws As Worksheet
    Dim fd As FileDialog
    Set ws = ThisWorkbook.Worksheets("Data")
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Select attachment file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then
            Module11.StoreAttachPath ws.Cells(r, "I"), .SelectedItems(1)
        End If
    End With
End Sub
Sub PickAttachmentFile_Y14()
    PickAttachmentFileRow 29
End Sub
Sub PickAttachmentFile_Y15()
    PickAttachmentFileRow 30
End Sub
Sub PickAttachmentFile_Y16()
    PickAttachmentFileRow 31
End Sub
Sub PickAttachmentFile_Y17()
    PickAttachmentFileRow 32
End Sub
Sub PickAttachmentFile_Y18()
    PickAttachmentFileRow 33
End Sub
Private Function MiniPercentFromValue(ByVal v As Variant, ByRef pct As Double) As Boolean
    Dim n As Double
    Dim s As String

    s = CleanText(v)

    If TryFirstNumber(v, n) Then
        If InStr(1, s, "%", vbTextCompare) > 0 Then
            pct = n
        ElseIf Abs(n) <= 1 Then
            pct = n * 100
        Else
            pct = n
        End If

        MiniPercentFromValue = True
    Else
        MiniPercentFromValue = False
    End If
End Function

Private Function MiniBhaNumber(ByVal v As Variant) As String
    Dim s As String

    s = CleanText(v)
    s = Replace(s, "BHA", "", , , vbTextCompare)
    s = Replace(s, "#", "")
    s = Trim$(s)

    If Len(s) = 0 Then s = "1"

    MiniBhaNumber = s
End Function

Private Function MiniDepth0(ByVal v As Variant) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        MiniDepth0 = Format$(n, "0") & "m"
    Else
        MiniDepth0 = "XXm"
    End If
End Function

Private Function MiniMeters2(ByVal n As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniMeters2 = Format$(n, "0.00") & "m"
    Else
        MiniMeters2 = "XXm"
    End If
End Function

Private Function MiniMeters0(ByVal n As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniMeters0 = Format$(n, "0") & "m"
    Else
        MiniMeters0 = "XXm"
    End If
End Function

Private Function MiniNumber0(ByVal n As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniNumber0 = Format$(n, "0")
    Else
        MiniNumber0 = "XX"
    End If
End Function

Private Function MiniPct0(ByVal n As Double, ByVal hasValue As Boolean) As String
    If hasValue Then
        MiniPct0 = Format$(n, "0.00") & "%"
    Else
        MiniPct0 = "XX%"
    End If
End Function

Private Function MiniHoursFromCell(ByVal v As Variant, ByVal cellAddress As String) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        If Abs(n - Round(n, 0)) < 0.005 Then
            MiniHoursFromCell = Format$(n, "0") & "Hrs"
        Else
            MiniHoursFromCell = Format$(n, "0.0") & "Hrs"
        End If
    Else
        MiniHoursFromCell = "[" & cellAddress & "]Hrs"
    End If
End Function

Private Function MiniMeters2FromCell(ByVal v As Variant, ByVal cellAddress As String) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        MiniMeters2FromCell = Format$(n, "0.00") & "m"
    Else
        MiniMeters2FromCell = "[" & cellAddress & "]m"
    End If
End Function

Private Function MiniDegrees2FromCell(ByVal v As Variant, ByVal cellAddress As String) As String
    Dim n As Double

    If TryFirstNumber(v, n) Then
        MiniDegrees2FromCell = Format$(n, "0.00") & ChrW(176)
    Else
        MiniDegrees2FromCell = "[" & cellAddress & "]" & ChrW(176)
    End If
End Function
