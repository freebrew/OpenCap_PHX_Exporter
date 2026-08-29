Attribute VB_Name = "MDL_PlanGauge"
Option Explicit

' ================================================================================
'  PLAN PROXIMITY GAUGE  (Slidesheet AA1:AC7 - dial left in AA:AB, metrics right, ending at AC)
'
'  Single bullseye: plan at centre, survey (filled) + PTB (hollow) on one dial.
'
'  GRAVITY scale fits the waypoint TVD corridor (+/- AB14 about waypoint TVD
'  on the dial Y axis). Markers may sit outside the ring (Option 1). Band:
'  horizontal green/red RECTANGLES that extend past the circle (not clipped
'  to the rim) so outside-circle markers can still sit in green. MAG: no band.
'
'  Frames (plan = origin):
'    GRAVITY  (Inc >= 5 deg)
'      Plan at same MD. Y = planTVD - actualTVD (+ = UP). X = LT/RT of plan Azm.
'    MAGNETIC (Inc <  5 deg)
'      Plan foot at actual TVD. Y = ahead/back, X = right/left; PRP = hypot.
'
'  Way points: AC14:AD33 (MD/TVD). Half-width: AB14.
'  Actual N/E from min-curvature on E/F/G; TVD prefers H. LOOKUP rows skipped.
'  PTB bit depth = column D on that survey row.
' ================================================================================

Private Const SS_SHEET As String = "Slidesheet"
Private Const PLAN_SHEET As String = "_OC_Survey"
Private Const GAUGE_RANGE As String = "AA1:AC7"
Private Const SHP_PREFIX As String = "OCG_"

Private Const FS_HDR As Single = 12
Private Const FS_VAL As Single = 11
Private Const FS_AXIS As Single = 10
Private Const FS_CAP As Single = 10
Private Const CROSSOVER_INC As Double = 5#
Private Const SURV_ROW_FIRST As Long = 13
Private Const SURV_ROW_LAST As Long = 320
Private Const WP_ROW_FIRST As Long = 14
Private Const WP_ROW_LAST As Long = 33
Private Const WP_COL_MD As Long = 29          ' AC
Private Const WP_COL_TVD As Long = 30         ' AD
Private Const WP_HALFWIDTH_ADDR As String = "AB14"

Private Const PI As Double = 3.14159265358979

Private Function cInk() As Long:    cInk = RGB(40, 40, 40):      End Function
Private Function cGrid() As Long:   cGrid = RGB(180, 180, 180):  End Function
' Dial ring + U/D/L/R — dark so they read on pastel bands
Private Function cAxis() As Long:   cAxis = RGB(30, 30, 30):     End Function
Private Function cPlan() As Long:   cPlan = RGB(0, 90, 90):      End Function
Private Function cAct() As Long:    cAct = RGB(192, 0, 0):       End Function
Private Function cPtb() As Long:    cPtb = RGB(0, 90, 190):      End Function
' Soft band fills (pastel); edge lines + OUT text stay saturated
Private Function cBandIn() As Long: cBandIn = RGB(200, 235, 210): End Function
Private Function cBandOut() As Long: cBandOut = RGB(255, 215, 215): End Function
Private Function cBandEdge() As Long: cBandEdge = RGB(35, 130, 65): End Function
Private Function cBandOutInk() As Long: cBandOutInk = RGB(180, 40, 40): End Function

Private Function Deg2Rad(ByVal d As Double) As Double
    Deg2Rad = d * PI / 180#
End Function

Private Function ClampDbl(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then
        ClampDbl = lo
    ElseIf v > hi Then
        ClampDbl = hi
    Else
        ClampDbl = v
    End If
End Function

' --------------------------------------------------------------------------------
'  Minimum curvature step between two stations -> delta TVD / North / East
' --------------------------------------------------------------------------------
Private Sub McStep(ByVal md1 As Double, ByVal i1 As Double, ByVal a1 As Double, _
                   ByVal md2 As Double, ByVal i2 As Double, ByVal a2 As Double, _
                   ByRef dv As Double, ByRef dN As Double, ByRef dE As Double)
    Dim r1 As Double, r2 As Double, b1 As Double, b2 As Double
    Dim cosDL As Double, beta As Double, H As Double
    r1 = Deg2Rad(i1): r2 = Deg2Rad(i2)
    b1 = Deg2Rad(a1): b2 = Deg2Rad(a2)

    cosDL = Cos(r2 - r1) - Sin(r1) * Sin(r2) * (1# - Cos(b2 - b1))
    If cosDL > 1# Then cosDL = 1#
    If cosDL < -1# Then cosDL = -1#
    beta = WorksheetFunction.Acos(cosDL)

    If beta > 0.0000001 Then
        H = (md2 - md1) / 2# * (2# / beta * Tan(beta / 2#))
    Else
        H = (md2 - md1) / 2#
    End If
    dv = H * (Cos(r1) + Cos(r2))
    dN = H * (Sin(r1) * Cos(b1) + Sin(r2) * Cos(b2))
    dE = H * (Sin(r1) * Sin(b1) + Sin(r2) * Sin(b2))
End Sub

' --------------------------------------------------------------------------------
'  Plan access: load MD / INC / AZI / TVD / NS / EW arrays from _OC_Survey
' --------------------------------------------------------------------------------
Private Function LoadPlan(ByRef pMD() As Double, ByRef pInc() As Double, _
                          ByRef pAzi() As Double, ByRef pTvd() As Double, _
                          ByRef pNS() As Double, ByRef pEW() As Double) As Long
    LoadPlan = 0
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PLAN_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim cMD As Long, cInc As Long, cAzi As Long, cTvd As Long, cNS As Long, cEW As Long
    Dim c As Long
    For c = 1 To 20
        Select Case UCase$(Trim$(CStr(ws.Cells(2, c).Value2 & "")))
            Case "MD":  cMD = c
            Case "INC": cInc = c
            Case "AZI", "AZM": cAzi = c
            Case "TVD": cTvd = c
            Case "NS":  cNS = c
            Case "EW":  cEW = c
        End Select
    Next c
    If cMD = 0 Or cInc = 0 Or cAzi = 0 Or cTvd = 0 Or cNS = 0 Or cEW = 0 Then Exit Function

    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, cMD).End(xlUp).Row
    If lastR < 3 Then Exit Function

    Dim n As Long: n = 0
    ReDim pMD(lastR): ReDim pInc(lastR): ReDim pAzi(lastR)
    ReDim pTvd(lastR): ReDim pNS(lastR): ReDim pEW(lastR)

    Dim r As Long
    For r = 3 To lastR
        If IsNumeric(ws.Cells(r, cMD).Value2) And Len(CStr(ws.Cells(r, cMD).Value2 & "")) > 0 Then
            Dim mdv As Double: mdv = CDbl(ws.Cells(r, cMD).Value2)
            Dim takeIt As Boolean
            If n = 0 Then takeIt = True Else takeIt = (mdv > pMD(n - 1))
            If takeIt Then
                pMD(n) = mdv
                pInc(n) = val(ws.Cells(r, cInc).Value2 & "")
                pAzi(n) = val(ws.Cells(r, cAzi).Value2 & "")
                pTvd(n) = val(ws.Cells(r, cTvd).Value2 & "")
                pNS(n) = val(ws.Cells(r, cNS).Value2 & "")
                pEW(n) = val(ws.Cells(r, cEW).Value2 & "")
                n = n + 1
            End If
        End If
    Next r
    LoadPlan = n
End Function

Private Sub PlanAt(ByVal md As Double, ByVal n As Long, _
                   pMD() As Double, pInc() As Double, pAzi() As Double, _
                   pTvd() As Double, pNS() As Double, pEW() As Double, _
                   ByRef outN As Double, ByRef outE As Double, ByRef outV As Double, _
                   ByRef outAzi As Double, ByRef outInc As Double)
    If md <= pMD(0) Then
        outN = pNS(0): outE = pEW(0): outV = pTvd(0)
        outAzi = pAzi(0): outInc = pInc(0)
        Exit Sub
    End If
    If md >= pMD(n - 1) Then
        outN = pNS(n - 1): outE = pEW(n - 1): outV = pTvd(n - 1)
        outAzi = pAzi(n - 1): outInc = pInc(n - 1)
        Exit Sub
    End If

    Dim i As Long
    For i = 0 To n - 2
        If md >= pMD(i) And md <= pMD(i + 1) Then Exit For
    Next i

    Dim f As Double
    If pMD(i + 1) > pMD(i) Then f = (md - pMD(i)) / (pMD(i + 1) - pMD(i)) Else f = 0#
    outInc = pInc(i) + f * (pInc(i + 1) - pInc(i))
    outN = pNS(i) + f * (pNS(i + 1) - pNS(i))
    outE = pEW(i) + f * (pEW(i + 1) - pEW(i))
    outV = pTvd(i) + f * (pTvd(i + 1) - pTvd(i))

    Dim s As Double, c As Double
    s = (1# - f) * Sin(Deg2Rad(pAzi(i))) + f * Sin(Deg2Rad(pAzi(i + 1)))
    c = (1# - f) * Cos(Deg2Rad(pAzi(i))) + f * Cos(Deg2Rad(pAzi(i + 1)))
    If Abs(s) < 0.0000000001 And Abs(c) < 0.0000000001 Then
        outAzi = pAzi(i)
    Else
        outAzi = WorksheetFunction.Atan2(c, s) * 180# / PI
        If outAzi < 0 Then outAzi = outAzi + 360#
    End If
End Sub

Private Sub PlanAtTvd(ByVal tvd As Double, ByVal n As Long, _
                      pMD() As Double, pInc() As Double, pAzi() As Double, _
                      pTvd() As Double, pNS() As Double, pEW() As Double, _
                      ByRef outN As Double, ByRef outE As Double, ByRef outV As Double, _
                      ByRef outAzi As Double, ByRef outInc As Double)
    Dim i As Long, f As Double
    Dim lo As Double, hi As Double
    Dim s As Double, c As Double

    For i = 0 To n - 2
        lo = pTvd(i): hi = pTvd(i + 1)
        If Abs(hi - lo) < 0.0001 Then GoTo NextTvdSeg
        If (tvd >= lo And tvd <= hi) Or (tvd >= hi And tvd <= lo) Then
            f = (tvd - lo) / (hi - lo)
            outInc = pInc(i) + f * (pInc(i + 1) - pInc(i))
            outN = pNS(i) + f * (pNS(i + 1) - pNS(i))
            outE = pEW(i) + f * (pEW(i + 1) - pEW(i))
            outV = tvd
            s = (1# - f) * Sin(Deg2Rad(pAzi(i))) + f * Sin(Deg2Rad(pAzi(i + 1)))
            c = (1# - f) * Cos(Deg2Rad(pAzi(i))) + f * Cos(Deg2Rad(pAzi(i + 1)))
            If Abs(s) < 0.0000000001 And Abs(c) < 0.0000000001 Then
                outAzi = pAzi(i)
            Else
                outAzi = WorksheetFunction.Atan2(c, s) * 180# / PI
                If outAzi < 0 Then outAzi = outAzi + 360#
            End If
            Exit Sub
        End If
NextTvdSeg:
    Next i

    Dim bestI As Long: bestI = 0
    Dim bestD As Double: bestD = Abs(pTvd(0) - tvd)
    For i = 1 To n - 1
        If Abs(pTvd(i) - tvd) < bestD Then
            bestD = Abs(pTvd(i) - tvd)
            bestI = i
        End If
    Next i
    outN = pNS(bestI): outE = pEW(bestI): outV = pTvd(bestI)
    outAzi = pAzi(bestI): outInc = pInc(bestI)
End Sub

' Linear waypoint TVD at MD from AC14:AD33 (same rule as GEO Window AB formulas).
' Returns False when MD is outside the waypoint MD span or fewer than 2 points.
Private Function WaypointTvdAtMd(ws As Worksheet, ByVal md As Double, _
                                 ByRef outTvd As Double) As Boolean
    WaypointTvdAtMd = False
    Dim wMD() As Double, wTvd() As Double
    Dim n As Long: n = 0
    Dim maxN As Long: maxN = WP_ROW_LAST - WP_ROW_FIRST + 8
    ReDim wMD(0 To maxN)
    ReDim wTvd(0 To maxN)

    Dim r As Long
    For r = WP_ROW_FIRST To WP_ROW_LAST
        Dim vM As Variant, vT As Variant
        vM = ws.Cells(r, WP_COL_MD).Value2
        vT = ws.Cells(r, WP_COL_TVD).Value2
        If IsArray(vM) Or IsArray(vT) Then GoTo NextWpRow
        If VarType(vM) = vbEmpty Or VarType(vT) = vbEmpty Then GoTo NextWpRow
        If Not IsNumeric(vM) Or Not IsNumeric(vT) Then GoTo NextWpRow
        If Len(Trim$(CStr(vM))) = 0 Or Len(Trim$(CStr(vT))) = 0 Then GoTo NextWpRow

        Dim mdv As Double: mdv = CDbl(vM)
        ' VBA does not short-circuit Or — never evaluate wMD(n-1) when n=0.
        Dim takeIt As Boolean
        If n = 0 Then
            takeIt = True
        Else
            takeIt = (mdv > wMD(n - 1))
        End If
        If takeIt Then
            If n > maxN Then GoTo NextWpRow
            wMD(n) = mdv
            wTvd(n) = CDbl(vT)
            n = n + 1
        End If
NextWpRow:
    Next r
    If n < 2 Then Exit Function
    If md < wMD(0) Or md > wMD(n - 1) Then Exit Function

    Dim i As Long
    For i = 0 To n - 2
        If md >= wMD(i) And md <= wMD(i + 1) Then Exit For
    Next i
    If i > n - 2 Then i = n - 2

    Dim f As Double
    If wMD(i + 1) > wMD(i) Then
        f = (md - wMD(i)) / (wMD(i + 1) - wMD(i))
    Else
        f = 0#
    End If
    outTvd = wTvd(i) + f * (wTvd(i + 1) - wTvd(i))
    WaypointTvdAtMd = True
End Function

Private Function WaypointHalfWidth(ws As Worksheet) As Double
    Dim v As Variant
    On Error Resume Next
    v = ws.Range(WP_HALFWIDTH_ADDR).Value2
    On Error GoTo 0
    If IsArray(v) Then
        WaypointHalfWidth = 0#
    ElseIf IsNumeric(v) And CDbl(v) > 0# Then
        WaypointHalfWidth = CDbl(v)
    Else
        WaypointHalfWidth = 0#
    End If
End Function

Private Function IsSurveySummaryRow(ws As Worksheet, ByVal r As Long) As Boolean
    Dim fE As String, fF As String, fd As String
    On Error Resume Next
    fE = UCase$(CStr(ws.Cells(r, 5).Formula & ""))
    fF = UCase$(CStr(ws.Cells(r, 6).Formula & ""))
    fd = UCase$(CStr(ws.Cells(r, 4).Formula & ""))
    On Error GoTo 0
    IsSurveySummaryRow = (InStr(1, fE, "LOOKUP", vbBinaryCompare) > 0) _
                      Or (InStr(1, fF, "LOOKUP", vbBinaryCompare) > 0) _
                      Or (InStr(1, fd, "LOOKUP", vbBinaryCompare) > 0)
End Function

Private Function ActualAtLastSurvey(ws As Worksheet, _
        ByRef sMD As Double, ByRef sInc As Double, ByRef sAzi As Double, _
        ByRef sn As Double, ByRef sE As Double, ByRef sV As Double, _
        ByRef lastRow As Long) As Long
    Dim n As Long: n = 0
    Dim prevMD As Double, prevInc As Double, prevAzi As Double
    Dim curN As Double, curE As Double, curV As Double
    prevMD = 0#: prevInc = 0#: prevAzi = 0#
    curN = 0#: curE = 0#: curV = 0#

    Dim r As Long
    For r = SURV_ROW_FIRST To SURV_ROW_LAST
        If IsSurveySummaryRow(ws, r) Then GoTo NextSurveyRow

        Dim vMd As Variant, vInc As Variant, vAzi As Variant
        vMd = ws.Cells(r, 5).Value2
        vInc = ws.Cells(r, 6).Value2
        vAzi = ws.Cells(r, 7).Value2

        If Not (IsNumeric(vMd) And IsNumeric(vInc) And IsNumeric(vAzi) _
                And Len(CStr(vMd & "")) > 0 And Len(CStr(vInc & "")) > 0 _
                And Len(CStr(vAzi & "")) > 0) Then
            GoTo NextSurveyRow
        End If

        Dim mdv As Double: mdv = CDbl(vMd)
        If mdv <= prevMD Then GoTo NextSurveyRow

        Dim dv As Double, dN As Double, dE As Double
        McStep prevMD, prevInc, prevAzi, mdv, CDbl(vInc), CDbl(vAzi), dv, dN, dE
        curN = curN + dN: curE = curE + dE: curV = curV + dv
        prevMD = mdv: prevInc = CDbl(vInc): prevAzi = CDbl(vAzi)
        lastRow = r
        n = n + 1
NextSurveyRow:
    Next r

    sMD = prevMD: sInc = prevInc: sAzi = prevAzi
    sn = curN: sE = curE: sV = curV
    ActualAtLastSurvey = n
End Function

Private Sub FrameComponents(ByVal gravityMode As Boolean, ByVal planAzi As Double, _
        ByVal dN As Double, ByVal dE As Double, ByVal dTvdUp As Double, _
        ByRef x As Double, ByRef y As Double)
    Dim right As Double, along As Double
    right = dE * Cos(Deg2Rad(planAzi)) - dN * Sin(Deg2Rad(planAzi))
    along = dN * Cos(Deg2Rad(planAzi)) + dE * Sin(Deg2Rad(planAzi))
    If gravityMode Then
        x = right
        y = dTvdUp
    Else
        x = right
        y = along
    End If
End Sub

Private Function ActualTvdForGauge(ws As Worksheet, ByVal lastRow As Long, _
                                   ByVal integratedTvd As Double) As Double
    Dim vH As Variant
    On Error Resume Next
    vH = ws.Cells(lastRow, 8).Value2
    On Error GoTo 0
    If IsNumeric(vH) And Len(CStr(vH & "")) > 0 Then
        ActualTvdForGauge = CDbl(vH)
    Else
        ActualTvdForGauge = integratedTvd
    End If
End Function

' --------------------------------------------------------------------------------
'  MAIN ENTRY
' --------------------------------------------------------------------------------
Public Sub RenderPlanGauge()
    On Error Resume Next
    RenderPlanGaugeCore
    On Error GoTo 0
End Sub

Public Sub RenderPlanGaugeStrict()
    RenderPlanGaugeCore
End Sub

Private Sub RenderPlanGaugeCore()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SS_SHEET)
    Dim pMD() As Double, pInc() As Double, pAzi() As Double
    Dim pTvd() As Double, pNS() As Double, pEW() As Double
    Dim nPlan As Long
    nPlan = LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
    Dim sMD As Double, sInc As Double, sAzi As Double
    Dim sn As Double, sE As Double, sV As Double
    Dim lastRow As Long, nSurv As Long
    nSurv = ActualAtLastSurvey(ws, sMD, sInc, sAzi, sn, sE, sV, lastRow)

    If nPlan < 2 Or nSurv < 1 Then
        DrawGauge ws, False, 0, 0, "", False, 0, 0, _
                  "NO DATA - Import Plan / enter surveys", False, 0, 0, 1#
        Exit Sub
    End If
    Dim actTvd As Double
    actTvd = ActualTvdForGauge(ws, lastRow, sV)

    Dim gravityMode As Boolean: gravityMode = (sInc >= CROSSOVER_INC)

    Dim plN As Double, plE As Double, plV As Double, plAzi As Double, plInc As Double
    If gravityMode Then
        PlanAt sMD, nPlan, pMD, pInc, pAzi, pTvd, pNS, pEW, plN, plE, plV, plAzi, plInc
    Else
        PlanAtTvd actTvd, nPlan, pMD, pInc, pAzi, pTvd, pNS, pEW, plN, plE, plV, plAzi, plInc
    End If

    Dim dN As Double, dE As Double, dTvdUp As Double
    dN = sn - plN
    dE = sE - plE
    dTvdUp = plV - actTvd

    Dim ax As Double, ay As Double
    FrameComponents gravityMode, plAzi, dN, dE, dTvdUp, ax, ay
    Dim showPtb As Boolean: showPtb = False
    Dim bx As Double, by As Double
    Dim tarStart As Variant: tarStart = ws.Range("U2").Value2
    Dim bitMD As Variant: bitMD = ws.Cells(lastRow, 4).Value2
    If IsNumeric(tarStart) And IsNumeric(bitMD) Then
        If sMD >= CDbl(tarStart) And CDbl(bitMD) > sMD Then
            Dim incB As Double, azB As Double
            Dim vW As Variant: vW = ws.Cells(lastRow, 23).Value2
            Dim vX As Variant: vX = ws.Cells(lastRow, 24).Value2
            If IsNumeric(vW) And Len(CStr(vW & "")) > 0 Then incB = CDbl(vW) Else incB = sInc
            If IsNumeric(vX) And Len(CStr(vX & "")) > 0 Then azB = CDbl(vX) Else azB = sAzi

            Dim stepV As Double, stepN As Double, stepE As Double
            McStep sMD, sInc, sAzi, CDbl(bitMD), incB, azB, stepV, stepN, stepE

            Dim bitTvd As Double: bitTvd = actTvd + stepV
            Dim pbN As Double, pbE As Double, pbV As Double, pbAzi As Double, pbInc As Double
            If gravityMode Then
                PlanAt CDbl(bitMD), nPlan, pMD, pInc, pAzi, pTvd, pNS, pEW, pbN, pbE, pbV, pbAzi, pbInc
            Else
                PlanAtTvd bitTvd, nPlan, pMD, pInc, pAzi, pTvd, pNS, pEW, pbN, pbE, pbV, pbAzi, pbInc
            End If

            Dim bDN As Double, bDE As Double, bUp As Double
            bDN = (sn + stepN) - pbN
            bDE = (sE + stepE) - pbE
            bUp = pbV - bitTvd
            FrameComponents gravityMode, pbAzi, bDN, bDE, bUp, bx, by
            showPtb = True
        End If
    End If
    Dim halfW As Double: halfW = WaypointHalfWidth(ws)
    Dim hasBand As Boolean: hasBand = False
    Dim yTop As Double, yBot As Double
    Dim scaleS As Double: scaleS = 0#
    Dim wpTvd As Double

    If gravityMode And halfW > 0# Then
        If WaypointTvdAtMd(ws, sMD, wpTvd) Then
            Dim yWp As Double
            yWp = plV - wpTvd
            yTop = yWp + halfW
            yBot = yWp - halfW
            hasBand = True
            ' Fit corridor in the circle (band may sit off plan centre).
            scaleS = halfW
            If Abs(yTop) > scaleS Then scaleS = Abs(yTop)
            If Abs(yBot) > scaleS Then scaleS = Abs(yBot)
            ' Do NOT pin band edges to the rim — that erased the red OUT caps and
            ' made a top-of-window breach look green. Leave margin so green is
            ' always surrounded by red (above + below) inside the dial.
            scaleS = scaleS * 1.35
            If scaleS < 0.5 Then scaleS = 0.5
        End If
    End If

    If Not hasBand Then
        scaleS = halfW
        If scaleS <= 0# Then scaleS = 1#
        Dim magA0 As Double, magB0 As Double
        magA0 = Sqr(ax * ax + ay * ay)
        magB0 = Sqr(bx * bx + by * by)
        If magA0 > scaleS Then scaleS = magA0
        If showPtb And magB0 > scaleS Then scaleS = magB0
        If scaleS < 0.5 Then scaleS = 0.5
    End If

    DrawGauge ws, True, ax, ay, IIf(gravityMode, "GRAV", "MAG"), showPtb, bx, by, _
              "", hasBand, yTop, yBot, scaleS

    ' Data "Position of Wellbore": LT/RT + UP/DN from plan; geo = hypot until lateral
    On Error Resume Next
    UpdateWellborePosition ax, True, ay, True, lastRow, sMD, sInc
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------------
'  Rendering
' --------------------------------------------------------------------------------
Private Sub DrawGauge(ws As Worksheet, ByVal hasData As Boolean, _
        ByVal ax As Double, ByVal ay As Double, ByVal modeTxt As String, _
        ByVal showPtb As Boolean, ByVal bx As Double, ByVal by As Double, _
        ByVal caption As String, ByVal hasBand As Boolean, _
        ByVal yTop As Double, ByVal yBot As Double, ByVal scaleS As Double)
    Dim wasProt As Boolean
    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo Reprotect
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).name, Len(SHP_PREFIX)) = SHP_PREFIX Then ws.Shapes(i).Delete
    Next i
    Dim area As Range: Set area = ws.Range(GAUGE_RANGE)
    Dim l As Double, t As Double, w As Double, H As Double
    l = area.Left: t = area.Top: w = area.Width: H = area.Height

    ' Dial hugs the left edge (AA into AB) and grows with the AA1:AC7 height
    ' (rows 1-7 at 15 pt). Tight metrics column sits to its right, ending at AC.
    Dim textW As Double: textW = 72#
    Dim gap As Double: gap = 3#
    Dim r As Double
    r = (H - 4#) / 2#
    Dim dialLeft As Double: dialLeft = l + 1#
    If dialLeft + 2# * r > l + w - textW - gap - 1# Then _
        r = (l + w - textW - gap - 1# - dialLeft) / 2#
    If r < 18# Then r = 18#

    Dim cx As Double, cy As Double
    cx = dialLeft + r
    cy = t + H / 2#

    Dim textX As Double: textX = l + w - textW - 1#
    If textX < dialLeft + 2# * r + gap Then textX = dialLeft + 2# * r + gap

    Dim boxL As Double, boxT As Double, boxR As Double, boxB As Double
    boxL = l + 1#
    boxT = t + 1#
    boxR = textX - 1#
    If boxR < dialLeft + 2# * r + 1# Then boxR = dialLeft + 2# * r + 1#
    boxB = t + H - 1#

    If Not hasData Then
        DrawCombinedDial ws, cx, cy, r, modeTxt, scaleS, False, 0, 0, False, 0, 0, _
                         False, 0, 0, boxL, boxT, boxR, boxB
        AddGaugeText ws, SHP_PREFIX & "Note", textX, t + 2, textW, FS_HDR + 4, _
                     caption, cInk(), FS_HDR, True
        GoTo Reprotect
    End If

    Dim magA As Double: magA = Sqr(ax * ax + ay * ay)
    Dim magB As Double: magB = Sqr(bx * bx + by * by)
    DrawCombinedDial ws, cx, cy, r, modeTxt, scaleS, True, ax, ay, showPtb, bx, by, _
                     hasBand, yTop, yBot, boxL, boxT, boxR, boxB

    Dim rowH As Double: rowH = FS_VAL + 1
    Dim y0 As Double: y0 = t + 1
    Dim aOut As Boolean, bOut As Boolean
    aOut = hasBand And (ay > yTop Or ay < yBot)
    bOut = hasBand And showPtb And (by > yTop Or by < yBot)

    AddGaugeText ws, SHP_PREFIX & "Mode", textX, y0, textW, FS_HDR + 1, _
                 IIf(modeTxt = "GRAV", "GRAVITY", "MAGNETIC"), cInk(), FS_HDR, True
    AddGaugeText ws, SHP_PREFIX & "RdY", textX, y0 + rowH, textW, FS_VAL + 1, _
                 AxisLabel(modeTxt, True, ay) & IIf(aOut, " OUT", ""), _
                 IIf(aOut, cBandOutInk(), cAct()), FS_VAL, False
    AddGaugeText ws, SHP_PREFIX & "RdX", textX, y0 + 2 * rowH, textW, FS_VAL + 1, _
                 AxisLabel(modeTxt, False, ax), cAct(), FS_VAL, False
    AddGaugeText ws, SHP_PREFIX & "RdT", textX, y0 + 3 * rowH, textW, FS_VAL + 1, _
                 IIf(modeTxt = "GRAV", "TOT", "PRP") & Space(1) & Format(magA, "0.00") & "m", _
                 cInk(), FS_VAL, False

    Dim yPtb As Double: yPtb = y0 + 4 * rowH + 2
    If showPtb Then
        AddGaugeText ws, SHP_PREFIX & "PtbHdr", textX, yPtb, textW, FS_HDR + 1, _
                     "PTB", cPtb(), FS_HDR, True
        AddGaugeText ws, SHP_PREFIX & "PtbY", textX, yPtb + rowH, textW, FS_VAL + 1, _
                     AxisLabel(modeTxt, True, by) & IIf(bOut, " OUT", ""), _
                     IIf(bOut, cBandOutInk(), cPtb()), FS_VAL, False
        AddGaugeText ws, SHP_PREFIX & "PtbX", textX, yPtb + 2 * rowH, textW, FS_VAL + 1, _
                     AxisLabel(modeTxt, False, bx), cPtb(), FS_VAL, False
        AddGaugeText ws, SHP_PREFIX & "PtbT", textX, yPtb + 3 * rowH, textW, FS_VAL + 1, _
                     IIf(modeTxt = "GRAV", "TOT", "PRP") & Space(1) & Format(magB, "0.00") & "m", _
                     cInk(), FS_VAL, False
    Else
        AddGaugeText ws, SHP_PREFIX & "PtbHdr", textX, yPtb, textW, FS_HDR + 1, _
                     "PTB", cGrid(), FS_HDR, True
        AddGaugeText ws, SHP_PREFIX & "PtbY", textX, yPtb + rowH, textW + 10, FS_VAL + 1, _
                     "off target", cGrid(), FS_VAL, False
    End If

Reprotect:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "DrawGauge", errDesc
End Sub

Private Sub DrawCombinedDial(ws As Worksheet, _
        ByVal cx As Double, ByVal cy As Double, ByVal r As Double, _
        ByVal modeTxt As String, ByVal scaleS As Double, _
        ByVal live As Boolean, ByVal ax As Double, ByVal ay As Double, _
        ByVal showPtb As Boolean, ByVal bx As Double, ByVal by As Double, _
        ByVal hasBand As Boolean, ByVal yTop As Double, ByVal yBot As Double, _
        ByVal boxL As Double, ByVal boxT As Double, ByVal boxR As Double, ByVal boxB As Double)

    Dim shp As Shape

    ' Bands first (behind): full-width horizontal rectangles that extend past
    ' the circle so a marker outside the rim can still sit in green.
    If hasBand And scaleS > 0# Then
        On Error Resume Next
        DrawWaypointBand ws, cx, cy, r, scaleS, yTop, yBot, boxL, boxT, boxR, boxB
        On Error GoTo 0
    End If

    Set shp = ws.Shapes.AddShape(msoShapeOval, cx - r, cy - r, 2# * r, 2# * r)
    StyleGaugeShape shp, SHP_PREFIX & "Circle"
    shp.Fill.Visible = msoFalse
    shp.line.ForeColor.RGB = cAxis(): shp.line.Weight = 1.75

    Set shp = ws.Shapes.AddLine(cx - r, cy, cx + r, cy)
    StyleGaugeShape shp, SHP_PREFIX & "AxX"
    shp.line.ForeColor.RGB = cAxis(): shp.line.Weight = 0.75: shp.line.DashStyle = msoLineDash
    Set shp = ws.Shapes.AddLine(cx, cy - r, cx, cy + r)
    StyleGaugeShape shp, SHP_PREFIX & "AxY"
    shp.line.ForeColor.RGB = cAxis(): shp.line.Weight = 0.75: shp.line.DashStyle = msoLineDash

    Dim tU As String, tD As String, tLf As String, tRt As String
    If modeTxt = "GRAV" Then
        tU = "U": tD = "D": tLf = "L": tRt = "R"
    Else
        tU = "A": tD = "B": tLf = "L": tRt = "R"
    End If
    Dim aw As Double: aw = FS_AXIS + 4
    AddGaugeText ws, SHP_PREFIX & "LblU", cx - aw / 2#, cy - r - 1, aw, aw, tU, cAxis(), FS_AXIS, True
    AddGaugeText ws, SHP_PREFIX & "LblD", cx - aw / 2#, cy + r - aw + 2, aw, aw, tD, cAxis(), FS_AXIS, True
    AddGaugeText ws, SHP_PREFIX & "LblL", cx - r + 1, cy - aw / 2#, aw, aw, tLf, cAxis(), FS_AXIS, True
    AddGaugeText ws, SHP_PREFIX & "LblR", cx + r - aw + 1, cy - aw / 2#, aw, aw, tRt, cAxis(), FS_AXIS, True

    Set shp = ws.Shapes.AddShape(msoShapeOval, cx - 3, cy - 3, 6, 6)
    StyleGaugeShape shp, SHP_PREFIX & "Plan"
    shp.Fill.ForeColor.RGB = cPlan(): shp.line.Visible = msoFalse

    If Not live Then Exit Sub
    If scaleS <= 0# Then scaleS = 1#

    If showPtb Then
        DrawMarker ws, "B", cx, cy, r, scaleS, bx, by, cPtb(), True, _
                   boxL, boxT, boxR, boxB
    End If
    DrawMarker ws, "A", cx, cy, r, scaleS, ax, ay, cAct(), False, _
               boxL, boxT, boxR, boxB
End Sub

' Horizontal waypoint corridor clipped to the circle's bounding square so the
' green/red fills stay aligned with the dial (not the wider AA1:AC7 cell).
' Green = inside [yBot, yTop]; red = above/below. Markers may still sit outside.
Private Sub DrawWaypointBand(ws As Worksheet, ByVal cx As Double, ByVal cy As Double, _
        ByVal r As Double, ByVal scaleS As Double, _
        ByVal yTop As Double, ByVal yBot As Double, _
        ByVal boxL As Double, ByVal boxT As Double, ByVal boxR As Double, ByVal boxB As Double)

    Dim yHi As Double, yLo As Double
    yHi = yTop: yLo = yBot
    If yHi < yLo Then
        Dim tmp As Double: tmp = yHi: yHi = yLo: yLo = tmp
    End If

    ' Exact circle bounds (ignore outer text/box — that was the visual misalignment).
    Dim xL As Double, xR As Double
    Dim topEdge As Double, botEdge As Double
    xL = cx - r
    xR = cx + r
    topEdge = cy - r
    botEdge = cy + r
    Dim stripW As Double: stripW = xR - xL
    If stripW < 4# Then Exit Sub

    Dim syHi As Double, syLo As Double
    syHi = cy - r * (yHi / scaleS)
    syLo = cy - r * (yLo / scaleS)

    ' Visible slice of the corridor inside the circle square (do not expand
    ' green to the rim when the true edge is outside — that hid the red caps).
    Dim gTop As Double, gBot As Double
    gTop = syHi
    gBot = syLo
    If gTop < topEdge Then gTop = topEdge
    If gTop > botEdge Then gTop = botEdge
    If gBot < topEdge Then gBot = topEdge
    If gBot > botEdge Then gBot = botEdge

    ' Red ABOVE window (shallower / toward U) — always when any circle area is above green
    If gTop > topEdge + 1# Then
        AddBandRect ws, "OutU", xL, topEdge, stripW, gTop - topEdge, cBandOut(), 0.35
    End If

    ' Green = inside geo window
    If Abs(gBot - gTop) > 1# Then
        AddBandRect ws, "BandIn", xL, gTop, stripW, gBot - gTop, cBandIn(), 0.35
    End If

    ' Red BELOW window (deeper / toward D)
    If botEdge > gBot + 1# Then
        AddBandRect ws, "OutD", xL, gBot, stripW, botEdge - gBot, cBandOut(), 0.35
    End If

    ' Bound lines at true corridor edges when they fall inside the circle
    If syHi > topEdge + 0.5 And syHi < botEdge - 0.5 Then
        AddBandHLine ws, "BandTop", xL, xR, syHi, cBandEdge()
    End If
    If syLo > topEdge + 0.5 And syLo < botEdge - 0.5 Then
        AddBandHLine ws, "BandBot", xL, xR, syLo, cBandEdge()
    End If
End Sub

Private Sub AddBandRect(ws As Worksheet, ByVal tag As String, _
        ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal H As Double, _
        ByVal clr As Long, ByVal transparency As Double)
    If w < 1# Or H < 1# Then Exit Sub
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, H)
    StyleGaugeShape shp, SHP_PREFIX & tag
    shp.Fill.ForeColor.RGB = clr
    shp.Fill.transparency = transparency
    shp.line.Visible = msoFalse
End Sub

Private Sub AddBandHLine(ws As Worksheet, ByVal tag As String, _
        ByVal x1 As Double, ByVal x2 As Double, ByVal y As Double, ByVal clr As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddLine(x1, y, x2, y)
    StyleGaugeShape shp, SHP_PREFIX & tag
    shp.line.ForeColor.RGB = clr
    shp.line.Weight = 1.5
End Sub

Private Sub DrawMarker(ws As Worksheet, ByVal tag As String, _
        ByVal cx As Double, ByVal cy As Double, ByVal r As Double, _
        ByVal scaleS As Double, ByVal vX As Double, ByVal vY As Double, _
        ByVal clr As Long, ByVal hollow As Boolean, _
        ByVal boxL As Double, ByVal boxT As Double, ByVal boxR As Double, ByVal boxB As Double)

    Dim mag As Double: mag = Sqr(vX * vX + vY * vY)
    Dim shp As Shape

    If mag <= 0.005 Then
        Set shp = ws.Shapes.AddShape(msoShapeOval, cx - 6, cy - 6, 12, 12)
        StyleGaugeShape shp, SHP_PREFIX & "Dot" & tag
        shp.Fill.Visible = msoFalse
        shp.line.ForeColor.RGB = clr: shp.line.Weight = 2#
        Exit Sub
    End If

    Dim dx As Double, dy As Double
    dx = cx + r * (vX / scaleS)
    dy = cy - r * (vY / scaleS)

    ' Clamp to gauge box so markers remain visible
    Dim margin As Double: margin = 6#
    dx = ClampDbl(dx, boxL + margin, boxR - margin)
    dy = ClampDbl(dy, boxT + margin, boxB - margin)

    Set shp = ws.Shapes.AddLine(cx, cy, dx, dy)
    StyleGaugeShape shp, SHP_PREFIX & "Vec" & tag
    shp.line.ForeColor.RGB = clr
    If hollow Then
        shp.line.Weight = 1.5
        shp.line.DashStyle = msoLineDash
    Else
        shp.line.Weight = 2#
    End If
    shp.line.EndArrowheadStyle = msoArrowheadTriangle
    shp.line.EndArrowheadLength = msoArrowheadShort
    shp.line.EndArrowheadWidth = msoArrowheadNarrow

    Set shp = ws.Shapes.AddShape(msoShapeOval, dx - 4.5, dy - 4.5, 9, 9)
    StyleGaugeShape shp, SHP_PREFIX & "Dot" & tag
    If hollow Then
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.line.ForeColor.RGB = clr: shp.line.Weight = 2#
    Else
        shp.Fill.ForeColor.RGB = clr
        shp.line.Visible = msoFalse
    End If
End Sub

Private Function AxisLabel(ByVal modeTxt As String, ByVal isVertical As Boolean, _
                           ByVal v As Double) As String
    Dim tag As String
    If modeTxt = "GRAV" Then
        If isVertical Then
            tag = IIf(v >= 0, "UP", "DN")
        Else
            tag = IIf(v >= 0, "RT", "LT")
        End If
    Else
        If isVertical Then
            tag = IIf(v >= 0, "AH", "BH")
        Else
            tag = IIf(v >= 0, "RT", "LT")
        End If
    End If
    AxisLabel = Left$(tag & "   ", 3) & Format(Abs(v), "0.00") & "m"
End Function

Private Sub StyleGaugeShape(shp As Shape, ByVal nm As String)
    shp.name = nm
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    shp.Shadow.Visible = msoFalse
    On Error GoTo 0
End Sub

Private Sub AddGaugeText(ws As Worksheet, ByVal nm As String, _
        ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal H As Double, _
        ByVal txt As String, ByVal clr As Long, ByVal sz As Single, ByVal bold As Boolean)
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, H)
    StyleGaugeShape shp, nm
    shp.Fill.Visible = msoFalse
    shp.line.Visible = msoFalse
    With shp.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        With .TextRange
            .text = txt
            .Font.name = "Consolas"
            .Font.Size = sz
            .Font.bold = IIf(bold, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = clr
        End With
    End With
End Sub

' ================================================================================
'  PUBLIC SHIMS
'
'  MDL_CorridorImage draws the same well against the same plan as the gauge above,
'  so it must use the same maths. Private members are only visible inside their own
'  module, so the alternative to these shims is a second copy of the projection
'  code — and two copies eventually disagree, which would put the emailed picture
'  and the on-screen gauge at odds with no way to tell which is wrong.
'
'  Everything below is additive. No procedure above this line was altered.
' ================================================================================

' Survey block extent and the inclination at which the gauge swaps reference
' frames, republished so callers scan exactly the rows the gauge scans.
Public Property Get PG_SurvRowFirst() As Long
    PG_SurvRowFirst = SURV_ROW_FIRST
End Property

Public Property Get PG_SurvRowLast() As Long
    PG_SurvRowLast = SURV_ROW_LAST
End Property

Public Property Get PG_CrossoverInc() As Double
    PG_CrossoverInc = CROSSOVER_INC
End Property

Public Property Get PG_SlidesheetName() As String
    PG_SlidesheetName = SS_SHEET
End Property

Public Function PG_LoadPlan(ByRef pMD() As Double, ByRef pInc() As Double, _
                            ByRef pAzi() As Double, ByRef pTvd() As Double, _
                            ByRef pNS() As Double, ByRef pEW() As Double) As Long
    PG_LoadPlan = LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
End Function

Public Sub PG_PlanAt(ByVal md As Double, ByVal n As Long, _
                     pMD() As Double, pInc() As Double, pAzi() As Double, _
                     pTvd() As Double, pNS() As Double, pEW() As Double, _
                     ByRef outN As Double, ByRef outE As Double, ByRef outV As Double, _
                     ByRef outAzi As Double, ByRef outInc As Double)
    PlanAt md, n, pMD, pInc, pAzi, pTvd, pNS, pEW, outN, outE, outV, outAzi, outInc
End Sub

Public Sub PG_McStep(ByVal md1 As Double, ByVal i1 As Double, ByVal a1 As Double, _
                     ByVal md2 As Double, ByVal i2 As Double, ByVal a2 As Double, _
                     ByRef dv As Double, ByRef dN As Double, ByRef dE As Double)
    McStep md1, i1, a1, md2, i2, a2, dv, dN, dE
End Sub

Public Sub PG_FrameComponents(ByVal gravityMode As Boolean, ByVal planAzi As Double, _
                              ByVal dN As Double, ByVal dE As Double, ByVal dTvdUp As Double, _
                              ByRef x As Double, ByRef y As Double)
    FrameComponents gravityMode, planAzi, dN, dE, dTvdUp, x, y
End Sub

Public Function PG_WaypointTvdAtMd(ws As Worksheet, ByVal md As Double, _
                                   ByRef outTvd As Double) As Boolean
    PG_WaypointTvdAtMd = WaypointTvdAtMd(ws, md, outTvd)
End Function

Public Function PG_WaypointHalfWidth(ws As Worksheet) As Double
    PG_WaypointHalfWidth = WaypointHalfWidth(ws)
End Function

Public Function PG_IsSurveySummaryRow(ws As Worksheet, ByVal r As Long) As Boolean
    PG_IsSurveySummaryRow = IsSurveySummaryRow(ws, r)
End Function


