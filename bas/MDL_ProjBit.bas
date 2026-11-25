Attribute VB_Name = "MDL_ProjBit"
Option Explicit

' ================================================================================
'  MDL_ProjBit — Bit projection UDFs (PROJBIT min-curvature / toolface math)
'  Used by Slidesheet V/W/X/AS (TVD @ BIT, INC @ BIT, AZM @ BIT, Meters To Slide)
' ================================================================================

Private Const PI_ As Double = 3.14159265358979
Private Const EPS As Double = 0.0000001
' Skip near plan stations that would spike BURR (tiny remaining MD/TVD).
Private Const MIN_AIM_MD As Double = 10#
Private Const MIN_AIM_TVD As Double = 5#
Private Const MIN_AIM_DINC As Double = 0.5
' Catch-up to the plan one course ahead (see CatchUpDls).
Private Const CATCHUP_PLAN_SHEET As String = "_OC_Survey"
Private Const CATCHUP_CURVE_MIN As Double = 1#
Private Const CATCHUP_LATERAL_INC As Double = 80#
Private Const CATCHUP_TVD_STANDS As Double = 3#
Private Const CATCHUP_TVD_MAX_DEG As Double = 3#
' A recommended slide shorter than this is rotate. Showing it just clutters Y.
Private Const SLIDE_IGNORE_M As Double = 1#

' Plan-point aim: step and reach for moving the aim point down the plan, and
' the cached plan path (PG_PlanAt) with each course's DLS.
Private Const AIM_WALK_STEP_M As Double = 3#
Private Const AIM_WALK_MAX_M As Double = 90#
Private m_planKey As String
Private m_planN As Long
Private m_pMD() As Double, m_pInc() As Double, m_pAzi() As Double
Private m_pTvd() As Double, m_pNS() As Double, m_pEW() As Double
Private m_pDls() As Double

' --- angle helpers -------------------------------------------------------------

Private Function Deg2Rad(ByVal d As Double) As Double
    Deg2Rad = d * PI_ / 180#
End Function

Private Function Rad2Deg(ByVal r As Double) As Double
    Rad2Deg = r * 180# / PI_
End Function

Private Function Wrap180(ByVal x As Double) As Double
    x = x - 360# * Int(x / 360#)
    If x > 180# Then x = x - 360#
    Wrap180 = x
End Function

Private Function Wrap360(ByVal x As Double) As Double
    x = x - 360# * Int(x / 360#)
    If x < 0# Then x = x + 360#
    Wrap360 = x
End Function

Private Function DoglegAngleDeg(ByVal i1 As Double, ByVal a1 As Double, _
                                ByVal i2 As Double, ByVal a2 As Double) As Double
    Dim r1 As Double, r2 As Double, dA As Double, c As Double
    r1 = Deg2Rad(i1)
    r2 = Deg2Rad(i2)
    dA = Deg2Rad(Wrap180(a2 - a1))
    c = Cos(r1) * Cos(r2) + Sin(r1) * Sin(r2) * Cos(dA)
    If c > 1# Then c = 1#
    If c < -1# Then c = -1#
    DoglegAngleDeg = Rad2Deg(Application.WorksheetFunction.Acos(c))
End Function

Private Function ToolfaceBetweenDeg(ByVal i1 As Double, ByVal a1 As Double, _
                                    ByVal i2 As Double, ByVal a2 As Double) As Double
    Dim r1 As Double, r2 As Double, dA As Double
    r1 = Deg2Rad(i1)
    r2 = Deg2Rad(i2)
    dA = Deg2Rad(Wrap180(a2 - a1))
    ' Excel Atan2(x, y); Python atan2(y, x)
    ToolfaceBetweenDeg = Rad2Deg(Application.WorksheetFunction.Atan2( _
        Sin(r2) * Cos(r1) * Cos(dA) - Sin(r1) * Cos(r2), _
        Sin(r2) * Sin(dA)))
End Function

' ================================================================================
'  PLAN-POINT AIM
'
'  Every required toolface / DLR / slide length (Slidesheet AR:AT, Y, Cont DI
'  V1:V4) comes from one curve: a circular arc from the bit (position AND
'  attitude) to a point on the plan path. The point is the plan position at
'  the next named target. When reaching it needs more than the plan's own
'  curve rate (or the rolling motor yield on a hold), the point moves down
'  the plan until it does not.
'
'  An attitude-only aim (match the target's Inc/Azm) ignores where the bit
'  is: 3.4 m left of plan it said L44, which holds the miss. The arc to the
'  plan point carries both the offset and the attitude difference.
'
'  Bit position = sheet cumulative N/E (AL/AM) + min-curve to the bit; plan
'  position = the gauge's min-curvature interpolation (PG_PlanAt), so the
'  aim and the proximity gauge use the same points in space.
' ================================================================================

' The plan path, reloaded only when the plan sheet changes.
Private Function PlanPathCount() As Long
    Dim ws As Worksheet
    Dim lastR As Long
    Dim key As String
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(CATCHUP_PLAN_SHEET)
    lastR = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    key = CStr(lastR) & "|" & CStr(ws.Cells(lastR, 1).Value2) & "|" & _
          CStr(ws.Cells(lastR, 5).Value2) & "|" & CStr(ws.Cells(lastR, 6).Value2) & "|" & _
          CStr(ws.Cells(3, 1).Value2) & "|" & CStr(ws.Cells(3, 2).Value2)
    If key <> m_planKey Or m_planN < 2 Then
        m_planN = MDL_PlanGauge.PG_LoadPlan(m_pMD, m_pInc, m_pAzi, m_pTvd, m_pNS, m_pEW)
        m_planKey = key
        If m_planN >= 2 Then
            Dim j As Long
            ReDim m_pDls(0 To m_planN - 1)
            For j = 0 To m_planN - 2
                If m_pMD(j + 1) - m_pMD(j) > EPS Then
                    m_pDls(j) = DoglegAngleDeg(m_pInc(j), m_pAzi(j), m_pInc(j + 1), m_pAzi(j + 1)) _
                                * 30# / (m_pMD(j + 1) - m_pMD(j))
                End If
            Next j
        End If
    End If
    PlanPathCount = m_planN
    Exit Function
Fail:
    PlanPathCount = 0
End Function

' Arc from attitude (ib, aB) through the point (dN, dE, dV from the bit,
' dV positive deeper). tfDeg: 0 high side, + right. dls: deg/30 m.
Private Function ArcToPoint(ByVal ib As Double, ByVal aB As Double, _
                            ByVal dN As Double, ByVal dE As Double, ByVal dV As Double, _
                            ByRef tfDeg As Double, ByRef dls As Double) As Boolean
    Dim ri As Double, ra As Double
    Dim hs As Double, rt As Double, al As Double
    Dim side As Double, chord As Double
    Dim theta As Double, arcLen As Double
    ArcToPoint = False
    ri = Deg2Rad(ib)
    ra = Deg2Rad(aB)
    hs = dN * Cos(ri) * Cos(ra) + dE * Cos(ri) * Sin(ra) - dV * Sin(ri)
    rt = -dN * Sin(ra) + dE * Cos(ra)
    al = dN * Sin(ri) * Cos(ra) + dE * Sin(ri) * Sin(ra) + dV * Cos(ri)
    chord = Sqr(dN * dN + dE * dE + dV * dV)
    If chord <= EPS Then Exit Function
    side = Sqr(hs * hs + rt * rt)
    If side <= 0.000001 Then
        tfDeg = 0#
        dls = 0#
        ArcToPoint = True
        Exit Function
    End If
    theta = Application.WorksheetFunction.Atan2(al, side)
    arcLen = chord * theta / Sin(theta)
    If arcLen <= EPS Then Exit Function
    tfDeg = Rad2Deg(Application.WorksheetFunction.Atan2(hs, rt))
    dls = Rad2Deg(2# * theta) * 30# / arcLen
    ArcToPoint = True
End Function

Private Function AimPlanPoint(ByVal bm As Double, ByVal ib As Double, ByVal aB As Double, _
                              ByVal bN As Double, ByVal bE As Double, ByVal bV As Double, _
                              ByVal motorOut As Double, _
                              ByVal tgtMd As Range, ByVal tgtInc As Range, _
                              ByVal tgtAzm As Range, ByVal tgtTvd As Range, _
                              ByRef tfDeg As Double, ByRef dls As Double, _
                              ByRef aimMd As Double, _
                              Optional ByVal fixedTgtMd As Variant) As Boolean
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long, k As Long, np As Long
    Dim q As Double, endMd As Double
    Dim ptN As Double, ptE As Double, ptV As Double, ptAz As Double, ptInc As Double
    Dim kt As Long

    AimPlanPoint = False
    n = LoadTargets(tgtMd, tgtInc, tgtAzm, tgtTvd, m, i, a, tv)
    If HasNum(fixedTgtMd) Then
        ' Target already chosen (Cont DI follows the as-drilled survey). A bit
        ' that has run past it aims just ahead of itself on the plan.
        q = CDbl(fixedTgtMd)
        If q < bm + AIM_WALK_STEP_M Then q = bm + AIM_WALK_STEP_M
    Else
        If n < 1 Then Exit Function
        k = FirstAimIndex(m, i, tv, n, bm, ib, bV)
        If k > n Then Exit Function
        q = m(k)
    End If
    np = PlanPathCount()
    If np < 2 Then Exit Function
    endMd = m_pMD(np - 1)

    ' Limit for the arc: the plan's own DLS once the stretch to the aim point
    ' includes a curve (same rule as the plan-rate slide), else the motor.
    ' A close target (or one just past a hold) moves the aim down the plan
    ' until the arc is within that limit, at most AIM_WALK_MAX_M past it.
    Dim jj As Long, planMax As Double, cap As Double, walkEnd As Double
    jj = 0
    Do While jj < np - 2 And m_pMD(jj + 1) <= bm
        jj = jj + 1
    Loop
    planMax = 0#
    If q > endMd Then q = endMd
    walkEnd = q + AIM_WALK_MAX_M
    If walkEnd > endMd Then walkEnd = endMd
    Do
        MDL_PlanGauge.PG_PlanAt q, np, m_pMD, m_pInc, m_pAzi, m_pTvd, m_pNS, m_pEW, _
                                ptN, ptE, ptV, ptAz, ptInc
        ' A saved target TVD (the Targets form) replaces the plan TVD at that
        ' station. Further down the plan the path's own TVD is used.
        For kt = 1 To n
            If Abs(m(kt) - q) < 0.05 Then
                If Abs(tv(kt) - ptV) > 0.5 Then ptV = tv(kt)
                Exit For
            End If
        Next kt
        If Not ArcToPoint(ib, aB, ptN - bN, ptE - bE, ptV - bV, tfDeg, dls) Then Exit Function
        Do While jj <= np - 2
            If m_pMD(jj) >= q Then Exit Do
            If m_pDls(jj) > planMax Then planMax = m_pDls(jj)
            jj = jj + 1
        Loop
        If planMax >= CATCHUP_CURVE_MIN Then
            cap = planMax
            If motorOut > EPS And motorOut < cap Then cap = motorOut
        Else
            cap = motorOut
        End If
        If cap <= EPS Then Exit Do
        If dls <= cap + EPS Then Exit Do
        If q >= walkEnd - EPS Then Exit Do
        q = q + AIM_WALK_STEP_M
        If q > walkEnd Then q = walkEnd
    Loop
    aimMd = q
    AimPlanPoint = True
End Function

' Bit position for the calling Slidesheet row: AL/AM (cumulative N/E at the
' survey) plus the min-curve step from the survey (E/F/G) to the bit.
Private Function CallerBitPosition(ByVal bm As Double, ByVal ib As Double, ByVal aB As Double, _
                                   ByVal tvdBit As Variant, _
                                   ByRef bN As Double, ByRef bE As Double, _
                                   ByRef bV As Double) As Boolean
    Dim c As Range
    Dim ws As Worksheet
    Dim r As Long
    Dim vE As Variant, vF As Variant, vG As Variant
    Dim vH As Variant, vN As Variant, vEast As Variant
    Dim dt As Double, dN As Double, dE As Double

    CallerBitPosition = False
    On Error Resume Next
    Set c = Application.Caller
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    Set ws = c.Worksheet
    If ws.name <> "Slidesheet" Then Exit Function
    r = c.row
    vE = ws.Cells(r, "E").Value2
    vF = ws.Cells(r, "F").Value2
    vG = ws.Cells(r, "G").Value2
    vH = ws.Cells(r, "H").Value2
    vN = ws.Cells(r, "AL").Value2
    vEast = ws.Cells(r, "AM").Value2
    If Not (HasNum(vE) And HasNum(vF) And HasNum(vG) And HasNum(vN) And HasNum(vEast)) Then Exit Function
    MinCurveStep CDbl(vE), CDbl(vF), CDbl(vG), bm, ib, aB, dt, dN, dE
    bN = CDbl(vN) + dN
    bE = CDbl(vEast) + dE
    If HasNum(tvdBit) Then
        bV = CDbl(tvdBit)
    ElseIf HasNum(vH) Then
        bV = CDbl(vH) + dt
    Else
        Exit Function
    End If
    CallerBitPosition = True
End Function

' Plan-point aim for the calling Slidesheet row. False -> callers keep the
' attitude-only path (no plan path, or not called from a survey row).
Private Function RowPlanPointAim(ByVal bm As Double, ByVal ib As Double, ByVal aB As Double, _
                                 ByVal motorOut As Double, ByVal tvdBit As Variant, _
                                 ByVal tgtMd As Range, ByVal tgtInc As Range, _
                                 ByVal tgtAzm As Range, ByVal tgtTvd As Range, _
                                 ByRef tfDeg As Double, ByRef dls As Double) As Boolean
    Dim bN As Double, bE As Double, bV As Double, aimMd As Double
    RowPlanPointAim = False
    If Not CallerBitPosition(bm, ib, aB, tvdBit, bN, bE, bV) Then Exit Function
    RowPlanPointAim = AimPlanPoint(bm, ib, aB, bN, bE, bV, motorOut, _
                                   tgtMd, tgtInc, tgtAzm, tgtTvd, tfDeg, dls, aimMd)
End Function

' Cont DI: the same arc from the Cont DI bit. The tie-on survey (A2:C2) must be
' a Slidesheet survey row; its AL/AM and rolling motor yield (AI) are used.
' aimTgtMd: the target MD shown on Cont DI (S1); blank = next one ahead of the bit.
' part: "TF" (R##/L##), "DLR" (>= 0), "BRR" (build part, + build / - drop),
' "TRR" (turn part, magnitude), "TRRS" (turn part, + right / - left).
Public Function ProjContDiAim(ByVal part As Variant, _
                              ByVal survMd As Variant, ByVal survInc As Variant, _
                              ByVal survAzm As Variant, _
                              ByVal bitMd As Variant, ByVal incBit As Variant, _
                              ByVal azmBit As Variant, ByVal tvdBit As Variant, _
                              ByVal tgtMd As Range, ByVal tgtInc As Range, _
                              ByVal tgtAzm As Range, ByVal tgtTvd As Range, _
                              Optional ByVal aimTgtMd As Variant) As Variant
    Dim ss As Worksheet
    Dim r As Long, rr As Long
    Dim sm As Double
    Dim bm As Double, ib As Double, aB As Double
    Dim bN As Double, bE As Double, bV As Double
    Dim dt As Double, dN As Double, dE As Double
    Dim mo As Double, tfDeg As Double, dls As Double, aimMd As Double
    Dim turn As Double

    On Error GoTo Fail
    ProjContDiAim = ""
    If Not (HasNum(survMd) And HasNum(survInc) And HasNum(survAzm)) Then Exit Function
    If Not (HasNum(bitMd) And HasNum(incBit) And HasNum(azmBit) And HasNum(tvdBit)) Then Exit Function
    sm = CDbl(survMd)
    bm = CDbl(bitMd): ib = CDbl(incBit): aB = CDbl(azmBit)

    Set ss = ThisWorkbook.Worksheets("Slidesheet")
    r = 0
    For rr = 505 To 13 Step -1
        If HasNum(ss.Cells(rr, "E").Value2) Then
            If Abs(CDbl(ss.Cells(rr, "E").Value2) - sm) < 0.01 Then
                r = rr
                Exit For
            End If
        End If
    Next rr
    If r = 0 Then Exit Function
    If Not (HasNum(ss.Cells(r, "AL").Value2) And HasNum(ss.Cells(r, "AM").Value2)) Then Exit Function

    MinCurveStep sm, CDbl(survInc), CDbl(survAzm), bm, ib, aB, dt, dN, dE
    bN = CDbl(ss.Cells(r, "AL").Value2) + dN
    bE = CDbl(ss.Cells(r, "AM").Value2) + dE
    bV = CDbl(tvdBit)
    mo = SafeNum(ss.Cells(r, "AI").Value2)

    If Not AimPlanPoint(bm, ib, aB, bN, bE, bV, mo, tgtMd, tgtInc, tgtAzm, tgtTvd, _
                        tfDeg, dls, aimMd, aimTgtMd) Then Exit Function

    Select Case UCase$(Trim$(CStr(part & "")))
        Case "TF"
            ProjContDiAim = FormatRequiredTfDisplay(tfDeg, ib, aB, 5#)
        Case "DLR"
            ProjContDiAim = dls
        Case "BRR"
            ProjContDiAim = dls * Cos(Deg2Rad(tfDeg))
        Case "TRR", "TRRS"
            If Sin(Deg2Rad(ib)) > 0.01 Then
                turn = dls * Sin(Deg2Rad(tfDeg)) / Sin(Deg2Rad(ib))
            Else
                turn = 0#
            End If
            If UCase$(Trim$(CStr(part & ""))) = "TRR" Then turn = Abs(turn)
            ProjContDiAim = turn
        Case "AIM"
            ProjContDiAim = aimMd
    End Select
    Exit Function
Fail:
    ProjContDiAim = ""
End Function

' Min-curvature step → ΔTVD, ΔN, ΔE
Private Sub MinCurveStep(ByVal md1 As Double, ByVal i1 As Double, ByVal a1 As Double, _
                         ByVal md2 As Double, ByVal i2 As Double, ByVal a2 As Double, _
                         ByRef dTvd As Double, ByRef dN As Double, ByRef dE As Double)
    Dim beta As Double, h As Double, r1 As Double, r2 As Double, b1 As Double, b2 As Double
    beta = Deg2Rad(DoglegAngleDeg(i1, a1, i2, a2))
    If beta > 0.0000001 Then
        h = (md2 - md1) / 2# * (2# / beta * Tan(beta / 2#))
    Else
        h = (md2 - md1) / 2#
    End If
    r1 = Deg2Rad(i1): r2 = Deg2Rad(i2)
    b1 = Deg2Rad(a1): b2 = Deg2Rad(a2)
    dTvd = h * (Cos(r1) + Cos(r2))
    dN = h * (Sin(r1) * Cos(b1) + Sin(r2) * Cos(b2))
    dE = h * (Sin(r1) * Sin(b1) + Sin(r2) * Sin(b2))
End Sub

Private Function SafeNum(ByVal v As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    If isError(v) Then
        SafeNum = defaultValue
    ElseIf IsNumeric(v) Then
        SafeNum = CDbl(v)
    Else
        SafeNum = defaultValue
    End If
End Function

Private Function HasNum(ByVal v As Variant) As Boolean
    ' VBA does not stop at the first False, so CStr of an error value
    ' (#REF! / #NUM!) used to throw and the caller turned that into #NUM!.
    HasNum = False
    If isError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(v)) = 0 Then Exit Function
    End If
    On Error Resume Next
    HasNum = IsNumeric(v)
    On Error GoTo 0
End Function

' Recent motor output (°/30 m of slide): the last 3 numeric Q walking up this
' row, WEIGHTED by metres seen (M). A stand counts when Q is a real number > 0.
' Q = C×J/M, so Q×M = C×J and the weighted mean is the true dogleg delivered
' per slid metre across those stands — a 2 m sliver (Q spiking to 20) no longer
' counts the same as a 19 m stand, and a change of motor behavior (turn-and-
' build stands to high side) shows up within one or two stands instead of
' being buried in a five-stand average.
' Need 2+ samples; else this-row Q if it is numeric > 0; else blank.
' mRng is passed so Excel dirties this cell when metres-seen changes (Q = C×J/M).
' qRng/mRng are the same-height columns ending on this row (Excel owns recalc).
Public Function ProjRollingMotorOut(ByVal qRng As Range, ByVal mRng As Range) As Variant
    Const N_WANT As Long = 3
    Const Q_LO As Double = 0#
    Dim n As Long
    Dim qArr As Variant
    Dim mArr As Variant
    Dim i As Long
    Dim taken As Long
    Dim tot As Double        ' sum of Q × weight
    Dim wTot As Double       ' sum of weights (metres seen)
    Dim wt As Double
    Dim qV As Variant
    Dim mv As Variant
    Dim thisOk As Boolean
    Dim thisQ As Double

    On Error GoTo Fail
    If qRng Is Nothing Or mRng Is Nothing Then
        ProjRollingMotorOut = ""
        Exit Function
    End If
    n = qRng.rows.Count
    If n < 1 Or mRng.rows.Count <> n Then
        ProjRollingMotorOut = ""
        Exit Function
    End If

    qArr = qRng.Value2
    mArr = mRng.Value2
    taken = 0
    tot = 0#
    wTot = 0#
    thisOk = False
    thisQ = 0#

    For i = n To 1 Step -1
        RollingSample qArr, mArr, i, n, qV, mv
        If YieldSampleOk(qV, Q_LO) Then
            wt = 1#
            If HasNum(mv) Then
                If CDbl(mv) > 0# Then wt = CDbl(mv)
            End If
            tot = tot + CDbl(qV) * wt
            wTot = wTot + wt
            taken = taken + 1
            If i = n Then
                thisOk = True
                thisQ = CDbl(qV)
            End If
            If taken >= N_WANT Then Exit For
        End If
    Next i

    If taken >= 2 And wTot > 0# Then
        ProjRollingMotorOut = tot / wTot
    ElseIf thisOk Then
        ProjRollingMotorOut = thisQ
    Else
        ProjRollingMotorOut = ""
    End If
    Exit Function
Fail:
    ProjRollingMotorOut = CVErr(xlErrNum)
End Function

Private Sub RollingSample(ByVal qArr As Variant, ByVal mArr As Variant, _
                          ByVal i As Long, ByVal n As Long, _
                          ByRef qV As Variant, ByRef mv As Variant)
    If n = 1 And Not IsArray(qArr) Then
        qV = qArr
        mv = mArr
    Else
        qV = qArr(i, 1)
        mv = mArr(i, 1)
    End If
End Sub

Private Function YieldSampleOk(ByVal qV As Variant, ByVal qLo As Double) As Boolean
    YieldSampleOk = False
    If Not HasNum(qV) Then Exit Function
    If CDbl(qV) <= qLo Then Exit Function
    YieldSampleOk = True
End Function

' Dogleg below survey from rolling motor output (°/30 m of slide).
' This-row M < 2 m → 0 (interpolated / no seen slide: stay on survey attitude).
Private Function DoglegBelowYield(ByVal motorOut As Double, _
                                  ByVal mSeen As Double, _
                                  ByVal mBelow As Double) As Double
    If motorOut <= EPS Or mSeen + 0.0000001 < 2# Or Abs(mBelow) < EPS Then
        DoglegBelowYield = 0#
    Else
        DoglegBelowYield = motorOut / 30# * mBelow
    End If
End Function

' --- public UDFs ---------------------------------------------------------------

' Parse toolface text: 190M (magnetic deg), R/L highside, -30 left, 30 right.
' Highside accepts both R40 / L40 and the sheet's usual 40R / 40L.
' HSR is high side (same dial as HS / R). HSL is low side (same as LS / L).
' A full-stand # of Metres (T = C) often carries HSR; leaving it unparsed
' made INC @ BIT #VALUE! and wiped the Y comment.
' Unparseable text -> #VALUE! so W/X stop instead of silently projecting TF 0.
Public Function ProjParseTF(ByVal tfText As Variant) As Variant
    Dim s As String, n As Double
    On Error GoTo Fail
    If isError(tfText) Then ProjParseTF = CVErr(xlErrValue): Exit Function
    s = UCase$(Trim$(Replace(CStr(tfText & ""), Chr$(160), " ")))
    s = Replace(s, " ", "")
    If Len(s) = 0 Then ProjParseTF = "": Exit Function

    If Right$(s, 1) = "M" Then
        s = Left$(s, Len(s) - 1)
        If IsNumeric(s) Then
            ProjParseTF = CDbl(s)
            Exit Function
        End If
    End If

    If s = "R" Or s = "HS" Or s = "HSR" Then ProjParseTF = 0#: Exit Function
    If s = "L" Or s = "LS" Or s = "HSL" Then ProjParseTF = 180#: Exit Function

    If Left$(s, 1) = "R" And IsNumeric(mid$(s, 2)) Then
        ProjParseTF = Abs(CDbl(mid$(s, 2)))
        Exit Function
    End If
    If Left$(s, 1) = "L" And IsNumeric(mid$(s, 2)) Then
        ProjParseTF = -Abs(CDbl(mid$(s, 2)))
        Exit Function
    End If
    If Len(s) > 1 Then
        If Right$(s, 1) = "R" And IsNumeric(Left$(s, Len(s) - 1)) Then
            ProjParseTF = Abs(CDbl(Left$(s, Len(s) - 1)))
            Exit Function
        End If
        If Right$(s, 1) = "L" And IsNumeric(Left$(s, Len(s) - 1)) Then
            ProjParseTF = -Abs(CDbl(Left$(s, Len(s) - 1)))
            Exit Function
        End If
    End If

    If IsNumeric(s) Then
        n = CDbl(s)
        ProjParseTF = n   ' leading - already means left
        Exit Function
    End If

Fail:
    ProjParseTF = CVErr(xlErrValue)
End Function

' Dogleg build below survey (degrees), matching PROJOUT db
Private Function DoglegBelow(ByVal dls As Double, ByVal course As Double, _
                             ByVal mSeen As Double, ByVal mBelow As Double) As Double
    If Abs(mSeen) < EPS Or Abs(mBelow) < EPS Then
        DoglegBelow = 0#
    Else
        DoglegBelow = dls / mSeen * course / 30# * mBelow
    End If
End Function

' ISCWSA constant-TF fill-in (ebook §8.2):
'   ΔInc = dB · cos(TF)
'   ΔAzm = dB · sin(TF) / sin(Inc)
' dB is the same DoglegBelow (seen yield) used by INC@BIT.
' Inc < 5° (magnetic): /sin(I) is unstable — planar walk dB·sin(TF) only.
Private Function AzmWalkDeg(ByVal survInc As Double, ByVal dB As Double, ByVal tf As Double) As Double
    Dim sI As Double
    If Abs(dB) < EPS Then
        AzmWalkDeg = 0#
        Exit Function
    End If
    sI = Sin(Deg2Rad(survInc))
    If survInc >= 5# And Abs(sI) > EPS Then
        AzmWalkDeg = dB * Sin(Deg2Rad(tf)) / sI
    Else
        AzmWalkDeg = dB * Sin(Deg2Rad(tf))
    End If
End Function

' Last surveyed course TF: N = degrees, O = L / R (HS / Mag → no walk).
Private Function SignedSeenTf(ByVal seenTf As Variant, ByVal seenLR As Variant) As Double
    Dim n As Double, lr As String
    If Not HasNum(seenTf) Then
        SignedSeenTf = 0#
        Exit Function
    End If
    n = Abs(CDbl(seenTf))
    lr = UCase$(Trim$(CStr(seenLR & "")))
    If lr = "L" Then
        SignedSeenTf = -n
    ElseIf lr = "R" Then
        SignedSeenTf = n
    Else
        SignedSeenTf = 0#
    End If
End Function

' Walk TF for the azm-at-bit projection, calibrated by demonstrated results.
' The dial TF (U/AK) is what the DD set; N/O is the effective TF the hole
' actually carved last course (reactive torque + rotary dilution eat the
' difference — a stated 30R has delivered ~5-27R effective on this well).
' Projecting the dial literally overshot the walk up to 10 deg/stand at low
' inc (/sin(I) magnifies), flipping the AT recommendation to the wrong side.
'   same sign     -> smaller magnitude of dial vs demonstrated (never project
'                    more turn than the hole has shown it delivers)
'   opposite sign / no N -> dial (demonstrated TF is rotary noise; trust intent)
'   dial 0 / HS   -> 0 walk (high side turns nothing)
' A BLANK this-row dial is not "hold the survey." Leftover metres to the bit
' were already drilled: replay each overlapping R/S slide at that row's U,
' and hold only the rotary gaps. N/O arrive as arguments so Excel owns the
' recalc dependency — never read via Application.Caller (no dependency edge
' → stale AZM on recalc).
Private Function ResolveWalkTf(ByVal tfDeg As Variant, _
                               ByVal seenTf As Variant, ByVal seenLR As Variant) As Double
    Dim tfD As Double, tfE As Double
    tfD = SafeNum(tfDeg)
    tfE = SignedSeenTf(seenTf, seenLR)
    If Abs(tfE) < 0.05 Then
        ResolveWalkTf = tfD
    ElseIf tfD * tfE > 0# And Abs(tfE) < Abs(tfD) Then
        ResolveWalkTf = tfE
    Else
        ResolveWalkTf = tfD
    End If
End Function

' Replay already-drilled unseen hole (survey MD → bit MD): each overlapping
' Slide From/To (R/S) uses that row's Toolface (U). Gaps are rotary holds.
' Later rows win on overlap. tfDep is an Excel dirty hook (usually $U$13:$U$505).
Private Sub TouchTfDep(ByVal tfDep As Variant)
    Dim dump As Variant
    On Error Resume Next
    If IsMissing(tfDep) Then Exit Sub
    If TypeName(tfDep) = "Range" Then dump = tfDep.Cells(1, 1).Value2
    On Error GoTo 0
End Sub

Public Function ProjWalkToBit(ByVal survMd As Double, ByVal survInc As Double, _
                              ByVal survAzm As Double, ByVal bitMd As Double, _
                              ByVal motorOut As Double, _
                              ByVal slideFrom As Range, ByVal slideTo As Range, _
                              ByVal tfCol As Range, _
                              ByRef outInc As Double, ByRef outAzm As Double, _
                              ByRef outDt As Double, ByRef outDn As Double, _
                              ByRef outDe As Double) As Boolean
    ProjWalkToBit = WalkUnseenToBit(survMd, survInc, survAzm, bitMd, motorOut, _
                                    slideFrom, slideTo, tfCol, _
                                    outInc, outAzm, outDt, outDn, outDe)
End Function

Private Function WalkUnseenToBit(ByVal survMd As Double, ByVal survInc As Double, _
                                 ByVal survAzm As Double, ByVal bitMd As Double, _
                                 ByVal motorOut As Double, _
                                 ByVal slideFrom As Range, ByVal slideTo As Range, _
                                 ByVal tfCol As Range, _
                                 ByRef outInc As Double, ByRef outAzm As Double, _
                                 ByRef outDt As Double, ByRef outDn As Double, _
                                 ByRef outDe As Double) As Boolean
    Dim nr As Long, i As Long, j As Long, nc As Long, nB As Long
    Dim a As Double, b As Double, mid As Double, md1 As Double, md2 As Double
    Dim tf As Double, dB As Double, i2 As Double, a2 As Double
    Dim sdt As Double, sdn As Double, sde As Double
    Dim bestRow As Long, parsed As Variant
    Dim vA As Variant, vb As Variant, vu As Variant
    Dim cA() As Double, cB() As Double, cTf() As Double, cRow() As Long
    Dim brk() As Double
    Dim inc As Double, azm As Double

    WalkUnseenToBit = False
    outInc = survInc: outAzm = survAzm
    outDt = 0#: outDn = 0#: outDe = 0#
    If bitMd <= survMd + 0.01 Then Exit Function
    If slideFrom Is Nothing Or slideTo Is Nothing Or tfCol Is Nothing Then Exit Function
    nr = slideFrom.rows.Count
    If nr < 1 Then Exit Function

    ReDim cA(1 To nr)
    ReDim cB(1 To nr)
    ReDim cTf(1 To nr)
    ReDim cRow(1 To nr)
    nc = 0
    For i = 1 To nr
        vA = slideFrom.Cells(i, 1).Value2
        vb = slideTo.Cells(i, 1).Value2
        If Not HasNum(vA) Or Not HasNum(vb) Then GoTo NextClip
        a = CDbl(vA): b = CDbl(vb)
        If b <= a + 0.001 Then GoTo NextClip
        If a < survMd Then a = survMd
        If b > bitMd Then b = bitMd
        If b <= a + 0.001 Then GoTo NextClip
        vu = tfCol.Cells(i, 1).Value2
        parsed = ProjParseTF(vu)
        If isError(parsed) Or Not HasNum(parsed) Then GoTo NextClip
        nc = nc + 1
        cA(nc) = a: cB(nc) = b: cTf(nc) = CDbl(parsed): cRow(nc) = i
NextClip:
    Next i

    nB = 2 + nc * 2
    ReDim brk(1 To nB)
    brk(1) = survMd: brk(2) = bitMd: nB = 2
    For i = 1 To nc
        nB = nB + 1: brk(nB) = cA(i)
        nB = nB + 1: brk(nB) = cB(i)
    Next i
    For i = 1 To nB - 1
        For j = i + 1 To nB
            If brk(j) < brk(i) Then
                mid = brk(i): brk(i) = brk(j): brk(j) = mid
            End If
        Next j
    Next i
    j = 1
    For i = 2 To nB
        If brk(i) - brk(j) > 0.0005 Then
            j = j + 1
            brk(j) = brk(i)
        End If
    Next i
    nB = j

    inc = survInc: azm = survAzm
    For i = 1 To nB - 1
        md1 = brk(i): md2 = brk(i + 1)
        If md2 - md1 < 0.001 Then GoTo NextSeg
        mid = (md1 + md2) / 2#
        bestRow = 0: tf = 0#
        For j = 1 To nc
            If cA(j) <= mid + 0.0005 And mid < cB(j) - 0.0000001 Then
                If cRow(j) >= bestRow Then
                    bestRow = cRow(j)
                    tf = cTf(j)
                End If
            End If
        Next j
        If bestRow > 0 And motorOut > EPS Then
            dB = motorOut / 30# * (md2 - md1)
            i2 = inc + dB * Cos(Deg2Rad(tf))
            a2 = Wrap360(azm + AzmWalkDeg(inc, dB, tf))
        Else
            i2 = inc: a2 = azm
        End If
        If i2 < 0# Then i2 = 0#
        If i2 > 180# Then i2 = 180#
        MinCurveStep md1, inc, azm, md2, i2, a2, sdt, sdn, sde
        outDt = outDt + sdt: outDn = outDn + sdn: outDe = outDe + sde
        inc = i2: azm = a2
NextSeg:
    Next i
    outInc = inc: outAzm = azm
    WalkUnseenToBit = True
End Function

Private Function TryWalkFromCaller(ByVal motorOut As Double, _
                                   ByRef outInc As Double, ByRef outAzm As Double, _
                                   ByRef outDt As Double, ByRef outDn As Double, _
                                   ByRef outDe As Double) As Boolean
    Dim c As Range, ws As Worksheet, r As Long
    Dim survMd As Double, bitMd As Double, survInc As Double, survAzm As Double
    Dim mo As Double
    On Error GoTo Fail
    TryWalkFromCaller = False
    Set c = Application.Caller
    If c Is Nothing Then Exit Function
    Set ws = c.parent
    r = c.row
    If Not HasNum(ws.Cells(r, 5).Value2) Or Not HasNum(ws.Cells(r, 4).Value2) Then Exit Function
    If Not HasNum(ws.Cells(r, 6).Value2) Or Not HasNum(ws.Cells(r, 7).Value2) Then Exit Function
    survMd = CDbl(ws.Cells(r, 5).Value2)
    bitMd = CDbl(ws.Cells(r, 4).Value2)
    survInc = CDbl(ws.Cells(r, 6).Value2)
    survAzm = CDbl(ws.Cells(r, 7).Value2)
    mo = motorOut
    If mo <= EPS And HasNum(ws.Cells(r, 35).Value2) Then mo = CDbl(ws.Cells(r, 35).Value2)
    TryWalkFromCaller = WalkUnseenToBit(survMd, survInc, survAzm, bitMd, mo, _
                                        ws.Range("R13:R505"), ws.Range("S13:S505"), _
                                        ws.Range("U13:U505"), _
                                        outInc, outAzm, outDt, outDn, outDe)
    Exit Function
Fail:
    TryWalkFromCaller = False
End Function

' tfDeg is this-row AK (next-stand dial). Unseen hole is walked from R/S/U
' first: leftover slide already has a recorded toolface. This-row TF only
' applies to leftover when no overlapping R/S exists (legacy single-TF path).
' An error (unparseable U text) is passed through so W/X show it.
Public Function ProjIncAtBit(ByVal survInc As Variant, ByVal dls As Variant, _
                             ByVal course As Variant, ByVal mSeen As Variant, _
                             ByVal mBelow As Variant, ByVal tfDeg As Variant, _
                             Optional ByVal motorOut As Variant, _
                             Optional ByVal tfDep As Variant) As Variant
    Dim cI As Double, dg As Double, co As Double, ms As Double, mb As Double, tf As Double
    Dim dB As Double, mo As Double
    Dim wI As Double, wA As Double, wt As Double, wN As Double, wE As Double
    On Error GoTo Fail
    TouchTfDep tfDep
    If Not HasNum(survInc) Then ProjIncAtBit = "": Exit Function
    cI = CDbl(survInc)
    If isError(tfDeg) Then ProjIncAtBit = tfDeg: Exit Function
    If HasNum(motorOut) Then mo = CDbl(motorOut) Else mo = SafeNum(dls)
    If TryWalkFromCaller(mo, wI, wA, wt, wN, wE) Then
        ProjIncAtBit = wI
        Exit Function
    End If
    If Not HasNum(tfDeg) Then ProjIncAtBit = cI: Exit Function
    dg = SafeNum(dls): co = SafeNum(course): ms = SafeNum(mSeen)
    mb = SafeNum(mBelow): tf = SafeNum(tfDeg)
    If HasNum(motorOut) Then
        dB = DoglegBelowYield(CDbl(motorOut), ms, mb)
    Else
        dB = DoglegBelow(dg, co, ms, mb)
    End If
    ProjIncAtBit = cI + dB * Cos(Deg2Rad(tf))
    Exit Function
Fail:
    ProjIncAtBit = CVErr(xlErrNum)
End Function

Public Function ProjAzmAtBit(ByVal survInc As Variant, ByVal survAzm As Variant, _
                             ByVal dls As Variant, ByVal course As Variant, _
                             ByVal mSeen As Variant, ByVal mBelow As Variant, _
                             ByVal tfDeg As Variant, _
                             Optional ByVal seenTf As Variant, _
                             Optional ByVal seenLR As Variant, _
                             Optional ByVal motorOut As Variant, _
                             Optional ByVal tfDep As Variant) As Variant
    Dim cI As Double, cA As Double, dg As Double, co As Double, ms As Double, mb As Double, tf As Double
    Dim dB As Double, mo As Double
    Dim wI As Double, wA As Double, wt As Double, wN As Double, wE As Double
    On Error GoTo Fail
    TouchTfDep tfDep
    If Not HasNum(survInc) Or Not HasNum(survAzm) Then ProjAzmAtBit = "": Exit Function
    cI = CDbl(survInc)
    cA = CDbl(survAzm)
    ' A #VALUE! toolface (unparseable AK) stays off this column.
    If isError(tfDeg) Then ProjAzmAtBit = "": Exit Function
    If HasNum(motorOut) Then mo = CDbl(motorOut) Else mo = SafeNum(dls)
    If TryWalkFromCaller(mo, wI, wA, wt, wN, wE) Then
        ProjAzmAtBit = wA
        Exit Function
    End If
    If Not HasNum(tfDeg) Then ProjAzmAtBit = cA: Exit Function
    dg = SafeNum(dls): co = SafeNum(course): ms = SafeNum(mSeen)
    mb = SafeNum(mBelow)
    tf = ResolveWalkTf(tfDeg, seenTf, seenLR)
    If HasNum(motorOut) Then
        dB = DoglegBelowYield(CDbl(motorOut), ms, mb)
    Else
        dB = DoglegBelow(dg, co, ms, mb)
    End If
    ProjAzmAtBit = Wrap360(cA + AzmWalkDeg(cI, dB, tf))
    Exit Function
Fail:
    ProjAzmAtBit = ""
End Function

Public Function ProjTvdAtBit(ByVal survMd As Variant, ByVal survInc As Variant, _
                             ByVal survAzm As Variant, ByVal survTvd As Variant, _
                             ByVal bitMd As Variant, ByVal incBit As Variant, _
                             ByVal azmBit As Variant, _
                             Optional ByVal tfDep As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double, mo As Double
    Dim wI As Double, wA As Double
    On Error GoTo Fail
    TouchTfDep tfDep
    If Not HasNum(survMd) Or Not HasNum(survInc) Or Not HasNum(survAzm) Or Not HasNum(survTvd) Then
        ProjTvdAtBit = "": Exit Function
    End If
    If Not HasNum(bitMd) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjTvdAtBit = "": Exit Function
    End If
    mo = 0#
    If TryWalkFromCaller(mo, wI, wA, dt, dN, dE) Then
        ProjTvdAtBit = CDbl(survTvd) + dt
        Exit Function
    End If
    MinCurveStep CDbl(survMd), CDbl(survInc), CDbl(survAzm), _
                 CDbl(bitMd), CDbl(incBit), CDbl(azmBit), dt, dN, dE
    ProjTvdAtBit = CDbl(survTvd) + dt
    Exit Function
Fail:
    ProjTvdAtBit = ""
End Function

' Cumulative N from prior N + min-curve ΔN between surveys.
' On the first survey row (isFirst = True), integrate from Row 12 tie-on
' (md1/i1/a1) or surface (0,0,0) when blank — matching Well Seeker X tie-on.
Public Function ProjCumN(ByVal prevN As Variant, ByVal md1 As Variant, ByVal i1 As Variant, _
                         ByVal a1 As Variant, ByVal md2 As Variant, ByVal i2 As Variant, _
                         ByVal a2 As Variant, ByVal isFirst As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    Dim m1 As Double, inc1 As Double, az1 As Double
    On Error GoTo Fail
    If Not HasNum(md2) Or Not HasNum(i2) Or Not HasNum(a2) Then ProjCumN = "": Exit Function
    If CBool(isFirst) Then
        If HasNum(md1) Then
            m1 = CDbl(md1): inc1 = SafeNum(i1, 0#): az1 = SafeNum(a1, 0#)
        Else
            m1 = 0#: inc1 = 0#: az1 = 0#
        End If
        If CDbl(md2) <= m1 Then
            ProjCumN = SafeNum(prevN)
            Exit Function
        End If
        MinCurveStep m1, inc1, az1, CDbl(md2), CDbl(i2), CDbl(a2), dt, dN, dE
        ProjCumN = SafeNum(prevN) + dN
        Exit Function
    End If
    If Not HasNum(md1) Or Not HasNum(i1) Or Not HasNum(a1) Then ProjCumN = "": Exit Function
    MinCurveStep CDbl(md1), CDbl(i1), CDbl(a1), CDbl(md2), CDbl(i2), CDbl(a2), dt, dN, dE
    ProjCumN = SafeNum(prevN) + dN
    Exit Function
Fail:
    ProjCumN = CVErr(xlErrNum)
End Function

Public Function ProjCumE(ByVal prevE As Variant, ByVal md1 As Variant, ByVal i1 As Variant, _
                         ByVal a1 As Variant, ByVal md2 As Variant, ByVal i2 As Variant, _
                         ByVal a2 As Variant, ByVal isFirst As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    Dim m1 As Double, inc1 As Double, az1 As Double
    On Error GoTo Fail
    If Not HasNum(md2) Or Not HasNum(i2) Or Not HasNum(a2) Then ProjCumE = "": Exit Function
    If CBool(isFirst) Then
        If HasNum(md1) Then
            m1 = CDbl(md1): inc1 = SafeNum(i1, 0#): az1 = SafeNum(a1, 0#)
        Else
            m1 = 0#: inc1 = 0#: az1 = 0#
        End If
        If CDbl(md2) <= m1 Then
            ProjCumE = SafeNum(prevE)
            Exit Function
        End If
        MinCurveStep m1, inc1, az1, CDbl(md2), CDbl(i2), CDbl(a2), dt, dN, dE
        ProjCumE = SafeNum(prevE) + dE
        Exit Function
    End If
    If Not HasNum(md1) Or Not HasNum(i1) Or Not HasNum(a1) Then ProjCumE = "": Exit Function
    MinCurveStep CDbl(md1), CDbl(i1), CDbl(a1), CDbl(md2), CDbl(i2), CDbl(a2), dt, dN, dE
    ProjCumE = SafeNum(prevE) + dE
    Exit Function
Fail:
    ProjCumE = CVErr(xlErrNum)
End Function

' Min-curvature INC/AZM at MD q on the target table (same slerp as PlanAt).
Private Sub PlanAtt(ByRef m() As Double, ByRef i() As Double, ByRef a() As Double, _
                    ByVal n As Long, ByVal q As Double, ByRef outI As Double, ByRef outA As Double)
    Dim j As Long, t As Double
    If n < 2 Then
        outI = i(1): outA = a(1): Exit Sub
    End If
    If q <= m(1) Then outI = i(1): outA = a(1): Exit Sub
    If q >= m(n) Then outI = i(n): outA = a(n): Exit Sub
    j = 2
    Do While j <= n And m(j) < q
        j = j + 1
    Loop
    If j > n Then j = n
    If Abs(m(j) - m(j - 1)) < EPS Then
        outI = i(j): outA = a(j): Exit Sub
    End If
    t = (q - m(j - 1)) / (m(j) - m(j - 1))
    MDL_PlanGauge.PG_InterpCourseDir i(j - 1), a(j - 1), i(j), a(j), t, outI, outA
End Sub

' Min-curvature TVD at MD q on the target table (McStep from the course start).
Private Function PlanTvdMc(ByRef m() As Double, ByRef i() As Double, ByRef a() As Double, _
                           ByRef tv() As Double, ByVal n As Long, ByVal q As Double) As Double
    Dim j As Long, t As Double
    Dim outI As Double, outA As Double
    Dim dTvd As Double, dN As Double, dE As Double
    If n < 1 Then PlanTvdMc = 0#: Exit Function
    If q <= m(1) Then PlanTvdMc = tv(1): Exit Function
    If q >= m(n) Then PlanTvdMc = tv(n): Exit Function
    j = 2
    Do While j <= n And m(j) < q
        j = j + 1
    Loop
    If j > n Then j = n
    If Abs(m(j) - m(j - 1)) < EPS Then
        PlanTvdMc = tv(j): Exit Function
    End If
    t = (q - m(j - 1)) / (m(j) - m(j - 1))
    MDL_PlanGauge.PG_InterpCourseDir i(j - 1), a(j - 1), i(j), a(j), t, outI, outA
    MinCurveStep m(j - 1), i(j - 1), a(j - 1), q, outI, outA, dTvd, dN, dE
    PlanTvdMc = tv(j - 1) + dTvd
End Function

' Load target columns from a vertical range (MD, INC, AZM, TVD) — each a column vector
Private Function LoadTargets(ByVal mdR As Range, ByVal incR As Range, ByVal azmR As Range, _
                             ByVal tvdR As Range, ByRef m() As Double, ByRef i() As Double, _
                             ByRef a() As Double, ByRef tv() As Double) As Long
    Dim r As Long, n As Long, v As Variant
    n = 0
    ReDim m(1 To 100)
    ReDim i(1 To 100)
    ReDim a(1 To 100)
    ReDim tv(1 To 100)
    For r = 1 To mdR.rows.Count
        v = mdR.Cells(r, 1).value
        If HasNum(v) Then
            n = n + 1
            If n > 100 Then Exit For
            m(n) = CDbl(v)
            i(n) = SafeNum(incR.Cells(r, 1).value)
            a(n) = SafeNum(azmR.Cells(r, 1).value)
            tv(n) = SafeNum(tvdR.Cells(r, 1).value)
        End If
    Next r
    If n > 0 Then
        ReDim Preserve m(1 To n)
        ReDim Preserve i(1 To n)
        ReDim Preserve a(1 To n)
        ReDim Preserve tv(1 To n)
    End If
    LoadTargets = n
End Function

' A station is used up only when its inclination is already made.
' Tiny leftover MD is not a reason to jump to the next (hold) station —
' ComputeAim then uses the current stand length as the BURR distance.
Private Function TargetIsExhausted(ByVal bitMd As Double, ByVal incBit As Double, _
                                   ByVal tgtMd As Double, ByVal tgtInc As Double, _
                                   ByVal tgtTvd As Double, ByVal tvdBit As Variant) As Boolean
    Dim dist As Double, dTvd As Double
    TargetIsExhausted = False
    If incBit + MIN_AIM_DINC >= tgtInc Then
        dist = tgtMd - bitMd
        If dist <= EPS Then
            TargetIsExhausted = True
            Exit Function
        End If
        If dist < MIN_AIM_MD Then
            TargetIsExhausted = True
            Exit Function
        End If
        If HasNum(tvdBit) Then
            dTvd = tgtTvd - CDbl(tvdBit)
            If dTvd > 0# And dTvd < MIN_AIM_TVD Then TargetIsExhausted = True
        End If
    End If
End Function

' Final target already achieved at the survey station:
' survey INC >= final INC, or survey TVD >= final TVD.
' (Bit projection can overshoot TVD/INC before the survey has landed.)
Private Function PastFinalTarget(ByVal survInc As Double, ByVal survTvd As Variant, _
                                 ByRef i() As Double, ByRef tv() As Double, _
                                 ByVal n As Long) As Boolean
    PastFinalTarget = False
    If n < 1 Then Exit Function
    If survInc >= i(n) Then
        PastFinalTarget = True
        Exit Function
    End If
    If HasNum(survTvd) Then
        If CDbl(survTvd) >= tv(n) Then PastFinalTarget = True
    End If
End Function

' First usable plan index with m(k) > bitMd (skips exhausted intermediate targets).
' Stations already behind the bit are never the aim — even if their Inc is
' not yet made. The last plan station is never skipped by exhaustion — only
' by MD past or PastFinalTarget.
Private Function FirstAimIndex(ByRef m() As Double, ByRef i() As Double, _
                               ByRef tv() As Double, ByVal n As Long, _
                               ByVal bitMd As Double, ByVal incBit As Double, _
                               ByVal tvdBit As Variant) As Long
    Dim k As Long
    k = 1
    Do While k <= n And m(k) <= bitMd
        k = k + 1
    Loop
    Do While k <= n
        If k = n Then
            FirstAimIndex = k
            Exit Function
        End If
        If Not TargetIsExhausted(bitMd, incBit, m(k), i(k), tv(k), tvdBit) Then
            FirstAimIndex = k
            Exit Function
        End If
        k = k + 1
    Loop
    FirstAimIndex = n + 1
End Function

' Active target MD ahead of bit (first MD > bitMd), or blank. Legacy MD-only picker.
Public Function ProjActiveTargetMd(ByVal bitMd As Variant, ByVal tgtMd As Range) As Variant
    Dim r As Long, bm As Double, v As Variant
    On Error GoTo Fail
    If Not HasNum(bitMd) Then ProjActiveTargetMd = "": Exit Function
    bm = CDbl(bitMd)
    For r = 1 To tgtMd.rows.Count
        v = tgtMd.Cells(r, 1).value
        If HasNum(v) Then
            If CDbl(v) > bm Then
                ProjActiveTargetMd = CDbl(v)
                Exit Function
            End If
        End If
    Next r
    ProjActiveTargetMd = ""
    Exit Function
Fail:
    ProjActiveTargetMd = CVErr(xlErrNum)
End Function

' Active target MD with full target table (skips exhausted near stations).
' Optional survInc/survTvd gate the "past final target" stop (survey F/H).
Public Function ProjActiveTargetMdEx(ByVal bitMd As Variant, ByVal incBit As Variant, _
                                     ByVal tvdBit As Variant, ByVal tgtMd As Range, _
                                     ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                                     ByVal tgtTvd As Range, _
                                     Optional ByVal survInc As Variant, _
                                     Optional ByVal survTvd As Variant) As Variant
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long, k As Long, bm As Double, ib As Double
    Dim gateInc As Double, gateTvd As Variant
    On Error GoTo Fail
    If Not HasNum(bitMd) Then ProjActiveTargetMdEx = "": Exit Function
    bm = CDbl(bitMd)
    ib = SafeNum(incBit)
    n = LoadTargets(tgtMd, tgtInc, tgtAzm, tgtTvd, m, i, a, tv)
    If n < 1 Then ProjActiveTargetMdEx = "": Exit Function
    If HasNum(survInc) Then
        gateInc = CDbl(survInc)
    Else
        gateInc = ib
    End If
    If HasNum(survTvd) Then
        gateTvd = survTvd
    Else
        gateTvd = tvdBit
    End If
    If PastFinalTarget(gateInc, gateTvd, i, tv, n) Then
        ProjActiveTargetMdEx = "": Exit Function
    End If
    k = FirstAimIndex(m, i, tv, n, bm, ib, tvdBit)
    If k > n Then
        ProjActiveTargetMdEx = ""
    Else
        ProjActiveTargetMdEx = m(k)
    End If
    Exit Function
Fail:
    ProjActiveTargetMdEx = CVErr(xlErrNum)
End Function

' Aim walk to next achievable plan target; returns False if none
' reqDls = full dogleg rate (°/30m); burr = build-up rate (°/30m); reqTf = TF bit→aim
Private Function ComputeAim(ByVal bitMd As Double, ByVal incBit As Double, _
                            ByVal azmBit As Double, ByVal motorOut As Double, _
                            ByVal course As Double, ByVal tgtMd As Range, _
                            ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                            ByVal tgtTvd As Range, ByRef dist As Double, _
                            ByRef reqDls As Double, ByRef burr As Double, _
                            ByRef reqTf As Double, ByRef aimInc As Double, _
                            ByRef aimTvd As Double, _
                            Optional ByVal allowWalk As Boolean = True, _
                            Optional ByVal tvdBit As Variant) As Boolean
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long, k As Long, q As Double, ti As Double, ta As Double
    Dim beta As Double

    ComputeAim = False
    dist = 0#: reqDls = 0#: burr = 0#: reqTf = 0#
    aimInc = 0#: aimTvd = 0#
    n = LoadTargets(tgtMd, tgtInc, tgtAzm, tgtTvd, m, i, a, tv)
    If n < 1 Then Exit Function
    ' Past-final stop is applied in ProjActiveTargetMdEx (survey F/H); AO blanks AR/AS/AT.

    k = FirstAimIndex(m, i, tv, n, bitMd, incBit, tvdBit)
    If k > n Then Exit Function

    q = m(k)
    ti = i(k)
    ta = a(k)
    dist = q - bitMd
    ' Same BURR formula as a normal stand: ΔI×30/dMD.
    ' When dMD is shorter than this stand (or the bit has passed the
    ' station MD) use course length from the sheet — not a made-up floor.
    If dist < course And incBit + MIN_AIM_DINC < ti Then
        If course > EPS Then dist = course
    End If
    If dist <= EPS Then Exit Function
    beta = DoglegAngleDeg(incBit, azmBit, ti, ta)
    reqDls = beta * 30# / dist

    ' Walk aim forward while required DLS exceeds motor (PROJBIT).
    If allowWalk And motorOut > EPS And course > EPS Then
        Do While q < m(n) And reqDls > motorOut + EPS
            q = q + course
            If q > m(n) Then q = m(n)
            PlanAtt m, i, a, n, q, ti, ta
            dist = q - bitMd
            If dist <= EPS Then Exit Do
            beta = DoglegAngleDeg(incBit, azmBit, ti, ta)
            reqDls = beta * 30# / dist
            If Abs(q - m(n)) < EPS And reqDls > motorOut Then Exit Do
        Loop
    End If

    If dist <= EPS Then Exit Function
    burr = (ti - incBit) * 30# / dist
    reqTf = ToolfaceBetweenDeg(incBit, azmBit, ti, ta)
    aimInc = ti
    aimTvd = PlanTvdMc(m, i, a, tv, n, q)
    ComputeAim = True
End Function

' --- Plan rate over the next course -----------------------------------------
' While the plan is curving, metres-to-slide and BURR use this rate, not the
' dogleg from the bit to the section end. That section aim divides whatever
' inclination is still left by the shrinking distance, so a bit a few degrees
' behind a 4°/30 m build gets told to slide 11–13 m. The plan rate on the
' same motor is 5–6 m, which is what actually held the curve.
' Toolface still aims from the bit toward the plan one course ahead.
' Only while the plan dogleg over the stand is > 1 deg/30 m and plan INC
' ahead is < 80 deg. Holds and the lateral stay on the section aim.

' Whole plan (MD/INC/AZI/TVD) in one sheet read; arrays are 1-based. 0 = none.
Private Function LoadFullPlan(ByRef m() As Double, ByRef i() As Double, _
                              ByRef a() As Double, ByRef tv() As Double) As Long
    Dim ws As Worksheet, hdr As Variant, v As Variant
    Dim cM As Long, cI As Long, cA As Long, cT As Long
    Dim c As Long, r As Long, n As Long, lastR As Long, nRows As Long
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(CATCHUP_PLAN_SHEET)
    hdr = ws.Range("A2:T2").Value2
    For c = 1 To 20
        Select Case UCase$(Trim$(CStr(hdr(1, c) & "")))
            Case "MD": cM = c
            Case "INC": cI = c
            Case "AZI", "AZM": cA = c
            Case "TVD": cT = c
        End Select
    Next c
    If cM = 0 Or cI = 0 Or cA = 0 Or cT = 0 Then Exit Function
    lastR = ws.Cells(ws.rows.Count, cM).End(xlUp).row
    If lastR < 4 Then Exit Function
    v = ws.Range(ws.Cells(3, 1), ws.Cells(lastR, 20)).Value2
    nRows = UBound(v, 1)
    ReDim m(1 To nRows): ReDim i(1 To nRows): ReDim a(1 To nRows): ReDim tv(1 To nRows)
    For r = 1 To nRows
        If HasNum(v(r, cM)) Then
            If n = 0 Then
                n = n + 1
            ElseIf CDbl(v(r, cM)) > m(n) Then
                n = n + 1
            Else
                GoTo NextRow
            End If
            m(n) = CDbl(v(r, cM))
            i(n) = SafeNum(v(r, cI))
            a(n) = SafeNum(v(r, cA))
            tv(n) = SafeNum(v(r, cT))
        End If
NextRow:
    Next r
    LoadFullPlan = n
    Exit Function
Fail:
    LoadFullPlan = 0
End Function

' Plan dogleg rate (°/30 m) over the next course, and the toolface from the
' bit toward the plan attitude at the end of that course. 0 when catch-up
' does not apply. tvdBit is unused; kept so callers do not change.
Private Function CatchUpDls(ByVal bitMd As Double, ByVal incBit As Double, _
                            ByVal azmBit As Double, ByVal course As Double, _
                            ByVal tvdBit As Variant, ByRef reqTf As Double) As Double
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long
    Dim i0 As Double, a0 As Double, i1 As Double, a1 As Double
    Dim beta As Double

    CatchUpDls = 0#: reqTf = 0#
    If course <= EPS Then Exit Function
    n = LoadFullPlan(m, i, a, tv)
    If n < 2 Then Exit Function
    If bitMd < m(1) Then Exit Function

    PlanAtt m, i, a, n, bitMd, i0, a0
    PlanAtt m, i, a, n, bitMd + course, i1, a1
    If i1 >= CATCHUP_LATERAL_INC Then Exit Function
    beta = DoglegAngleDeg(i0, a0, i1, a1)
    If beta * 30# / course <= CATCHUP_CURVE_MIN Then Exit Function

    CatchUpDls = beta * 30# / course
    reqTf = ToolfaceBetweenDeg(incBit, azmBit, i1, a1)
End Function

' Same-row number on the calling Slidesheet cell (course C), so ProjRequiredTf
' can use the same stand length as Meters To Slide without a new argument.
Private Function CallerRowNum(ByVal colName As String) As Double
    Dim c As Range
    On Error Resume Next
    Set c = Application.Caller
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    On Error Resume Next
    If c.Worksheet.name = "Slidesheet" Then
        CallerRowNum = SafeNum(c.Worksheet.Range(colName & CStr(c.row)).Value2)
    End If
    On Error GoTo 0
End Function

' Effective motor toward required TF: cos(slideTF − reqTF). Blank TF → full motor (recommend).
Private Function TfMotorFactor(ByVal tfDeg As Variant, ByVal reqTf As Double) As Double
    Dim slideTf As Double, dlt As Double, eff As Double
    If Not HasNum(tfDeg) Then
        TfMotorFactor = 1#
        Exit Function
    End If
    slideTf = CDbl(tfDeg)
    dlt = Wrap180(slideTf - reqTf)
    eff = Cos(Deg2Rad(dlt))
    If eff < 0.05 Then
        TfMotorFactor = 0#   ' TF nearly orthogonal / opposite — cannot progress
    Else
        TfMotorFactor = eff
    End If
End Function

' Meters to slide: (|reqDLS| / motorOut) * course
' reqDLS is the plan-point arc (AimPlanPoint), the same curve as AR and AT.
' Without a plan path: plan dogleg on a curve, else dogleg to the section.
' incBit / azmBit are the PROJECTED bit attitude (Slidesheet W / X).
' Under 1 m returns 0. Y uses this AS value as-is.
Public Function ProjMetersToSlide(ByVal bitMd As Variant, ByVal incBit As Variant, _
                                  ByVal azmBit As Variant, ByVal motorOut As Variant, _
                                  ByVal course As Variant, ByVal tgtMd As Range, _
                                  ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                                  ByVal tgtTvd As Range, _
                                  Optional ByVal tvdBit As Variant) As Variant
    Dim bm As Double, ib As Double, aB As Double, mo As Double, co As Double
    Dim dist As Double, req As Double, burr As Double, reqTf As Double
    Dim aimInc As Double, aimTvd As Double, cu As Double, cuTf As Double
    On Error GoTo Fail

    If Not HasNum(bitMd) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjMetersToSlide = "": Exit Function
    End If
    bm = CDbl(bitMd): ib = CDbl(incBit): aB = CDbl(azmBit)
    mo = SafeNum(motorOut): co = SafeNum(course)
    If mo <= EPS Or co <= EPS Then ProjMetersToSlide = "": Exit Function

    If Not ComputeAim(bm, ib, aB, mo, co, tgtMd, tgtInc, tgtAzm, tgtTvd, _
                      dist, req, burr, reqTf, aimInc, aimTvd, False, tvdBit) Then
        ProjMetersToSlide = "": Exit Function
    End If

    Dim tfP As Double, dlsP As Double
    If RowPlanPointAim(bm, ib, aB, mo, tvdBit, tgtMd, tgtInc, tgtAzm, tgtTvd, tfP, dlsP) Then
        req = dlsP
    Else
        ' No plan path: slide the plan's own rate on a curve.
        cu = CatchUpDls(bm, ib, aB, co, tvdBit, cuTf)
        If cu > EPS Then req = cu
    End If

    If Abs(req) <= EPS Then
        ProjMetersToSlide = 0#
        Exit Function
    End If

    ' 3D dogleg required (°/30 m) / motor × this stand. Cannot slide more than C.
    ' Under 1 m is rotate — a 0.12 m call is noise, not a slide.
    ProjMetersToSlide = Abs(req) / mo * co
    If ProjMetersToSlide > co Then ProjMetersToSlide = co
    If ProjMetersToSlide < SLIDE_IGNORE_M Then ProjMetersToSlide = 0#
    Exit Function
Fail:
    ProjMetersToSlide = CVErr(xlErrNum)
End Function

' Meters remain to rotate in the current drill set (course length, col C):
'   max(0, courseLen - metersToSlide)
Public Function ProjMetersRemainToRotate(ByVal bitMd As Variant, ByVal incBit As Variant, _
                                         ByVal azmBit As Variant, ByVal motorOut As Variant, _
                                         ByVal course As Variant, ByVal tgtMd As Range, _
                                         ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                                         ByVal tgtTvd As Range, _
                                         Optional ByVal tvdBit As Variant) As Variant
    Dim co As Double
    Dim slide As Variant
    Dim remain As Double
    On Error GoTo Fail

    co = SafeNum(course)
    If co <= EPS Then
        ProjMetersRemainToRotate = "": Exit Function
    End If

    slide = ProjMetersToSlide(bitMd, incBit, azmBit, motorOut, course, _
                              tgtMd, tgtInc, tgtAzm, tgtTvd, tvdBit)
    If isError(slide) Then
        ProjMetersRemainToRotate = slide
        Exit Function
    End If
    If Not HasNum(slide) Then
        ProjMetersRemainToRotate = "": Exit Function
    End If

    remain = co - CDbl(slide)
    If remain < 0# Then remain = 0#
    ProjMetersRemainToRotate = remain
    Exit Function
Fail:
    ProjMetersRemainToRotate = CVErr(xlErrNum)
End Function

' Dogleg required (DLR, °/30 m): the plan-point arc from the bit. Without a
' plan path, the dogleg to the next station's attitude over its remaining
' MD. ProjBurr is kept as the public function name so
' existing workbook formulas remain compatible. The plan's nominal DLS must
' not overwrite this live, projected-bit-to-target requirement.
Public Function ProjBurr(ByVal bitMd As Variant, ByVal incBit As Variant, _
                         ByVal azmBit As Variant, ByVal motorOut As Variant, _
                         ByVal course As Variant, ByVal tgtMd As Range, _
                         ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                         ByVal tgtTvd As Range, Optional ByVal tvdBit As Variant) As Variant
    Dim bm As Double, ib As Double, aB As Double, mo As Double, co As Double
    Dim dist As Double, req As Double, burr As Double, reqTf As Double
    Dim aimInc As Double, aimTvd As Double
    On Error GoTo Fail

    If Not HasNum(bitMd) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjBurr = "": Exit Function
    End If
    bm = CDbl(bitMd): ib = CDbl(incBit): aB = CDbl(azmBit)
    mo = SafeNum(motorOut): co = SafeNum(course)

    If Not ComputeAim(bm, ib, aB, mo, co, tgtMd, tgtInc, tgtAzm, tgtTvd, _
                      dist, req, burr, reqTf, aimInc, aimTvd, False, tvdBit) Then
        ProjBurr = "": Exit Function
    End If

    Dim tfP As Double, dlsP As Double
    If RowPlanPointAim(bm, ib, aB, mo, tvdBit, tgtMd, tgtInc, tgtAzm, tgtTvd, tfP, dlsP) Then
        req = dlsP
    End If
    ProjBurr = req
    Exit Function
Fail:
    ProjBurr = CVErr(xlErrNum)
End Function

' Build rate (°/30 m) to LAND on the next plan station's TVD (Cont DI BRR).
'
' When the aim station is a landing (Inc >= landInc, default 88°) the number
' that matters is the circular arc that reaches the station's inclination
' exactly at its TVD - MD is a by-product:
'     dTVD = R * (sin I_tgt - sin I_bit)      ->   BUR = (30 * 180/pi) / R
' ProjBurr's 3D dogleg-over-plan-MD rate answers "match the target attitude
' by the plan's MD" instead; when the bit is behind plan that lands high
' (4.07 vs 3.74 at 2870 m on 35780: 11 m above target TVD).
' Intermediate (non-landing) stations, or a bit already at / below the
' station's Inc or TVD, fall back to ProjBurr so the cell never goes blank
' mid-curve.
Public Function ProjLandingBurr(ByVal bitMd As Variant, ByVal incBit As Variant, _
                                ByVal azmBit As Variant, ByVal tvdBit As Variant, _
                                ByVal course As Variant, ByVal tgtMd As Range, _
                                ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                                ByVal tgtTvd As Range, _
                                Optional ByVal landInc As Variant) As Variant
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long, k As Long
    Dim bm As Double, ib As Double, tb As Double, li As Double
    Dim dTvd As Double, dSin As Double, radius As Double
    On Error GoTo Fail

    If Not HasNum(bitMd) Or Not HasNum(incBit) Or Not HasNum(tvdBit) Then
        ProjLandingBurr = "": Exit Function
    End If
    bm = CDbl(bitMd): ib = CDbl(incBit): tb = CDbl(tvdBit)
    li = SafeNum(landInc, 88#)

    n = LoadTargets(tgtMd, tgtInc, tgtAzm, tgtTvd, m, i, a, tv)
    If n < 1 Then ProjLandingBurr = "": Exit Function
    k = FirstAimIndex(m, i, tv, n, bm, ib, tvdBit)
    If k > n Then ProjLandingBurr = "": Exit Function

    If i(k) >= li Then
        dTvd = tv(k) - tb
        dSin = Sin(Deg2Rad(i(k))) - Sin(Deg2Rad(ib))
        If dTvd > EPS And dSin > EPS Then
            radius = dTvd / dSin
            ProjLandingBurr = Rad2Deg(30# / radius)
            Exit Function
        End If
    End If

    ' Not a landing aim (or geometry already past it): plan-MD dogleg rate.
    ProjLandingBurr = ProjBurr(bitMd, incBit, azmBit, 0, course, _
                               tgtMd, tgtInc, tgtAzm, tgtTvd, tvdBit)
    Exit Function
Fail:
    ProjLandingBurr = CVErr(xlErrNum)
End Function

' Hole metres to put in the Y comment: AS = |reqDLS| / Q_avg × C, hard-capped
' at this stand's C. Two-decimal sheet precision — never 0.25 rounding.
Public Function ProjInstructedSlideM(ByVal metersToSlide As Variant, ByVal tfText As Variant, _
                                     Optional ByVal maxSlide As Variant) As Double
    Dim m As Double, cap As Double

    m = 0#
    If HasNum(metersToSlide) Then
        m = CDbl(metersToSlide)
        If m < 0# Then m = 0#
    End If

    cap = 0#
    If HasNum(maxSlide) Then
        cap = CDbl(maxSlide)
        If cap < 0# Then cap = 0#
    End If

    If cap > 0# And m > cap Then m = cap

    m = Application.WorksheetFunction.Round(m, 2)
    If cap > 0# Then
        cap = Application.WorksheetFunction.Round(cap, 2)
        If m > cap Then m = cap
    End If
    ProjInstructedSlideM = m
End Function

' Bit Northing = survey CumN + min-curve ΔN (surv → bit)
Public Function ProjBitN(ByVal survMd As Variant, ByVal survInc As Variant, _
                         ByVal survAzm As Variant, ByVal cumN As Variant, _
                         ByVal bitMd As Variant, ByVal incBit As Variant, _
                         ByVal azmBit As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    On Error GoTo Fail
    If Not HasNum(survMd) Or Not HasNum(survInc) Or Not HasNum(survAzm) Then
        ProjBitN = "": Exit Function
    End If
    If Not HasNum(cumN) Or Not HasNum(bitMd) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjBitN = "": Exit Function
    End If
    MinCurveStep CDbl(survMd), CDbl(survInc), CDbl(survAzm), _
                 CDbl(bitMd), CDbl(incBit), CDbl(azmBit), dt, dN, dE
    ProjBitN = CDbl(cumN) + dN
    Exit Function
Fail:
    ProjBitN = CVErr(xlErrNum)
End Function

' Bit Easting = survey CumE + min-curve ΔE (surv → bit)
Public Function ProjBitE(ByVal survMd As Variant, ByVal survInc As Variant, _
                         ByVal survAzm As Variant, ByVal cumE As Variant, _
                         ByVal bitMd As Variant, ByVal incBit As Variant, _
                         ByVal azmBit As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    On Error GoTo Fail
    If Not HasNum(survMd) Or Not HasNum(survInc) Or Not HasNum(survAzm) Then
        ProjBitE = "": Exit Function
    End If
    If Not HasNum(cumE) Or Not HasNum(bitMd) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjBitE = "": Exit Function
    End If
    MinCurveStep CDbl(survMd), CDbl(survInc), CDbl(survAzm), _
                 CDbl(bitMd), CDbl(incBit), CDbl(azmBit), dt, dN, dE
    ProjBitE = CDbl(cumE) + dE
    Exit Function
Fail:
    ProjBitE = CVErr(xlErrNum)
End Function

' Required toolface to intercept the next plan target.
'
' 1) Start from the PROJECTED bit attitude (Inc@Bit / Azm@Bit, W/X), which
'    already embeds the slide drilled below the survey at the user toolface.
' 2) Aim at the NEXT NAMED PLAN TARGET ahead of the bit (ProjTargets_*,
'    FirstAimIndex) — the same station BURR (AR) and Metres To Slide (AS)
'    aim at, so the three numbers in Y describe one manoeuvre. Aiming the TF
'    at a 15 m _OC_Survey plan station instead (Aug-30 change) told the DD to
'    turn left onto a station the projected bit had already passed while the
'    slide metres were still sized for the named target to the right.
' 3) From bit attitude to aim attitude, use the industry-standard
'    (Well Seeker) spherical toolface — the angle from high side to the
'    dogleg plane that carries (incBit, azmBit) onto (aimInc, aimAzm):
'       TF = Atan2( sinI2·cosI1·cosΔA − sinI1·cosI2 ,  sinI2·sinΔA )
'    (ToolfaceBetweenDeg). A degree of azimuth walk only moves the bit
'    sin(Inc) as far as a degree of build, so raw Atan2(ΔInc, ΔAzm) is wrong —
'    it over-weights the turn (~2.4x at Inc 25°) and disagreed with the office.
'
' Display: Gravity → R## / L## / 0 ; Magnetic (low inc) → ###M
Public Function ProjRequiredTf(ByVal bitMd As Variant, _
                               ByVal incAtBit As Variant, ByVal azmAtBit As Variant, _
                               ByVal tgtMd As Range, ByVal tgtInc As Range, _
                               ByVal tgtAzm As Range, ByVal tgtTvd As Range, _
                               Optional ByVal thresholdDeg As Variant, _
                               Optional ByVal tvdBit As Variant) As Variant
    Dim bm As Double, ib As Double, aB As Double
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long, k As Long
    Dim ti As Double, ta As Double
    Dim reqTf As Double, co As Double, cu As Double, cuTf As Double
    Dim dist As Double, req As Double, burr As Double, wayTf As Double
    Dim aimInc As Double, aimTvd As Double
    On Error GoTo Fail

    If Not HasNum(bitMd) Or Not HasNum(incAtBit) Or Not HasNum(azmAtBit) Then
        ProjRequiredTf = "": Exit Function
    End If
    bm = CDbl(bitMd)
    ib = CDbl(incAtBit)
    aB = CDbl(azmAtBit)

    n = LoadTargets(tgtMd, tgtInc, tgtAzm, tgtTvd, m, i, a, tv)
    If n < 1 Then
        ProjRequiredTf = "": Exit Function
    End If
    k = FirstAimIndex(m, i, tv, n, bm, ib, tvdBit)
    If k > n Then
        ProjRequiredTf = "": Exit Function
    End If
    ti = i(k)
    ta = a(k)

    Dim tfP As Double, dlsP As Double
    If RowPlanPointAim(bm, ib, aB, CallerRowNum("AI"), tvdBit, _
                       tgtMd, tgtInc, tgtAzm, tgtTvd, tfP, dlsP) Then
        ProjRequiredTf = FormatRequiredTfDisplay(tfP, ib, aB, thresholdDeg)
        Exit Function
    End If

    ' Spherical TF needs no MD: it is the direction of the dogleg plane from
    ' bit attitude to aim attitude, exactly what Well Seeker tabulates.
    If DoglegAngleDeg(ib, aB, ti, ta) < 0.000001 Then
        reqTf = 0#   ' already pointed at the aim attitude
    Else
        reqTf = ToolfaceBetweenDeg(ib, aB, ti, ta)
    End If

    ' Same floor as Meters To Slide: when catching up to the plan one course
    ' ahead is harder than the waypoint aim, steer to that attitude instead.
    co = CallerRowNum("C")
    If co > EPS Then
        cu = CatchUpDls(bm, ib, aB, co, tvdBit, cuTf)
        If cu > EPS Then
            If ComputeAim(bm, ib, aB, 0#, co, tgtMd, tgtInc, tgtAzm, tgtTvd, _
                          dist, req, burr, wayTf, aimInc, aimTvd, False, tvdBit) Then
                If cu > Abs(req) Then reqTf = cuTf
            End If
        End If
    End If

    ProjRequiredTf = FormatRequiredTfDisplay(reqTf, ib, aB, thresholdDeg)
    Exit Function
Fail:
    ProjRequiredTf = CVErr(xlErrNum)
End Function

' Cont DI tab: required toolface from a bit attitude straight to a target
' attitude — the same spherical TF (ToolfaceBetweenDeg) and R##/L##/0/###M
' display as Slidesheet AT, so the two tabs agree.
Public Function ProjTfToTarget(ByVal incBit As Variant, ByVal azmBit As Variant, _
                               ByVal tgtInc As Variant, ByVal tgtAzm As Variant, _
                               Optional ByVal thresholdDeg As Variant) As Variant
    Dim ib As Double, aB As Double, ti As Double, ta As Double, reqTf As Double
    On Error GoTo Fail
    If Not HasNum(incBit) Or Not HasNum(azmBit) Or Not HasNum(tgtInc) Or Not HasNum(tgtAzm) Then
        ProjTfToTarget = "": Exit Function
    End If
    ib = CDbl(incBit): aB = CDbl(azmBit): ti = CDbl(tgtInc): ta = CDbl(tgtAzm)
    If DoglegAngleDeg(ib, aB, ti, ta) < 0.000001 Then
        reqTf = 0#
    Else
        reqTf = ToolfaceBetweenDeg(ib, aB, ti, ta)
    End If
    ProjTfToTarget = FormatRequiredTfDisplay(reqTf, ib, aB, thresholdDeg)
    Exit Function
Fail:
    ProjTfToTarget = CVErr(xlErrNum)
End Function

' 3D dogleg angle (degrees) between two attitudes — the numerator of the
' Slidesheet's required DLS (ProjBurr: beta × 30 / dMD).
Public Function ProjDoglegDeg(ByVal inc1 As Variant, ByVal azm1 As Variant, _
                              ByVal inc2 As Variant, ByVal azm2 As Variant) As Variant
    On Error GoTo Fail
    If Not HasNum(inc1) Or Not HasNum(azm1) Or Not HasNum(inc2) Or Not HasNum(azm2) Then
        ProjDoglegDeg = "": Exit Function
    End If
    ProjDoglegDeg = DoglegAngleDeg(CDbl(inc1), CDbl(azm1), CDbl(inc2), CDbl(azm2))
    Exit Function
Fail:
    ProjDoglegDeg = CVErr(xlErrNum)
End Function

' Gravity highside: R/L/0. Magnetic: 0-359M (Wrap360 — left of north is e.g. 348M, not "Left").
' Magnetic TF = bearing of the push: high side's horizontal projection points
' along the hole azimuth, so bearing = azmBit + highside TF.
Private Function FormatRequiredTfDisplay(ByVal tfDeg As Double, ByVal incBit As Double, _
                                         ByVal azmBit As Double, _
                                         ByVal thresholdDeg As Variant) As String
    Dim mode As Variant
    Dim a As Double
    mode = ProjTfMode("", incBit, thresholdDeg)
    If isError(mode) Then mode = "Gravity"

    If CStr(mode) = "Magnetic" Then
        a = Wrap360(azmBit + tfDeg)
        FormatRequiredTfDisplay = Format$(Application.WorksheetFunction.Round(a, 0), "0") & "M"
    Else
        a = Application.WorksheetFunction.Round(Wrap180(tfDeg), 0)
        If a = 0# Then
            FormatRequiredTfDisplay = "0"
        ElseIf a > 0# Then
            FormatRequiredTfDisplay = "R" & Format$(a, "0")
        Else
            FormatRequiredTfDisplay = "L" & Format$(Abs(a), "0")
        End If
    End If
End Function

' Magnetic if TF ends with M; Gravity if L/R/signed HS; else IncBit vs threshold
Public Function ProjTfMode(ByVal tfText As Variant, ByVal incBit As Variant, _
                           ByVal thresholdDeg As Variant) As Variant
    Dim s As String, thr As Double, inc As Double
    On Error GoTo Fail
    thr = SafeNum(thresholdDeg, 5#)
    If thr <= 0# Then thr = 5#
    s = UCase$(Trim$(Replace(CStr(tfText & ""), Chr$(160), " ")))

    If Len(s) > 0 Then
        If Right$(s, 1) = "M" Then
            ProjTfMode = "Magnetic"
            Exit Function
        End If
        If s = "R" Or s = "L" Or s = "HS" Or s = "LS" Or s = "HSR" Or s = "HSL" Then
            ProjTfMode = "Gravity"
            Exit Function
        End If
        If (Left$(s, 1) = "R" Or Left$(s, 1) = "L") And IsNumeric(mid$(s, 2)) Then
            ProjTfMode = "Gravity"
            Exit Function
        End If
        If Len(s) > 1 Then
            If (Right$(s, 1) = "R" Or Right$(s, 1) = "L") And IsNumeric(Left$(s, Len(s) - 1)) Then
                ProjTfMode = "Gravity"
                Exit Function
            End If
        End If
        If IsNumeric(s) Then
            ' Unsigned / signed number without M → highside (Gravity)
            ProjTfMode = "Gravity"
            Exit Function
        End If
    End If

    If Not HasNum(incBit) Then
        ProjTfMode = "Magnetic"
        Exit Function
    End If
    inc = Abs(CDbl(incBit))
    If inc < thr Then
        ProjTfMode = "Magnetic"
    Else
        ProjTfMode = "Gravity"
    End If
    Exit Function
Fail:
    ProjTfMode = CVErr(xlErrValue)
End Function

' Comments: left & Chr(1) & DLR for the Y cell.
'   left  = "Sliding <this stand C, 2 dp> @ <TF>"
'   right = "DLR 0.00"
' Slide metres = ProjInstructedSlideM (capped at C). Leftover rotate is column Z = next-row C − slide.
Public Function ProjSlideComment(ByVal metersToSlide As Variant, ByVal tfText As Variant, _
                                 ByVal burr As Variant, _
                                 Optional ByVal widthChars As Double = 0#, _
                                 Optional ByVal maxSlide As Variant) As Variant
    Dim m As Double, tfShow As String, b As Double
    Dim leftPart As String, rightPart As String
    On Error GoTo Fail

    If Not HasNum(burr) Then ProjSlideComment = "": Exit Function
    b = CDbl(burr)
    m = ProjInstructedSlideM(metersToSlide, tfText, maxSlide)

    tfShow = Trim$(CStr(tfText & ""))
    If Len(tfShow) = 0 Then tfShow = "-"

    ' Under 1 m is not a slide. Leave Y blank instead of "Sliding 0.12m".
    If m < SLIDE_IGNORE_M Then
        ProjSlideComment = ""
        Exit Function
    End If
    leftPart = "Sliding " & Format$(m, "0.00") & "m @ " & tfShow
    rightPart = "DLR " & Format$(b, "0.00")
    ProjSlideComment = leftPart & Chr$(1) & rightPart
    Exit Function
Fail:
    ProjSlideComment = CVErr(xlErrValue)
End Function


' Slide metres drilled between two MDs (e.g. survey depth -> bit depth).

' Sums how much each [Slide From, Slide To] interval overlaps [fromMd, toMd].

' Used for MtrBelow: the slide footage the survey has not seen yet.
Public Function ProjSlideMetersBetween(ByVal fromMd As Variant, ByVal toMd As Variant, _
                                       ByVal slideFrom As Range, ByVal slideTo As Range, _
                                       Optional ByVal tfDep As Variant) As Variant
    Dim lo As Double, hi As Double, tot As Double
    Dim r As Long, a As Double, b As Double
    Dim vA As Variant, vb As Variant
    On Error GoTo Fail
    TouchTfDep tfDep
    If Not HasNum(fromMd) Or Not HasNum(toMd) Then
        ProjSlideMetersBetween = 0#
        Exit Function
    End If
    lo = CDbl(fromMd)
    hi = CDbl(toMd)
    If hi <= lo Then
        ProjSlideMetersBetween = 0#
        Exit Function
    End If
    tot = 0#
    For r = 1 To slideFrom.rows.Count
        vA = slideFrom.Cells(r, 1).value
        vb = slideTo.Cells(r, 1).value
        If HasNum(vA) And HasNum(vb) Then
            a = CDbl(vA)
            b = CDbl(vb)
            If a < lo Then a = lo
            If b > hi Then b = hi
            If b > a Then tot = tot + (b - a)
        End If
    Next r
    ProjSlideMetersBetween = tot
    Exit Function
Fail:
    ProjSlideMetersBetween = CVErr(xlErrNum)
End Function









































