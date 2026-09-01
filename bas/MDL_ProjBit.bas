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
' Required-TF aim: nearest full-plan station at least this far past the bit
' (half the 30 m plan grid, so the aim lands 15-45 m ahead).
Private Const TF_LOOKAHEAD_MIN As Double = 15#

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

' Min-curvature step → ΔTVD, ΔN, ΔE
Private Sub MinCurveStep(ByVal md1 As Double, ByVal i1 As Double, ByVal a1 As Double, _
                         ByVal md2 As Double, ByVal i2 As Double, ByVal a2 As Double, _
                         ByRef dTvd As Double, ByRef dN As Double, ByRef dE As Double)
    Dim beta As Double, H As Double, r1 As Double, r2 As Double, b1 As Double, b2 As Double
    beta = Deg2Rad(DoglegAngleDeg(i1, a1, i2, a2))
    If beta > 0.0000001 Then
        H = (md2 - md1) / 2# * (2# / beta * Tan(beta / 2#))
    Else
        H = (md2 - md1) / 2#
    End If
    r1 = Deg2Rad(i1): r2 = Deg2Rad(i2)
    b1 = Deg2Rad(a1): b2 = Deg2Rad(a2)
    dTvd = H * (Cos(r1) + Cos(r2))
    dN = H * (Sin(r1) * Cos(b1) + Sin(r2) * Cos(b2))
    dE = H * (Sin(r1) * Sin(b1) + Sin(r2) * Sin(b2))
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
    HasNum = (Not isError(v)) And IsNumeric(v) And Len(Trim$(CStr(v & ""))) > 0
End Function

' Last-5 numeric Q (motor output, °/30 m of slide) walking up this row.
' A stand counts when Q is a real number > 0 — same five values you see in Q.
' Need 2+ samples; else this-row Q if it is numeric > 0; else blank.
' mRng is passed so Excel dirties this cell when metres-seen changes (Q = C×J/M).
' qRng/mRng are the same-height columns ending on this row (Excel owns recalc).
Public Function ProjRollingMotorOut(ByVal qRng As Range, ByVal mRng As Range) As Variant
    Const N_WANT As Long = 5
    Const Q_LO As Double = 0#
    Dim n As Long
    Dim qArr As Variant
    Dim mArr As Variant
    Dim i As Long
    Dim taken As Long
    Dim tot As Double
    Dim qv As Variant
    Dim mv As Variant
    Dim thisOk As Boolean
    Dim thisQ As Double

    On Error GoTo Fail
    If qRng Is Nothing Or mRng Is Nothing Then
        ProjRollingMotorOut = ""
        Exit Function
    End If
    n = qRng.Rows.Count
    If n < 1 Or mRng.Rows.Count <> n Then
        ProjRollingMotorOut = ""
        Exit Function
    End If

    qArr = qRng.Value2
    mArr = mRng.Value2
    taken = 0
    tot = 0#
    thisOk = False
    thisQ = 0#

    For i = n To 1 Step -1
        RollingSample qArr, mArr, i, n, qv, mv
        If YieldSampleOk(qv, Q_LO) Then
            tot = tot + CDbl(qv)
            taken = taken + 1
            If i = n Then
                thisOk = True
                thisQ = CDbl(qv)
            End If
            If taken >= N_WANT Then Exit For
        End If
    Next i

    If taken >= 2 Then
        ProjRollingMotorOut = tot / CDbl(taken)
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
                          ByRef qv As Variant, ByRef mv As Variant)
    If n = 1 And Not IsArray(qArr) Then
        qv = qArr
        mv = mArr
    Else
        qv = qArr(i, 1)
        mv = mArr(i, 1)
    End If
End Sub

Private Function YieldSampleOk(ByVal qv As Variant, ByVal qLo As Double) As Boolean
    YieldSampleOk = False
    If Not HasNum(qv) Then Exit Function
    If CDbl(qv) <= qLo Then Exit Function
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

' Parse toolface text: 190M (magnetic deg), R/L highside, -30 left, 30 right
Public Function ProjParseTF(ByVal tfText As Variant) As Variant
    Dim s As String, n As Double
    On Error GoTo Fail
    If isError(tfText) Then ProjParseTF = CVErr(xlErrValue): Exit Function
    s = UCase$(Trim$(Replace(CStr(tfText & ""), Chr$(160), " ")))
    If Len(s) = 0 Then ProjParseTF = "": Exit Function

    If right$(s, 1) = "M" Then
        s = Left$(s, Len(s) - 1)
        If IsNumeric(s) Then
            ProjParseTF = CDbl(s)
            Exit Function
        End If
    End If

    If s = "R" Or s = "HS" Then ProjParseTF = 0#: Exit Function
    If s = "L" Or s = "LS" Then ProjParseTF = 180#: Exit Function

    If Left$(s, 1) = "R" And IsNumeric(mid$(s, 2)) Then
        ProjParseTF = Abs(CDbl(mid$(s, 2)))
        Exit Function
    End If
    If Left$(s, 1) = "L" And IsNumeric(mid$(s, 2)) Then
        ProjParseTF = -Abs(CDbl(mid$(s, 2)))
        Exit Function
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
'   dial blank/0  -> demonstrated N/O (rotary stand, old behavior)
'   same sign     -> smaller magnitude of dial vs demonstrated (never project
'                    more turn than the hole has shown it delivers)
'   opposite sign / no N -> dial (demonstrated TF is rotary noise; trust intent)
' N/O arrive as arguments so Excel owns the recalc dependency — never read
' via Application.Caller (no dependency edge → stale AZM on recalc).
Private Function ResolveWalkTf(ByVal tfDeg As Variant, _
                               ByVal seenTf As Variant, ByVal seenLR As Variant) As Double
    Dim tfD As Double, tfE As Double
    tfD = SafeNum(tfDeg)
    tfE = SignedSeenTf(seenTf, seenLR)
    If Abs(tfD) < 0.05 Then
        ResolveWalkTf = tfE
    ElseIf Abs(tfE) < 0.05 Then
        ResolveWalkTf = tfD
    ElseIf tfD * tfE > 0# And Abs(tfE) < Abs(tfD) Then
        ResolveWalkTf = tfE
    Else
        ResolveWalkTf = tfD
    End If
End Function

Public Function ProjIncAtBit(ByVal survInc As Variant, ByVal dls As Variant, _
                             ByVal course As Variant, ByVal mSeen As Variant, _
                             ByVal mBelow As Variant, ByVal tfDeg As Variant, _
                             Optional ByVal motorOut As Variant) As Variant
    Dim ci As Double, dg As Double, co As Double, ms As Double, mb As Double, tf As Double
    Dim dB As Double
    On Error GoTo Fail
    If Not HasNum(survInc) Then ProjIncAtBit = "": Exit Function
    ci = CDbl(survInc)
    dg = SafeNum(dls): co = SafeNum(course): ms = SafeNum(mSeen)
    mb = SafeNum(mBelow): tf = SafeNum(tfDeg)
    If HasNum(motorOut) Then
        dB = DoglegBelowYield(CDbl(motorOut), ms, mb)
    Else
        dB = DoglegBelow(dg, co, ms, mb)
    End If
    ProjIncAtBit = ci + dB * Cos(Deg2Rad(tf))
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
                             Optional ByVal motorOut As Variant) As Variant
    Dim ci As Double, ca As Double, dg As Double, co As Double, ms As Double, mb As Double, tf As Double
    Dim dB As Double
    On Error GoTo Fail
    If Not HasNum(survInc) Or Not HasNum(survAzm) Then ProjAzmAtBit = "": Exit Function
    ci = CDbl(survInc)
    ca = CDbl(survAzm)
    dg = SafeNum(dls): co = SafeNum(course): ms = SafeNum(mSeen)
    mb = SafeNum(mBelow)
    tf = ResolveWalkTf(tfDeg, seenTf, seenLR)
    If HasNum(motorOut) Then
        dB = DoglegBelowYield(CDbl(motorOut), ms, mb)
    Else
        dB = DoglegBelow(dg, co, ms, mb)
    End If
    ProjAzmAtBit = Wrap360(ca + AzmWalkDeg(ci, dB, tf))
    Exit Function
Fail:
    ProjAzmAtBit = CVErr(xlErrNum)
End Function

Public Function ProjTvdAtBit(ByVal survMd As Variant, ByVal survInc As Variant, _
                             ByVal survAzm As Variant, ByVal survTvd As Variant, _
                             ByVal bitMD As Variant, ByVal incBit As Variant, _
                             ByVal azmBit As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    On Error GoTo Fail
    If Not HasNum(survMd) Or Not HasNum(survInc) Or Not HasNum(survAzm) Or Not HasNum(survTvd) Then
        ProjTvdAtBit = "": Exit Function
    End If
    If Not HasNum(bitMD) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjTvdAtBit = "": Exit Function
    End If
    MinCurveStep CDbl(survMd), CDbl(survInc), CDbl(survAzm), _
                 CDbl(bitMD), CDbl(incBit), CDbl(azmBit), dt, dN, dE
    ProjTvdAtBit = CDbl(survTvd) + dt
    Exit Function
Fail:
    ProjTvdAtBit = CVErr(xlErrNum)
End Function

' Cumulative N from prior N + min-curve ΔN between surveys (first survey → 0)
Public Function ProjCumN(ByVal prevN As Variant, ByVal md1 As Variant, ByVal i1 As Variant, _
                         ByVal a1 As Variant, ByVal md2 As Variant, ByVal i2 As Variant, _
                         ByVal a2 As Variant, ByVal isFirst As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    On Error GoTo Fail
    If Not HasNum(md2) Or Not HasNum(i2) Or Not HasNum(a2) Then ProjCumN = "": Exit Function
    If CBool(isFirst) Then
        ProjCumN = 0#
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
    On Error GoTo Fail
    If Not HasNum(md2) Or Not HasNum(i2) Or Not HasNum(a2) Then ProjCumE = "": Exit Function
    If CBool(isFirst) Then
        ProjCumE = 0#
        Exit Function
    End If
    If Not HasNum(md1) Or Not HasNum(i1) Or Not HasNum(a1) Then ProjCumE = "": Exit Function
    MinCurveStep CDbl(md1), CDbl(i1), CDbl(a1), CDbl(md2), CDbl(i2), CDbl(a2), dt, dN, dE
    ProjCumE = SafeNum(prevE) + dE
    Exit Function
Fail:
    ProjCumE = CVErr(xlErrNum)
End Function

' Linear INC/AZM interpolate on plan (PROJBIT att)
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
    t = (q - m(j - 1)) / (m(j) - m(j - 1))
    outI = i(j - 1) + t * (i(j) - i(j - 1))
    outA = Wrap360(a(j - 1) + t * Wrap180(a(j) - a(j - 1)))
End Sub

Private Function PlanTvdLin(ByRef m() As Double, ByRef tv() As Double, _
                            ByVal n As Long, ByVal q As Double) As Double
    Dim j As Long, f As Double
    If n < 1 Then PlanTvdLin = 0#: Exit Function
    If q <= m(1) Then PlanTvdLin = tv(1): Exit Function
    If q >= m(n) Then PlanTvdLin = tv(n): Exit Function
    j = 2
    Do While j <= n And m(j) < q
        j = j + 1
    Loop
    If j > n Then j = n
    f = (q - m(j - 1)) / (m(j) - m(j - 1))
    PlanTvdLin = tv(j - 1) + f * (tv(j) - tv(j - 1))
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
    For r = 1 To mdR.Rows.Count
        v = mdR.Cells(r, 1).Value
        If HasNum(v) Then
            n = n + 1
            If n > 100 Then Exit For
            m(n) = CDbl(v)
            i(n) = SafeNum(incR.Cells(r, 1).Value)
            a(n) = SafeNum(azmR.Cells(r, 1).Value)
            tv(n) = SafeNum(tvdR.Cells(r, 1).Value)
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
Private Function TargetIsExhausted(ByVal bitMD As Double, ByVal incBit As Double, _
                                   ByVal tgtMd As Double, ByVal tgtInc As Double, _
                                   ByVal tgtTvd As Double, ByVal tvdBit As Variant) As Boolean
    Dim dist As Double, dTvd As Double
    TargetIsExhausted = False
    If incBit + MIN_AIM_DINC >= tgtInc Then
        dist = tgtMd - bitMD
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
                               ByVal bitMD As Double, ByVal incBit As Double, _
                               ByVal tvdBit As Variant) As Long
    Dim k As Long
    k = 1
    Do While k <= n And m(k) <= bitMD
        k = k + 1
    Loop
    Do While k <= n
        If k = n Then
            FirstAimIndex = k
            Exit Function
        End If
        If Not TargetIsExhausted(bitMD, incBit, m(k), i(k), tv(k), tvdBit) Then
            FirstAimIndex = k
            Exit Function
        End If
        k = k + 1
    Loop
    FirstAimIndex = n + 1
End Function

' Active target MD ahead of bit (first MD > bitMd), or blank. Legacy MD-only picker.
Public Function ProjActiveTargetMd(ByVal bitMD As Variant, ByVal tgtMd As Range) As Variant
    Dim r As Long, bm As Double, v As Variant
    On Error GoTo Fail
    If Not HasNum(bitMD) Then ProjActiveTargetMd = "": Exit Function
    bm = CDbl(bitMD)
    For r = 1 To tgtMd.Rows.Count
        v = tgtMd.Cells(r, 1).Value
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
Public Function ProjActiveTargetMdEx(ByVal bitMD As Variant, ByVal incBit As Variant, _
                                     ByVal tvdBit As Variant, ByVal tgtMd As Range, _
                                     ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                                     ByVal tgtTvd As Range, _
                                     Optional ByVal survInc As Variant, _
                                     Optional ByVal survTvd As Variant) As Variant
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long, k As Long, bm As Double, ib As Double
    Dim gateInc As Double, gateTvd As Variant
    On Error GoTo Fail
    If Not HasNum(bitMD) Then ProjActiveTargetMdEx = "": Exit Function
    bm = CDbl(bitMD)
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
Private Function ComputeAim(ByVal bitMD As Double, ByVal incBit As Double, _
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

    k = FirstAimIndex(m, i, tv, n, bitMD, incBit, tvdBit)
    If k > n Then Exit Function

    q = m(k)
    ti = i(k)
    ta = a(k)
    dist = q - bitMD
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
            dist = q - bitMD
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
    aimTvd = PlanTvdLin(m, tv, n, q)
    ComputeAim = True
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
' reqDLS is the 3D dogleg °/30 m to the next plan station ahead of the bit.
' Y comment uses this AS value as-is (no required-TF cosine inflate).
Public Function ProjMetersToSlide(ByVal bitMD As Variant, ByVal incBit As Variant, _
                                  ByVal azmBit As Variant, ByVal motorOut As Variant, _
                                  ByVal course As Variant, ByVal tgtMd As Range, _
                                  ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                                  ByVal tgtTvd As Range, _
                                  Optional ByVal tvdBit As Variant) As Variant
    Dim bm As Double, ib As Double, aB As Double, mo As Double, co As Double
    Dim dist As Double, req As Double, burr As Double, reqTf As Double
    Dim aimInc As Double, aimTvd As Double
    On Error GoTo Fail

    If Not HasNum(bitMD) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjMetersToSlide = "": Exit Function
    End If
    bm = CDbl(bitMD): ib = CDbl(incBit): aB = CDbl(azmBit)
    mo = SafeNum(motorOut): co = SafeNum(course)
    If mo <= EPS Or co <= EPS Then ProjMetersToSlide = "": Exit Function

    If Not ComputeAim(bm, ib, aB, mo, co, tgtMd, tgtInc, tgtAzm, tgtTvd, _
                      dist, req, burr, reqTf, aimInc, aimTvd, False, tvdBit) Then
        ProjMetersToSlide = "": Exit Function
    End If

    If Abs(req) <= EPS Then
        ProjMetersToSlide = 0#
        Exit Function
    End If

    ' 3D dogleg required (°/30 m) / motor × this stand. Cannot slide more than C.
    ProjMetersToSlide = Abs(req) / mo * co
    If ProjMetersToSlide > co Then ProjMetersToSlide = co
    Exit Function
Fail:
    ProjMetersToSlide = CVErr(xlErrNum)
End Function

' Meters remain to rotate in the current drill set (course length, col C):
'   max(0, courseLen - metersToSlide)
Public Function ProjMetersRemainToRotate(ByVal bitMD As Variant, ByVal incBit As Variant, _
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

    slide = ProjMetersToSlide(bitMD, incBit, azmBit, motorOut, course, _
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

' 3D dogleg required (°/30 m) to the next plan station ahead of the bit.
' Same rate Meters To Slide uses: |reqDLS| / motor × course.
' Aim is not walked forward. Inc-only / TVD-arc build is not this number.
Public Function ProjBurr(ByVal bitMD As Variant, ByVal incBit As Variant, _
                         ByVal azmBit As Variant, ByVal motorOut As Variant, _
                         ByVal course As Variant, ByVal tgtMd As Range, _
                         ByVal tgtInc As Range, ByVal tgtAzm As Range, _
                         ByVal tgtTvd As Range, Optional ByVal tvdBit As Variant) As Variant
    Dim bm As Double, ib As Double, aB As Double, mo As Double, co As Double
    Dim dist As Double, req As Double, burr As Double, reqTf As Double
    Dim aimInc As Double, aimTvd As Double
    On Error GoTo Fail

    If Not HasNum(bitMD) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjBurr = "": Exit Function
    End If
    bm = CDbl(bitMD): ib = CDbl(incBit): aB = CDbl(azmBit)
    mo = SafeNum(motorOut): co = SafeNum(course)

    If Not ComputeAim(bm, ib, aB, mo, co, tgtMd, tgtInc, tgtAzm, tgtTvd, _
                      dist, req, burr, reqTf, aimInc, aimTvd, False, tvdBit) Then
        ProjBurr = "": Exit Function
    End If

    ProjBurr = req
    Exit Function
Fail:
    ProjBurr = CVErr(xlErrNum)
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
                         ByVal bitMD As Variant, ByVal incBit As Variant, _
                         ByVal azmBit As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    On Error GoTo Fail
    If Not HasNum(survMd) Or Not HasNum(survInc) Or Not HasNum(survAzm) Then
        ProjBitN = "": Exit Function
    End If
    If Not HasNum(cumN) Or Not HasNum(bitMD) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjBitN = "": Exit Function
    End If
    MinCurveStep CDbl(survMd), CDbl(survInc), CDbl(survAzm), _
                 CDbl(bitMD), CDbl(incBit), CDbl(azmBit), dt, dN, dE
    ProjBitN = CDbl(cumN) + dN
    Exit Function
Fail:
    ProjBitN = CVErr(xlErrNum)
End Function

' Bit Easting = survey CumE + min-curve ΔE (surv → bit)
Public Function ProjBitE(ByVal survMd As Variant, ByVal survInc As Variant, _
                         ByVal survAzm As Variant, ByVal cumE As Variant, _
                         ByVal bitMD As Variant, ByVal incBit As Variant, _
                         ByVal azmBit As Variant) As Variant
    Dim dt As Double, dN As Double, dE As Double
    On Error GoTo Fail
    If Not HasNum(survMd) Or Not HasNum(survInc) Or Not HasNum(survAzm) Then
        ProjBitE = "": Exit Function
    End If
    If Not HasNum(cumE) Or Not HasNum(bitMD) Or Not HasNum(incBit) Or Not HasNum(azmBit) Then
        ProjBitE = "": Exit Function
    End If
    MinCurveStep CDbl(survMd), CDbl(survInc), CDbl(survAzm), _
                 CDbl(bitMD), CDbl(incBit), CDbl(azmBit), dt, dN, dE
    ProjBitE = CDbl(cumE) + dE
    Exit Function
Fail:
    ProjBitE = CVErr(xlErrNum)
End Function

' Required toolface to intercept the next plan target.
'
' 1) Project INC/AZM at bit from the survey using the USER toolface (tfText),
'    so the bit attitude reflects the slide already drilled below the survey.
' 2) Aim at the NEAR-TERM PLAN ATTITUDE: the nearest _OC_Survey plan station
'    at least TF_LOOKAHEAD_MIN past the bit. Tracking the plan means matching
'    its attitude station by station — azimuth gets corrected now, where
'    turning is cheap — NOT the lazy great-circle to a distant named target
'    (that let azimuth converge only at the end and recommended R8 in a build
'    the driller had to slide at 30R). Named targets are only the fallback
'    when the plan sheet is empty.
' 3) From bit attitude to aim attitude, use the industry-standard
'    (Well Seeker) spherical toolface — the angle from high side to the
'    dogleg plane that carries (incBit, azmBit) onto (aimInc, aimAzm):
'       TF = Atan2( sinI2·cosI1·cosΔA − sinI1·cosI2 ,  sinI2·sinΔA )
'    (ToolfaceBetweenDeg). A degree of azimuth walk only moves the bit
'    sin(Inc) as far as a degree of build, so raw Atan2(ΔInc, ΔAzm) is wrong —
'    it over-weights the turn (~2.4x at Inc 25°) and disagreed with the office.
'
' Display: Gravity → R## / L## / 0 ; Magnetic (low inc) → ###M
' Required TF from bit → aim, using projected Inc@Bit / Azm@Bit
' (those already embed the user toolface via ProjIncAtBit / ProjAzmAtBit).
Public Function ProjRequiredTf(ByVal bitMD As Variant, _
                               ByVal incAtBit As Variant, ByVal azmAtBit As Variant, _
                               ByVal tgtMd As Range, ByVal tgtInc As Range, _
                               ByVal tgtAzm As Range, ByVal tgtTvd As Range, _
                               Optional ByVal thresholdDeg As Variant, _
                               Optional ByVal tvdBit As Variant) As Variant
    Dim bm As Double, ib As Double, aB As Double
    Dim m() As Double, i() As Double, a() As Double, tv() As Double
    Dim n As Long, k As Long
    Dim reqTf As Double
    On Error GoTo Fail

    If Not HasNum(bitMD) Or Not HasNum(incAtBit) Or Not HasNum(azmAtBit) Then
        ProjRequiredTf = "": Exit Function
    End If
    bm = CDbl(bitMD)
    ib = CDbl(incAtBit)
    aB = CDbl(azmAtBit)

    ' Preferred aim: nearest full-plan station past bit + lookahead.
    Dim pMD() As Double, pInc() As Double, pAzi() As Double
    Dim pTvd() As Double, pNS() As Double, pEW() As Double
    Dim np As Long, idx As Long
    Dim ti As Double, ta As Double, haveAim As Boolean
    haveAim = False
    np = MDL_PlanGauge.PG_LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
    If np >= 2 Then
        idx = 0
        Do While idx < np - 1 And pMD(idx) < bm + TF_LOOKAHEAD_MIN
            idx = idx + 1
        Loop
        ti = pInc(idx)
        ta = pAzi(idx)
        haveAim = True
    End If

    ' Fallback: next named plan target (old behavior) if no plan is loaded.
    If Not haveAim Then
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
    End If

    ' Spherical TF needs no MD: it is the direction of the dogleg plane from
    ' bit attitude to aim attitude, exactly what Well Seeker tabulates.
    If DoglegAngleDeg(ib, aB, ti, ta) < 0.000001 Then
        reqTf = 0#   ' already pointed at the aim attitude
    Else
        reqTf = ToolfaceBetweenDeg(ib, aB, ti, ta)
    End If

    ProjRequiredTf = FormatRequiredTfDisplay(reqTf, ib, aB, thresholdDeg)
    Exit Function
Fail:
    ProjRequiredTf = CVErr(xlErrNum)
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
        If right$(s, 1) = "M" Then
            ProjTfMode = "Magnetic"
            Exit Function
        End If
        If s = "R" Or s = "L" Or s = "HS" Or s = "LS" Then
            ProjTfMode = "Gravity"
            Exit Function
        End If
        If (Left$(s, 1) = "R" Or Left$(s, 1) = "L") And IsNumeric(mid$(s, 2)) Then
            ProjTfMode = "Gravity"
            Exit Function
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

' Comments: left & Chr(1) & BURR for the Y cell.
'   left  = "Sliding <this stand C, 2 dp> @ <TF>"
'   right = "BURR 0.00"
' Slide metres = ProjInstructedSlideM (capped at C). Leftover rotate is column Z = C − slide.
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

    leftPart = "Sliding " & Format$(m, "0.00") & "m @ " & tfShow
    rightPart = "BURR " & Format$(b, "0.00")
    ProjSlideComment = leftPart & Chr$(1) & rightPart
    Exit Function
Fail:
    ProjSlideComment = CVErr(xlErrValue)
End Function


' Slide metres drilled between two MDs (e.g. survey depth -> bit depth).

' Sums how much each [Slide From, Slide To] interval overlaps [fromMd, toMd].

' Used for MtrBelow: the slide footage the survey has not seen yet.
Public Function ProjSlideMetersBetween(ByVal fromMd As Variant, ByVal toMd As Variant, _
                                       ByVal slideFrom As Range, ByVal slideTo As Range) As Variant
    Dim lo As Double, hi As Double, tot As Double
    Dim r As Long, a As Double, b As Double
    Dim va As Variant, vb As Variant
    On Error GoTo Fail
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
    For r = 1 To slideFrom.Rows.Count
        va = slideFrom.Cells(r, 1).Value
        vb = slideTo.Cells(r, 1).Value
        If HasNum(va) And HasNum(vb) Then
            a = CDbl(va)
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









