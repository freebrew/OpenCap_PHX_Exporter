Attribute VB_Name = "MDL_TDCalc"
Option Explicit

' ================================================================================
'  MODULE: MDL_TDCalc
'  TD Calculator block on the Data sheet (H49:M55)
'
'    Actual      I53:M53  last Slidesheet survey, projected to bit MD
'                         INC/AZM from survey (or bit W/X if present)
'                         NS/EW/TVD via minimum-curvature integration
'    Planned TD  I54:M54  written by Import Plan from the last _OC_Survey station
'    Actual TD   I55:M55  original sheet formula (merged MD): current MD
'                         + remaining NS/EW distance * RF from heading vs
'                         remaining-vector azimuth. Not a copy of Planned TD.
' ================================================================================

Private Const SS_SHEET As String = "Slidesheet"
Private Const PLAN_SHEET As String = "_OC_Survey"
Private Const PLANSEC_SHEET As String = "_OC_PlanSec"
Private Const DATA_SHEET As String = "Data"

Private Const SURV_ROW_FIRST As Long = 13
Private Const SURV_ROW_LAST As Long = 320
Private Const TGT_FIRST As Long = 2
Private Const TGT_LAST As Long = 5

Private Const SS_COL_BIT_MD As Long = 4    ' D
Private Const SS_COL_MD As Long = 5        ' E
Private Const SS_COL_INC As Long = 6       ' F
Private Const SS_COL_AZM As Long = 7       ' G
Private Const SS_COL_TVD As Long = 8       ' H
Private Const SS_COL_BIT_INC As Long = 23  ' W
Private Const SS_COL_BIT_AZM As Long = 24  ' X

Private Const TD_ROW_ACTUAL As Long = 53
Private Const TD_ROW_PLANNED As Long = 54
Private Const TD_ROW_EST As Long = 55
Private Const TD_COL_FIRST As Long = 9     ' I

Private Const SURVEY_BLOCK As String = "Slidesheet!$D$13:$AQ$320"
Private Const TGT_BLOCK As String = "Slidesheet!$T$2:$Y$5"
Private Const PLANNED_BLOCK As String = "Data!$I$54:$M$54"

' Recovered from Field/Demo before TdEstFinal overwrote I55 (merged I55:M55).
Private Const ACTUAL_TD_FORMULA As String = _
    "=ROUND($I$53+SQRT(($L$54-$L$53)^2+($M$54-$M$53)^2)*LET(" & _
    "dN,$L$54-$L$53,dE,$M$54-$M$53," & _
    "AZt,IF(dN=0,IF(dE>=0,PI()/2,3*PI()/2),MOD(ATAN(dE/dN)+IF(dN<0,PI(),0)+2*PI(),2*PI()))," & _
    "dAZ,MOD(AZt-($K$53*PI()/180)+PI(),2*PI())-PI()," & _
    "DL,ABS(dAZ),IF(DL=0,1,DL/(2*SIN(DL/2)))),2)"

Private Const PI_ As Double = 3.14159265358979
Private Const EPS As Double = 0.0000001

' ================================================================================
'  ACTUAL ROW
' ================================================================================

Public Function TdActualBit(ByVal field As String, _
                            Optional ByVal surveyBlock As Range, _
                            Optional ByVal tgtBlock As Range) As Variant
    Dim md As Double, inc As Double, azm As Double
    Dim n As Double, e As Double, tvd As Double
    Dim ok As Boolean

    If Not surveyBlock Is Nothing Then
        ' Argument exists so Excel dirty-tracks the survey block.
    End If
    If Not tgtBlock Is Nothing Then
    End If

    ok = LastActualAtBit(md, inc, azm, n, e, tvd)
    If Not ok Then
        TdActualBit = ""
        Exit Function
    End If

    Select Case UCase$(Trim$(field))
        Case "MD":         TdActualBit = md
        Case "INC":        TdActualBit = inc
        Case "AZM", "AZI": TdActualBit = azm
        Case "NS", "N":    TdActualBit = n
        Case "EW", "E":    TdActualBit = e
        Case "TVD":        TdActualBit = tvd
        Case Else:         TdActualBit = CVErr(xlErrValue)
    End Select
End Function

' Estimated final TD after min-curve: bit -> next named T2:Y5 target -> Planned TD.
Public Function TdEstFinal(ByVal field As String, _
                           Optional ByVal surveyBlock As Range, _
                           Optional ByVal tgtBlock As Range, _
                           Optional ByVal plannedBlock As Range) As Variant
    Dim aMd As Double, aInc As Double, aAzm As Double
    Dim aN As Double, aE As Double, aV As Double
    Dim tMd As Double, tInc As Double, tAzm As Double
    Dim tN As Double, tE As Double, tv As Double
    Dim pMD As Double, pInc As Double, pAzm As Double
    Dim pN As Double, pE As Double, pV As Double
    Dim hasTgt As Boolean
    Dim c1 As Double, c2 As Double
    Dim estMd As Double

    If Not surveyBlock Is Nothing Then
    End If
    If Not tgtBlock Is Nothing Then
    End If
    If Not plannedBlock Is Nothing Then
    End If

    If Not LastActualAtBit(aMd, aInc, aAzm, aN, aE, aV) Then
        TdEstFinal = ""
        Exit Function
    End If
    If Not ReadPlannedTd(pMD, pInc, pAzm, pN, pE) Then
        TdEstFinal = ""
        Exit Function
    End If
    pV = PlanTvdAtMd(pMD)
    If pV = 0# And aV <> 0# Then pV = aV

    hasTgt = NextNamedTarget(aMd, tMd, tInc, tAzm, tN, tE, tv)
    If hasTgt And tMd > aMd + 0.005 Then
        If tv = 0# Then tv = PlanTvdAtMd(tMd)
        c1 = CourseBetween(aMd, aInc, aAzm, aN, aE, aV, tMd, tInc, tAzm, tN, tE, tv)
        c2 = CourseBetween(tMd, tInc, tAzm, tN, tE, tv, pMD, pInc, pAzm, pN, pE, pV)
        estMd = aMd + c1 + c2
    Else
        c2 = CourseBetween(aMd, aInc, aAzm, aN, aE, aV, pMD, pInc, pAzm, pN, pE, pV)
        estMd = aMd + c2
    End If

    Select Case UCase$(Trim$(field))
        Case "MD":         TdEstFinal = estMd
        Case "INC":        TdEstFinal = pInc
        Case "AZM", "AZI": TdEstFinal = pAzm
        Case "NS", "N":    TdEstFinal = pN
        Case "EW", "E":    TdEstFinal = pE
        Case "TVD":        TdEstFinal = pV
        Case Else:         TdEstFinal = CVErr(xlErrValue)
    End Select
End Function

Public Sub InstallTdActualFormulas()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim prevSU As Boolean
    Dim fields(0 To 4) As String
    Dim i As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    wasProt = SheetUnprotectForVba(ws)

    On Error Resume Next
    ws.Range("I55:M55").UnMerge
    On Error GoTo ErrHandler

    fields(0) = "MD": fields(1) = "INC": fields(2) = "AZM"
    fields(3) = "NS": fields(4) = "EW"

    For i = 0 To 4
        ws.Cells(TD_ROW_ACTUAL, TD_COL_FIRST + i).Formula = _
            "=IFERROR(TdActualBit(""" & fields(i) & """," & SURVEY_BLOCK & "," & TGT_BLOCK & "),"""")"
    Next i

    ws.Range("J55:M55").ClearContents
    ws.Range("I55").Formula2 = ACTUAL_TD_FORMULA
    ws.Range("I55:M55").Merge

    SheetReprotectAfterVba ws, wasProt
    Application.ScreenUpdating = prevSU
    Exit Sub

ErrHandler:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshTdPlannedFromPlan()
    Dim plan As Worksheet
    Dim ws As Worksheet
    Dim cMD As Long, cInc As Long, cAzi As Long, cNS As Long, cEW As Long
    Dim c As Long
    Dim lastR As Long
    Dim wasProt As Boolean
    Dim prevSU As Boolean

    On Error Resume Next
    Set plan = ThisWorkbook.Worksheets(PLAN_SHEET)
    On Error GoTo 0
    If plan Is Nothing Then Exit Sub

    For c = 1 To 20
        Select Case UCase$(Trim$(CStr(plan.Cells(2, c).Value2 & "")))
            Case "MD":         cMD = c
            Case "INC":        cInc = c
            Case "AZI", "AZM": cAzi = c
            Case "NS":         cNS = c
            Case "EW":         cEW = c
        End Select
    Next c
    If cMD = 0 Or cInc = 0 Or cAzi = 0 Or cNS = 0 Or cEW = 0 Then Exit Sub

    lastR = plan.Cells(plan.Rows.Count, cMD).End(xlUp).Row
    Do While lastR >= 3
        If IsNumeric(plan.Cells(lastR, cMD).Value2) _
           And Len(CStr(plan.Cells(lastR, cMD).Value2 & "")) > 0 Then Exit Do
        lastR = lastR - 1
    Loop
    If lastR < 3 Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo PlannedDone

    ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 0).Value = PlanNum(plan, lastR, cMD)
    ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 1).Value = PlanNum(plan, lastR, cInc)
    ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 2).Value = PlanNum(plan, lastR, cAzi)
    ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 3).Value = PlanNum(plan, lastR, cNS)
    ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 4).Value = PlanNum(plan, lastR, cEW)

PlannedDone:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.ScreenUpdating = prevSU
End Sub

Public Sub InstallTdCalc()
    Dim prevSU As Boolean
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    InstallTdActualFormulas
    RefreshTdPlannedFromPlan
    Application.ScreenUpdating = prevSU
End Sub

' ================================================================================
'  Position at bit: integrate surveys, then min-curve survey -> bit
' ================================================================================

Private Function LastActualAtBit(ByRef md As Double, ByRef inc As Double, _
        ByRef azm As Double, ByRef n As Double, ByRef e As Double, _
        ByRef tvd As Double) As Boolean
    Dim ws As Worksheet
    Dim r As Long, lastR As Long
    Dim prevMD As Double, prevInc As Double, prevAzi As Double
    Dim curN As Double, curE As Double, curV As Double
    Dim nSurv As Long
    Dim vMd As Variant, vInc As Variant, vAzi As Variant
    Dim mdv As Double
    Dim dv As Double, dN As Double, dE As Double
    Dim bitMd As Double, bitInc As Double, bitAzm As Double
    Dim vD As Variant, vW As Variant, vX As Variant, vH As Variant

    LastActualAtBit = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    prevMD = 0#: prevInc = 0#: prevAzi = 0#
    curN = 0#: curE = 0#: curV = 0#
    lastR = 0: nSurv = 0

    For r = SURV_ROW_FIRST To SURV_ROW_LAST
        If IsSurveySummaryRow(ws, r) Then GoTo NextSurvey
        vMd = ws.Cells(r, SS_COL_MD).Value2
        vInc = ws.Cells(r, SS_COL_INC).Value2
        vAzi = ws.Cells(r, SS_COL_AZM).Value2
        If Not (IsNumeric(vMd) And IsNumeric(vInc) And IsNumeric(vAzi)) Then GoTo NextSurvey
        If Len(CStr(vMd & "")) = 0 Or Len(CStr(vInc & "")) = 0 Or Len(CStr(vAzi & "")) = 0 Then GoTo NextSurvey
        mdv = CDbl(vMd)
        If mdv <= prevMD Then GoTo NextSurvey
        McStep prevMD, prevInc, prevAzi, mdv, CDbl(vInc), CDbl(vAzi), dv, dN, dE
        curN = curN + dN: curE = curE + dE: curV = curV + dv
        prevMD = mdv: prevInc = CDbl(vInc): prevAzi = CDbl(vAzi)
        lastR = r
        nSurv = nSurv + 1
NextSurvey:
    Next r
    If nSurv < 1 Then Exit Function

    vH = ws.Cells(lastR, SS_COL_TVD).Value2
    If IsNumeric(vH) And Len(CStr(vH & "")) > 0 Then curV = CDbl(vH)

    md = prevMD: inc = prevInc: azm = prevAzi
    n = curN: e = curE: tvd = curV

    vD = ws.Cells(lastR, SS_COL_BIT_MD).Value2
    If IsNumeric(vD) And Len(CStr(vD & "")) > 0 Then
        bitMd = CDbl(vD)
    Else
        bitMd = prevMD
    End If

    vW = ws.Cells(lastR, SS_COL_BIT_INC).Value2
    vX = ws.Cells(lastR, SS_COL_BIT_AZM).Value2
    If IsNumeric(vW) And Len(CStr(vW & "")) > 0 Then bitInc = CDbl(vW) Else bitInc = prevInc
    If IsNumeric(vX) And Len(CStr(vX & "")) > 0 Then bitAzm = CDbl(vX) Else bitAzm = prevAzi

    If bitMd > prevMD + 0.005 Then
        McStep prevMD, prevInc, prevAzi, bitMd, bitInc, bitAzm, dv, dN, dE
        n = n + dN: e = e + dE: tvd = tvd + dv
        md = bitMd: inc = bitInc: azm = bitAzm
    ElseIf bitMd > 0# Then
        md = bitMd
        inc = bitInc
        azm = bitAzm
    End If

    LastActualAtBit = True
End Function

Private Function NextNamedTarget(ByVal bitMd As Double, _
        ByRef tMd As Double, ByRef tInc As Double, ByRef tAzm As Double, _
        ByRef tN As Double, ByRef tE As Double, ByRef tv As Double) As Boolean
    Dim ss As Worksheet
    Dim r As Long
    Dim vMd As Variant, vNm As String

    NextNamedTarget = False
    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    On Error GoTo 0
    If ss Is Nothing Then Exit Function

    For r = TGT_FIRST To TGT_LAST
        vMd = ss.Cells(r, "U").Value2
        vNm = Trim$(CStr(ss.Cells(r, "Y").Value2 & ""))
        If Len(vNm) = 0 Then GoTo NextTgt
        If Not IsNumeric(vMd) Then GoTo NextTgt
        If CDbl(vMd) <= bitMd + 0.005 Then GoTo NextTgt
        tMd = CDbl(vMd)
        tInc = NumOr(ss.Cells(r, "V").Value2, 0#)
        tAzm = NumOr(ss.Cells(r, "W").Value2, 0#)
        tv = NumOr(ss.Cells(r, "X").Value2, 0#)
        If Not TargetCoords(tMd, tN, tE, tv) Then
            tN = 0#: tE = 0#
        End If
        NextNamedTarget = True
        Exit Function
NextTgt:
    Next r
End Function

Private Function TargetCoords(ByVal md As Double, ByRef n As Double, _
        ByRef e As Double, ByRef tvd As Double) As Boolean
    Dim ps As Worksheet
    Dim lastR As Long, r As Long
    Dim vMd As Variant
    Dim bestR As Long
    Dim bestD As Double, d As Double

    TargetCoords = False
    On Error Resume Next
    Set ps = ThisWorkbook.Worksheets(PLANSEC_SHEET)
    On Error GoTo 0
    If Not ps Is Nothing Then
        lastR = ps.Cells(ps.Rows.Count, 1).End(xlUp).Row
        bestD = 1E+30: bestR = 0
        For r = 3 To lastR
            vMd = ps.Cells(r, 1).Value2
            If IsNumeric(vMd) Then
                d = Abs(CDbl(vMd) - md)
                If d < bestD Then
                    bestD = d
                    bestR = r
                End If
            End If
        Next r
        If bestR > 0 And bestD < 1.5 Then
            n = NumOr(ps.Cells(bestR, 5).Value2, 0#)
            e = NumOr(ps.Cells(bestR, 6).Value2, 0#)
            If tvd = 0# Then tvd = NumOr(ps.Cells(bestR, 4).Value2, 0#)
            TargetCoords = True
            Exit Function
        End If
    End If

    TargetCoords = PlanCoordsAtMd(md, n, e, tvd)
End Function

Private Function PlanCoordsAtMd(ByVal md As Double, ByRef n As Double, _
        ByRef e As Double, ByRef tvd As Double) As Boolean
    Dim inc As Double, azm As Double
    PlanCoordsAtMd = PlanAtMd(md, inc, azm, n, e, tvd)
End Function

Private Function PlanTvdAtMd(ByVal md As Double) As Double
    Dim inc As Double, azm As Double, n As Double, e As Double, v As Double
    If PlanAtMd(md, inc, azm, n, e, v) Then
        PlanTvdAtMd = v
    Else
        PlanTvdAtMd = 0#
    End If
End Function

Private Function PlanAtMd(ByVal md As Double, ByRef inc As Double, ByRef azm As Double, _
        ByRef n As Double, ByRef e As Double, ByRef tvd As Double) As Boolean
    Dim plan As Worksheet
    Dim cMD As Long, cInc As Long, cAzi As Long, cTvd As Long, cNS As Long, cEW As Long
    Dim c As Long, lastR As Long, r As Long
    Dim md1 As Double, md2 As Double, f As Double
    Dim i As Long

    PlanAtMd = False
    On Error Resume Next
    Set plan = ThisWorkbook.Worksheets(PLAN_SHEET)
    On Error GoTo 0
    If plan Is Nothing Then Exit Function

    For c = 1 To 20
        Select Case UCase$(Trim$(CStr(plan.Cells(2, c).Value2 & "")))
            Case "MD":         cMD = c
            Case "INC":        cInc = c
            Case "AZI", "AZM": cAzi = c
            Case "TVD":        cTvd = c
            Case "NS":         cNS = c
            Case "EW":         cEW = c
        End Select
    Next c
    If cMD = 0 Or cNS = 0 Or cEW = 0 Then Exit Function

    lastR = plan.Cells(plan.Rows.Count, cMD).End(xlUp).Row
    If lastR < 4 Then Exit Function

    Dim pMD() As Double, pInc() As Double, pAzi() As Double
    Dim pTvd() As Double, pNS() As Double, pEW() As Double
    Dim k As Long
    ReDim pMD(1 To lastR): ReDim pInc(1 To lastR): ReDim pAzi(1 To lastR)
    ReDim pTvd(1 To lastR): ReDim pNS(1 To lastR): ReDim pEW(1 To lastR)
    k = 0
    For r = 3 To lastR
        If IsNumeric(plan.Cells(r, cMD).Value2) Then
            k = k + 1
            pMD(k) = CDbl(plan.Cells(r, cMD).Value2)
            If cInc > 0 Then pInc(k) = NumOr(plan.Cells(r, cInc).Value2, 0#)
            If cAzi > 0 Then pAzi(k) = NumOr(plan.Cells(r, cAzi).Value2, 0#)
            If cTvd > 0 Then pTvd(k) = NumOr(plan.Cells(r, cTvd).Value2, 0#)
            pNS(k) = NumOr(plan.Cells(r, cNS).Value2, 0#)
            pEW(k) = NumOr(plan.Cells(r, cEW).Value2, 0#)
        End If
    Next r
    If k < 1 Then Exit Function

    If md <= pMD(1) Then
        inc = pInc(1): azm = pAzi(1): n = pNS(1): e = pEW(1): tvd = pTvd(1)
        PlanAtMd = True
        Exit Function
    End If
    If md >= pMD(k) Then
        inc = pInc(k): azm = pAzi(k): n = pNS(k): e = pEW(k): tvd = pTvd(k)
        PlanAtMd = True
        Exit Function
    End If
    For i = 1 To k - 1
        If md >= pMD(i) And md <= pMD(i + 1) Then
            md1 = pMD(i): md2 = pMD(i + 1)
            If Abs(md2 - md1) < EPS Then
                f = 0#
            Else
                f = (md - md1) / (md2 - md1)
            End If
            inc = pInc(i) + f * (pInc(i + 1) - pInc(i))
            azm = pAzi(i) + f * (pAzi(i + 1) - pAzi(i))
            n = pNS(i) + f * (pNS(i + 1) - pNS(i))
            e = pEW(i) + f * (pEW(i + 1) - pEW(i))
            tvd = pTvd(i) + f * (pTvd(i + 1) - pTvd(i))
            PlanAtMd = True
            Exit Function
        End If
    Next i
End Function

Private Function ReadPlannedTd(ByRef md As Double, ByRef inc As Double, _
        ByRef azm As Double, ByRef n As Double, ByRef e As Double) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    ReadPlannedTd = False
    If ws Is Nothing Then Exit Function
    If Not IsNumeric(ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST).Value2) Then Exit Function
    md = CDbl(ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST).Value2)
    inc = NumOr(ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 1).Value2, 0#)
    azm = NumOr(ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 2).Value2, 0#)
    n = NumOr(ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 3).Value2, 0#)
    e = NumOr(ws.Cells(TD_ROW_PLANNED, TD_COL_FIRST + 4).Value2, 0#)
    ReadPlannedTd = True
End Function

' Inverse min-curve: MD course from two positions + attitudes.
Private Function CourseBetween(ByVal md1 As Double, ByVal i1 As Double, ByVal a1 As Double, _
        ByVal n1 As Double, ByVal e1 As Double, ByVal v1 As Double, _
        ByVal md2 As Double, ByVal i2 As Double, ByVal a2 As Double, _
        ByVal n2 As Double, ByVal e2 As Double, ByVal v2 As Double) As Double
    Dim chord As Double, beta As Double, inv As Double, planD As Double
    chord = Sqr((n2 - n1) ^ 2 + (e2 - e1) ^ 2 + (v2 - v1) ^ 2)
    beta = Deg2Rad(DoglegAngleDeg(i1, a1, i2, a2))
    If beta > EPS Then
        inv = chord * beta / (2# * Sin(beta / 2#))
    Else
        inv = chord
    End If
    planD = md2 - md1
    If planD < 0# Then planD = 0#
    If inv < planD Then
        CourseBetween = planD
    Else
        CourseBetween = inv
    End If
End Function

Private Function IsSurveySummaryRow(ws As Worksheet, ByVal r As Long) As Boolean
    Dim fd As String, fE As String, fF As String
    On Error Resume Next
    fd = UCase$(CStr(ws.Cells(r, SS_COL_BIT_MD).Formula & ""))
    fE = UCase$(CStr(ws.Cells(r, SS_COL_MD).Formula & ""))
    fF = UCase$(CStr(ws.Cells(r, SS_COL_INC).Formula & ""))
    On Error GoTo 0
    IsSurveySummaryRow = (InStr(1, fd, "LOOKUP", vbBinaryCompare) > 0) _
                      Or (InStr(1, fE, "LOOKUP", vbBinaryCompare) > 0) _
                      Or (InStr(1, fF, "LOOKUP", vbBinaryCompare) > 0)
End Function

Private Function PlanNum(plan As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    Dim v As Variant: v = plan.Cells(r, c).Value2
    If IsNumeric(v) And Len(CStr(v & "")) > 0 Then
        PlanNum = CDbl(v)
    Else
        PlanNum = ""
    End If
End Function

Private Function NumOr(ByVal v As Variant, ByVal defaultValue As Double) As Double
    If IsNumeric(v) And Len(CStr(v & "")) > 0 Then
        NumOr = CDbl(v)
    Else
        NumOr = defaultValue
    End If
End Function

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

Private Sub McStep(ByVal md1 As Double, ByVal i1 As Double, ByVal a1 As Double, _
                   ByVal md2 As Double, ByVal i2 As Double, ByVal a2 As Double, _
                   ByRef dv As Double, ByRef dN As Double, ByRef dE As Double)
    Dim r1 As Double, r2 As Double, b1 As Double, b2 As Double
    Dim cosDL As Double, beta As Double, h As Double
    r1 = Deg2Rad(i1): r2 = Deg2Rad(i2)
    b1 = Deg2Rad(a1): b2 = Deg2Rad(a2)
    cosDL = Cos(r2 - r1) - Sin(r1) * Sin(r2) * (1# - Cos(b2 - b1))
    If cosDL > 1# Then cosDL = 1#
    If cosDL < -1# Then cosDL = -1#
    beta = Application.WorksheetFunction.Acos(cosDL)
    If beta > EPS Then
        h = (md2 - md1) / 2# * (2# / beta * Tan(beta / 2#))
    Else
        h = (md2 - md1) / 2#
    End If
    dv = h * (Cos(r1) + Cos(r2))
    dN = h * (Sin(r1) * Cos(b1) + Sin(r2) * Cos(b2))
    dE = h * (Sin(r1) * Sin(b1) + Sin(r2) * Sin(b2))
End Sub





