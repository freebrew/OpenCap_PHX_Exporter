Attribute VB_Name = "MDL_AcPoc_Check"
Option Explicit

' Self-checks. Any FAIL increments gAcPocFailCount via AcPoc_Log.

Public Sub AcPoc_SelfCheck()
    AcPoc_CheckPlan
    AcPoc_CheckQualifyTable
    AcPoc_CheckDetailCoverage
    AcPoc_CheckSummaryVsDetail
    AcPoc_CheckFrame
End Sub

Private Sub AcPoc_CheckPlan()
    Dim pl As Worksheet
    Dim sec As Worksheet
    Dim nPl As Long, nSec As Long
    Dim lastMd As Double, td As Double
    Dim i As Long
    Dim heelOk As Boolean

    Set pl = AcPoc_EnsureSheet("_POC_Plan", True)
    Set sec = AcPoc_EnsureSheet("_POC_PlanSec", True)
    nPl = AcPoc_LastDataRow(pl, 1) - 1
    nSec = AcPoc_LastDataRow(sec, 1) - 1

    If nPl <= 200 Then
        AcPoc_Log "Check plan", "FAIL", "Planned Survey count " & nPl & " (need > 200)"
    Else
        AcPoc_Log "Check plan", "OK", "Planned Survey count " & nPl
    End If

    lastMd = CDbl(pl.Cells(nPl + 1, 1).Value2)
    td = 0#
    For i = 2 To nSec + 1
        If CDbl(sec.Cells(i, 1).Value2) > td Then td = CDbl(sec.Cells(i, 1).Value2)
        If Abs(CDbl(sec.Cells(i, 1).Value2) - 3187.02) < 0.02 Then heelOk = True
        If StrComp(CStr(sec.Cells(i, 7).Value2 & ""), "HEEL", vbTextCompare) = 0 Then heelOk = True
    Next i

    If Abs(lastMd - td) > 0.05 Then
        AcPoc_Log "Check plan", "FAIL", "Last survey MD " & Format$(lastMd, "0.00") & _
            " != Plan Sections TD " & Format$(td, "0.00")
    Else
        AcPoc_Log "Check plan", "OK", "Last MD matches Plan Sections TD " & Format$(td, "0.00")
    End If

    If Not heelOk Then
        AcPoc_Log "Check plan", "FAIL", "Heel 3187.02 not found in Plan Sections"
    Else
        AcPoc_Log "Check plan", "OK", "Heel station present"
    End If
End Sub

Private Sub AcPoc_CheckQualifyTable()
    Dim q As Worksheet
    Dim n As Long
    Dim r As Long
    Dim well As String
    Dim md As Double, sf As Double

    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    n = AcPoc_LastDataRow(q, 1) - 1
    If n < 1 Then
        AcPoc_Log "Check qualify", "FAIL", "Qualify table empty"
        Exit Sub
    End If
    ' First kept row by depth on 35781 is 35780 @ 633.33 SF 2.182
    well = CStr(q.Cells(2, 1).Value2 & "")
    md = CDbl(q.Cells(2, 2).Value2)
    sf = CDbl(q.Cells(2, 4).Value2)
    If InStr(1, well, "35780", vbTextCompare) = 0 Then
        AcPoc_Log "Check qualify", "WARN", "First depth row is not 35780: " & well
    End If
    If Abs(md - 633.33) > 0.05 Then
        AcPoc_Log "Check qualify", "WARN", "First RefMD=" & Format$(md, "0.00") & " (expected 633.33 on 35781)"
    End If

    Dim saw35782 As Boolean
    For r = 2 To n + 1
        If InStr(1, CStr(q.Cells(r, 1).Value2 & ""), "35782", vbTextCompare) > 0 Then
            saw35782 = True
            If Abs(CDbl(q.Cells(r, 4).Value2) - 1.619) > 0.005 Then
                AcPoc_Log "Check qualify", "FAIL", "35782 SF=" & q.Cells(r, 4).Value2 & " expected 1.619"
            Else
                AcPoc_Log "Check qualify", "OK", "35782 P3 kept at SF 1.619"
            End If
        End If
    Next r
    If Not saw35782 Then AcPoc_Log "Check qualify", "FAIL", "35782 P3 missing from qualify table"
    AcPoc_Log "Check qualify", "OK", n & " qualify rows (capacity " & AcPoc_Capacity() & ")"
End Sub

Private Sub AcPoc_CheckDetailCoverage()
    Dim q As Worksheet
    Dim d As Worksheet
    Dim nQ As Long, nD As Long
    Dim i As Long, j As Long
    Dim well As String
    Dim found As Boolean
    Dim missing As Long

    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    Set d = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    nQ = AcPoc_LastDataRow(q, 1) - 1
    nD = AcPoc_LastDataRow(d, 1) - 1
    If nD < 1 Then
        AcPoc_Log "Check coverage", "FAIL", "No detail rows"
        Exit Sub
    End If

    For i = 2 To nQ + 1
        If StrComp(CStr(q.Cells(i, 7).Value2 & ""), "Y", vbTextCompare) <> 0 Then GoTo NextQ
        well = CStr(q.Cells(i, 1).Value2 & "")
        found = False
        For j = 2 To nD + 1
            If AcPoc_WellsMatch(well, CStr(d.Cells(j, 1).Value2 & "")) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then
            missing = missing + 1
            AcPoc_Log "Check coverage", "FAIL", "No detail page for " & well
        End If
NextQ:
    Next i
    If missing = 0 Then AcPoc_Log "Check coverage", "OK", "Every plotted well has a detail page"
End Sub

Private Sub AcPoc_CheckSummaryVsDetail()
    Dim q As Worksheet
    Dim d As Worksheet
    Dim nQ As Long, nD As Long
    Dim i As Long, j As Long
    Dim well As String
    Dim md As Double, sf As Double, bc As Double
    Dim hit As Boolean
    Dim bestSf As Double
    Dim rowSf As Double

    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    Set d = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    nQ = AcPoc_LastDataRow(q, 1) - 1
    nD = AcPoc_LastDataRow(d, 1) - 1

    For i = 2 To nQ + 1
        well = CStr(q.Cells(i, 1).Value2 & "")
        md = CDbl(q.Cells(i, 2).Value2)
        sf = CDbl(q.Cells(i, 4).Value2)
        bc = CDbl(q.Cells(i, 5).Value2)
        hit = False
        For j = 2 To nD + 1
            If Not AcPoc_WellsMatch(well, CStr(d.Cells(j, 1).Value2 & "")) Then GoTo SkipDetMatch
            If Abs(CDbl(d.Cells(j, 2).Value2) - md) > 0.05 Then GoTo SkipDetMatch
            rowSf = 0#
            If IsNumeric(d.Cells(j, 12).Value2) Then rowSf = CDbl(d.Cells(j, 12).Value2)
            If rowSf > 0# Then
                If Abs(rowSf - sf) > 0.02 Then
                    AcPoc_Log "Check SF row", "FAIL", well & " RefMD " & Format$(md, "0.00") & _
                        " detail SF=" & Format$(rowSf, "0.000") & " summary SF=" & Format$(sf, "0.000")
                End If
            End If
            If IsNumeric(d.Cells(j, 9).Value2) Then
                If Abs(CDbl(d.Cells(j, 9).Value2) - bc) > 0.05 Then
                    AcPoc_Log "Check SF row", "WARN", well & " C2C detail=" & d.Cells(j, 9).Value2 & _
                        " summary=" & bc
                End If
            End If
            hit = True
            Exit For
SkipDetMatch:
        Next j
        If Not hit Then
            AcPoc_Log "Check SF row", "WARN", "No detail station at RefMD " & Format$(md, "0.00") & " for " & well
        End If
    Next i
    AcPoc_Log "Check SF row", "OK", "Summary vs detail cross-check finished"
End Sub

Private Sub AcPoc_CheckFrame()
    Dim wsPlan As Worksheet
    Dim wsDet As Worksheet
    Dim wsQ As Worksheet
    Dim nPl As Long, nD As Long, nQ As Long
    Dim pMd() As Double, pN() As Double, pE() As Double, pT() As Double
    Dim i As Long, j As Long, nKeep As Long
    Dim well As String
    Dim plotIt As Boolean
    Dim okRow As Boolean
    Dim refMd As Double, offN As Double, offE As Double, offT As Double, bc As Double
    Dim planN As Double, planE As Double, planT As Double
    Dim dist As Double, resid As Double
    Dim worst As Double, nChk As Long, nBad As Long
    Dim wWorst As Double
    Dim wName As String
    Dim wk As String
    Dim qKey() As String
    Dim det As Variant
    Dim tmp As Double

    AcPoc_Log "Frame proof", "START", "Entered frame proof"
    Set wsPlan = ThisWorkbook.Worksheets("_POC_Plan")
    Set wsDet = ThisWorkbook.Worksheets("_POC_AC_Detail")
    Set wsQ = ThisWorkbook.Worksheets("_POC_Qualify")

    nPl = AcPoc_LastDataRow(wsPlan, 1) - 1
    nD = AcPoc_LastDataRow(wsDet, 1) - 1
    nQ = AcPoc_LastDataRow(wsQ, 1) - 1
    AcPoc_Log "Frame proof", "START", "Comparing " & nD & " detail stations to " & nPl & " plan stations"
    If nPl < 2 Or nD < 2 Or nQ < 1 Then
        AcPoc_Log "Frame proof", "FAIL", "Not enough plan/detail/qualify rows"
        Exit Sub
    End If

    ReDim pMd(1 To nPl)
    ReDim pN(1 To nPl)
    ReDim pE(1 To nPl)
    ReDim pT(1 To nPl)
    nKeep = 0
    For i = 1 To nPl
        If AcPoc_TryDbl(wsPlan.Cells(i + 1, 1).Value2, tmp) Then
            nKeep = nKeep + 1
            pMd(nKeep) = tmp
            If Not AcPoc_TryDbl(wsPlan.Cells(i + 1, 5).Value2, pN(nKeep)) Then pN(nKeep) = 0#
            If Not AcPoc_TryDbl(wsPlan.Cells(i + 1, 6).Value2, pE(nKeep)) Then pE(nKeep) = 0#
            If Not AcPoc_TryDbl(wsPlan.Cells(i + 1, 4).Value2, pT(nKeep)) Then pT(nKeep) = 0#
        End If
    Next i
    If nKeep < 2 Then
        AcPoc_Log "Frame proof", "FAIL", "Plan stations could not be read as numbers"
        Exit Sub
    End If
    nPl = nKeep

    ReDim qKey(1 To nQ)
    For j = 1 To nQ
        qKey(j) = AcPoc_WellKey(CStr(wsQ.Cells(j + 1, 1).Value2 & ""))
    Next j
    det = wsDet.Range(wsDet.Cells(2, 1), wsDet.Cells(nD + 1, 13)).Value2

    worst = 0#
    For i = 1 To nD
        If (i And 127) = 0 Then
            Application.StatusBar = "[STEP 5/6] Frame proof [ " & i & "/" & nD & " ]"
        End If
        well = CStr(det(i, 1) & "")
        plotIt = False
        wk = AcPoc_WellKey(well)
        For j = 1 To nQ
            If wk = qKey(j) Then plotIt = True: Exit For
        Next j
        okRow = plotIt
        If okRow Then okRow = AcPoc_TryDbl(det(i, 7), offN)
        If okRow Then okRow = AcPoc_TryDbl(det(i, 8), offE)
        If okRow Then okRow = AcPoc_TryDbl(det(i, 9), bc)
        If okRow Then okRow = AcPoc_TryDbl(det(i, 2), refMd)
        If okRow Then okRow = AcPoc_TryDbl(det(i, 5), offT)
        If okRow Then okRow = (bc > 0#)
        If okRow Then okRow = AcPoc_InterpPlan(pMd, pN, pE, pT, nPl, refMd, planN, planE, planT)
        If okRow Then
            dist = Sqr((planN - offN) * (planN - offN) + (planE - offE) * (planE - offE) + (planT - offT) * (planT - offT))
            resid = Abs(dist - bc)
            nChk = nChk + 1
            If resid > worst Then
                worst = resid
                wWorst = resid
                wName = well & " @ " & Format$(refMd, "0.00")
            End If
            If resid > 0.5 Then
                nBad = nBad + 1
                If nBad <= 8 Then
                    AcPoc_Log "Frame proof", "FAIL", well & " RefMD " & Format$(refMd, "0.00") & _
                        " |3D-C2C|=" & Format$(resid, "0.00") & " m  (3D=" & Format$(dist, "0.00") & _
                        " C2C=" & Format$(bc, "0.00") & ")"
                End If
            End If
        End If
    Next i

    If nChk = 0 Then
        AcPoc_Log "Frame proof", "FAIL", "No detail rows could be compared to the plan"
    ElseIf nBad = 0 Then
        AcPoc_Log "Frame proof", "OK", nChk & " stations; worst residual " & _
            Format$(worst, "0.000") & " m at " & wName
    Else
        AcPoc_Log "Frame proof", "FAIL", nBad & " of " & nChk & " stations exceed 0.5 m (worst " & _
            Format$(wWorst, "0.00") & " m)"
    End If
End Sub

Public Function AcPoc_TryDbl(ByVal v As Variant, ByRef out As Double) As Boolean
    Dim s As String
    AcPoc_TryDbl = False
    Select Case VarType(v)
        Case vbEmpty, vbNull, vbError
            Exit Function
        Case vbDouble, vbSingle, vbCurrency, vbInteger, vbLong, vbByte
            out = CDbl(v)
            AcPoc_TryDbl = True
        Case vbString
            s = Trim$(CStr(v))
            If Len(s) = 0 Then Exit Function
            If Not IsNumeric(s) Then Exit Function
            out = CDbl(s)
            AcPoc_TryDbl = True
        Case Else
            Exit Function
    End Select
End Function

Public Function AcPoc_InterpPlan(pMd() As Double, pN() As Double, pE() As Double, pT() As Double, _
        ByVal n As Long, ByVal md As Double, _
        ByRef outN As Double, ByRef outE As Double, ByRef outT As Double) As Boolean
    Dim i As Long
    Dim t As Double
    If n < 2 Then Exit Function
    If md <= pMd(1) Then
        outN = pN(1): outE = pE(1): outT = pT(1)
        AcPoc_InterpPlan = True
        Exit Function
    End If
    If md >= pMd(n) Then
        outN = pN(n): outE = pE(n): outT = pT(n)
        AcPoc_InterpPlan = True
        Exit Function
    End If
    For i = 1 To n - 1
        If md >= pMd(i) And md <= pMd(i + 1) Then
            If Abs(pMd(i + 1) - pMd(i)) < 0.0001 Then
                t = 0#
            Else
                t = (md - pMd(i)) / (pMd(i + 1) - pMd(i))
            End If
            outN = pN(i) + t * (pN(i + 1) - pN(i))
            outE = pE(i) + t * (pE(i + 1) - pE(i))
            outT = pT(i) + t * (pT(i + 1) - pT(i))
            AcPoc_InterpPlan = True
            Exit Function
        End If
    Next i
End Function
