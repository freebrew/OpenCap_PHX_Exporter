Attribute VB_Name = "MDL_AcPoc_Render"
Option Explicit

Private Const SHP_PREFIX As String = "ACP_"
Private Const VS_AZM As Double = 149.46
Private Const PI As Double = 3.14159265358979

Private mWs As Worksheet
Private mNames() As String
Private mNShp As Long
Private mSeq As Long

Public Sub AcPoc_Render()
    Dim outDir As String
    outDir = AcPoc_OutDir()
    On Error GoTo Fail
    AcPoc_DrawCorridor3d
    AcPoc_ExportSheetPng "Corridor", outDir & "corridor_3d.png"
    AcPoc_DrawPlanVs
    AcPoc_ExportSheetPng "Corridor", outDir & "corridor_plan_vs.png"
    AcPoc_DrawProfile
    AcPoc_ExportSheetPng "Corridor", outDir & "ac_profile.png"
    AcPoc_BuildCharts
    AcPoc_Log "Render", "OK", "Wrote corridor_3d.png, corridor_plan_vs.png, ac_profile.png + Charts"
    Exit Sub
Fail:
    AcPoc_Log "Render", "FAIL", "Render error " & Err.Number & ": " & Err.Description
End Sub

Private Sub AcPoc_BeginDraw(ByVal title As String)
    Dim i As Long
    Set mWs = AcPoc_EnsureSheet("Corridor", False)
    On Error Resume Next
    For i = mWs.Shapes.Count To 1 Step -1
        If Left$(mWs.Shapes(i).name, Len(SHP_PREFIX)) = SHP_PREFIX Then mWs.Shapes(i).Delete
    Next i
    On Error GoTo 0
    ReDim mNames(0 To 4000)
    mNShp = 0
    mSeq = 0
    mWs.Activate
    AcPoc_Rect 0, 0, 1100, 720, RGB(255, 255, 255), RGB(200, 200, 200), 0.75
    AcPoc_Tx 12, 8, title, 14, RGB(20, 20, 20), True
    AcPoc_Tx 12, 28, "Offset tracks are Compass closest-approach loci, not full offset surveys.", _
        8, RGB(90, 90, 90), False
End Sub

Private Sub AcPoc_DrawCorridor3d()
    Dim pl As Worksheet
    Dim nPl As Long
    Dim i As Long
    Dim nMin As Double, nMax As Double, eMin As Double, eMax As Double
    Dim tMin As Double, tMax As Double
    Dim sc As Double, scT As Double
    Dim ox As Double, oy As Double
    Dim xs() As Double, ys() As Double
    Dim nPt As Long
    Dim heelMd As Double

    AcPoc_BeginDraw "WELLBORE CORRIDOR  —  plan + qualifying AC offsets (oblique)"
    Set pl = AcPoc_EnsureSheet("_POC_Plan", True)
    nPl = AcPoc_LastDataRow(pl, 1) - 1
    If nPl < 2 Then Exit Sub

    heelMd = 3187.02
    nMin = 1E+99: nMax = -1E+99
    eMin = 1E+99: eMax = -1E+99
    tMin = 1E+99: tMax = -1E+99
    AcPoc_BoundsPlan pl, nPl, nMin, nMax, eMin, eMax, tMin, tMax
    AcPoc_BoundsOffsets nMin, nMax, eMin, eMax, tMin, tMax

    ox = 80#: oy = 70#
    If (eMax - eMin) < 1 Then eMax = eMin + 1
    If (nMax - nMin) < 1 Then nMax = nMin + 1
    If (tMax - tMin) < 1 Then tMax = tMin + 1
    sc = 820# / ((eMax - eMin) + 0.45 * (nMax - nMin))
    scT = 540# / (tMax - tMin)
    If scT < sc * 0.35 Then sc = scT / 0.35

    AcPoc_PlanPoly3d pl, nPl, heelMd, nMin, eMin, tMin, sc, scT, ox, oy
    AcPoc_OffsetPolys3d nMin, eMin, tMin, sc, scT, ox, oy
    AcPoc_Tx 12, 690, "N/E equal scale  ·  TVD down  ·  colour = worst SF of that well (red <1.5, yellow <2, green ≥2)", _
        8, RGB(80, 80, 80), False
End Sub

Private Sub AcPoc_DrawPlanVs()
    Dim pl As Worksheet
    Dim nPl As Long
    Dim nMin As Double, nMax As Double, eMin As Double, eMax As Double
    Dim tMin As Double, tMax As Double
    Dim ox As Double, oy As Double, sc As Double
    Dim caz As Double, saz As Double

    AcPoc_BeginDraw "PLAN VIEW (left) and VERTICAL SECTION 149.46° (right)"
    Set pl = AcPoc_EnsureSheet("_POC_Plan", True)
    nPl = AcPoc_LastDataRow(pl, 1) - 1
    If nPl < 2 Then Exit Sub
    nMin = 1E+99: nMax = -1E+99
    eMin = 1E+99: eMax = -1E+99
    tMin = 1E+99: tMax = -1E+99
    AcPoc_BoundsPlan pl, nPl, nMin, nMax, eMin, eMax, tMin, tMax
    AcPoc_BoundsOffsets nMin, nMax, eMin, eMax, tMin, tMax

    ox = 70#: oy = 80#
    If (eMax - eMin) < 1 Then eMax = eMin + 1
    If (nMax - nMin) < 1 Then nMax = nMin + 1
    sc = WorksheetFunction.Min(460# / (eMax - eMin), 560# / (nMax - nMin))
    AcPoc_Tx ox, oy - 18, "East vs North", 10, RGB(30, 30, 30), True
    AcPoc_PlanPolyNE pl, nPl, nMin, eMin, sc, ox, oy + 560, 3187.02
    AcPoc_OffsetPolysNE nMin, eMin, sc, ox, oy + 560

    caz = Cos(VS_AZM * PI / 180#)
    saz = Sin(VS_AZM * PI / 180#)
    ox = 600#: oy = 80#
    AcPoc_Tx ox, oy - 18, "VS vs TVD", 10, RGB(30, 30, 30), True
    AcPoc_PlanPolyVS pl, nPl, caz, saz, tMin, sc * 0.85, ox, oy, 3187.02
    AcPoc_OffsetPolysVS caz, saz, tMin, sc * 0.85, ox, oy
End Sub

Private Sub AcPoc_DrawProfile()
    Dim q As Worksheet
    Dim d As Worksheet
    Dim nQ As Long, nD As Long
    Dim i As Long, j As Long
    Dim well As String
    Dim mdMax As Double
    Dim ox As Double, oy As Double
    Dim scX As Double, scY As Double
    Dim lastX As Double, lastY As Double
    Dim have As Boolean
    Dim sf As Double, md As Double
    Dim ink As Long

    AcPoc_BeginDraw "SEPARATION FACTOR and C2C vs reference MD"
    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    Set d = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    nQ = AcPoc_LastDataRow(q, 1) - 1
    nD = AcPoc_LastDataRow(d, 1) - 1
    mdMax = 1#
    For i = 2 To nD + 1
        If CDbl(d.Cells(i, 2).Value2) > mdMax Then mdMax = CDbl(d.Cells(i, 2).Value2)
    Next i
    ox = 70#: oy = 80#
    scX = 980# / mdMax
    scY = 70#
    Dim yBase As Double
    yBase = oy + 490#

    AcPoc_Ln ox, yBase - 1.5 * scY, ox + 980, yBase - 1.5 * scY, RGB(198, 40, 40), 1#, msoLineDash
    AcPoc_Ln ox, yBase - 2# * scY, ox + 980, yBase - 2# * scY, RGB(249, 168, 37), 1#, msoLineDash
    AcPoc_Tx ox + 984, yBase - 1.5 * scY, "SF 1.5", 8, RGB(198, 40, 40), False
    AcPoc_Tx ox + 984, yBase - 2# * scY, "SF 2.0", 8, RGB(180, 130, 20), False

    For i = 2 To nQ + 1
        If StrComp(CStr(q.Cells(i, 7).Value2 & ""), "Y", vbTextCompare) <> 0 Then GoTo NextW
        well = CStr(q.Cells(i, 1).Value2 & "")
        ink = AcPoc_InkColor(CDbl(q.Cells(i, 4).Value2))
        have = False
        For j = 2 To nD + 1
            If Not AcPoc_WellsMatch(well, CStr(d.Cells(j, 1).Value2 & "")) Then GoTo NextR
            If Not IsNumeric(d.Cells(j, 12).Value2) Then GoTo NextR
            sf = CDbl(d.Cells(j, 12).Value2)
            If sf <= 0# Then GoTo NextR
            md = CDbl(d.Cells(j, 2).Value2)
            If sf > 8# Then sf = 8#
            If have Then
                AcPoc_Ln lastX, lastY, ox + md * scX, yBase - sf * scY, ink, 1.4, msoLineSolid
            End If
            lastX = ox + md * scX
            lastY = yBase - sf * scY
            have = True
NextR:
        Next j
        AcPoc_Tx ox + 8, oy + 520 + (i * 12), AcPoc_ShortWell(well), 8, ink, False
NextW:
    Next i
    AcPoc_Tx ox, 680, "X = reference MD (m)   Y = Separation Factor", 8, RGB(80, 80, 80), False
End Sub

Private Sub AcPoc_BuildCharts()
    Dim ch As Worksheet
    Dim dat As Worksheet
    Dim pl As Worksheet
    Dim nPl As Long
    Dim nQ As Long
    Dim i As Long
    Dim co As ChartObject
    Dim q As Worksheet

    Set ch = AcPoc_EnsureSheet("Charts", False)
    Set dat = AcPoc_EnsureSheet("_POC_ChartData", True)
    Set pl = AcPoc_EnsureSheet("_POC_Plan", True)
    nPl = AcPoc_LastDataRow(pl, 1) - 1
    AcPoc_ClearSheet dat
    dat.Range("A1").Value2 = "PlanE"
    dat.Range("B1").Value2 = "PlanN"
    dat.Range("C1").Value2 = "PlanVS"
    dat.Range("D1").Value2 = "PlanTVD"
    dat.Range("E1").Value2 = "PlanMD"
    dat.Range("F1").Value2 = "QualMD"
    dat.Range("G1").Value2 = "QualSF"
    For i = 1 To nPl
        dat.Cells(i + 1, 1).Value2 = pl.Cells(i + 1, 6).Value2
        dat.Cells(i + 1, 2).Value2 = pl.Cells(i + 1, 5).Value2
        dat.Cells(i + 1, 3).Value2 = pl.Cells(i + 1, 7).Value2
        dat.Cells(i + 1, 4).Value2 = pl.Cells(i + 1, 4).Value2
        dat.Cells(i + 1, 5).Value2 = pl.Cells(i + 1, 1).Value2
    Next i
    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    nQ = AcPoc_LastDataRow(q, 1) - 1
    For i = 1 To nQ
        dat.Cells(i + 1, 6).Value2 = q.Cells(i + 1, 2).Value2
        dat.Cells(i + 1, 7).Value2 = q.Cells(i + 1, 4).Value2
    Next i

    On Error Resume Next
    For i = ch.ChartObjects.Count To 1 Step -1
        ch.ChartObjects(i).Delete
    Next i
    On Error GoTo 0
    ch.Range("A1").Value2 = "Native Excel charts — plan view (E vs N) and VS vs TVD. Zoom / hover in Excel."
    ch.Range("A1").Font.Bold = True

    Set co = ch.ChartObjects.Add(20, 30, 480, 360)
    co.Chart.ChartType = xlXYScatterSmoothNoMarkers
    co.Chart.SetSourceData dat.Range(dat.Cells(1, 1), dat.Cells(nPl + 1, 2))
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Plan view  E vs N"

    Set co = ch.ChartObjects.Add(520, 30, 480, 360)
    co.Chart.ChartType = xlXYScatterSmoothNoMarkers
    co.Chart.SetSourceData dat.Range(dat.Cells(1, 3), dat.Cells(nPl + 1, 4))
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "Vertical section  VS vs TVD"

    If nQ >= 1 Then
        Set co = ch.ChartObjects.Add(20, 410, 980, 280)
        co.Chart.ChartType = xlXYScatter
        co.Chart.SetSourceData dat.Range(dat.Cells(1, 6), dat.Cells(nQ + 1, 7))
        co.Chart.HasTitle = True
        co.Chart.ChartTitle.Text = "Qualify SF vs Ref MD"
    End If
End Sub

Private Sub AcPoc_BoundsPlan(ByVal pl As Worksheet, ByVal nPl As Long, _
        ByRef nMin As Double, ByRef nMax As Double, ByRef eMin As Double, ByRef eMax As Double, _
        ByRef tMin As Double, ByRef tMax As Double)
    Dim i As Long
    Dim n As Double, e As Double, t As Double
    For i = 2 To nPl + 1
        n = CDbl(pl.Cells(i, 5).Value2)
        e = CDbl(pl.Cells(i, 6).Value2)
        t = CDbl(pl.Cells(i, 4).Value2)
        If n < nMin Then nMin = n
        If n > nMax Then nMax = n
        If e < eMin Then eMin = e
        If e > eMax Then eMax = e
        If t < tMin Then tMin = t
        If t > tMax Then tMax = t
    Next i
End Sub

Private Sub AcPoc_BoundsOffsets(ByRef nMin As Double, ByRef nMax As Double, _
        ByRef eMin As Double, ByRef eMax As Double, ByRef tMin As Double, ByRef tMax As Double)
    Dim d As Worksheet
    Dim q As Worksheet
    Dim nD As Long, nQ As Long
    Dim i As Long, j As Long
    Dim well As String
    Dim n As Double, e As Double, t As Double
    Set d = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    nD = AcPoc_LastDataRow(d, 1) - 1
    nQ = AcPoc_LastDataRow(q, 1) - 1
    For i = 2 To nD + 1
        well = CStr(d.Cells(i, 1).Value2 & "")
        For j = 2 To nQ + 1
            If AcPoc_WellsMatch(well, CStr(q.Cells(j, 1).Value2 & "")) Then
                If IsNumeric(d.Cells(i, 7).Value2) Then
                    n = CDbl(d.Cells(i, 7).Value2)
                    e = CDbl(d.Cells(i, 8).Value2)
                    t = CDbl(d.Cells(i, 5).Value2)
                    If n < nMin Then nMin = n
                    If n > nMax Then nMax = n
                    If e < eMin Then eMin = e
                    If e > eMax Then eMax = e
                    If t < tMin Then tMin = t
                    If t > tMax Then tMax = t
                End If
                Exit For
            End If
        Next j
    Next i
End Sub

Private Sub AcPoc_PlanPoly3d(ByVal pl As Worksheet, ByVal nPl As Long, ByVal heelMd As Double, _
        ByVal n0 As Double, ByVal e0 As Double, ByVal t0 As Double, _
        ByVal sc As Double, ByVal scT As Double, ByVal ox As Double, ByVal oy As Double)
    Dim i As Long, k As Long
    Dim xs() As Double, ys() As Double
    Dim md As Double, inc As Double
    Dim clr As Long
    ReDim xs(0 To nPl)
    ReDim ys(0 To nPl)
    k = -1
    For i = 2 To nPl + 1
        md = CDbl(pl.Cells(i, 1).Value2)
        inc = CDbl(pl.Cells(i, 2).Value2)
        k = k + 1
        xs(k) = AcPoc_X3(CDbl(pl.Cells(i, 6).Value2), CDbl(pl.Cells(i, 5).Value2), e0, n0, sc, ox)
        ys(k) = AcPoc_Y3(CDbl(pl.Cells(i, 4).Value2), CDbl(pl.Cells(i, 5).Value2), t0, n0, sc, scT, oy)
    Next i
    AcPoc_Poly xs, ys, k + 1, RGB(26, 58, 92), 2.2
    AcPoc_Tx xs(0), ys(0) - 10, "SLOT", 8, RGB(26, 58, 92), True
    AcPoc_Tx xs(k), ys(k) + 8, "TD", 8, RGB(26, 58, 92), True
End Sub

Private Sub AcPoc_OffsetPolys3d(ByVal n0 As Double, ByVal e0 As Double, ByVal t0 As Double, _
        ByVal sc As Double, ByVal scT As Double, ByVal ox As Double, ByVal oy As Double)
    Dim q As Worksheet
    Dim d As Worksheet
    Dim nQ As Long, nD As Long
    Dim i As Long, j As Long, k As Long
    Dim well As String
    Dim ink As Long
    Dim xs() As Double, ys() As Double
    Dim n As Double, e As Double, t As Double

    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    Set d = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    nQ = AcPoc_LastDataRow(q, 1) - 1
    nD = AcPoc_LastDataRow(d, 1) - 1
    For i = 2 To nQ + 1
        If StrComp(CStr(q.Cells(i, 7).Value2 & ""), "Y", vbTextCompare) <> 0 Then GoTo NextQ
        well = CStr(q.Cells(i, 1).Value2 & "")
        ink = AcPoc_InkColor(CDbl(q.Cells(i, 4).Value2))
        ReDim xs(0 To 400)
        ReDim ys(0 To 400)
        k = -1
        For j = 2 To nD + 1
            If Not AcPoc_WellsMatch(well, CStr(d.Cells(j, 1).Value2 & "")) Then GoTo NextD
            If Not IsNumeric(d.Cells(j, 7).Value2) Then GoTo NextD
            n = CDbl(d.Cells(j, 7).Value2)
            e = CDbl(d.Cells(j, 8).Value2)
            t = CDbl(d.Cells(j, 5).Value2)
            k = k + 1
            If k > 400 Then Exit For
            xs(k) = AcPoc_X3(e, n, e0, n0, sc, ox)
            ys(k) = AcPoc_Y3(t, n, t0, n0, sc, scT, oy)
            If Len(Trim$(CStr(d.Cells(j, 13).Value2 & ""))) > 0 Then
                AcPoc_Dot xs(k), ys(k), 3.2, ink, RGB(20, 20, 20), 0.8
            End If
NextD:
        Next j
        If k >= 1 Then AcPoc_Poly xs, ys, k + 1, ink, 1.6
        If k >= 0 Then
            AcPoc_Tx xs(k) + 6, ys(k), AcPoc_ShortWell(well) & "  SF " & _
                Format$(CDbl(q.Cells(i, 4).Value2), "0.000"), 8, ink, False
        End If
NextQ:
    Next i
End Sub

Private Sub AcPoc_PlanPolyNE(ByVal pl As Worksheet, ByVal nPl As Long, _
        ByVal n0 As Double, ByVal e0 As Double, ByVal sc As Double, _
        ByVal ox As Double, ByVal oy As Double, ByVal heelMd As Double)
    Dim i As Long, k As Long
    Dim xs() As Double, ys() As Double
    ReDim xs(0 To nPl)
    ReDim ys(0 To nPl)
    k = -1
    For i = 2 To nPl + 1
        k = k + 1
        xs(k) = ox + (CDbl(pl.Cells(i, 6).Value2) - e0) * sc
        ys(k) = oy - (CDbl(pl.Cells(i, 5).Value2) - n0) * sc
    Next i
    AcPoc_Poly xs, ys, k + 1, RGB(26, 58, 92), 2#
End Sub

Private Sub AcPoc_OffsetPolysNE(ByVal n0 As Double, ByVal e0 As Double, _
        ByVal sc As Double, ByVal ox As Double, ByVal oy As Double)
    Dim q As Worksheet, d As Worksheet
    Dim nQ As Long, nD As Long, i As Long, j As Long, k As Long
    Dim well As String, ink As Long
    Dim xs() As Double, ys() As Double
    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    Set d = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    nQ = AcPoc_LastDataRow(q, 1) - 1
    nD = AcPoc_LastDataRow(d, 1) - 1
    For i = 2 To nQ + 1
        If StrComp(CStr(q.Cells(i, 7).Value2 & ""), "Y", vbTextCompare) <> 0 Then GoTo NextQ
        well = CStr(q.Cells(i, 1).Value2 & "")
        ink = AcPoc_InkColor(CDbl(q.Cells(i, 4).Value2))
        ReDim xs(0 To 400): ReDim ys(0 To 400)
        k = -1
        For j = 2 To nD + 1
            If Not AcPoc_WellsMatch(well, CStr(d.Cells(j, 1).Value2 & "")) Then GoTo NextD
            If Not IsNumeric(d.Cells(j, 7).Value2) Then GoTo NextD
            k = k + 1
            If k > 400 Then Exit For
            xs(k) = ox + (CDbl(d.Cells(j, 8).Value2) - e0) * sc
            ys(k) = oy - (CDbl(d.Cells(j, 7).Value2) - n0) * sc
NextD:
        Next j
        If k >= 1 Then AcPoc_Poly xs, ys, k + 1, ink, 1.4
NextQ:
    Next i
End Sub

Private Sub AcPoc_PlanPolyVS(ByVal pl As Worksheet, ByVal nPl As Long, _
        ByVal caz As Double, ByVal saz As Double, ByVal t0 As Double, _
        ByVal sc As Double, ByVal ox As Double, ByVal oy As Double, ByVal heelMd As Double)
    Dim i As Long, k As Long
    Dim xs() As Double, ys() As Double
    Dim vs As Double
    Dim vs0 As Double
    vs0 = CDbl(pl.Cells(2, 5).Value2) * caz + CDbl(pl.Cells(2, 6).Value2) * saz
    ReDim xs(0 To nPl): ReDim ys(0 To nPl)
    k = -1
    For i = 2 To nPl + 1
        vs = CDbl(pl.Cells(i, 5).Value2) * caz + CDbl(pl.Cells(i, 6).Value2) * saz
        k = k + 1
        xs(k) = ox + (vs - vs0) * sc
        ys(k) = oy + (CDbl(pl.Cells(i, 4).Value2) - t0) * sc
    Next i
    AcPoc_Poly xs, ys, k + 1, RGB(26, 58, 92), 2#
End Sub

Private Sub AcPoc_OffsetPolysVS(ByVal caz As Double, ByVal saz As Double, ByVal t0 As Double, _
        ByVal sc As Double, ByVal ox As Double, ByVal oy As Double)
    Dim q As Worksheet, d As Worksheet, pl As Worksheet
    Dim nQ As Long, nD As Long, i As Long, j As Long, k As Long
    Dim well As String, ink As Long
    Dim xs() As Double, ys() As Double
    Dim vs As Double, vs0 As Double
    Set pl = AcPoc_EnsureSheet("_POC_Plan", True)
    vs0 = CDbl(pl.Cells(2, 5).Value2) * caz + CDbl(pl.Cells(2, 6).Value2) * saz
    Set q = AcPoc_EnsureSheet("_POC_Qualify", True)
    Set d = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    nQ = AcPoc_LastDataRow(q, 1) - 1
    nD = AcPoc_LastDataRow(d, 1) - 1
    For i = 2 To nQ + 1
        If StrComp(CStr(q.Cells(i, 7).Value2 & ""), "Y", vbTextCompare) <> 0 Then GoTo NextQ
        well = CStr(q.Cells(i, 1).Value2 & "")
        ink = AcPoc_InkColor(CDbl(q.Cells(i, 4).Value2))
        ReDim xs(0 To 400): ReDim ys(0 To 400)
        k = -1
        For j = 2 To nD + 1
            If Not AcPoc_WellsMatch(well, CStr(d.Cells(j, 1).Value2 & "")) Then GoTo NextD
            If Not IsNumeric(d.Cells(j, 7).Value2) Then GoTo NextD
            vs = CDbl(d.Cells(j, 7).Value2) * caz + CDbl(d.Cells(j, 8).Value2) * saz
            k = k + 1
            If k > 400 Then Exit For
            xs(k) = ox + (vs - vs0) * sc
            ys(k) = oy + (CDbl(d.Cells(j, 5).Value2) - t0) * sc
NextD:
        Next j
        If k >= 1 Then AcPoc_Poly xs, ys, k + 1, ink, 1.4
NextQ:
    Next i
End Sub

Private Function AcPoc_X3(ByVal e As Double, ByVal n As Double, ByVal e0 As Double, _
        ByVal n0 As Double, ByVal sc As Double, ByVal ox As Double) As Double
    AcPoc_X3 = ox + (e - e0) * sc + (n - n0) * 0.35 * sc
End Function

Private Function AcPoc_Y3(ByVal tvd As Double, ByVal n As Double, ByVal t0 As Double, _
        ByVal n0 As Double, ByVal sc As Double, ByVal scT As Double, ByVal oy As Double) As Double
    AcPoc_Y3 = oy + (tvd - t0) * scT - (n - n0) * 0.22 * sc
End Function

Private Function AcPoc_ShortWell(ByVal s As String) As String
    Dim p As Long
    s = AcPoc_CleanWellName(s)
    p = InStrRev(s, "- ")
    If p > 0 Then
        AcPoc_ShortWell = Trim$(Mid$(s, p + 2))
    Else
        AcPoc_ShortWell = Right$(s, 28)
    End If
End Function

Private Function AcPoc_NextName() As String
    mSeq = mSeq + 1
    AcPoc_NextName = SHP_PREFIX & Format$(mSeq, "00000")
End Function

Private Sub AcPoc_Remember(ByVal nm As String)
    If mNShp > UBound(mNames) Then ReDim Preserve mNames(0 To mNShp + 500)
    mNames(mNShp) = nm
    mNShp = mNShp + 1
End Sub

Private Sub AcPoc_Rect(ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal hgt As Double, _
        ByVal fillClr As Long, ByVal lineClr As Long, ByVal wt As Double)
    Dim shp As Shape
    If w < 0.1 Then w = 0.1
    If hgt < 0.1 Then hgt = 0.1
    Set shp = mWs.Shapes.AddShape(msoShapeRectangle, l, t, w, hgt)
    shp.name = AcPoc_NextName(): AcPoc_Remember shp.name
    shp.Fill.Visible = msoTrue
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = fillClr
    shp.Line.ForeColor.RGB = lineClr
    shp.Line.Weight = wt
    shp.Shadow.Visible = msoFalse
End Sub

Private Sub AcPoc_Ln(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal clr As Long, ByVal wt As Double, ByVal dash As Long)
    Dim shp As Shape
    Set shp = mWs.Shapes.AddLine(x1, y1, x2, y2)
    shp.name = AcPoc_NextName(): AcPoc_Remember shp.name
    shp.Line.ForeColor.RGB = clr
    shp.Line.Weight = wt
    shp.Line.DashStyle = dash
End Sub

Private Sub AcPoc_Dot(ByVal cx As Double, ByVal cy As Double, ByVal r As Double, _
        ByVal fillClr As Long, ByVal lineClr As Long, ByVal wt As Double)
    Dim shp As Shape
    Set shp = mWs.Shapes.AddShape(msoShapeOval, cx - r, cy - r, 2 * r, 2 * r)
    shp.name = AcPoc_NextName(): AcPoc_Remember shp.name
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = fillClr
    shp.Line.ForeColor.RGB = lineClr
    shp.Line.Weight = wt
    shp.Shadow.Visible = msoFalse
End Sub

Private Sub AcPoc_Poly(xs() As Double, ys() As Double, ByVal n As Long, ByVal clr As Long, ByVal wt As Double)
    Dim fb As FreeformBuilder
    Dim shp As Shape
    Dim i As Long
    Dim stepN As Long
    If n < 2 Then Exit Sub
    stepN = 1
    If n > 240 Then stepN = Int(n / 200)
    Set fb = mWs.Shapes.BuildFreeform(msoEditingCorner, xs(0), ys(0))
    i = stepN
    Do While i < n - 1
        fb.AddNodes msoSegmentLine, msoEditingAuto, xs(i), ys(i)
        i = i + stepN
    Loop
    fb.AddNodes msoSegmentLine, msoEditingAuto, xs(n - 1), ys(n - 1)
    Set shp = fb.ConvertToShape
    shp.name = AcPoc_NextName(): AcPoc_Remember shp.name
    shp.Fill.Visible = msoFalse
    shp.Line.ForeColor.RGB = clr
    shp.Line.Weight = wt
    shp.Shadow.Visible = msoFalse
End Sub

Private Sub AcPoc_Tx(ByVal x As Double, ByVal y As Double, ByVal s As String, _
        ByVal sz As Single, ByVal clr As Long, ByVal bold As Boolean)
    Dim shp As Shape
    Set shp = mWs.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, 420, sz + 8)
    shp.name = AcPoc_NextName(): AcPoc_Remember shp.name
    shp.TextFrame.Characters.Text = s
    shp.TextFrame.Characters.Font.Size = sz
    shp.TextFrame.Characters.Font.Color = clr
    shp.TextFrame.Characters.Font.Bold = bold
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
End Sub

Private Sub AcPoc_ExportSheetPng(ByVal sheetName As String, ByVal outPath As String)
    Dim ws As Worksheet
    Dim grp As Shape
    Dim co As ChartObject
    Dim prevVis As Boolean
    Dim names() As String
    Dim i As Long
    Dim n As Long

    Set ws = ThisWorkbook.Worksheets(sheetName)
    n = 0
    ReDim names(1 To ws.Shapes.Count)
    For i = 1 To ws.Shapes.Count
        If Left$(ws.Shapes(i).name, Len(SHP_PREFIX)) = SHP_PREFIX Then
            n = n + 1
            names(n) = ws.Shapes(i).name
        End If
    Next i
    If n < 1 Then
        AcPoc_Log "Render", "FAIL", "No shapes to export for " & outPath
        Exit Sub
    End If

    prevVis = Application.Visible
    If Not prevVis Then Application.Visible = True
    ws.Activate
    If n = 1 Then
        Set grp = ws.Shapes(names(1))
    Else
        ReDim Preserve names(1 To n)
        Set grp = ws.Shapes.Range(names).Group
        grp.name = SHP_PREFIX & "GRP"
    End If

    On Error GoTo Clean
    Set co = ws.ChartObjects.Add(0, 0, 1100, 720)
    co.Chart.ChartArea.Border.LineStyle = xlNone
    grp.CopyPicture xlScreen, xlPicture
    co.Activate
    co.Chart.Paste
    If co.Chart.Shapes.Count = 0 Then
        grp.CopyPicture xlScreen, xlBitmap
        co.Chart.Paste
    End If
    If co.Chart.Shapes.Count = 0 Then Err.Raise 5, , "nothing pasted into the chart"
    co.Chart.Export Filename:=outPath, FilterName:="PNG"
    co.Delete
    If n > 1 Then
        On Error Resume Next
        ws.Shapes(SHP_PREFIX & "GRP").Ungroup
        On Error GoTo 0
    End If
    If Not prevVis Then Application.Visible = prevVis
    Exit Sub
Clean:
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    ws.Shapes(SHP_PREFIX & "GRP").Ungroup
    If Not prevVis Then Application.Visible = prevVis
    AcPoc_Log "Render", "FAIL", "PNG export: " & Err.Description
End Sub
