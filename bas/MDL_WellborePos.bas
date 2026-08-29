Attribute VB_Name = "MDL_WellborePos"
Option Explicit

' Sync Data "Position of Wellbore (as of last survey)":
'   R/L from Plan     <- gauge frame X (ax): >=0 Right, <0 Left
'   Above/Below Plan  <- gauge frame Y (ay): plan TVD - as-drilled TVD (>=0 Above/UP, <0 Below/DN)
'   Distance from Geo
'     Vertical / curve (Inc < ~90 or MD before sail waypoints AC14:AD33):
'       hypot(E/W, N/S) from plan — there is no geo target yet.
'     Lateral (Inc >= 80 and MD at/after first sail waypoint):
'       Slidesheet GEO Window column AB on the last survey row
'       (skip AB14 header; >=0 Above geo, <0 Below geo)

Private Const SH_DATA As String = "Data"
Private Const SH_SS As String = "Slidesheet"
Private Const SURV_FIRST As Long = 13
Private Const SURV_LAST As Long = 320
Private Const COL_GEO_AB As Long = 28  ' AB GEO Window (skip AB14)
Private Const COL_SAIL_MD As Long = 29 ' AC sail / way-point MD
Private Const SAIL_ROW1 As Long = 14
Private Const SAIL_ROW2 As Long = 33
Private Const INC_LATERAL As Double = 80#  ' ~90 lateral; below this use plan hypot

Private Const ROW_RL As Long = 13
Private Const ROW_AB As Long = 14
Private Const ROW_GEO As Long = 15
Private Const COL_VAL As Long = 5         ' E
Private Const COL_DIR As Long = 6         ' F

' latM / planUpM: same frame as OCG gauge markers (FrameComponents).
' survRow: Slidesheet row of last real survey (from ActualAtLastSurvey); 0 = scan.
Public Sub UpdateWellborePosition(ByVal latM As Double, ByVal hasLat As Boolean, _
                                  ByVal planUpM As Double, ByVal hasPlanUp As Boolean, _
                                  Optional ByVal survRow As Long = 0, _
                                  Optional ByVal sMD As Double = -1#, _
                                  Optional ByVal sInc As Double = -1#)
    Dim wsD As Worksheet
    Dim wasProt As Boolean
    Dim latAbs As Double
    Dim latDir As String
    Dim planAbs As Double
    Dim planDir As String
    Dim geoM As Double
    Dim hasGeo As Boolean
    Dim geoAbs As Double
    Dim geoDir As String

    Set wsD = ThisWorkbook.Worksheets(SH_DATA)

    If UseSailGeoWindow(sMD, sInc) Then
        hasGeo = GeoWindowAtSurvey(survRow, geoM)
    ElseIf hasLat And hasPlanUp Then
        geoM = Sqr(latM * latM + planUpM * planUpM)
        If planUpM < 0# Then geoM = -geoM
        hasGeo = True
    Else
        hasGeo = False
    End If

    wasProt = SheetUnprotectForVba(wsD)
    On Error GoTo FailWrite

    If hasLat Then
        latAbs = Abs(latM)
        latDir = IIf(latM >= 0#, "Right", "Left")
        wsD.Cells(ROW_RL, COL_VAL).Value = Format$(latAbs, "0.00") & "m"
        wsD.Cells(ROW_RL, COL_DIR).Value = latDir
    End If

    If hasPlanUp Then
        planAbs = Abs(planUpM)
        If planUpM < 0# Then
            planDir = "Below"
        Else
            planDir = "Above"
        End If
        wsD.Cells(ROW_AB, COL_VAL).Value = Format$(planAbs, "0.00") & "m"
        wsD.Cells(ROW_AB, COL_DIR).Value = planDir
    End If

    If hasGeo Then
        geoAbs = Abs(geoM)
        If geoM < 0# Then
            geoDir = "Below"
        Else
            geoDir = "Above"
        End If
        wsD.Cells(ROW_GEO, COL_VAL).Value = Format$(geoAbs, "0.00") & "m"
        wsD.Cells(ROW_GEO, COL_DIR).Value = geoDir
    End If

    SheetReprotectAfterVba wsD, wasProt
    Exit Sub

FailWrite:
    SheetReprotectAfterVba wsD, wasProt
End Sub

' Geo target exists only once we are in the lateral at/after the sail table MD.
Private Function UseSailGeoWindow(ByVal sMD As Double, ByVal sInc As Double) As Boolean
    Dim sailMd As Double
    UseSailGeoWindow = False
    If sInc < INC_LATERAL Then Exit Function
    If Not FirstSailMd(sailMd) Then Exit Function
    If sMD + 0.0001 >= sailMd Then UseSailGeoWindow = True
End Function

Private Function FirstSailMd(ByRef sailMd As Double) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim v As Variant
    FirstSailMd = False
    sailMd = 0#
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = SAIL_ROW1 To SAIL_ROW2
        v = ws.Cells(r, COL_SAIL_MD).Value2
        If IsNumeric(v) And Len(CStr(v & "")) > 0 Then
            sailMd = CDbl(v)
            FirstSailMd = True
            Exit Function
        End If
    Next r
End Function

' GEO Window (AB) on the gauge's last survey row; fallback = last real survey with numeric AB.
Private Function GeoWindowAtSurvey(ByVal survRow As Long, ByRef geoM As Double) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim fE As String, fF As String, fd As String

    GeoWindowAtSurvey = False
    geoM = 0#
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    If survRow >= SURV_FIRST And survRow <= SURV_LAST And survRow <> 14 Then
        If ReadNumericAb(ws, survRow, geoM) Then
            GeoWindowAtSurvey = True
            Exit Function
        End If
    End If

    For r = SURV_LAST To SURV_FIRST Step -1
        If r = 14 Then GoTo NextR
        On Error Resume Next
        fE = UCase$(CStr(ws.Cells(r, 5).Formula & ""))
        fF = UCase$(CStr(ws.Cells(r, 6).Formula & ""))
        fd = UCase$(CStr(ws.Cells(r, 4).Formula & ""))
        On Error GoTo 0
        If InStr(1, fE, "LOOKUP", vbBinaryCompare) > 0 Then GoTo NextR
        If InStr(1, fF, "LOOKUP", vbBinaryCompare) > 0 Then GoTo NextR
        If InStr(1, fd, "LOOKUP", vbBinaryCompare) > 0 Then GoTo NextR
        If Not IsNumeric(ws.Cells(r, 5).Value2) Then GoTo NextR
        If Not IsNumeric(ws.Cells(r, 6).Value2) Then GoTo NextR
        If Not IsNumeric(ws.Cells(r, 7).Value2) Then GoTo NextR
        If Not ReadNumericAb(ws, r, geoM) Then GoTo NextR
        GeoWindowAtSurvey = True
        Exit Function
NextR:
    Next r
End Function

Private Function ReadNumericAb(ws As Worksheet, ByVal r As Long, ByRef geoM As Double) As Boolean
    Dim v As Variant
    Dim t As String
    ReadNumericAb = False
    v = ws.Cells(r, COL_GEO_AB).Value2
    t = Trim$(CStr(ws.Cells(r, COL_GEO_AB).text & ""))
    If Len(t) = 0 Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If Not IsNumeric(t) Then Exit Function   ' reject "Last H/L" etc.
    geoM = CDbl(v)
    ReadNumericAb = True
End Function




