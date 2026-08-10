Attribute VB_Name = "MDL_WellborePos"
Option Explicit

' Sync Data "Position of Wellbore (as of last survey)" from gauge lateral + Sail Up/down.

Private Const SH_DATA As String = "Data"
Private Const SH_SS As String = "Slidesheet"
Private Const SURV_FIRST As Long = 13
Private Const SURV_LAST As Long = 320
Private Const COL_SAIL_UPDN As Long = 28  ' AB (skip AB14 geo window — only survey rows)

Private Const ROW_RL As Long = 13
Private Const ROW_AB As Long = 14
Private Const ROW_TOT As Long = 15
Private Const COL_VAL As Long = 5         ' E
Private Const COL_DIR As Long = 6         ' F

' latM: gauge frame X (same sign as OCG_RdX: >=0 Right, <0 Left).
' survRow: Slidesheet row of last real survey (from ActualAtLastSurvey); 0 = scan.
Public Sub UpdateWellborePosition(ByVal latM As Double, ByVal hasLat As Boolean, _
                                  Optional ByVal survRow As Long = 0)
    Dim wsD As Worksheet
    Dim wasProt As Boolean
    Dim latAbs As Double
    Dim latDir As String
    Dim vertM As Double
    Dim hasVert As Boolean
    Dim vertAbs As Double
    Dim vertDir As String
    Dim totM As Double

    Set wsD = ThisWorkbook.Worksheets(SH_DATA)

    hasVert = SailUpDownAtSurvey(survRow, vertM)

    wasProt = SheetUnprotectForVba(wsD)
    On Error GoTo FailWrite

    If hasLat Then
        latAbs = Abs(latM)
        latDir = IIf(latM >= 0#, "Right", "Left")
        wsD.Cells(ROW_RL, COL_VAL).Value = Format$(latAbs, "0.00") & "m"
        wsD.Cells(ROW_RL, COL_DIR).Value = latDir
    End If

    If hasVert Then
        vertAbs = Abs(vertM)
        If vertM < 0# Then
            vertDir = "Below"
        Else
            vertDir = "Above"
        End If
        wsD.Cells(ROW_AB, COL_VAL).Value = Format$(vertAbs, "0.00") & "m"
        wsD.Cells(ROW_AB, COL_DIR).Value = vertDir
    End If

    If hasLat And hasVert Then
        totM = Sqr(latAbs * latAbs + vertAbs * vertAbs)
        wsD.Cells(ROW_TOT, COL_VAL).Value = Format$(totM, "0.00") & "m"
        wsD.Cells(ROW_TOT, COL_DIR).Value = "Total"
    ElseIf hasLat And Not hasVert Then
        wsD.Cells(ROW_TOT, COL_VAL).Value = Format$(latAbs, "0.00") & "m"
        wsD.Cells(ROW_TOT, COL_DIR).Value = "Total"
    ElseIf hasVert And Not hasLat Then
        wsD.Cells(ROW_TOT, COL_VAL).Value = Format$(vertAbs, "0.00") & "m"
        wsD.Cells(ROW_TOT, COL_DIR).Value = "Total"
    End If

    SheetReprotectAfterVba wsD, wasProt
    Exit Sub

FailWrite:
    SheetReprotectAfterVba wsD, wasProt
End Sub

' Sail AB Up/down on the gauge's last survey row; fallback = last real survey with numeric AB.
Private Function SailUpDownAtSurvey(ByVal survRow As Long, ByRef upDownM As Double) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim fE As String, fF As String, fd As String

    SailUpDownAtSurvey = False
    upDownM = 0#
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    If survRow >= SURV_FIRST And survRow <= SURV_LAST And survRow <> 14 Then
        If ReadNumericAb(ws, survRow, upDownM) Then
            SailUpDownAtSurvey = True
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
        ' Real survey stations have MD/Inc/Azm in E/F/G (same as PlanGauge).
        If Not IsNumeric(ws.Cells(r, 5).Value2) Then GoTo NextR
        If Not IsNumeric(ws.Cells(r, 6).Value2) Then GoTo NextR
        If Not IsNumeric(ws.Cells(r, 7).Value2) Then GoTo NextR
        If Not ReadNumericAb(ws, r, upDownM) Then GoTo NextR
        SailUpDownAtSurvey = True
        Exit Function
NextR:
    Next r
End Function

Private Function ReadNumericAb(ws As Worksheet, ByVal r As Long, ByRef upDownM As Double) As Boolean
    Dim v As Variant
    Dim t As String
    ReadNumericAb = False
    v = ws.Cells(r, COL_SAIL_UPDN).Value2
    t = Trim$(CStr(ws.Cells(r, COL_SAIL_UPDN).text & ""))
    If Len(t) = 0 Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If Not IsNumeric(t) Then Exit Function   ' reject "Last H/L" etc.
    upDownM = CDbl(v)
    ReadNumericAb = True
End Function
