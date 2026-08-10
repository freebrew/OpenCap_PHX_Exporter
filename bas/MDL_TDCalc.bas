Attribute VB_Name = "MDL_TDCalc"
Option Explicit

' ================================================================================
'  MODULE: MDL_TDCalc
'  TD Calculator block on the Data sheet (H49:M55)
'
'    Actual      I53:M53  live formulas -> last Slidesheet survey row, BIT columns
'                         MD=D, INC=W, AZM=X, NS=AP, EW=AQ
'    Planned TD  I54:M54  written by Import Plan from the last _OC_Survey station
'    Actual TD   I55      existing formula, untouched
'
'  The Actual row stays live because TdActualBit takes the survey block as an
'  argument: Excel recalculates it whenever any cell in that block changes.
' ================================================================================

Private Const SS_SHEET   As String = "Slidesheet"
Private Const PLAN_SHEET As String = "_OC_Survey"
Private Const DATA_SHEET As String = "Data"

Private Const SURV_ROW_FIRST As Long = 13
Private Const SURV_ROW_LAST  As Long = 320

' Slidesheet columns
Private Const SS_COL_BIT_MD As Long = 4    ' D
Private Const SS_COL_MD     As Long = 5    ' E
Private Const SS_COL_INC    As Long = 6    ' F
Private Const SS_COL_AZM    As Long = 7    ' G
Private Const SS_COL_BIT_INC As Long = 23  ' W
Private Const SS_COL_BIT_AZM As Long = 24  ' X
Private Const SS_COL_BIT_N   As Long = 42  ' AP
Private Const SS_COL_BIT_E   As Long = 43  ' AQ

' Data TD Calculator block
Private Const TD_ROW_ACTUAL  As Long = 53
Private Const TD_ROW_PLANNED As Long = 54
Private Const TD_COL_FIRST   As Long = 9   ' I (MD)

Private Const SURVEY_BLOCK As String = "Slidesheet!$D$13:$AQ$320"

' ================================================================================
'  ACTUAL ROW  --  worksheet function
' ================================================================================

' Bit-projected value at the last Slidesheet survey station.
'   field: MD | INC | AZM | NS | EW
'   surveyBlock: pass Slidesheet!$D$13:$AQ$320 so Excel tracks the dependency.
Public Function TdActualBit(ByVal field As String, _
                            Optional ByVal surveyBlock As Range) As Variant
    Dim ws As Worksheet
    If Not surveyBlock Is Nothing Then
        Set ws = surveyBlock.Worksheet
    Else
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(SS_SHEET)
        On Error GoTo 0
    End If
    If ws Is Nothing Then
        TdActualBit = CVErr(xlErrRef)
        Exit Function
    End If

    Dim col As Long
    Select Case UCase$(Trim$(field))
        Case "MD":          col = SS_COL_BIT_MD
        Case "INC":         col = SS_COL_BIT_INC
        Case "AZM", "AZI":  col = SS_COL_BIT_AZM
        Case "NS", "N":     col = SS_COL_BIT_N
        Case "EW", "E":     col = SS_COL_BIT_E
        Case Else
            TdActualBit = CVErr(xlErrValue)
            Exit Function
    End Select

    Dim r As Long: r = LastSurveyRow(ws)
    If r = 0 Then
        TdActualBit = ""
        Exit Function
    End If

    Dim v As Variant: v = ws.Cells(r, col).Value2
    If IsNumeric(v) And Len(CStr(v & "")) > 0 Then
        TdActualBit = CDbl(v)
    Else
        TdActualBit = ""
    End If
End Function

' Row of the deepest valid survey station.  Same rules as the plan gauge:
' skip LOOKUP summary rows, require numeric MD/Inc/Azm, MD must increase.
Private Function LastSurveyRow(ws As Worksheet) As Long
    Dim r As Long
    Dim deepest As Double
    Dim best As Long
    Dim vMD As Variant, vInc As Variant, vAzm As Variant

    For r = SURV_ROW_FIRST To SURV_ROW_LAST
        If Not IsSurveySummaryRow(ws, r) Then
            vMD = ws.Cells(r, SS_COL_MD).Value2
            vInc = ws.Cells(r, SS_COL_INC).Value2
            vAzm = ws.Cells(r, SS_COL_AZM).Value2
            If IsNumeric(vMD) And IsNumeric(vInc) And IsNumeric(vAzm) Then
                If Len(CStr(vMD & "")) > 0 And Len(CStr(vInc & "")) > 0 _
                   And Len(CStr(vAzm & "")) > 0 Then
                    If CDbl(vMD) > deepest Then
                        deepest = CDbl(vMD)
                        best = r
                    End If
                End If
            End If
        End If
    Next r

    LastSurveyRow = best
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

' Install the live Actual formulas on Data I53:M53.
Public Sub InstallTdActualFormulas()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim fields(0 To 4) As String
    fields(0) = "MD": fields(1) = "INC": fields(2) = "AZM"
    fields(3) = "NS": fields(4) = "EW"

    Dim wasProt As Boolean
    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo InstallDone

    Dim i As Long
    For i = 0 To 4
        ws.Cells(TD_ROW_ACTUAL, TD_COL_FIRST + i).Formula = _
            "=IFERROR(TdActualBit(""" & fields(i) & """," & SURVEY_BLOCK & "),"""")"
    Next i

InstallDone:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    On Error GoTo 0
End Sub

' ================================================================================
'  PLANNED TD ROW  --  filled from the imported plan
' ================================================================================

' Copy the last _OC_Survey station into Data I54:M54.  Called by Import Plan.
Public Sub RefreshTdPlannedFromPlan()
    Dim plan As Worksheet
    On Error Resume Next
    Set plan = ThisWorkbook.Worksheets(PLAN_SHEET)
    On Error GoTo 0
    If plan Is Nothing Then Exit Sub

    Dim cMD As Long, cInc As Long, cAzi As Long, cNS As Long, cEW As Long
    Dim c As Long
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

    ' Deepest station in the plan
    Dim lastR As Long: lastR = plan.Cells(plan.Rows.Count, cMD).End(xlUp).Row
    Do While lastR >= 3
        If IsNumeric(plan.Cells(lastR, cMD).Value2) _
           And Len(CStr(plan.Cells(lastR, cMD).Value2 & "")) > 0 Then Exit Do
        lastR = lastR - 1
    Loop
    If lastR < 3 Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim wasProt As Boolean
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
    On Error GoTo 0
End Sub

Private Function PlanNum(plan As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    Dim v As Variant: v = plan.Cells(r, c).Value2
    If IsNumeric(v) And Len(CStr(v & "")) > 0 Then
        PlanNum = CDbl(v)
    Else
        PlanNum = ""
    End If
End Function

' One-shot apply: live Actual formulas + Planned TD from the current plan.
Public Sub InstallTdCalc()
    InstallTdActualFormulas
    RefreshTdPlannedFromPlan
End Sub
