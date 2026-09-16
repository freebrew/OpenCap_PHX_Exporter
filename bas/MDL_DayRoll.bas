Attribute VB_Name = "MDL_DayRoll"
Option Explicit

' Autofill a new day-row in Data!H4:V22 from the RAW BHA SUMMARY mirror
' at Data!H35:M47 (BHA totals). Meters Sliding + Time Sliding are deltas
' against rows already entered. Bit/Circ hours come from a Pason-totals form.
'
' Hours units:
'   OpenCap / mirror K column = decimal hours (fraction = hundredths of an hour,
'   e.g. 15.33h). Column V formulas use M + N/60 (minutes as sixtieths).
'   When writing M:N we convert decimal hours -> hours + minutes via *60.

Private Const SH_DATA As String = "Data"
Private Const ROW_FIRST As Long = 4
Private Const ROW_LAST As Long = 22
Private Const ROW_BHA As Long = 3
Private Const COL_PERIOD As Long = 8    ' H
Private Const COL_START As Long = 9     ' I
Private Const COL_END As Long = 10      ' J
Private Const COL_SLIDE_M As Long = 12  ' L
Private Const COL_HRS As Long = 13      ' M
Private Const COL_MIN As Long = 14      ' N
Private Const COL_BIT As Long = 16      ' P
Private Const COL_CIRC As Long = 17     ' Q
Private Const COL_SLIDE_DEC As Long = 22 ' V

Private Const SRC_FIRST As Long = 36
Private Const SRC_LAST As Long = 47
Private Const SRC_BHA As Long = 8       ' H
Private Const SRC_MTRS_SLD As Long = 9  ' I
Private Const SRC_HRS_SLD As Long = 11  ' K

' Context shared with frmDayHours
Public gDayRoll_TargetRow As Long
Public gDayRoll_BHA As Long
Public gDayRoll_Period As String
Public gDayRoll_PriorBit As Double
Public gDayRoll_PriorCirc As Double
Public gDayRoll_SlideHrs As Double
Public gDayRoll_StartDepth As Double
Public gDayRoll_HasStartDepth As Boolean
Public gDayRoll_Applied As Boolean

' Smoke-test hook: when True, skip the modal form after writing I/L/M/N.
Public gDayRoll_SuppressForm As Boolean

Public Sub DayRoll_OnDataChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim area As Range
    Dim cell As Range
    Dim r As Long
    Dim period As String

    On Error GoTo CleanFail
    If Target Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    Set area = Intersect(Target, ws.Range(ws.Cells(ROW_FIRST, COL_PERIOD), _
                                          ws.Cells(ROW_LAST, COL_PERIOD)))
    If area Is Nothing Then Exit Sub

    For Each cell In area.Cells
        r = cell.Row
        period = Trim$(CStr(cell.Value & ""))
        If IsValidPeriod(period) Then
            ' Always show the End Depth / Pason hours form when H is set.
            ' FillSlideMetrics only writes blank L/M/N; it does not overwrite.
            FillNewDayRow ws, r, period
        End If
    Next cell
    Exit Sub

CleanFail:
    ' Never raise out of Worksheet_Change
End Sub

' Writes Start Depth / Meters Sliding / Time Sliding for row r.
' Returns True when at least one of those cells was written.
Public Function DayRoll_FillSlideMetrics(ByVal r As Long) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    DayRoll_FillSlideMetrics = FillSlideMetrics(ws, r)
End Function

' Called from RefreshData after OpenCap / DD Tools rebuild.
' Fills blank L/M/N on daily rows for the active BHA# (H3) from the mirror
' deltas. Never overwrites non-blank slide cells; never opens the Pason form.
Public Sub DayRoll_BackfillSlideMetrics()
    Dim ws As Worksheet
    Dim r As Long
    Dim period As String
    Dim wasProt As Boolean
    Dim wrote As Boolean

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    wasProt = SheetUnprotectForVba(ws)

    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail

    For r = ROW_FIRST To ROW_LAST
        period = Trim$(CStr(ws.Cells(r, COL_PERIOD).Value & ""))
        If IsValidPeriod(period) Then
            If NeedsSlideBackfill(ws, r) Then
                wrote = FillSlideMetrics(ws, r)
                If wrote Then
                    On Error Resume Next
                    Application.Calculate
                    On Error GoTo Fail
                End If
            End If
        End If
    Next r

    SheetReprotectAfterVba ws, wasProt
    Exit Sub

Fail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    On Error GoTo 0
End Sub

' Apply Pason running totals into P/Q and End Depth into J for gDayRoll_TargetRow.
' Validates deltas + end depth; returns False with errMsg on failure (no write).
Public Function DayRoll_ApplyPasonHours(ByVal bitTotal As Double, _
                                       ByVal circTotal As Double, _
                                       ByVal endDepth As Double, _
                                       ByRef errMsg As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim bitDelta As Double
    Dim circDelta As Double
    Dim prevEvents As Boolean

    errMsg = ""
    DayRoll_ApplyPasonHours = False

    r = gDayRoll_TargetRow
    If r < ROW_FIRST Or r > ROW_LAST Then
        errMsg = "No target row selected."
        Exit Function
    End If

    If Not ValidateEndDepth(endDepth, errMsg) Then Exit Function

    bitDelta = Round(bitTotal - gDayRoll_PriorBit, 4)
    circDelta = Round(circTotal - gDayRoll_PriorCirc, 4)

    If Not ValidateHourDeltas(bitDelta, circDelta, gDayRoll_SlideHrs, _
                             gDayRoll_Period, errMsg) Then
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo RestoreEvents

    ws.Cells(r, COL_END).Value = endDepth
    ws.Cells(r, COL_BIT).Value = bitDelta
    ws.Cells(r, COL_CIRC).Value = circDelta
    ' No bit hours today -> no sliding either. Fill blank L/M/N with zeros so
    ' the row's ROP / % formulas resolve instead of sitting empty.
    If bitDelta = 0 Then
        If CellBlank(ws.Cells(r, COL_SLIDE_M)) Then ws.Cells(r, COL_SLIDE_M).Value = 0
        If CellBlank(ws.Cells(r, COL_HRS)) Then ws.Cells(r, COL_HRS).Value = 0
        If CellBlank(ws.Cells(r, COL_MIN)) Then WriteMinutesCell ws.Cells(r, COL_MIN), 0
    End If
    On Error Resume Next
    Application.Calculate
    ' Events are off here, so push the new Q24 into tool / motor Current Hours.
    ToolHours_Sync
    On Error GoTo RestoreEvents
    gDayRoll_Applied = True
    DayRoll_ApplyPasonHours = True

RestoreEvents:
    Application.EnableEvents = prevEvents
End Function

' Shared validation used by the form preview and Apply.
Public Function DayRoll_ValidateHourDeltas(ByVal bitDelta As Double, _
                                          ByVal circDelta As Double, _
                                          ByVal slideHrs As Double, _
                                          ByVal period As String, _
                                          ByRef errMsg As String) As Boolean
    DayRoll_ValidateHourDeltas = ValidateHourDeltas(bitDelta, circDelta, _
                                                    slideHrs, period, errMsg)
End Function

' ========================================================================
'  Internals
' ========================================================================

Private Sub FillNewDayRow(ByVal ws As Worksheet, ByVal r As Long, _
                          ByVal period As String)
    Dim wrote As Boolean

    wrote = FillSlideMetrics(ws, r)
    If Not wrote And CellBlank(ws.Cells(r, COL_SLIDE_M)) _
                  And CellBlank(ws.Cells(r, COL_HRS)) Then
        ' Nothing useful from source; still allow Pason hours entry.
    End If

    PrepareFormContext ws, r, period

    If gDayRoll_SuppressForm Then Exit Sub
    On Error Resume Next
    frmDayHours.Show vbModal
    On Error GoTo 0
End Sub

Private Function FillSlideMetrics(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim srcRow As Long
    Dim bha As Long
    Dim mtrsSld As Double
    Dim hrsSld As Double
    Dim priorM As Double
    Dim priorH As Double
    Dim deltaM As Double
    Dim deltaH As Double
    Dim hrsPart As Long
    Dim minPart As Long
    Dim prevEvents As Boolean
    Dim wrote As Boolean

    FillSlideMetrics = False
    wrote = False

    bha = ResolveActiveBha(ws)
    If bha = 0 Then Exit Function
    srcRow = FindSourceRow(ws, bha)
    If srcRow = 0 Then Exit Function

    mtrsSld = GetNum(ws.Cells(srcRow, SRC_MTRS_SLD))
    hrsSld = GetNum(ws.Cells(srcRow, SRC_HRS_SLD))
    priorM = SumNumeric(ws, COL_SLIDE_M, ROW_FIRST, r - 1)
    ' Prefer H:M (min/60) so backfill works even if col V formulas are stale
    ' under Calculation=Manual during Refresh.
    priorH = SumSlideHoursDecimal(ws, ROW_FIRST, r - 1)
    deltaM = Round(mtrsSld - priorM, 2)
    deltaH = Round(hrsSld - priorH, 4)

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo RestoreSlide

    ' Start Depth from previous End Depth
    If r > ROW_FIRST Then
        If CellBlank(ws.Cells(r, COL_START)) Then
            If Not CellBlank(ws.Cells(r - 1, COL_END)) Then
                If IsNumeric(ws.Cells(r - 1, COL_END).Value) Then
                    ws.Cells(r, COL_START).Value = ws.Cells(r - 1, COL_END).Value
                    wrote = True
                End If
            End If
        End If
    End If

    If deltaM >= 0 And CellBlank(ws.Cells(r, COL_SLIDE_M)) Then
        ws.Cells(r, COL_SLIDE_M).Value = deltaM
        wrote = True
    End If

    If deltaH >= 0 And CellBlank(ws.Cells(r, COL_HRS)) _
                   And CellBlank(ws.Cells(r, COL_MIN)) Then
        SplitDecimalHours deltaH, hrsPart, minPart
        ws.Cells(r, COL_HRS).Value = hrsPart
        WriteMinutesCell ws.Cells(r, COL_MIN), minPart
        wrote = True
    End If

    FillSlideMetrics = wrote
    On Error Resume Next
    Application.Calculate
    On Error GoTo RestoreSlide

RestoreSlide:
    Application.EnableEvents = prevEvents
End Function

Private Sub PrepareFormContext(ByVal ws As Worksheet, ByVal r As Long, _
                               ByVal period As String)
    Dim slideDec As Double

    gDayRoll_TargetRow = r
    gDayRoll_Period = period
    gDayRoll_Applied = False
    gDayRoll_PriorBit = SumNumeric(ws, COL_BIT, ROW_FIRST, r - 1)
    gDayRoll_PriorCirc = SumNumeric(ws, COL_CIRC, ROW_FIRST, r - 1)

    If IsNumeric(ws.Cells(ROW_BHA, COL_PERIOD).Value) Then
        gDayRoll_BHA = CLng(ws.Cells(ROW_BHA, COL_PERIOD).Value)
    Else
        gDayRoll_BHA = 0
    End If

    If IsNumeric(ws.Cells(r, COL_START).Value) Then
        gDayRoll_StartDepth = CDbl(ws.Cells(r, COL_START).Value)
        gDayRoll_HasStartDepth = True
    Else
        gDayRoll_StartDepth = 0
        gDayRoll_HasStartDepth = False
    End If

    slideDec = 0
    If IsNumeric(ws.Cells(r, COL_SLIDE_DEC).Value) Then
        slideDec = CDbl(ws.Cells(r, COL_SLIDE_DEC).Value)
    ElseIf IsNumeric(ws.Cells(r, COL_HRS).Value) Or IsNumeric(ws.Cells(r, COL_MIN).Value) Then
        slideDec = GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
    End If
    gDayRoll_SlideHrs = Round(slideDec, 4)
End Sub

Private Function ValidateEndDepth(ByVal endDepth As Double, _
                                  ByRef errMsg As String) As Boolean
    ValidateEndDepth = False
    errMsg = ""

    If endDepth <= 0 Then
        errMsg = "End Depth must be a positive number."
        Exit Function
    End If
    If gDayRoll_HasStartDepth Then
        If endDepth + 0.0000001 < gDayRoll_StartDepth Then
            errMsg = "End Depth must be >= Start Depth (" & _
                     Format$(gDayRoll_StartDepth, "0.##") & ")."
            Exit Function
        End If
    End If

    ValidateEndDepth = True
End Function

Private Function ValidateHourDeltas(ByVal bitDelta As Double, _
                                   ByVal circDelta As Double, _
                                   ByVal slideHrs As Double, _
                                   ByVal period As String, _
                                   ByRef errMsg As String) As Boolean
    ValidateHourDeltas = False
    errMsg = ""

    ' Zero is a valid day (no drilling / no circulating). Only a running
    ' total that went backwards is an error.
    If bitDelta < 0 Or circDelta < 0 Then
        errMsg = "Bit and Circ deltas cannot be negative (running totals must not go backwards)."
        Exit Function
    End If
    If circDelta + 0.0000001 < bitDelta Then
        errMsg = "Circ Hours must be >= Bit Hours for this row."
        Exit Function
    End If
    If bitDelta + 0.0000001 < slideHrs Then
        errMsg = "Bit Hours must be >= Sliding Hours (" & _
                 Format$(slideHrs, "0.##") & ")."
        Exit Function
    End If
    If StrComp(period, "Midnight", vbTextCompare) = 0 Then
        If circDelta > 24# + 0.0000001 Then
            errMsg = "Circ Hours cannot exceed 24 for a Midnight row."
            Exit Function
        End If
    End If

    ValidateHourDeltas = True
End Function

' H3 BHA# must exist in the H36:H47 mirror. After a new-job import the header
' is often still the previous well's BHA (e.g. 6) while the mirror only has 1.
Private Function ResolveActiveBha(ByVal ws As Worksheet) As Long
    Dim headerBha As Long
    Dim lastBha As Long
    ResolveActiveBha = 0
    lastBha = LastSourceBha(ws)
    If IsNumeric(ws.Cells(ROW_BHA, COL_PERIOD).Value) Then
        headerBha = CLng(ws.Cells(ROW_BHA, COL_PERIOD).Value)
        If FindSourceRow(ws, headerBha) > 0 Then
            ResolveActiveBha = headerBha
            Exit Function
        End If
    End If
    If lastBha > 0 Then
        ws.Cells(ROW_BHA, COL_PERIOD).Value = lastBha
        ResolveActiveBha = lastBha
    End If
End Function

Private Function LastSourceBha(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim v As Variant
    LastSourceBha = 0
    For r = SRC_FIRST To SRC_LAST
        v = ws.Cells(r, SRC_BHA).Value
        If IsNumeric(v) And v <> "" Then LastSourceBha = CLng(v)
    Next r
End Function

Private Function FindSourceRow(ByVal ws As Worksheet, ByVal bha As Long) As Long
    Dim r As Long
    Dim v As Variant
    FindSourceRow = 0
    For r = SRC_FIRST To SRC_LAST
        v = ws.Cells(r, SRC_BHA).Value
        If IsNumeric(v) Then
            If CLng(v) = bha Then
                FindSourceRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function RowIsEmptyForFill(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    RowIsEmptyForFill = CellBlank(ws.Cells(r, COL_START)) _
                    And CellBlank(ws.Cells(r, COL_SLIDE_M)) _
                    And CellBlank(ws.Cells(r, COL_HRS)) _
                    And CellBlank(ws.Cells(r, COL_MIN)) _
                    And CellBlank(ws.Cells(r, COL_BIT)) _
                    And CellBlank(ws.Cells(r, COL_CIRC))
End Function

Private Function NeedsSlideBackfill(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    NeedsSlideBackfill = CellBlank(ws.Cells(r, COL_SLIDE_M)) _
                     Or (CellBlank(ws.Cells(r, COL_HRS)) _
                         And CellBlank(ws.Cells(r, COL_MIN)))
End Function

Private Function IsValidPeriod(ByVal s As String) As Boolean
    Select Case s
        Case "6am", "4pm", "Midnight", "End of Run"
            IsValidPeriod = True
        Case Else
            IsValidPeriod = False
    End Select
End Function

Private Sub SplitDecimalHours(ByVal decHrs As Double, _
                              ByRef hrsPart As Long, _
                              ByRef minPart As Long)
    Dim totalMin As Long
    If decHrs < 0 Then
        hrsPart = 0
        minPart = 0
        Exit Sub
    End If
    totalMin = CLng(Round(decHrs * 60#, 0))
    hrsPart = totalMin \ 60
    minPart = totalMin Mod 60
    If minPart = 60 Then
        hrsPart = hrsPart + 1
        minPart = 0
    End If
End Sub

' Honour existing NumberFormat on N: Text (@) gets zero-padded "05", else numeric.
Private Sub WriteMinutesCell(ByVal cell As Range, ByVal minutes As Long)
    Dim fmt As String
    fmt = CStr(cell.numberFormat & "")
    If InStr(1, fmt, "@", vbBinaryCompare) > 0 Then
        cell.Value = Format$(minutes, "00")
    Else
        cell.Value = minutes
    End If
End Sub

Private Function SumNumeric(ByVal ws As Worksheet, ByVal col As Long, _
                            ByVal r1 As Long, ByVal r2 As Long) As Double
    Dim r As Long
    Dim tot As Double
    tot = 0
    If r2 < r1 Then
        SumNumeric = 0
        Exit Function
    End If
    For r = r1 To r2
        If IsNumeric(ws.Cells(r, col).Value) Then
            tot = tot + CDbl(ws.Cells(r, col).Value)
        End If
    Next r
    SumNumeric = tot
End Function

' Decimal hours from daily rows: hours + minutes/60 (not minutes/100).
Private Function SumSlideHoursDecimal(ByVal ws As Worksheet, _
                                      ByVal r1 As Long, _
                                      ByVal r2 As Long) As Double
    Dim r As Long
    Dim tot As Double
    Dim dec As Double
    tot = 0
    If r2 < r1 Then
        SumSlideHoursDecimal = 0
        Exit Function
    End If
    For r = r1 To r2
        dec = 0
        If IsNumeric(ws.Cells(r, COL_SLIDE_DEC).Value) Then
            dec = CDbl(ws.Cells(r, COL_SLIDE_DEC).Value)
        ElseIf IsNumeric(ws.Cells(r, COL_HRS).Value) _
            Or IsNumeric(ws.Cells(r, COL_MIN).Value) Then
            dec = GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
        End If
        tot = tot + dec
    Next r
    SumSlideHoursDecimal = tot
End Function

Private Function GetNum(ByVal cell As Range) As Double
    If IsNumeric(cell.Value) Then
        GetNum = CDbl(cell.Value)
    Else
        GetNum = 0
    End If
End Function

Private Function CellBlank(ByVal cell As Range) As Boolean
    Dim v As Variant
    v = cell.Value
    If isError(v) Then
        CellBlank = True
    ElseIf IsNull(v) Then
        CellBlank = True
    ElseIf IsEmpty(v) Then
        CellBlank = True
    ElseIf Len(Trim$(CStr(v & ""))) = 0 Then
        CellBlank = True
    Else
        CellBlank = False
    End If
End Function








