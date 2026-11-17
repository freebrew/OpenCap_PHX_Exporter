Attribute VB_Name = "MDL_DayRoll"
Option Explicit

' Autofill the current day-row in Data!H4:V22 from the RAW BHA SUMMARY mirror
' at Data!H35:M47 (BHA totals). Meters Sliding + Time Sliding are deltas
' against prior rows, unless the Slidesheet has Slide From/To intervals (R/S).
' Those intervals say where the metres were slid, so each daily row keeps only
' the overlap with its own Start/End depth. The by-day total has no depth and
' must not land on the first Midnight row. The current row is overwritten when
' refreshed so stale zeros do not block updated OpenCap totals. Bit/Circ hours
' come from a form.
'
' Hours units:
'   OpenCap / mirror K column = decimal hours (fraction = hundredths of an hour,
'   e.g. 15.33h). Column V formulas use M + N/60 (minutes as sixtieths).
'   When writing M:N we convert decimal hours -> hours + minutes via *60.

Private Const SH_DATA As String = "Data"
Private Const SH_OC_SLIDE As String = "_OC_Slide"
Private Const SH_OC_BHA As String = "_OC_BHA"
Private Const SH_FC_BHA As String = "_FC_BHA"
Private Const SH_SS As String = "Slidesheet"
Private Const SS_FIRST As Long = 13
Private Const SS_LAST As Long = 505
Private Const COL_SLIDE_FROM As Long = 18   ' R Slide From
Private Const COL_SLIDE_TO As Long = 19     ' S Slide To
Private Const ROW_FIRST As Long = 4
Private Const ROW_LAST As Long = 22
Private Const ROW_BHA As Long = 3
Private Const COL_PERIOD As Long = 8    ' H
Private Const COL_START As Long = 9     ' I
Private Const COL_END As Long = 10      ' J
Private Const COL_TOTAL As Long = 11    ' K
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
Public gDayRoll_PriorSlide As Double
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
        r = cell.row
        period = Trim$(CStr(cell.Value & ""))
        If IsValidPeriod(period) Then
            ' Always show the End Depth / Pason hours form when H is set.
            ' Refresh current-row L/M/N from the active BHA mirror first.
            FillNewDayRow ws, r, period
        End If
    Next cell
    Exit Sub

CleanFail:
    ' Never raise out of Worksheet_Change
End Sub

' Writes Start Depth / Meters Sliding / Time Sliding for row r.
' Explicit calls overwrite the current row's slide values by default.
' Returns True when at least one of those cells was written.
Public Function DayRoll_FillSlideMetrics(ByVal r As Long, _
        Optional ByVal overwriteCurrent As Boolean = True) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    DayRoll_FillSlideMetrics = FillSlideMetrics(ws, r, overwriteCurrent)
End Function

' Called from RefreshData after OpenCap / DD Tools rebuild.
' Fills blank historical L/M/N and overwrites the latest period row for the
' active BHA# (H3) from mirror deltas. Never opens the Pason form.
Public Sub DayRoll_BackfillSlideMetrics()
    Dim ws As Worksheet
    Dim r As Long
    Dim activeRow As Long
    Dim period As String
    Dim wasProt As Boolean
    Dim wrote As Boolean
    Dim bha As Long
    Dim hasDay As Boolean
    Dim dayM As Double
    Dim dayH As Double

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    wasProt = SheetUnprotectForVba(ws)

    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail

    activeRow = LastValidPeriodRow(ws)
    bha = ResolveActiveBha(ws)
    ' Depth windows re-place every period row. A Midnight row can already hold
    ' the whole by-day total while the slide sits in a later End of Run window.
    Dim depthOn As Boolean
    depthOn = HasSlideIntervals()
    For r = ROW_FIRST To ROW_LAST
        period = Trim$(CStr(ws.Cells(r, COL_PERIOD).Value & ""))
        If IsValidPeriod(period) Then
            ' A day the OpenCap by-day CSV covers is always re-posted for THAT
            ' row (even if it already holds a wrong value). Only rows with no
            ' by-day data fall back to the blank / active-row delta rules.
            hasDay = DaySlideForRow(ws, bha, r, dayM, dayH)
            If depthOn Or hasDay Or r = activeRow Or NeedsSlideBackfill(ws, r) Then
                wrote = FillSlideMetrics(ws, r, depthOn Or (r = activeRow) Or hasDay)
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
    ' End Depth just landed, so re-place slide metres across every row. The
    ' fill before the form had no window and may have parked them on an earlier line.
    DayRoll_BackfillSlideMetrics
    On Error Resume Next
    Application.Calculate
    ' Events are off here, so push the new Q24 into tool / motor Current Hours.
    ToolHours_Sync
    DDTools_SyncActiveBhaRow
    Application.Calculate
    On Error GoTo RestoreEvents
    gDayRoll_Applied = True
    DayRoll_ApplyPasonHours = True
    If Not gDayRoll_SuppressForm Then
        Dim warnMsg As String
        warnMsg = DayRoll_SlideHoursWarning()
        If warnMsg <> "" Then MsgBox warnMsg, vbExclamation, "Time Sliding"
    End If

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

    wrote = FillSlideMetrics(ws, r, True)
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

Private Function FillSlideMetrics(ByVal ws As Worksheet, ByVal r As Long, _
                                  Optional ByVal overwriteSlide As Boolean = False) As Boolean
    Dim srcRow As Long
    Dim bha As Long
    Dim mtrsSld As Double
    Dim hrsSld As Double
    Dim priorM As Double
    Dim priorH As Double
    Dim deltaM As Double
    Dim deltaH As Double
    Dim dayM As Double
    Dim dayH As Double
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

    ' FieldCap's own BHA totals. The mirror row of the active BHA is rebuilt
    ' from these day rows, so reading it here would feed the table back to itself.
    If Not FieldCapBhaTotals(bha, mtrsSld, hrsSld) Then
        mtrsSld = GetNum(ws.Cells(srcRow, SRC_MTRS_SLD))
        hrsSld = GetNum(ws.Cells(srcRow, SRC_HRS_SLD))
    End If
    priorM = SumNumeric(ws, COL_SLIDE_M, ROW_FIRST, r - 1)
    ' Prefer H:M (min/60) so backfill works even if col V formulas are stale
    ' under Calculation=Manual during Refresh.
    priorH = SumSlideHoursDecimal(ws, ROW_FIRST, r - 1)
    deltaM = Round(mtrsSld - priorM, 2)
    deltaH = Round(hrsSld - priorH, 4)

    ' Per-day OpenCap figures beat the "BHA total minus everything above" delta,
    ' which dumps every earlier mistake into this one row.
    Dim hasDay As Boolean
    hasDay = DaySlideForRow(ws, bha, r, dayM, dayH)
    If hasDay Then
        deltaM = dayM
        deltaH = dayH
        overwriteSlide = True
    End If

    ' Depth beats the by-day bucket for metres. The bucket has no MD, so a
    ' whole day's slide was landing on the first closing row. Hours stay the
    ' day's own: a share of the whole BHA's hours by metres moved hours
    ' between days with different slide ROP.
    ' The Slidesheet gets a slide when the next survey goes in, so a row that
    ' ends past its last slide can be missing the newest ones. FieldCap's day
    ' already has them.
    Dim depthM As Double, depthH As Double
    If DepthSlideForRow(ws, bha, r, mtrsSld, hrsSld, depthM, depthH) Then
        deltaM = depthM
        If hasDay Then
            If dayM > depthM + 0.005 And RowEndsPastSlides(ws, r) Then deltaM = dayM
        Else
            deltaH = depthH
        End If
        overwriteSlide = True
    End If

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

    If deltaM >= 0 And (overwriteSlide Or CellBlank(ws.Cells(r, COL_SLIDE_M))) Then
        ws.Cells(r, COL_SLIDE_M).Value = deltaM
        wrote = True
    End If

    ' The depth share and the BHA delta know nothing of this row's bit
    ' hours, and slide time past them prints negative rotating time.
    Dim rowBit As Double
    rowBit = GetNum(ws.Cells(r, COL_BIT))
    If rowBit > 0 And deltaH > rowBit Then deltaH = rowBit

    ' A source with no slide hours (stale or partial _OC_Slide) must not
    ' replace hours the DD typed with 0:00.
    Dim hrsBlank As Boolean
    hrsBlank = CellBlank(ws.Cells(r, COL_HRS)) And CellBlank(ws.Cells(r, COL_MIN))
    If deltaH >= 0 And (deltaH > 0 Or hrsBlank) And (overwriteSlide Or hrsBlank) Then
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
    gDayRoll_PriorSlide = SumSlideHoursDecimal(ws, ROW_FIRST, r - 1)

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

Private Function IsClosingPeriod(ByVal s As String) As Boolean
    IsClosingPeriod = (s = "Midnight" Or s = "End of Run")
End Function

' Slide metres / hours for the report day that daily-table row r closes, taken
' from the OpenCap "slide-rotate by day" CSV (hidden _OC_Slide) for BHA bha.
' Days are matched in date order to the Midnight / End of Run rows of the table
' (the table has no date column). 6am / 4pm rows inside a day are subtracted
' so the closing row carries only the remainder. False = no by-day data for
' this row (caller keeps the old delta rules).
' FieldCap logs a slide stand as one Sliding block, connections and surveys
' included, so its Sliding + Drilling time runs past Pason's on-bottom hours.
' The day keeps FieldCap's slide-to-rotate split of the bit hours instead,
' and slide time can never exceed bit time.
Private Function DaySlideForRow(ByVal ws As Worksheet, ByVal bha As Long, _
                                ByVal r As Long, ByRef outM As Double, _
                                ByRef outH As Double) As Boolean
    Dim wsS As Worksheet
    Dim lastR As Long, i As Long, j As Long, n As Long, pos As Long
    Dim keys() As Double, mt() As Double, hr() As Double, ro() As Double
    Dim v As Variant
    Dim k As Double, m As Double, h As Double, rh As Double
    Dim bitDay As Double, dayH As Double
    Dim period As String
    Dim dayIdx As Long, rr As Long, firstOfDay As Long

    DaySlideForRow = False
    outM = 0
    outH = 0
    If bha <= 0 Or r < ROW_FIRST Or r > ROW_LAST Then Exit Function
    period = Trim$(CStr(ws.Cells(r, COL_PERIOD).Value & ""))
    If Not IsClosingPeriod(period) Then Exit Function

    On Error Resume Next
    Set wsS = ThisWorkbook.Worksheets(SH_OC_SLIDE)
    On Error GoTo 0
    If wsS Is Nothing Then Exit Function
    lastR = wsS.Cells(wsS.rows.Count, 3).End(xlUp).row
    If lastR < 2 Then Exit Function

    ReDim keys(1 To lastR)
    ReDim mt(1 To lastR)
    ReDim hr(1 To lastR)
    ReDim ro(1 To lastR)
    n = 0
    For i = 2 To lastR
        v = wsS.Cells(i, 3).Value2
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CLng(v) = bha Then
                v = wsS.Cells(i, 2).Value2
                If IsNumeric(v) And Not IsEmpty(v) Then
                    k = CDbl(v)
                ElseIf IsDate(v) Then
                    k = CDbl(CDate(v))
                Else
                    k = i
                End If
                m = GetNum(wsS.Cells(i, 4))
                h = GetNum(wsS.Cells(i, 6))
                rh = GetNum(wsS.Cells(i, 7))
                pos = 0
                For j = 1 To n
                    If keys(j) = k Then pos = j: Exit For
                Next j
                If pos > 0 Then
                    mt(pos) = mt(pos) + m
                    hr(pos) = hr(pos) + h
                    ro(pos) = ro(pos) + rh
                Else
                    pos = n + 1
                    For j = 1 To n
                        If keys(j) > k Then pos = j: Exit For
                    Next j
                    For j = n To pos Step -1
                        keys(j + 1) = keys(j): mt(j + 1) = mt(j): hr(j + 1) = hr(j): ro(j + 1) = ro(j)
                    Next j
                    keys(pos) = k: mt(pos) = m: hr(pos) = h: ro(pos) = rh
                    n = n + 1
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function

    ' Which report day does row r close?
    dayIdx = 0
    firstOfDay = ROW_FIRST
    For rr = ROW_FIRST To r
        If IsClosingPeriod(Trim$(CStr(ws.Cells(rr, COL_PERIOD).Value & ""))) Then
            dayIdx = dayIdx + 1
            If rr < r Then firstOfDay = rr + 1
        End If
    Next rr
    If dayIdx < 1 Or dayIdx > n Then Exit Function

    dayH = hr(dayIdx)
    bitDay = SumNumeric(ws, COL_BIT, firstOfDay, r)
    If bitDay > 0 And dayH + ro(dayIdx) > bitDay Then
        dayH = dayH * bitDay / (dayH + ro(dayIdx))
    End If
    outM = mt(dayIdx) - SumNumeric(ws, COL_SLIDE_M, firstOfDay, r - 1)
    outH = dayH - SumSlideHoursDecimal(ws, firstOfDay, r - 1)
    If outM < 0 Then outM = 0
    If outH < 0 Then outH = 0
    outM = Round(outM, 2)
    outH = Round(outH, 4)
    DaySlideForRow = True
End Function

' True when Slidesheet R/S holds at least one real slide interval.
Private Function HasSlideIntervals() As Boolean
    Dim ss As Worksheet
    Dim r As Long
    Dim va As Variant, vb As Variant
    HasSlideIntervals = False
    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SH_SS)
    On Error GoTo 0
    If ss Is Nothing Then Exit Function
    For r = SS_FIRST To SS_LAST
        va = ss.Cells(r, COL_SLIDE_FROM).Value2
        vb = ss.Cells(r, COL_SLIDE_TO).Value2
        If IsNumeric(va) And IsNumeric(vb) Then
            If CDbl(vb) > CDbl(va) Then
                HasSlideIntervals = True
                Exit Function
            End If
        End If
    Next r
End Function

' True when row r ends deeper than the last Slide To on the Slidesheet.
Private Function RowEndsPastSlides(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim ss As Worksheet
    Dim rr As Long
    Dim va As Variant, vb As Variant
    Dim lo As Double, hi As Double, deepest As Double
    RowEndsPastSlides = False
    If Not RowDepthWindow(ws, r, lo, hi) Then Exit Function
    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SH_SS)
    On Error GoTo 0
    If ss Is Nothing Then Exit Function
    deepest = 0#
    For rr = SS_FIRST To SS_LAST
        va = ss.Cells(rr, COL_SLIDE_FROM).Value2
        vb = ss.Cells(rr, COL_SLIDE_TO).Value2
        If IsNumeric(va) And IsNumeric(vb) And Not IsEmpty(va) And Not IsEmpty(vb) Then
            If CDbl(vb) > CDbl(va) And CDbl(vb) > deepest Then deepest = CDbl(vb)
        End If
    Next rr
    RowEndsPastSlides = (hi > deepest + 0.5)
End Function

' FieldCap's slide metres / hours for one BHA, from the OpenCap BHA CSV.
Private Function FieldCapBhaTotals(ByVal bha As Long, ByRef outM As Double, _
                                   ByRef outH As Double) As Boolean
    Dim names As Variant, nm As Variant
    Dim wsB As Worksheet
    Dim cB As Long, cM As Long, cH As Long, c As Long, i As Long, lastR As Long
    Dim hdr As String
    Dim v As Variant
    FieldCapBhaTotals = False
    outM = 0#
    outH = 0#
    If bha <= 0 Then Exit Function
    names = Array(SH_OC_BHA, SH_FC_BHA)
    For Each nm In names
        Set wsB = Nothing
        On Error Resume Next
        Set wsB = ThisWorkbook.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not wsB Is Nothing Then
            cB = 0: cM = 0: cH = 0
            For c = 1 To 80
                hdr = Trim$(CStr(wsB.Cells(1, c).Value & ""))
                If hdr = "BHA #" Then cB = c
                If hdr = "BHA Mtrs Slid" Then cM = c
                If hdr = "BHA Hrs Slid" Then cH = c
            Next c
            If cB > 0 And (cM > 0 Or cH > 0) Then
                lastR = wsB.Cells(wsB.rows.Count, cB).End(xlUp).row
                For i = 2 To lastR
                    v = wsB.Cells(i, cB).Value2
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        If CLng(v) = bha Then
                            If cM > 0 Then outM = GetNum(wsB.Cells(i, cM))
                            If cH > 0 Then outH = GetNum(wsB.Cells(i, cH))
                            FieldCapBhaTotals = True
                            Exit Function
                        End If
                    End If
                Next i
            End If
        End If
    Next nm
End Function

' Totals of the Data day rows (H4:V22) for the BHA in Data!H3: the numbers
' the email's BHA lines print, so the BHA table can show the same ones.
' Rotate hours are bit hours less slide hours. False = no day rows.
Public Function DayRoll_ActiveBhaTotals(ByRef bha As Long, ByRef metres As Double, _
                                        ByRef slideM As Double, ByRef slideH As Double, _
                                        ByRef rotH As Double, ByRef circH As Double) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim bit As Double
    Dim lo As Double, hi As Double
    DayRoll_ActiveBhaTotals = False
    bha = 0: metres = 0#: slideM = 0#: slideH = 0#: rotH = 0#: circH = 0#
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    If Not IsNumeric(ws.Cells(ROW_BHA, COL_PERIOD).Value) Then Exit Function
    If CellBlank(ws.Cells(ROW_BHA, COL_PERIOD)) Then Exit Function
    bha = CLng(ws.Cells(ROW_BHA, COL_PERIOD).Value)
    For r = ROW_FIRST To ROW_LAST
        If IsValidPeriod(Trim$(CStr(ws.Cells(r, COL_PERIOD).Value & ""))) Then
            If RowDepthWindow(ws, r, lo, hi) Then
                metres = metres + (hi - lo)
            Else
                metres = metres + GetNum(ws.Cells(r, COL_TOTAL))
            End If
            slideM = slideM + GetNum(ws.Cells(r, COL_SLIDE_M))
            slideH = slideH + GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
            bit = bit + GetNum(ws.Cells(r, COL_BIT))
            circH = circH + GetNum(ws.Cells(r, COL_CIRC))
            DayRoll_ActiveBhaTotals = True
        End If
    Next r
    rotH = bit - slideH
    If rotH < 0# Then rotH = 0#
End Function

' Day rows whose slide time is more than their bit time (rotating time would
' go negative in the email). Empty string = none.
Public Function DayRoll_SlideHoursWarning() As String
    Dim ws As Worksheet
    Dim r As Long
    Dim slideH As Double, bit As Double
    Dim msg As String
    DayRoll_SlideHoursWarning = ""
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = ROW_FIRST To ROW_LAST
        If IsValidPeriod(Trim$(CStr(ws.Cells(r, COL_PERIOD).Value & ""))) Then
            If IsNumeric(ws.Cells(r, COL_BIT).Value) And Not CellBlank(ws.Cells(r, COL_BIT)) Then
                bit = CDbl(ws.Cells(r, COL_BIT).Value)
                slideH = GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
                If slideH > bit + 0.01 Then
                    msg = msg & vbLf & "  Row " & r & " (" & ws.Cells(r, COL_PERIOD).Value & "): " & _
                          Format$(slideH, "0.00") & " h sliding, " & Format$(bit, "0.00") & " bit hours"
                End If
            End If
        End If
    Next r
    If msg <> "" Then
        DayRoll_SlideHoursWarning = "Time Sliding is more than the bit hours on:" & msg & vbLf & vbLf & _
            "Rotating time would print negative in the email. Fix the slide time (M:N) or the bit hours (P) on those rows."
    End If
End Function

Private Function RowDepthWindow(ByVal ws As Worksheet, ByVal r As Long, _
                                ByRef lo As Double, ByRef hi As Double) As Boolean
    Dim vLo As Variant, vHi As Variant
    RowDepthWindow = False
    vLo = ws.Cells(r, COL_START).Value2
    vHi = ws.Cells(r, COL_END).Value2
    If Not IsNumeric(vLo) Or Not IsNumeric(vHi) Then Exit Function
    If Len(CStr(vLo & "")) = 0 Or Len(CStr(vHi & "")) = 0 Then Exit Function
    lo = CDbl(vLo)
    hi = CDbl(vHi)
    RowDepthWindow = (hi > lo)
End Function

' Sum of OpenCap by-day slide metres / hours for this BHA. False = no by-day rows.
Private Function BhaByDayPool(ByVal bha As Long, ByRef poolM As Double, _
                              ByRef poolH As Double) As Boolean
    Dim wsS As Worksheet
    Dim lastR As Long, i As Long
    Dim v As Variant
    poolM = 0#
    poolH = 0#
    BhaByDayPool = False
    If bha <= 0 Then Exit Function
    On Error Resume Next
    Set wsS = ThisWorkbook.Worksheets(SH_OC_SLIDE)
    On Error GoTo 0
    If wsS Is Nothing Then Exit Function
    lastR = wsS.Cells(wsS.rows.Count, 3).End(xlUp).row
    If lastR < 2 Then Exit Function
    For i = 2 To lastR
        v = wsS.Cells(i, 3).Value2
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CLng(v) = bha Then
                poolM = poolM + GetNum(wsS.Cells(i, 4))
                poolH = poolH + GetNum(wsS.Cells(i, 6))
                BhaByDayPool = True
            End If
        End If
    Next i
End Function

' Slide metres inside this row's Start/End, from Slidesheet R/S overlap.
' Hours follow the metres: this row's share of the by-day slide hours
' (BHA mirror hours when the by-day file has no rows for this BHA).
' False when there is no interval to validate, or this row has no depth window.
Private Function DepthSlideForRow(ByVal ws As Worksheet, ByVal bha As Long, _
                                  ByVal r As Long, ByVal mirrorM As Double, _
                                  ByVal mirrorH As Double, _
                                  ByRef outM As Double, ByRef outH As Double) As Boolean
    Dim ss As Worksheet
    Dim lo As Double, hi As Double
    Dim vM As Variant
    Dim poolM As Double, poolH As Double
    DepthSlideForRow = False
    outM = 0#
    outH = 0#
    If Not HasSlideIntervals() Then Exit Function
    If Not RowDepthWindow(ws, r, lo, hi) Then Exit Function
    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SH_SS)
    On Error GoTo 0
    If ss Is Nothing Then Exit Function
    vM = MDL_ProjBit.ProjSlideMetersBetween(lo, hi, _
            ss.Range(ss.Cells(SS_FIRST, COL_SLIDE_FROM), ss.Cells(SS_LAST, COL_SLIDE_FROM)), _
            ss.Range(ss.Cells(SS_FIRST, COL_SLIDE_TO), ss.Cells(SS_LAST, COL_SLIDE_TO)))
    If isError(vM) Then Exit Function
    If Not IsNumeric(vM) Then Exit Function
    outM = Round(CDbl(vM), 2)
    If Not BhaByDayPool(bha, poolM, poolH) Then
        poolM = mirrorM
        poolH = mirrorH
    End If
    If poolM > 0.0001 And outM > 0# And poolH > 0# Then
        outH = Round(poolH * (outM / poolM), 4)
    Else
        outH = 0#
    End If
    If outH < 0# Then outH = 0#
    DepthSlideForRow = True
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

Private Function LastValidPeriodRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    LastValidPeriodRow = 0
    For r = ROW_LAST To ROW_FIRST Step -1
        If IsValidPeriod(Trim$(CStr(ws.Cells(r, COL_PERIOD).Value & ""))) Then
            LastValidPeriodRow = r
            Exit Function
        End If
    Next r
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





















