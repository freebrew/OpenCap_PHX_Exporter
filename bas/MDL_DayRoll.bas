Attribute VB_Name = "MDL_DayRoll"
Option Explicit

' Autofill the current day-row in Data!H4:V22 from the RAW BHA SUMMARY mirror
' at Data!H35:M47 (BHA totals). A day the FieldCap by-day CSV covers takes
' FieldCap's slide metres and slide time as logged (L, M:N) on every REFRESH;
' everything below is the fallback for days it does not cover.
' Without a FieldCap day, Meters Sliding + Time Sliding are deltas
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
'   Time Rotating (C10, C17, S) is Pason bit hours minus slide hours (P-V).
'   Total metres K = End-Start (J-I). Rotated O = K-L. Email C4:C20 reads
'   the daily table only (rows 23/24). No sidecar AJ:AM helper columns.

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
Private Const COL_ROTATED As Long = 15  ' O
Private Const COL_BIT As Long = 16      ' P
Private Const COL_CIRC As Long = 17     ' Q
Private Const COL_SLIDE_DEC As Long = 22 ' V
' Retired sidecar AJ:AM (FieldCap Total / Rotated / Rot Hrs / Rotating Hrs).
' Cleared on every refresh so they cannot come back.
Private Const COL_HELPER_FIRST As Long = 36  ' AJ
Private Const COL_HELPER_LAST As Long = 39   ' AM
Private Const ROW_TODAY As Long = 23
Private Const ROW_TOTALS As Long = 24
' Date serials 2000-01-01 .. 2099-12-31; a by-day key outside is a row index, not a day.
Private Const DATE_SERIAL_MIN As Double = 36526
Private Const DATE_SERIAL_MAX As Double = 73050
Private Const MAX_FILL_DAYS As Long = 400

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

' One report day as FieldCap logged it, less the 6am / 4pm rows above it in that day.
Private Type FcDay
    slideM As Double
    slideH As Double
    rotM As Double
    rotH As Double
    totalM As Double
End Type

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
        period = Trim$(CStr(cell.value & ""))
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
    Dim day As FcDay
    Dim prevEvents As Boolean

    prevEvents = Application.EnableEvents
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    wasProt = SheetUnprotectForVba(ws)

    Application.EnableEvents = False
    EnsureFieldCapFormulas ws
    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail

    activeRow = LastValidPeriodRow(ws)
    bha = ResolveActiveBha(ws)
    ' Depth windows re-place every period row. A Midnight row can already hold
    ' the whole by-day total while the slide sits in a later End of Run window.
    Dim depthOn As Boolean
    depthOn = HasSlideIntervals()
    ' Rows FieldCap no longer covers go back to depth and bit hours before any
    ' closing row subtracts the 6am / 4pm rows above it.
    For r = ROW_FIRST To ROW_LAST
        If Not DaySlideForRow(ws, bha, r, day) Then ClearFieldCapDay ws, r
    Next r
    Application.EnableEvents = prevEvents
    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail

    For r = ROW_FIRST To ROW_LAST
        period = Trim$(CStr(ws.Cells(r, COL_PERIOD).value & ""))
        If IsValidPeriod(period) Then
            hasDay = DaySlideForRow(ws, bha, r, day)
            ' A day the OpenCap by-day CSV covers is always re-posted for THAT
            ' row (even if it already holds a wrong value). Only rows with no
            ' by-day data fall back to the blank / active-row delta rules.
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
    Application.EnableEvents = prevEvents
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

    bitDelta = RoundHalfUp(bitTotal - gDayRoll_PriorBit, 4)
    circDelta = RoundHalfUp(circTotal - gDayRoll_PriorCirc, 4)

    If Not ValidateHourDeltas(bitDelta, circDelta, gDayRoll_SlideHrs, _
                             gDayRoll_Period, errMsg) Then
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo RestoreEvents

    ws.Cells(r, COL_END).value = endDepth
    ws.Cells(r, COL_BIT).value = bitDelta
    ws.Cells(r, COL_CIRC).value = circDelta
    ' No bit hours today -> no sliding either. Fill blank L/M/N with zeros so
    ' the row's ROP / % formulas resolve instead of sitting empty.
    If bitDelta = 0 Then
        If CellBlank(ws.Cells(r, COL_SLIDE_M)) Then ws.Cells(r, COL_SLIDE_M).value = 0
        If CellBlank(ws.Cells(r, COL_HRS)) Then ws.Cells(r, COL_HRS).value = 0
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
    Dim day As FcDay
    Dim hrsPart As Long
    Dim minPart As Long
    Dim prevEvents As Boolean
    Dim wrote As Boolean

    FillSlideMetrics = False
    wrote = False

    bha = ResolveActiveBha(ws)
    If bha = 0 Then Exit Function
    srcRow = FindSourceRow(ws, bha)
    Dim hasDay As Boolean
    hasDay = DaySlideForRow(ws, bha, r, day)

    ' FieldCap's own BHA totals. The mirror row of the active BHA is rebuilt
    ' from these day rows, so reading it here would feed the table back to itself.
    ' A FieldCap day needs neither, so an empty DD Tools mirror cannot stop it.
    If Not FieldCapBhaTotals(bha, mtrsSld, hrsSld) Then
        If srcRow = 0 And Not hasDay Then Exit Function
        If srcRow > 0 Then
            mtrsSld = GetNum(ws.Cells(srcRow, SRC_MTRS_SLD))
            hrsSld = GetNum(ws.Cells(srcRow, SRC_HRS_SLD))
        End If
    End If
    priorM = SumNumeric(ws, COL_SLIDE_M, ROW_FIRST, r - 1)
    ' Prefer H:M (min/60) so backfill works even if col V formulas are stale
    ' under Calculation=Manual during Refresh.
    priorH = SumSlideHoursDecimal(ws, ROW_FIRST, r - 1)
    deltaM = RoundHalfUp(mtrsSld - priorM, 2)
    deltaH = RoundHalfUp(hrsSld - priorH, 4)

    ' FieldCap's day is the source of truth for slide metres and slide time
    ' (L, M:N). Total / rotated metres stay End-Start and End-Start-slide (K, O).
    ' Only rows FieldCap does not cover fall back to the Slidesheet intervals.
    If hasDay Then
        deltaM = day.slideM
        deltaH = day.slideH
        overwriteSlide = True
    Else
        Dim depthM As Double, depthH As Double
        If DepthSlideForRow(ws, bha, r, mtrsSld, hrsSld, depthM, depthH) Then
            deltaM = depthM
            deltaH = depthH
            overwriteSlide = True
        End If
    End If

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo RestoreSlide

    ' Start Depth from previous End Depth
    If r > ROW_FIRST Then
        If CellBlank(ws.Cells(r, COL_START)) Then
            If Not CellBlank(ws.Cells(r - 1, COL_END)) Then
                If IsNumeric(ws.Cells(r - 1, COL_END).value) Then
                    ws.Cells(r, COL_START).value = ws.Cells(r - 1, COL_END).value
                    wrote = True
                End If
            End If
        End If
    End If

    If deltaM >= 0 And (overwriteSlide Or CellBlank(ws.Cells(r, COL_SLIDE_M))) Then
        ws.Cells(r, COL_SLIDE_M).value = deltaM
        wrote = True
    End If

    ' Without a FieldCap day, a source with no slide hours (stale BHA mirror)
    ' must not replace hours the DD typed with 0:00.
    Dim hrsBlank As Boolean
    hrsBlank = CellBlank(ws.Cells(r, COL_HRS)) And CellBlank(ws.Cells(r, COL_MIN))
    If deltaH >= 0 And (hasDay Or deltaH > 0 Or hrsBlank) And (overwriteSlide Or hrsBlank) Then
        SplitDecimalHours deltaH, hrsPart, minPart
        ws.Cells(r, COL_HRS).value = hrsPart
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

    If IsNumeric(ws.Cells(ROW_BHA, COL_PERIOD).value) Then
        gDayRoll_BHA = CLng(ws.Cells(ROW_BHA, COL_PERIOD).value)
    Else
        gDayRoll_BHA = 0
    End If

    If IsNumeric(ws.Cells(r, COL_START).value) Then
        gDayRoll_StartDepth = CDbl(ws.Cells(r, COL_START).value)
        gDayRoll_HasStartDepth = True
    Else
        gDayRoll_StartDepth = 0
        gDayRoll_HasStartDepth = False
    End If

    slideDec = 0
    If IsNumeric(ws.Cells(r, COL_SLIDE_DEC).value) Then
        slideDec = CDbl(ws.Cells(r, COL_SLIDE_DEC).value)
    ElseIf IsNumeric(ws.Cells(r, COL_HRS).value) Or IsNumeric(ws.Cells(r, COL_MIN).value) Then
        slideDec = GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
    End If
    gDayRoll_SlideHrs = RoundHalfUp(slideDec, 4)
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
    If IsNumeric(ws.Cells(ROW_BHA, COL_PERIOD).value) Then
        headerBha = CLng(ws.Cells(ROW_BHA, COL_PERIOD).value)
        If FindSourceRow(ws, headerBha) > 0 Then
            ResolveActiveBha = headerBha
            Exit Function
        End If
    End If
    If lastBha > 0 Then
        ws.Cells(ROW_BHA, COL_PERIOD).value = lastBha
        ResolveActiveBha = lastBha
    ElseIf headerBha > 0 Then
        ' Empty mirror (DD Tools not built): keep the BHA the DD picked.
        ResolveActiveBha = headerBha
    End If
End Function

Private Function LastSourceBha(ByVal ws As Worksheet) As Long
    Dim r As Long
    Dim v As Variant
    LastSourceBha = 0
    For r = SRC_FIRST To SRC_LAST
        v = ws.Cells(r, SRC_BHA).value
        If IsNumeric(v) And v <> "" Then LastSourceBha = CLng(v)
    Next r
End Function

Private Function FindSourceRow(ByVal ws As Worksheet, ByVal bha As Long) As Long
    Dim r As Long
    Dim v As Variant
    FindSourceRow = 0
    For r = SRC_FIRST To SRC_LAST
        v = ws.Cells(r, SRC_BHA).value
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

' FieldCap's report day that daily-table row r closes, from the OpenCap
' "slide-rotate by day" CSV (hidden _OC_Slide) for BHA bha. Days are matched in
' date order to the Midnight / End of Run rows of the table (the table has no
' date column). 6am / 4pm rows inside a day are subtracted so the closing row
' carries only the remainder. False = no by-day data for this row.
' FieldCap is the source of truth for slide metres/time and rotate metres, even
' past Pason's bit hours (a slide block includes connections and surveys).
' Rotate clock hours still land in AL; report Time Rotating is P-V, not AL.
Private Function DaySlideForRow(ByVal ws As Worksheet, ByVal bha As Long, _
                                ByVal r As Long, ByRef out As FcDay) As Boolean
    Dim wsS As Worksheet
    Dim lastR As Long, i As Long, j As Long, n As Long, pos As Long
    Dim keys() As Double, mt() As Double, hr() As Double, rm() As Double, ro() As Double
    Dim v As Variant
    Dim k As Double, m As Double, h As Double, rmv As Double, rh As Double
    Dim period As String
    Dim dayIdx As Long, rr As Long, firstOfDay As Long
    Dim blankDay As FcDay

    DaySlideForRow = False
    out = blankDay
    If bha <= 0 Or r < ROW_FIRST Or r > ROW_LAST Then Exit Function
    period = Trim$(CStr(ws.Cells(r, COL_PERIOD).value & ""))
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
    ReDim rm(1 To lastR)
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
                rmv = GetNum(wsS.Cells(i, 5))
                h = GetNum(wsS.Cells(i, 6))
                rh = GetNum(wsS.Cells(i, 7))
                pos = 0
                For j = 1 To n
                    If keys(j) = k Then pos = j: Exit For
                Next j
                If pos > 0 Then
                    mt(pos) = mt(pos) + m
                    rm(pos) = rm(pos) + rmv
                    hr(pos) = hr(pos) + h
                    ro(pos) = ro(pos) + rh
                Else
                    pos = n + 1
                    For j = 1 To n
                        If keys(j) > k Then pos = j: Exit For
                    Next j
                    For j = n To pos Step -1
                        keys(j + 1) = keys(j): mt(j + 1) = mt(j): rm(j + 1) = rm(j)
                        hr(j + 1) = hr(j): ro(j + 1) = ro(j)
                    Next j
                    keys(pos) = k: mt(pos) = m: rm(pos) = rmv: hr(pos) = h: ro(pos) = rh
                    n = n + 1
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    FillMissingDays keys, mt, hr, rm, ro, n

    ' Which report day does row r close?
    dayIdx = 0
    firstOfDay = ROW_FIRST
    For rr = ROW_FIRST To r
        If IsClosingPeriod(Trim$(CStr(ws.Cells(rr, COL_PERIOD).value & ""))) Then
            dayIdx = dayIdx + 1
            If rr < r Then firstOfDay = rr + 1
        End If
    Next rr
    If dayIdx < 1 Or dayIdx > n Then Exit Function

    out.slideM = RemainderOf(mt(dayIdx) - SumNumeric(ws, COL_SLIDE_M, firstOfDay, r - 1), 2)
    out.slideH = RemainderOf(hr(dayIdx) - SumSlideHoursDecimal(ws, firstOfDay, r - 1), 4)
    out.rotM = RemainderOf(rm(dayIdx), 2)
    out.rotH = RemainderOf(ro(dayIdx), 4)
    out.totalM = RemainderOf(mt(dayIdx) + rm(dayIdx), 2)
    DaySlideForRow = True
End Function

Private Function RemainderOf(ByVal x As Double, ByVal digits As Long) As Double
    If x < 0 Then x = 0
    RemainderOf = RoundHalfUp(x, digits)
End Function

' VBA Round sends a half to the even digit (0.125 -> 0.12). Reports round halves up.
Private Function RoundHalfUp(ByVal x As Double, ByVal digits As Long) As Double
    RoundHalfUp = Application.WorksheetFunction.Round(x, digits)
End Function

' FieldCap writes a zero row for a day the BHA was in the hole without drilling.
' An older by-day file skips that day, so the dates between the BHA's first and
' last FieldCap day are filled with zero days here. Otherwise every later day
' would close the table row above its own.
Private Sub FillMissingDays(ByRef keys() As Double, ByRef mt() As Double, ByRef hr() As Double, _
                            ByRef rm() As Double, ByRef ro() As Double, ByRef n As Long)
    Dim first As Long, span As Long, i As Long, idx As Long
    Dim fk() As Double, fm() As Double, fh() As Double, frm() As Double, fro() As Double
    If n < 2 Then Exit Sub
    If keys(1) < DATE_SERIAL_MIN Or keys(n) > DATE_SERIAL_MAX Then Exit Sub
    first = CLng(Int(keys(1)))
    span = CLng(Int(keys(n))) - first + 1
    If span <= n Or span > MAX_FILL_DAYS Then Exit Sub
    ReDim fk(1 To span)
    ReDim fm(1 To span)
    ReDim fh(1 To span)
    ReDim frm(1 To span)
    ReDim fro(1 To span)
    For i = 1 To span
        fk(i) = first + i - 1
    Next i
    For i = 1 To n
        idx = CLng(Int(keys(i))) - first + 1
        fm(idx) = fm(idx) + mt(i)
        fh(idx) = fh(idx) + hr(i)
        frm(idx) = frm(idx) + rm(i)
        fro(idx) = fro(idx) + ro(i)
    Next i
    keys = fk: mt = fm: hr = fh: rm = frm: ro = fro
    n = span
End Sub

' Wipe retired AJ:AM sidecar on one day row (leftover FieldCap Total / Rot Hrs).
Private Sub ClearFieldCapDay(ByVal ws As Worksheet, ByVal r As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, COL_HELPER_FIRST), ws.Cells(r, COL_HELPER_LAST))
    If Application.WorksheetFunction.CountA(rng) > 0 Then rng.ClearContents
End Sub

Private Sub ClearHelperBand(ByVal ws As Worksheet)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(2, COL_HELPER_FIRST), ws.Cells(ROW_TOTALS, COL_HELPER_LAST))
    If Application.WorksheetFunction.CountA(rng) > 0 Then rng.ClearContents
End Sub

' Daily-table formulas: K = J-I, O = J-I-L, S = O/(P-V). Email C9/C10/C16/C17
' read V/P/S/T/U on rows 23/24. AJ:AM sidecar is cleared, never rewritten.
Private Sub EnsureFieldCapFormulas(ByVal ws As Worksheet)
    Dim r As Long
    Dim a As String
    ClearHelperBand ws
    For r = ROW_FIRST To ROW_LAST
        a = CStr(r)
        SetFormulaIfDiff ws.Cells(r, COL_TOTAL), _
            "=IFERROR(IF(ISBLANK(J" & a & "),"""",J" & a & "-I" & a & "),"""")"
        SetFormulaIfDiff ws.Cells(r, COL_ROTATED), _
            "=IFERROR(IF(ISBLANK(J" & a & "),"""",J" & a & "-I" & a & "-L" & a & "),"""")"
        SetFormulaIfDiff ws.Cells(r, 19), _
            "=IFERROR(IF(ISBLANK(P" & a & "),"""",ROUND(O" & a & "/(P" & a & "-V" & a & "),2)),"""")"
    Next r
    SetFormulaIfDiff ws.Cells(ROW_TOTALS, 19), "=IFERROR(ROUND(O24/(P24-V24),2),"""")"
    ' BHA slide time from V24: SUM(N) skips minutes typed as text and never carries past 60.
    SetFormulaIfDiff ws.Cells(ROW_TOTALS, COL_HRS), "=IFERROR(INT(ROUND(V24*60,0)/60),"""")"
    SetFormulaIfDiff ws.Cells(ROW_TOTALS, COL_MIN), "=IFERROR(MOD(ROUND(V24*60,0),60),"""")"
    SetFormulaIfDiff ws.Range("C9"), _
        "=IFERROR(IF(IFERROR(N(V23),0)=0,""0Hrs"",ROUND(V23,2)&""Hrs / ""&ROUND(T23,4)*100&""% / ""&ROUND(R23,2)&""m/hr""),"""")"
    SetFormulaIfDiff ws.Range("C10"), _
        "=IFERROR(IF(U23=0,""0Hrs"",ROUND(P23-V23,2)&""Hrs / ""&ROUND(U23,4)*100&""% / ""&ROUND(S23,2)&""m/hr""),"""")"
    SetFormulaIfDiff ws.Range("C16"), _
        "=IFERROR(IF(IFERROR(N(V24),0)=0,""0Hrs"",ROUND(V24,2)&""Hrs / ""&ROUND(T24,4)*100&""% / ""&ROUND(R24,2)&""m/hr""),"""")"
    SetFormulaIfDiff ws.Range("C17"), _
        "=IFERROR(IF(U24=0,""0Hrs"",ROUND(P24-V24,2)&""Hrs / ""&ROUND(U24,4)*100&""% / ""&ROUND(S24,2)&""m/hr""),"""")"
    ' Slide metres print whole (ROUND halves up). Rotated is total less slide (O).
    SetFormulaIfDiff ws.Range("C7"), "=IFERROR(ROUND(IF(ISERROR(L23),0,L23),0),"""")"
    SetFormulaIfDiff ws.Range("C8"), "=IFERROR(ROUND(O23,2),"""")"
    SetFormulaIfDiff ws.Range("C14"), "=IFERROR(ROUND(L24,0),"""")"
    SetFormulaIfDiff ws.Range("C15"), "=IFERROR(ROUND(O24,2),"""")"
    SetFormulaIfDiff ws.Range("C18"), _
        "=IFERROR(IF(U24=0,""0Hrs"",ROUND(K24/P24,2)&""m/hr""),"""")"
    SetFormulaIfDiff ws.Range("C19"), "=IFERROR(P24,"""")"
    SetFormulaIfDiff ws.Range("C20"), "=IFERROR(Q24,"""")"
    SetFormatIfDiff ws.Range("C7"), "0""m"""
    SetFormatIfDiff ws.Range("C14"), "0""m"""
    SetFormatIfDiff ws.Range(ws.Cells(ROW_FIRST, COL_SLIDE_M), ws.Cells(ROW_TOTALS, COL_SLIDE_M)), "0"
End Sub

' Formula-only: K/O/S/C9-C17 from the daily table; wipe AJ:AM. No L/M/N or Pason writes.
Public Sub DayRoll_EnsureReportHourFormulas()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    wasProt = SheetUnprotectForVba(ws)
    Application.EnableEvents = False
    EnsureFieldCapFormulas ws
    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    Application.EnableEvents = prevEvents
    SheetReprotectAfterVba ws, wasProt
End Sub

Private Sub SetFormulaIfDiff(ByVal cell As Range, ByVal f As String)
    If CStr(cell.formula) <> f Then cell.formula = f
End Sub

Private Sub SetFormatIfDiff(ByVal rng As Range, ByVal fmt As String)
    If IsNull(rng.numberFormat) Then
        rng.numberFormat = fmt
    ElseIf rng.numberFormat <> fmt Then
        rng.numberFormat = fmt
    End If
End Sub

' True when Slidesheet R/S holds at least one real slide interval.
Private Function HasSlideIntervals() As Boolean
    Dim ss As Worksheet
    Dim r As Long
    Dim vA As Variant, vb As Variant
    HasSlideIntervals = False
    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SH_SS)
    On Error GoTo 0
    If ss Is Nothing Then Exit Function
    For r = SS_FIRST To SS_LAST
        vA = ss.Cells(r, COL_SLIDE_FROM).Value2
        vb = ss.Cells(r, COL_SLIDE_TO).Value2
        If IsNumeric(vA) And IsNumeric(vb) Then
            If CDbl(vb) > CDbl(vA) Then
                HasSlideIntervals = True
                Exit Function
            End If
        End If
    Next r
End Function

' FieldCap's slide metres / hours for one BHA, from the OpenCap BHA CSV.
Private Function FieldCapBhaTotals(ByVal bha As Long, ByRef outM As Double, _
                                   ByRef outH As Double) As Boolean
    Dim names As Variant, nm As Variant
    Dim wsB As Worksheet
    Dim cB As Long, cM As Long, ch As Long, c As Long, i As Long, lastR As Long
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
            cB = 0: cM = 0: ch = 0
            For c = 1 To 80
                hdr = Trim$(CStr(wsB.Cells(1, c).value & ""))
                If hdr = "BHA #" Then cB = c
                If hdr = "BHA Mtrs Slid" Then cM = c
                If hdr = "BHA Hrs Slid" Then ch = c
            Next c
            If cB > 0 And (cM > 0 Or ch > 0) Then
                lastR = wsB.Cells(wsB.rows.Count, cB).End(xlUp).row
                For i = 2 To lastR
                    v = wsB.Cells(i, cB).Value2
                    If IsNumeric(v) And Not IsEmpty(v) Then
                        If CLng(v) = bha Then
                            If cM > 0 Then outM = GetNum(wsB.Cells(i, cM))
                            If ch > 0 Then outH = GetNum(wsB.Cells(i, ch))
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
' Metres are column K (End-Start). Rotate hours are always bit hours less
' slide hours, matching Time Rotating on Data. False = no day rows.
Public Function DayRoll_ActiveBhaTotals(ByRef bha As Long, ByRef metres As Double, _
                                        ByRef slideM As Double, ByRef slideH As Double, _
                                        ByRef rotH As Double, ByRef circH As Double) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim rowSlideH As Double
    Dim lo As Double, hi As Double
    DayRoll_ActiveBhaTotals = False
    bha = 0: metres = 0#: slideM = 0#: slideH = 0#: rotH = 0#: circH = 0#
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    If Not IsNumeric(ws.Cells(ROW_BHA, COL_PERIOD).value) Then Exit Function
    If CellBlank(ws.Cells(ROW_BHA, COL_PERIOD)) Then Exit Function
    bha = CLng(ws.Cells(ROW_BHA, COL_PERIOD).value)
    For r = ROW_FIRST To ROW_LAST
        If IsValidPeriod(Trim$(CStr(ws.Cells(r, COL_PERIOD).value & ""))) Then
            If IsNumeric(ws.Cells(r, COL_TOTAL).value) And Not CellBlank(ws.Cells(r, COL_TOTAL)) Then
                metres = metres + CDbl(ws.Cells(r, COL_TOTAL).value)
            ElseIf RowDepthWindow(ws, r, lo, hi) Then
                metres = metres + (hi - lo)
            End If
            rowSlideH = GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
            slideM = slideM + GetNum(ws.Cells(r, COL_SLIDE_M))
            slideH = slideH + rowSlideH
            If Not CellBlank(ws.Cells(r, COL_BIT)) Then
                rotH = rotH + GetNum(ws.Cells(r, COL_BIT)) - rowSlideH
            End If
            circH = circH + GetNum(ws.Cells(r, COL_CIRC))
            DayRoll_ActiveBhaTotals = True
        End If
    Next r
    If rotH < 0# Then rotH = 0#
End Function

' After REFRESH: day rows that do not match FieldCap's day, and rows FieldCap
' does not cover whose slide time is more than their bit time (rotating time
' there is bit hours less slide hours and would print negative). "" = none.
Public Function DayRoll_SlideHoursWarning() As String
    Dim ws As Worksheet
    Dim r As Long
    Dim bha As Long
    Dim slideH As Double, bit As Double
    Dim day As FcDay
    Dim off As String
    Dim msgFc As String, msgBit As String
    DayRoll_SlideHoursWarning = ""
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    If IsNumeric(ws.Cells(ROW_BHA, COL_PERIOD).value) And Not CellBlank(ws.Cells(ROW_BHA, COL_PERIOD)) Then
        bha = CLng(ws.Cells(ROW_BHA, COL_PERIOD).value)
    End If
    For r = ROW_FIRST To ROW_LAST
        If IsValidPeriod(Trim$(CStr(ws.Cells(r, COL_PERIOD).value & ""))) Then
            slideH = GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
            If DaySlideForRow(ws, bha, r, day) Then
                off = ""
                If Abs(GetNum(ws.Cells(r, COL_SLIDE_M)) - day.slideM) > 0.005 Then off = off & " slide m"
                If Abs(slideH - day.slideH) > 0.01 Then off = off & " slide time"
                If off <> "" Then
                    msgFc = msgFc & vbLf & "  Row " & r & " (" & ws.Cells(r, COL_PERIOD).value & "):" & off & _
                            "  (FieldCap S " & Format$(day.slideM, "0.00") & " m / " & Format$(day.slideH, "0.00") & " h)"
                End If
            ElseIf IsNumeric(ws.Cells(r, COL_BIT).value) And Not CellBlank(ws.Cells(r, COL_BIT)) Then
                bit = CDbl(ws.Cells(r, COL_BIT).value)
                If slideH > bit + 0.01 Then
                    msgBit = msgBit & vbLf & "  Row " & r & " (" & ws.Cells(r, COL_PERIOD).value & "): " & _
                             Format$(slideH, "0.00") & " h sliding, " & Format$(bit, "0.00") & " bit hours"
                End If
            End If
        End If
    Next r
    If msgFc <> "" Then
        DayRoll_SlideHoursWarning = "These day rows do not match FieldCap's day:" & msgFc & vbLf & vbLf & _
            "REFRESH writes FieldCap's numbers. If a row still differs, its cells are locked or hold text."
    End If
    If msgBit <> "" Then
        If DayRoll_SlideHoursWarning <> "" Then DayRoll_SlideHoursWarning = DayRoll_SlideHoursWarning & vbLf & vbLf
        DayRoll_SlideHoursWarning = DayRoll_SlideHoursWarning & _
            "FieldCap has no day for these rows, and Time Sliding is more than the bit hours:" & msgBit & vbLf & vbLf & _
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
    outM = RoundHalfUp(CDbl(vM), 2)
    If Not BhaByDayPool(bha, poolM, poolH) Then
        poolM = mirrorM
        poolH = mirrorH
    End If
    If poolM > 0.0001 And outM > 0# And poolH > 0# Then
        outH = RoundHalfUp(poolH * (outM / poolM), 4)
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
        If IsValidPeriod(Trim$(CStr(ws.Cells(r, COL_PERIOD).value & ""))) Then
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
    totalMin = CLng(RoundHalfUp(decHrs * 60#, 0))
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
        cell.value = Format$(minutes, "00")
    Else
        cell.value = minutes
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
        If IsNumeric(ws.Cells(r, col).value) Then
            tot = tot + CDbl(ws.Cells(r, col).value)
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
        If IsNumeric(ws.Cells(r, COL_SLIDE_DEC).value) Then
            dec = CDbl(ws.Cells(r, COL_SLIDE_DEC).value)
        ElseIf IsNumeric(ws.Cells(r, COL_HRS).value) _
            Or IsNumeric(ws.Cells(r, COL_MIN).value) Then
            dec = GetNum(ws.Cells(r, COL_HRS)) + GetNum(ws.Cells(r, COL_MIN)) / 60#
        End If
        tot = tot + dec
    Next r
    SumSlideHoursDecimal = tot
End Function

Private Function GetNum(ByVal cell As Range) As Double
    If IsNumeric(cell.value) Then
        GetNum = CDbl(cell.value)
    Else
        GetNum = 0
    End If
End Function

Private Function CellBlank(ByVal cell As Range) As Boolean
    Dim v As Variant
    v = cell.value
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

























