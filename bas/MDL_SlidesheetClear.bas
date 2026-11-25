Attribute VB_Name = "MDL_SlidesheetClear"
Option Explicit

Private Const SS_SHEET As String = "Slidesheet"
Private Const BTN_NAME As String = "btnClearSlidesheetRanges"
Private Const BTN_CAPTION As String = "Clear Ranges"
' Merged Z1:Z3 - left of the plan proximity gauge which owns merged AA1:AC7
Private Const BTN_ANCHOR As String = "Z1:Z3"
Private Const Y_FIRST As Long = 12
Private Const Y_LAST As Long = 505
Private Const Y_DATA_FIRST As Long = 13
Private Const Y_LABEL As String = "RKB"

' survey-row formula repair state (see RepairSurveyRowFormulas)
Private m_repairDeclined As Boolean
Private m_repairBusy As Boolean
' Sheet comment yellow (same as Module16 pale yellow). Not Excel idx-6 65535.
Private Const COMMENT_YELLOW As Long = 13434879  ' RGB(255,255,204)
Private Const TGT_FIRST As Long = 2
Private Const TGT_LAST As Long = 5
Private Const PLANSEC_SHEET As String = "_OC_PlanSec"
Private Const TGT_NAMES As String = "NUDGE,VERTICAL,KOP,TANGENT,SOT,EOT,ICP,BUILD,HEEL,TD"
' Full named-target list for BURR / slide UDFs (T2:Y5 is display-only).
Private Const FULL_TGT_MD_COL As Long = 14
Private Const FULL_TGT_FIRST As Long = 3
Private Const FULL_TGT_ROWS As Long = 80
' User edits (the Targets form, or a typed T2:Y5 number). Blank = use the plan.
' A new plan import clears the sheet, so these do not outlive the plan.
Private Const OVR_MD As Long = 19
Private Const OVR_INC As Long = 20
Private Const OVR_AZM As Long = 21
Private Const OVR_TVD As Long = 22
Private Const BTN_TGT As String = "btnEditTargets"
Private m_teN As Long
Private m_teRow() As Long
Private m_teNm() As String
Private m_teMd() As Double
Private m_teInc() As Double
Private m_teAzm() As Double
Private m_teTvd() As Double
Private m_updatingY As Boolean
Private m_hlTgtRow As Long   ' last highlighted target row (0 = none)
Private m_newestRow As Long  ' row the dark newest-survey bar was last painted on (0 = none)
Private m_syncingTgt As Boolean

' -- Fast text measurement (RefreshSlideComments) -------------------------------
' The old MeasureTextWidthPts created and deleted a textbox shape per string
' (~25 ms each). One refresh measures ~600 strings (~16 s), and a survey entry
' fires the refresh from both Worksheet_Change and Worksheet_Calculate, which
' is how one Inc/Azm keystroke cost about two minutes in the field.
' Fix: one reusable hidden textbox per pass plus a width cache, and a short
' debounce so the Change/Calculate pair of a single edit refreshes once.
Private m_measureShape As Shape     ' live only between Begin/EndTextMeasure
Private m_measureCache As Object    ' Scripting.Dictionary: text|font|size|bold -> pts
Private m_lastRefreshAt As Double   ' Timer stamp of last completed auto refresh

' Clear Slidesheet survey/entry ranges:
'   D,F,G,T,U rows 13:505
'   Sail waypoints AC14:AD33 (display AA=Inc Next, AB=Geo Window; helper AE)
'   Targets U2:X5
'   Comments Y12:Y505 (then restore RKB + auto slide comments)
Public Sub ClearSlidesheetRanges()
    Dim ss As Worksheet

    On Error GoTo Fail

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ss.Range("D13:D" & Y_LAST).ClearContents
    ss.Range("F13:F" & Y_LAST).ClearContents
    ss.Range("G13:G" & Y_LAST).ClearContents
    ss.Range("T13:T" & Y_LAST).ClearContents
    ss.Range("U13:U" & Y_LAST).ClearContents
    ss.Range("AC14:AD33").ClearContents
    ss.Range("U2:X5").ClearContents
    ss.Range("Y12:Y" & Y_LAST).ClearContents
    ss.Range("Z13:Z" & Y_LAST).ClearContents
    SyncPlanTargetWindowOnSheet ss
    ss.Range("Y12").Value2 = Y_LABEL

    Application.EnableEvents = True
    RefreshSlideComments forceAll:=True
    HighlightActiveTarget

    Application.ScreenUpdating = True

    MsgBox "Cleared:" & vbCrLf & _
           "  D/F/G/T/U 13:" & Y_LAST & vbCrLf & _
           "  Sail waypoints AC14:AD33" & vbCrLf & _
           "  Targets U2:X5" & vbCrLf & _
           "  Comments Y12:Y" & Y_LAST & " / ROT Z13:Z" & Y_LAST & " (auto text restored where inputs exist)", _
           vbInformation
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    SheetEnsureLocked ss
    On Error GoTo 0
    MsgBox "ClearSlidesheetRanges failed: " & Err.Description, vbCritical
End Sub

' True when a Slidesheet edit should rewrite Y. Other sheets never qualify.
' Y only exists when AR (BURR) or AO (active target MD) is numeric — most
' surveys (vertical / hold / past TD) never get a comment.
' Target-window edits (U2:X5) always refresh: they change BURR for aiming rows.
' Existing auto Y text also qualifies so a row that just left the window clears.
Public Function SlideCommentsNeedRefresh(ByVal Target As Range) As Boolean
    Dim ss As Worksheet
    Dim area As Range
    Dim cell As Range
    Dim r As Long
    Dim seen As Object
    Dim yTxt As String

    SlideCommentsNeedRefresh = False
    If Target Is Nothing Then Exit Function
    On Error GoTo Fail
    Set ss = Target.Worksheet
    If StrComp(ss.name, SS_SHEET, vbTextCompare) <> 0 Then Exit Function

    If Not Intersect(Target, ss.Range("U2:X5")) Is Nothing Then
        SlideCommentsNeedRefresh = True
        Exit Function
    End If

    Set area = Intersect(Target, ss.Range("D13:G" & Y_LAST & ",T13:U" & Y_LAST))
    If area Is Nothing Then Exit Function

    Set seen = CreateObject("Scripting.Dictionary")
    For Each cell In area.Cells
        r = cell.row
        If Not seen.Exists(r) Then
            seen.Add r, True
            If RowHasSlideAim(ss, r) Then
                SlideCommentsNeedRefresh = True
                Exit Function
            End If
            yTxt = Trim$(CStr(ss.Cells(r, "Y").Value2 & ""))
            If Len(yTxt) > 0 Then
                If IsAutoSlideCommentText(yTxt) Then
                    SlideCommentsNeedRefresh = True
                    Exit Function
                End If
            End If
        End If
    Next cell
    Exit Function
Fail:
    SlideCommentsNeedRefresh = False
End Function

Private Function RowHasSlideAim(ByVal ss As Worksheet, ByVal r As Long) As Boolean
    RowHasSlideAim = IsNumberValue(ss.Cells(r, "AO").Value2) _
                  Or IsNumberValue(ss.Cells(r, "AR").Value2)
End Function

' Rewrite Z on every stand row (next-row C − this-row T). Used after the
' formula change so leftover-rotate values catch up without a survey keystroke.
Public Sub RefreshAllRotateMetres()
    Dim ss As Worksheet
    Dim r As Long
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    On Error GoTo Clean
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    For r = Y_DATA_FIRST To Y_LAST
        WriteRotateMetresOnSheet ss, r
    Next r
Clean:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

' Z[r] = C[r+1] − T[r]. C/D/T edits call this without a full Y rebuild.
Public Sub RefreshRotateMetres(ByVal Target As Range)
    Dim ss As Worksheet
    Dim area As Range
    Dim cell As Range
    Dim r As Long
    Dim seen As Object
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    If Target Is Nothing Then Exit Sub
    On Error GoTo Clean
    Set ss = Target.Worksheet
    If StrComp(ss.name, SS_SHEET, vbTextCompare) <> 0 Then Exit Sub
    Set area = Intersect(Target, ss.Range("C13:D" & Y_LAST & ",T13:T" & Y_LAST))
    If area Is Nothing Then Exit Sub

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    Set seen = CreateObject("Scripting.Dictionary")
    For Each cell In area.Cells
        r = cell.row
        If r >= Y_DATA_FIRST And r <= Y_LAST Then
            ' T on this row feeds leftover Z on this row.
            If cell.Column = 20 Then
                If Not seen.Exists(r) Then
                    seen.Add r, True
                    WriteRotateMetresOnSheet ss, r
                End If
            End If
            ' C on this row is the course for leftover Z on the row above.
            If cell.Column = 3 Then
                If r - 1 >= Y_DATA_FIRST Then
                    If Not seen.Exists(r - 1) Then
                        seen.Add r - 1, True
                        WriteRotateMetresOnSheet ss, r - 1
                    End If
                End If
            End If
            ' D (MD) can change C on this row and C on the next row.
            If cell.Column = 4 Then
                If r - 1 >= Y_DATA_FIRST Then
                    If Not seen.Exists(r - 1) Then
                        seen.Add r - 1, True
                        WriteRotateMetresOnSheet ss, r - 1
                    End If
                End If
                If Not seen.Exists(r) Then
                    seen.Add r, True
                    WriteRotateMetresOnSheet ss, r
                End If
            End If
        End If
    Next cell

Clean:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

' Leftover rotate on this row: next stand's course (C one row down) minus
' already-slid (T on this row). Blank T = 0. Leaves a hand-typed Z alone.
' Clamps at 0 when T > next C.
' sheetFresh:=True when the caller just ran ss.Calculate. A single-cell
' Calculate on C drags the whole UDF chain behind it (~3.8 s per row, measured),
' so skip it there.
Private Sub WriteRotateMetresOnSheet(ByVal ss As Worksheet, ByVal r As Long, _
                                     Optional ByVal sheetFresh As Boolean = False)
    Dim course As Double
    Dim slid As Double
    Dim remain As Double
    Dim rotTxt As String
    Dim zCell As Range
    Dim hasC As Boolean
    Dim courseRow As Long

    Set zCell = ss.Cells(r, "Z")
    If Not IsAutoRotText(Trim$(CStr(zCell.Value2 & ""))) Then Exit Sub

    courseRow = r + 1
    On Error Resume Next
    If courseRow <= Y_LAST And Not sheetFresh Then ss.Cells(courseRow, "C").Calculate
    On Error GoTo 0

    hasC = (courseRow <= Y_LAST) And IsNumberValue(ss.Cells(courseRow, "C").Value2)
    If Not hasC Then
        If Len(Trim$(CStr(zCell.Value2 & ""))) > 0 Then zCell.ClearContents
        Exit Sub
    End If

    course = CDbl(ss.Cells(courseRow, "C").Value2)
    slid = 0#
    If IsNumberValue(ss.Cells(r, "T").Value2) Then slid = CDbl(ss.Cells(r, "T").Value2)

    remain = Application.WorksheetFunction.Round(course, 2) - _
             Application.WorksheetFunction.Round(slid, 2)
    If remain < 0# Then remain = 0#
    rotTxt = Format$(remain, "0.00")
    If CStr(zCell.Value2 & "") <> rotTxt Then
        With zCell
            .Value2 = rotTxt
            .WrapText = False
            .numberFormat = "0.00"
            .HorizontalAlignment = xlCenter
        End With
    End If
End Sub

' Write ProjSlideComment results into Y as values.
' forceAll:=True overwrites every data row (used on clear / initial convert).
' forceAll:=False only updates blank cells or prior auto "Sliding ... | BURR ..." notes.
Public Sub RefreshSlideComments(Optional ByVal forceAll As Boolean = False)
    Dim ss As Worksheet
    Dim r As Long
    Dim i As Long
    Dim autoTxt As String
    Dim curTxt As String
    Dim result As Variant
    Dim xArr As Variant
    Dim uArr As Variant
    Dim arArr As Variant
    Dim yArr As Variant
    Dim cArr As Variant
    Dim slideCap As Double
    Dim asVal As Variant
    Dim anyFormula As Boolean
    Dim writeIt As Boolean
    Dim wasProt As Boolean
    Dim leftPart As String
    Dim rightPart As String
    Dim p As Long
    Dim protoY As Range
    Dim spacePts As Double
    Dim prevEvents As Boolean
    Dim prevCalc As XlCalculation

    If m_updatingY Then Exit Sub
    ' Debounce leftover from when Change + Calculate both asked for a refresh.
    ' Calculate no longer calls this; keep the window so a Change that writes
    ' Y/Z cannot re-enter via a cascaded handler. forceAll always runs.
    If Not forceAll Then
        If m_lastRefreshAt > 0# And Abs(Timer - m_lastRefreshAt) < 0.5 Then Exit Sub
    End If
    ScreenBeginBusy
    m_updatingY = True

    ' Restore the caller's state on exit — automation runs with events off
    ' (Field Workbook_Open guard) and must stay that way after this sub.
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    On Error GoTo Clean

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    Application.EnableEvents = False

    ' Manual calc for the whole pass: under automatic calc every Y/Z write in
    ' the loop below triggers an incremental recalc — measured ~15-35 s per
    ' refresh on a full sheet vs ~0.4 s under manual. One sheet calc up front
    ' keeps the AS/AR/AT formula inputs fresh; Clean restores the caller's
    ' mode, which recalcs the written cells once.
    Application.Calculation = xlCalculationManual
    wasProt = SheetUnprotectForVba(ss)
    EnsureProjAimFormulasOnSheet ss
    EnsureLastSurveyAzmLookup ss
    ss.Calculate

    ' Keep / restore label row
    If Len(Trim$(CStr(ss.Cells(Y_FIRST, "Y").text))) = 0 Then
        ss.Cells(Y_FIRST, "Y").Value2 = Y_LABEL
    End If

    ' Meters To Slide in hidden AS; required TF (to target) in hidden AT.
    ' Visible V/W/X = TVD @ Bit / INC @ Bit / AZM @ Bit. Y uses AT (not user U).
    xArr = ss.Range(ss.Cells(Y_DATA_FIRST, "AS"), ss.Cells(Y_LAST, "AS")).Value2
    uArr = ss.Range(ss.Cells(Y_DATA_FIRST, "AT"), ss.Cells(Y_LAST, "AT")).Value2
    arArr = ss.Range(ss.Cells(Y_DATA_FIRST, "AR"), ss.Cells(Y_LAST, "AR")).Value2
    yArr = ss.Range(ss.Cells(Y_DATA_FIRST, "Y"), ss.Cells(Y_LAST, "Y")).Value2
    cArr = ss.Range(ss.Cells(Y_DATA_FIRST, "C"), ss.Cells(Y_LAST, "C")).Value2

    ' Only pay for per-cell HasFormula checks if the column still holds formulas.
    anyFormula = True
    If VarType(ss.Range(ss.Cells(Y_DATA_FIRST, "Y"), ss.Cells(Y_LAST, "Y")).HasFormula) = vbBoolean Then
        anyFormula = ss.Range(ss.Cells(Y_DATA_FIRST, "Y"), ss.Cells(Y_LAST, "Y")).HasFormula
    End If

    Set protoY = ss.Cells(Y_DATA_FIRST, "Y")
    BeginTextMeasure protoY          ' one reusable textbox for the whole pass
    spacePts = MeasureTextWidthPts(" ", protoY)
    If spacePts < 0.25 Then spacePts = 0.25

    For i = 1 To UBound(xArr, 1)
        r = Y_DATA_FIRST + i - 1

        autoTxt = ""
        slideCap = 0#
        If IsNumberValue(cArr(i, 1)) Then slideCap = CDbl(cArr(i, 1))

        ' Same gate as the survey math: Inc/Azm must be the survey green.
        ' Yellow projections and any other fill do not get a Y comment.
        If MDL_ContDI.IsSlidesheetGoodSurveyRowOn(ss, r) _
           And IsNumberValue(arArr(i, 1)) Then
            If IsNumberValue(xArr(i, 1)) Then
                asVal = xArr(i, 1)
            Else
                asVal = 0#
            End If
            result = ProjSlideComment(asVal, uArr(i, 1), arArr(i, 1), 0#, slideCap)
        Else
            result = ""
        End If

        If Not isError(result) Then
            autoTxt = CStr(result & "")
            p = InStr(1, autoTxt, Chr$(1), vbBinaryCompare)
            If p > 0 Then
                leftPart = Left$(autoTxt, p - 1)
                rightPart = mid$(autoTxt, p + 1)
                autoTxt = FitLeftRightInCell(leftPart, rightPart, ss.Cells(r, "Y"), spacePts)
            End If
        End If

        ' Z: leftover rotate = next-row C minus this-row T (already slid).
        WriteRotateMetresOnSheet ss, r, True

        curTxt = Trim$(CStr(yArr(i, 1) & ""))

        writeIt = forceAll
        If Not writeIt Then
            writeIt = IsAutoSlideCommentText(curTxt)
            If Not writeIt And anyFormula Then
                writeIt = ss.Cells(r, "Y").HasFormula
            End If
        End If

        If writeIt Then
            If curTxt <> autoTxt Or (anyFormula And ss.Cells(r, "Y").HasFormula) Then
                If Len(autoTxt) = 0 Then
                    ss.Cells(r, "Y").ClearContents
                Else
                    With ss.Cells(r, "Y")
                        .WrapText = False
                        .IndentLevel = 0
                        .HorizontalAlignment = xlLeft
                        .Interior.Pattern = xlSolid
                        .Interior.Color = COMMENT_YELLOW
                        .Value2 = autoTxt
                    End With
                End If
            End If
        End If
    Next i

    ' Y + Z data rows: one faded comment yellow. Blank Y still shows
    ' the sheet CF green (ISBLANK / ISFORMULA).
    With ss.Range(ss.Cells(Y_DATA_FIRST, "Y"), ss.Cells(Y_LAST, "Z")).Interior
        .Pattern = xlSolid
        .Color = COMMENT_YELLOW
    End With

    SyncPlanTargetWindowOnSheet ss
    HighlightActiveTargetOnSheet ss
    ' After the Y/Z yellow paint, so the newest green survey stays readable.
    HighlightNewestSurveyOnSheet ss

    m_lastRefreshAt = Timer          ' stamp only on a completed refresh

Clean:
    On Error Resume Next
    EndTextMeasure                   ' delete the reusable textbox (needs sheet still unprotected)
    SheetReprotectAfterVba ss, wasProt
    SheetEnsureLocked ss
    ' Restore calc while events are still off so the settle recalc cannot
    ' re-enter via Worksheet_Calculate.
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    m_updatingY = False
    ScreenEndBusy
End Sub

' Hidden AR (DLR) / AS (Metres To Slide) must start from the PROJECTED bit
' (W = Inc@Bit, X = Azm@Bit) — the same attitude AT (required TF) uses — not
' the survey station (F / G). Older sheets carry
'   ProjBurr(D13,F13,G13,...) / ProjMetersToSlide(D13,F13,G13,...)
' which sized the slide from the survey while the TF was aimed from the bit,
' so Y paired metres and toolface that belonged to two different manoeuvres.
' Rewrites only the F{r},G{r} argument pair; every other token is preserved.
'
' W / X (Inc@Bit / Azm@Bit) used to pass IF(AK="",0,AK): a blank dial became
' TF 0 = high side, so a rotary stand was projected as a full build with the
' demonstrated N/O walk. Pass AK through so the UDFs can see a blank dial.
' Unseen hole is still walked from R/S/U (already-drilled leftover). $U$13:$U$505
' is a dirty hook so a prior-row TF edit recalcs W/X/AJ.
' Caller owns protection / EnableEvents / calc mode.
Private Sub EnsureProjAimFormulasOnSheet(ByVal ss As Worksheet)
    Dim col As Variant
    Dim rng As Range
    Dim arr As Variant
    Dim i As Long
    Dim r As Long
    Dim f As String
    Dim oldTok As String
    Dim newTok As String
    Dim dirty As Boolean
    Dim uHook As String

    On Error GoTo Done
    uHook = "$U$" & Y_DATA_FIRST & ":$U$" & Y_LAST
    For Each col In Array("AR", "AS", "W", "X", "AJ")
        Set rng = ss.Range(ss.Cells(Y_DATA_FIRST, col), ss.Cells(Y_LAST, col))
        arr = rng.formula
        dirty = False
        For i = 1 To UBound(arr, 1)
            r = Y_DATA_FIRST + i - 1
            f = CStr(arr(i, 1) & "")
            If Len(f) > 1 And Left$(f, 1) = "=" Then
                If col = "W" Or col = "X" Then
                    oldTok = "IF(AK" & r & "="""",0,AK" & r & ")"
                    newTok = "AK" & r
                    If InStr(1, f, oldTok, vbBinaryCompare) > 0 Then
                        f = Replace(f, oldTok, newTok)
                        dirty = True
                    End If
                ElseIf col = "AR" Or col = "AS" Then
                    oldTok = "(D" & r & ",F" & r & ",G" & r & ","
                    newTok = "(D" & r & ",W" & r & ",X" & r & ","
                    If InStr(1, f, oldTok, vbBinaryCompare) > 0 Then
                        f = Replace(f, oldTok, newTok)
                        dirty = True
                    End If
                End If
                If InStr(1, f, uHook, vbBinaryCompare) = 0 Then
                    If col = "W" Or col = "X" Then
                        oldTok = "AI" & r & ")"
                        newTok = "AI" & r & "," & uHook & ")"
                        If InStr(1, f, oldTok, vbBinaryCompare) > 0 Then
                            f = Replace(f, oldTok, newTok)
                            dirty = True
                        End If
                    ElseIf col = "AJ" Then
                        oldTok = "$S$" & Y_DATA_FIRST & ":$S$" & Y_LAST & ")"
                        newTok = "$S$" & Y_DATA_FIRST & ":$S$" & Y_LAST & "," & uHook & ")"
                        If InStr(1, f, oldTok, vbBinaryCompare) > 0 Then
                            f = Replace(f, oldTok, newTok)
                            dirty = True
                        End If
                    End If
                End If
                arr(i, 1) = f
            End If
        Next i
        If dirty Then rng.formula = arr
    Next col
Done:
End Sub

Public Sub SyncProjAimFormulas()
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    On Error GoTo Clean
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    EnsureProjAimFormulasOnSheet ss
    On Error Resume Next
    ss.Range("AJ" & Y_DATA_FIRST & ":AJ" & Y_LAST).Calculate
    ss.Range("W" & Y_DATA_FIRST & ":X" & Y_LAST).Calculate
    ss.Range("V" & Y_DATA_FIRST & ":V" & Y_LAST).Calculate
    On Error GoTo Clean
Clean:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

' Public so the gauge / EMAIL path can refresh the strip without a Y rebuild.
Public Sub SyncLastGreenSurveyStrip()
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    On Error GoTo Clean
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    EnsureLastSurveyAzmLookup ss
Clean:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
    EnsureSurveyDepthMetres
End Sub

' Deepest row someone entered: MD, Inc and Azm are numbers, and Inc and Azm
' are not formulas. Yellow projection rows calculate the angles, so they are
' not the last survey. The green bar stays on LastSlidesheetSurveyRow.
Private Function LastEnteredSurveyRow(ByVal ss As Worksheet) As Long
    Dim r As Long
    LastEnteredSurveyRow = 0
    For r = Y_LAST To Y_DATA_FIRST Step -1
        If IsNumberValue(ss.Cells(r, "E").Value2) Then
            If IsNumberValue(ss.Cells(r, "F").Value2) Then
                If IsNumberValue(ss.Cells(r, "G").Value2) Then
                    If Not ss.Cells(r, "F").HasFormula Then
                        If Not ss.Cells(r, "G").HasFormula Then
                            LastEnteredSurveyRow = r
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next r
End Function

' Data!E7 is metres. Inclination and azimuth keep the degree format.
Private Sub EnsureSurveyDepthMetres()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim fmt As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error GoTo DoneDepth
    wasProt = SheetUnprotectForVba(ws)
    fmt = CStr(ws.Range("E7").NumberFormat & "")
    If InStr(1, fmt, ChrW(176), vbBinaryCompare) > 0 Then
        If InStr(1, fmt, "30", vbBinaryCompare) = 0 Then
            ws.Range("E7").NumberFormat = "#####.00""m"""
        End If
    End If
DoneDepth:
    SheetReprotectAfterVba ws, wasProt
End Sub

' Data!E7:E11 read Slidesheet E/F/G/H/J506. Those cells point at the latest
' entered survey (MD + Inc + Azm), not a yellow projection row.
Private Sub EnsureLastSurveyAzmLookup(ByVal ss As Worksheet)
    Dim r As Long
    Dim addr As String
    r = LastEnteredSurveyRow(ss)
    If r < Y_DATA_FIRST Then
        ss.Range("E506:H506,J506").ClearContents
        Exit Sub
    End If
    addr = CStr(r)
    If StrComp(CStr(ss.Range("E506").formula & ""), "=E" & addr, vbTextCompare) <> 0 Then
        ss.Range("E506").formula = "=E" & addr
    End If
    If StrComp(CStr(ss.Range("F506").formula & ""), "=F" & addr, vbTextCompare) <> 0 Then
        ss.Range("F506").formula = "=F" & addr
    End If
    If StrComp(CStr(ss.Range("G506").formula & ""), "=G" & addr, vbTextCompare) <> 0 Then
        ss.Range("G506").formula = "=G" & addr
    End If
    If StrComp(CStr(ss.Range("H506").formula & ""), "=H" & addr, vbTextCompare) <> 0 Then
        ss.Range("H506").formula = "=H" & addr
    End If
    If StrComp(CStr(ss.Range("J506").formula & ""), "=J" & addr, vbTextCompare) <> 0 Then
        ss.Range("J506").formula = "=J" & addr
    End If
End Sub

' Pad left/right text with spaces so the right part sits on the cell's right edge.
' Uses point widths (font-accurate), not ColumnWidth character units.
Private Function FitLeftRightInCell(ByVal leftPart As String, ByVal rightPart As String, _
                                    ByVal cell As Range, ByVal spacePts As Double) As String
    Dim avail As Double
    Dim leftPts As Double
    Dim rightPts As Double
    Dim PAD As Long
    ' Excel keeps a small built-in cell pad; 2pt leaves DLR flush to the right grid.
    Const CELL_MARGIN_PTS As Double = 2#

    leftPts = MeasureTextWidthPts(leftPart, cell)
    rightPts = MeasureTextWidthPts(rightPart, cell)
    avail = cell.width - CELL_MARGIN_PTS
    If spacePts <= 0# Then spacePts = 1#

    PAD = CLng(Application.WorksheetFunction.Round((avail - leftPts - rightPts) / spacePts, 0))
    If PAD < 2 Then PAD = 2
    FitLeftRightInCell = leftPart & Space$(PAD) & rightPart
End Function

' Sliding left, DLR, then leftover-rotate flush to the right grid.
Private Function FitCommentParts(ByVal leftPart As String, ByVal midPart As String, _
                                 ByVal rotPart As String, ByVal cell As Range, _
                                 ByVal spacePts As Double) As String
    Dim avail As Double
    Dim leftPts As Double, midPts As Double, rotPts As Double
    Dim gapPts As Double
    Dim PAD As Long
    Const CELL_MARGIN_PTS As Double = 2#

    If spacePts <= 0# Then spacePts = 1#
    leftPts = MeasureTextWidthPts(leftPart, cell)
    midPts = MeasureTextWidthPts(midPart, cell)
    rotPts = MeasureTextWidthPts(rotPart, cell)
    avail = cell.width - CELL_MARGIN_PTS
    gapPts = spacePts * 2#

    If Len(leftPart) = 0 And Len(midPart) = 0 Then
        PAD = CLng(Application.WorksheetFunction.Round((avail - rotPts) / spacePts, 0))
        If PAD < 0 Then PAD = 0
        FitCommentParts = Space$(PAD) & rotPart
        Exit Function
    End If

    PAD = CLng(Application.WorksheetFunction.Round( _
        (avail - leftPts - midPts - rotPts - gapPts) / spacePts, 0))
    If PAD < 2 Then PAD = 2
    If Len(midPart) = 0 Then
        FitCommentParts = leftPart & Space$(PAD) & rotPart
    Else
        FitCommentParts = leftPart & Space$(PAD) & midPart & Space$(2) & rotPart
    End If
End Function

' Open a measuring pass: one hidden autosized textbox reused for every
' MeasureTextWidthPts call until EndTextMeasure, plus a width cache.
Private Sub BeginTextMeasure(ByVal proto As Range)
    On Error Resume Next
    EndTextMeasure
    Set m_measureCache = CreateObject("Scripting.Dictionary")
    Set m_measureShape = proto.Worksheet.Shapes.AddTextbox(1, -2000, -2000, 20, 20)
    With m_measureShape
        .Visible = False
        With .TextFrame
            .AutoSize = True
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
        End With
    End With
    On Error GoTo 0
End Sub

Private Sub EndTextMeasure()
    On Error Resume Next
    If Not m_measureShape Is Nothing Then m_measureShape.Delete
    Set m_measureShape = Nothing
    Set m_measureCache = Nothing
    On Error GoTo 0
End Sub

' Measure rendered text width in points via a hidden autosized textbox.
' Inside a Begin/EndTextMeasure pass the shape is reused and widths cached;
' standalone calls keep the original create/measure/delete behavior.
Private Function MeasureTextWidthPts(ByVal txt As String, ByVal proto As Range) As Double
    Dim shp As Shape
    Dim nm As String
    Dim sz As Double
    Dim bld As Boolean
    Dim ownShape As Boolean
    Dim key As String

    If Len(txt) = 0 Then
        MeasureTextWidthPts = 0#
        Exit Function
    End If

    On Error Resume Next
    bld = CBool(proto.Font.bold)
    On Error GoTo Fail

    nm = proto.Font.name
    sz = proto.Font.size
    If sz <= 0# Then sz = 11#

    If Not m_measureCache Is Nothing Then
        key = txt & "|" & nm & "|" & sz & "|" & bld
        If m_measureCache.Exists(key) Then
            MeasureTextWidthPts = m_measureCache(key)
            Exit Function
        End If
    End If

    If m_measureShape Is Nothing Then
        Set shp = proto.Worksheet.Shapes.AddTextbox(1, -2000, -2000, 20, 20) ' msoTextOrientationHorizontal
        ownShape = True
        With shp
            .Visible = False
            With .TextFrame
                .AutoSize = True
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
            End With
        End With
    Else
        Set shp = m_measureShape
    End If

    With shp.TextFrame.Characters
        .text = txt
        .Font.name = nm
        .Font.size = sz
        .Font.bold = bld
    End With
    MeasureTextWidthPts = shp.width
    If ownShape Then shp.Delete

    If Not m_measureCache Is Nothing Then m_measureCache(key) = MeasureTextWidthPts
    Exit Function
Fail:
    On Error Resume Next
    If ownShape And Not shp Is Nothing Then shp.Delete
    ' Fallback: ~half font-size points per character
    MeasureTextWidthPts = Len(txt) * sz * 0.5
End Function

' Highlight T:Y of the plan target the last surveyed row is aiming at.
' Public entrypoint for sheet events / button handlers.
Public Sub HighlightActiveTarget()
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    On Error GoTo Fail
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    SyncPlanTargetWindowOnSheet ss
    HighlightActiveTargetOnSheet ss
    SheetReprotectAfterVba ss, wasProt
    SheetEnsureLocked ss
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    SheetEnsureLocked ss
    Application.EnableEvents = prevEvents
End Sub

' Newest survey = the bottom row whose Inc (F) and Azm (G) are the survey
' green and nothing else. H:Z on that row uses the same light green as F:G
' so the latest entry stays readable. The bar moves when the next green
' survey is entered; the row it leaves goes back to the column fills
' (gray H/I, yellow calc, gray inputs, green T/U).
' Caller owns EnableEvents and protection.
Private Sub HighlightNewestSurveyOnSheet(ByVal ss As Worksheet)
    Dim r As Long
    Dim i As Long
    Dim fill As Long

    On Error GoTo Done
    fill = NewestSurveyFill()
    r = MDL_ContDI.LastSlidesheetSurveyRow()
    EnsureNewestSurveyRule ss, r

    For i = Y_DATA_FIRST To Y_LAST
        If i <> r Then
            If ss.Cells(i, "J").Interior.Color = fill Then
                RestoreSurveyRowChrome ss, i
            End If
        End If
    Next i

    If r >= Y_DATA_FIRST Then ApplySurveyRowChrome ss, r
    m_newestRow = r
Done:
End Sub

' Recoloring F/G (green -> yellow, or back) raises no Excel event, so the bar
' cannot follow on its own. Sheet48 calls this when the selection leaves F:G.
' Cheap when nothing moved: one scan up from the bottom, no repaint.
Public Sub RefreshNewestSurveyIfMoved()
    Dim r As Long

    On Error GoTo Done
    r = MDL_ContDI.LastSlidesheetSurveyRow()
    If r = m_newestRow Then Exit Sub
    HighlightNewestSurveyRow
Done:
End Sub

Public Sub HighlightNewestSurveyRow()
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    On Error GoTo Fail
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    HighlightNewestSurveyOnSheet ss
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

Private Function NewestSurveyFill() As Long
    ' Same light green as the Inc / Azm entry columns F:G.
    NewestSurveyFill = RGB(204, 255, 204)
End Function

Private Function NewestSurveyFont() As Long
    NewestSurveyFont = RGB(0, 0, 0)
End Function

' H/I (TVD, S.S) and P/R/S are the gray inputs. J:O and Q and V:Z are the
' pale calc yellow. T/U stay the survey green the DD paints on a real station.
Private Function SurveyColRestFill(ByVal col As Long) As Long
    Select Case col
        Case 8, 9, 16, 18, 19
            SurveyColRestFill = RGB(242, 242, 242)
        Case 20, 21
            SurveyColRestFill = RGB(204, 255, 204)
        Case Else
            SurveyColRestFill = COMMENT_YELLOW
    End Select
End Function

Private Sub ApplySurveyRowChrome(ByVal ss As Worksheet, ByVal r As Long)
    Dim col As Long
    For col = 8 To 26
        With ss.Cells(r, col)
            .Interior.Pattern = xlSolid
            .Interior.Color = NewestSurveyFill()
            .Font.Color = NewestSurveyFont()
        End With
    Next col
End Sub

Private Sub RestoreSurveyRowChrome(ByVal ss As Worksheet, ByVal r As Long)
    Dim col As Long
    For col = 8 To 26
        With ss.Cells(r, col)
            .Interior.Pattern = xlSolid
            .Interior.Color = SurveyColRestFill(col)
            .Font.Color = RGB(0, 0, 0)
        End With
    Next col
End Sub

' Beats the Y-column "blank = green" rule, which would otherwise cover
' the newest-survey fill on an empty comment.
Private Sub EnsureNewestSurveyRule(ByVal ss As Worksheet, ByVal r As Long)
    Dim area As Range
    Dim fc As Object
    Dim i As Long
    Dim formula As String
    Dim keep As Long
    Dim wasExact As Boolean

    If r < Y_DATA_FIRST Then formula = "=ROW()=0" Else formula = "=ROW()=" & CStr(r)
    ' Drop a rule that still covers only J:Z so H:I are not left outside it.
    Set area = ss.Range(ss.Cells(Y_DATA_FIRST, "J"), ss.Cells(Y_LAST, "Z"))
    On Error Resume Next
    For i = area.FormatConditions.Count To 1 Step -1
        Set fc = area.FormatConditions(i)
        If InStr(1, CStr(fc.Formula1), "ROW()=", vbTextCompare) > 0 Then fc.Delete
    Next i
    On Error GoTo 0
    Set area = ss.Range(ss.Cells(Y_DATA_FIRST, "H"), ss.Cells(Y_LAST, "Z"))

    ' An existing rule's formula is read-only and a later edit can split its
    ' range, so a rule left over from an earlier row is deleted, not edited.
    ' Only the rule that is already this row over the whole H13:Z505 block is kept.
    On Error Resume Next
    For i = area.FormatConditions.Count To 1 Step -1
        Set fc = area.FormatConditions(i)
        If InStr(1, CStr(fc.Formula1), "ROW()=", vbTextCompare) > 0 Then
            wasExact = (StrComp(CStr(fc.Formula1), formula, vbTextCompare) = 0) _
                       And (fc.AppliesTo.Address(False, False) = area.Address(False, False))
            If wasExact And keep = 0 Then
                keep = 1
            Else
                fc.Delete
            End If
        End If
    Next i
    On Error GoTo 0
    If keep = 1 Then
        On Error Resume Next
        For i = 1 To area.FormatConditions.Count
            Set fc = area.FormatConditions(i)
            If StrComp(CStr(fc.Formula1), formula, vbTextCompare) = 0 Then
                fc.Interior.Color = NewestSurveyFill()
                fc.Font.Color = NewestSurveyFont()
            End If
        Next i
        On Error GoTo 0
        Exit Sub
    End If

    Set fc = area.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    fc.Interior.Color = NewestSurveyFill()
    fc.Font.Color = NewestSurveyFont()
    fc.StopIfTrue = True
    fc.SetFirstPriority
End Sub

' Assumes the caller already owns EnableEvents / protection.
Private Sub HighlightActiveTargetOnSheet(ByVal ss As Worksheet)
    Dim r As Long
    Dim bitMd As Double
    Dim aimMd As Double
    Dim tgtMd As Variant
    Dim activeRow As Long
    Dim lastSurvRow As Long
    Dim hlColor As Long
    Dim rng As Range
    Dim vF As Variant
    Dim vAo As Variant
    Dim vD As Variant

    On Error GoTo Done
    hlColor = RGB(198, 239, 206)

    ' Use the last GREEN surveyed row, not a yellow projection or bit-only tally.
    ' Bit-only depths past a target MD would otherwise advance the highlight
    ' while BURR is still computed against the prior survey's aim (e.g. TAR2 / 8.37).
    lastSurvRow = 0
    bitMd = 0#
    aimMd = 0#
    For r = 12 To Y_LAST
        vF = ss.Cells(r, "F").Value2
        If IsNumeric(vF) Then
            If Len(Trim$(CStr(vF & ""))) > 0 Then
                If Not MDL_ContDI.IsSlidesheetGoodSurveyRowOn(ss, r) Then GoTo NextAimRow
                lastSurvRow = r
                vD = ss.Cells(r, "D").Value2
                If IsNumeric(vD) Then bitMd = CDbl(vD)
                vAo = ss.Cells(r, "AO").Value2
                If IsNumeric(vAo) Then
                    aimMd = CDbl(vAo)
                Else
                    aimMd = 0#
                End If
            End If
        End If
NextAimRow:
    Next r

    activeRow = 0
    If lastSurvRow > 0 Then
        ' Prefer the row's AO (ProjActiveTargetMd) so highlight matches BURR's aim.
        If aimMd > 0# Then
            For r = TGT_FIRST To TGT_LAST
                tgtMd = ss.Cells(r, "U").Value2
                If IsNumeric(tgtMd) Then
                    If Abs(CDbl(tgtMd) - aimMd) < 0.005 Then
                        activeRow = r
                        Exit For
                    End If
                End If
            Next r
        End If
        ' Fallback: first plan MD ahead of that survey's bit depth.
        If activeRow = 0 And bitMd > 0# Then
            For r = TGT_FIRST To TGT_LAST
                tgtMd = ss.Cells(r, "U").Value2
                If IsNumeric(tgtMd) Then
                    If CDbl(tgtMd) > bitMd Then
                        activeRow = r
                        Exit For
                    End If
                End If
            Next r
        End If
    End If

    ' Skip paint work if nothing changed.
    If activeRow = m_hlTgtRow Then
        If activeRow = 0 Then GoTo Done
        If ss.Cells(activeRow, "T").Interior.Color = hlColor _
           And ss.Cells(activeRow, "Y").Interior.Color = hlColor Then GoTo Done
    End If

    ' Clear previous highlight on T2:Y5.
    Set rng = ss.Range(ss.Cells(TGT_FIRST, "T"), ss.Cells(TGT_LAST, "Y"))
    rng.Interior.Pattern = xlNone

    If activeRow >= TGT_FIRST And activeRow <= TGT_LAST Then
        ss.Range(ss.Cells(activeRow, "T"), ss.Cells(activeRow, "Y")).Interior.Color = hlColor
    End If
    m_hlTgtRow = activeRow

Done:
End Sub

' One pass after a code update: rebuild the target list (adds the name
' column), relink Cont DI, recalc the aim columns, rewrite auto Y text,
' redraw the gauge.
Public Sub ApplyPlanPointAim()
    Dim ss As Worksheet
    Dim nm As Variant
    Dim wasProt As Boolean
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    On Error Resume Next
    For Each nm In Array("SlideAimMD", "SlideAimINC", "SlideAimAZM", _
                         "SlideAimTVD", "SlideAimNAME", "SlideAimTF")
        ThisWorkbook.names(CStr(nm)).Delete
    Next nm
    On Error GoTo 0
    SyncPlanTargetWindow
    MDL_ContDI.ContDI_FollowSlideTarget
    wasProt = SheetUnprotectForVba(ss)
    ss.Range(ss.Cells(Y_DATA_FIRST, "AR"), ss.Cells(Y_LAST, "AT")).dirty
    ss.Calculate
    SheetReprotectAfterVba ss, wasProt
    ThisWorkbook.Worksheets(MDL_ContDI.SH_CONTDI).Calculate
    RefreshSlideComments
    MDL_PlanGauge.RenderPlanGaugeStrict
End Sub

' True for a real number; Empty and "" (formula blank) must not count.
Private Function IsNumberValue(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            IsNumberValue = True
        Case Else
            IsNumberValue = False
    End Select
End Function

' True if Y is empty or looks like previously generated slide text.
Private Function IsAutoSlideCommentText(ByVal t As String) As Boolean
    Dim s As String
    s = Trim$(t)
    If Len(s) = 0 Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    ' Accept old BURR text and current DLR text so a refresh upgrades both.
    If s Like "Sliding *| BURR *" Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    If s Like "Sliding *| DLR *" Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    ' Generated text that lost its DLR part ("Sliding 2.28m @ R53").
    If s Like "Sliding #*m @ *" Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    If (Left$(s, 8) = "Sliding ") _
       And ((InStr(1, s, "BURR ", vbTextCompare) > 0) _
         Or (InStr(1, s, "DLR ", vbTextCompare) > 0)) Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    ' Leftover "Nm rot" that used to live in Y — clear those back to slide/DLR only.
    IsAutoSlideCommentText = (InStr(1, s, "m rot", vbTextCompare) > 0)
End Function

Private Function IsAutoRotText(ByVal t As String) As Boolean
    Dim s As String
    s = Trim$(t)
    If Len(s) = 0 Then
        IsAutoRotText = True
        Exit Function
    End If
    If (Right$(s, 5) = "m ROT") Or (Right$(s, 5) = "m rot") Then
        IsAutoRotText = True
        Exit Function
    End If
    IsAutoRotText = IsNumeric(s)
End Function

Public Sub EnsureSlidesheetClearButton()
    Dim ss As Worksheet
    Dim rng As Range
    Dim btn As Button
    Dim shp As Shape

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    Set rng = ss.Range(BTN_ANCHOR)
    If rng.Cells.Count = 1 Then Set rng = rng.MergeArea

    On Error Resume Next
    ss.Buttons(BTN_NAME).Delete
    ss.Shapes(BTN_NAME).Delete
    On Error GoTo 0

    Set btn = ss.Buttons.Add(rng.Left, rng.Top, rng.width, rng.Height)
    btn.name = BTN_NAME
    btn.caption = BTN_CAPTION
    btn.OnAction = "'" & ThisWorkbook.name & "'!ClearSlidesheetRanges"

    Set shp = ss.Shapes(BTN_NAME)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.size = 10
    On Error GoTo 0
End Sub

' Targets button on Z6. Created once; an existing button is left where it is.
Public Sub EnsureTargetButton()
    Dim ss As Worksheet
    Dim rng As Range
    Dim btn As Button
    Dim shp As Shape
    Dim wasProt As Boolean

    On Error GoTo Fail
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    Set rng = ss.Range("Z6")
    If rng.MergeCells Then Set rng = rng.MergeArea
    On Error Resume Next
    Set shp = ss.Shapes(BTN_TGT)
    On Error GoTo Fail
    If Not shp Is Nothing Then
        ' The first build was taller than Z6. Snap it back onto that cell.
        If Abs(shp.Height - rng.Height) > 2# Or Abs(shp.width - rng.width) > 2# _
                Or Abs(shp.Top - rng.Top) > 2# Or Abs(shp.Left - rng.Left) > 2# Then
            wasProt = SheetUnprotectForVba(ss)
            shp.Left = rng.Left
            shp.Top = rng.Top
            shp.width = rng.width
            shp.Height = rng.Height
            SheetReprotectAfterVba ss, wasProt
        End If
        Exit Sub
    End If

    wasProt = SheetUnprotectForVba(ss)
    Set btn = ss.Buttons.Add(rng.Left, rng.Top, rng.width, rng.Height)
    btn.name = BTN_TGT
    btn.caption = "Targets"
    btn.OnAction = "'" & ThisWorkbook.name & "'!ShowTargetEditor"
    Set shp = ss.Shapes(BTN_TGT)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.size = 10
    On Error GoTo Fail
    SheetReprotectAfterVba ss, wasProt
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
End Sub

Public Sub ShowTargetEditor()
    frmTargets.Show vbModal
End Sub

' Only creates the button when it is missing.  An existing button is left exactly
' where it is (anchor Z1:Z3, clear of the plan gauge in AA1:AC7).
Public Sub ResizeSlidesheetClearButton()
    Dim ss As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    If ss Is Nothing Then Exit Sub
    Set shp = ss.Shapes(BTN_NAME)
    If shp Is Nothing Then EnsureSlidesheetClearButton
    On Error GoTo 0
End Sub

' Blank T2:Y5 so a new Import Plan cannot keep the previous job's TAR rows.
Public Sub WipePlanTargetWindow()
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    On Error GoTo Fail
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    ss.Range(ss.Cells(TGT_FIRST, "T"), ss.Cells(TGT_LAST, "Y")).ClearContents
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

' Show 4 named Plan Section targets on T2:Y5. Hidden later stations slide in
' as the bit passes the highlighted (next) target. BURR / slide UDFs read the
' full named list on _OC_PlanSec N:Q (ProjTargets_*), not the 4-row window.
Public Sub SyncPlanTargetWindow()
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    On Error GoTo Fail
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    SyncPlanTargetWindowOnSheet ss
    HighlightActiveTargetOnSheet ss
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

Private Sub SyncPlanTargetWindowOnSheet(ByVal ss As Worksheet)
    Dim ps As Worksheet
    Dim n As Long
    Dim md() As Double, inc() As Double, azm() As Double, tvd() As Double
    Dim nm() As String
    Dim bitMd As Double
    Dim aimI As Long
    Dim startI As Long
    Dim i As Long
    Dim vis As Long
    Dim r As Long
    Dim changed As Boolean
    Dim newMd As Variant, newInc As Variant, newAzm As Variant
    Dim newTvd As Variant, newNm As String

    If m_syncingTgt Then Exit Sub
    m_syncingTgt = True
    On Error GoTo Done

    If Not PlanSecSheetExists() Then GoTo Done
    Set ps = ThisWorkbook.Worksheets(PLANSEC_SHEET)
    n = LoadNamedPlanTargets(ps, md, inc, azm, tvd, nm)
    If n < 1 Then GoTo Done

    EnsureProjTargetNames
    WriteFullProjTargetTable ps, md, inc, azm, tvd, nm, n

    bitMd = LastSurveyBitMd(ss)
    aimI = 1
    For i = 1 To n
        If md(i) > bitMd + 0.005 Then
            aimI = i
            Exit For
        End If
        If i = n Then aimI = n
    Next i

    startI = aimI
    If startI + 3 > n Then startI = n - 3
    If startI < 1 Then startI = 1

    ' VBA IsNumeric("") is True and CDbl("") is 0, so a blank Inc/Azm cell
    ' used to compare equal to a real 0 and never get written.
    ' DisplayZeros is a Window property, not Worksheet — do not set it here
    ' (compile error: method or data member not found). WriteTgtNumber
    ' already formats zeros as 0.00.

    changed = False
    For vis = 0 To 3
        r = TGT_FIRST + vis
        If startI + vis <= n Then
            i = startI + vis
            newMd = md(i): newInc = inc(i): newAzm = azm(i)
            newTvd = tvd(i): newNm = nm(i)
            WriteTgtLabel ss.Cells(r, "T"), "TAR" & CStr(i), changed
        Else
            newMd = "": newInc = "": newAzm = ""
            newTvd = "": newNm = ""
            WriteTgtLabel ss.Cells(r, "T"), "", changed
        End If
        WriteTgtNumber ss.Cells(r, "U"), newMd, changed
        WriteTgtNumber ss.Cells(r, "V"), newInc, changed
        WriteTgtNumber ss.Cells(r, "W"), newAzm, changed
        WriteTgtNumber ss.Cells(r, "X"), newTvd, changed
        If Trim$(CStr(ss.Cells(r, "Y").Value2 & "")) <> newNm Then
            ss.Cells(r, "Y").Value2 = newNm
            changed = True
        End If
    Next vis

    If changed Then ApplyTargetNameDropdown ss

Done:
    m_syncingTgt = False
End Sub

' Workbook names the V/W/X/AO/AR/AS/AT formulas already call.
Private Sub EnsureProjTargetNames()
    Dim mdRef As String, incRef As String, azmRef As String, tvdRef As String
    Dim nameRef As String
    Dim startRef As String
    Dim lastR As Long

    lastR = FULL_TGT_FIRST + FULL_TGT_ROWS - 1
    mdRef = "='" & PLANSEC_SHEET & "'!$N$" & FULL_TGT_FIRST & ":$N$" & lastR
    incRef = "='" & PLANSEC_SHEET & "'!$O$" & FULL_TGT_FIRST & ":$O$" & lastR
    azmRef = "='" & PLANSEC_SHEET & "'!$P$" & FULL_TGT_FIRST & ":$P$" & lastR
    tvdRef = "='" & PLANSEC_SHEET & "'!$Q$" & FULL_TGT_FIRST & ":$Q$" & lastR
    nameRef = "='" & PLANSEC_SHEET & "'!$R$" & FULL_TGT_FIRST & ":$R$" & lastR
    startRef = "='" & PLANSEC_SHEET & "'!$N$1"

    SetWbName "ProjTargets_MD", mdRef
    SetWbName "ProjTargets_INC", incRef
    SetWbName "ProjTargets_AZM", azmRef
    SetWbName "ProjTargets_TVD", tvdRef
    SetWbName "ProjTargets_NAME", nameRef
    SetWbName "ProjBuildStartMd", startRef
End Sub

Private Sub SetWbName(ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    ThisWorkbook.names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.names.Add name:=nm, refersTo:=refersTo
End Sub

' N1 = first KOP/TANGENT MD (BURR start). N3:Q = named plan stations up to HEEL.
Private Sub WriteFullProjTargetTable(ByVal ps As Worksheet, _
        ByRef md() As Double, ByRef inc() As Double, ByRef azm() As Double, _
        ByRef tvd() As Double, ByRef nm() As String, ByVal n As Long)
    Dim i As Long
    Dim startMd As Double
    Dim lastR As Long
    Dim wasProt As Boolean

    wasProt = SheetUnprotectForVba(ps)

    startMd = 0#
    For i = 1 To n
        If nm(i) = "KOP" Or nm(i) = "TANGENT" Then
            startMd = md(i)
            Exit For
        End If
    Next i
    ps.Cells(1, FULL_TGT_MD_COL).Value2 = startMd

    ' HEEL is the last steering station: past it the lateral is steered on the
    ' geo corridor, and aiming slide/BURR at plan TD kept Y comments alive.
    lastR = FULL_TGT_FIRST + FULL_TGT_ROWS - 1
    ps.Range(ps.Cells(FULL_TGT_FIRST, 14), ps.Cells(lastR, 18)).ClearContents
    For i = 1 To n
        If i > FULL_TGT_ROWS Then Exit For
        ps.Cells(FULL_TGT_FIRST + i - 1, 14).Value2 = md(i)
        ps.Cells(FULL_TGT_FIRST + i - 1, 15).Value2 = inc(i)
        ps.Cells(FULL_TGT_FIRST + i - 1, 16).Value2 = azm(i)
        ps.Cells(FULL_TGT_FIRST + i - 1, 17).Value2 = tvd(i)
        ps.Cells(FULL_TGT_FIRST + i - 1, 18).Value2 = nm(i)
        If nm(i) = "HEEL" Then Exit For
    Next i

    SheetReprotectAfterVba ps, wasProt
End Sub

Private Function TgtCellBlank(ByVal v As Variant) As Boolean
    If isError(v) Then TgtCellBlank = True: Exit Function
    If IsNull(v) Then TgtCellBlank = True: Exit Function
    If IsEmpty(v) Then TgtCellBlank = True: Exit Function
    TgtCellBlank = (Len(Trim$(CStr(v & ""))) = 0)
End Function

Private Function SameTgtCell(ByVal cur As Variant, ByVal neu As Variant) As Boolean
    If TgtCellBlank(cur) And TgtCellBlank(neu) Then
        SameTgtCell = True
        Exit Function
    End If
    If TgtCellBlank(cur) Or TgtCellBlank(neu) Then
        SameTgtCell = False
        Exit Function
    End If
    If IsNumeric(cur) And IsNumeric(neu) Then
        SameTgtCell = (Abs(CDbl(cur) - CDbl(neu)) < 0.005)
        Exit Function
    End If
    SameTgtCell = (Trim$(CStr(cur & "")) = Trim$(CStr(neu & "")))
End Function

Private Sub WriteTgtNumber(ByVal cell As Range, ByVal neu As Variant, ByRef changed As Boolean)
    If TgtCellBlank(neu) Then
        If Not TgtCellBlank(cell.Value2) Then
            cell.ClearContents
            changed = True
        End If
        cell.numberFormat = "0.00"
        Exit Sub
    End If
    If Not SameTgtCell(cell.Value2, neu) Then
        cell.Value2 = CDbl(neu)
        changed = True
    End If
    If CStr(cell.numberFormat & "") <> "0.00" Then cell.numberFormat = "0.00"
End Sub

Private Sub WriteTgtLabel(ByVal cell As Range, ByVal neu As String, ByRef changed As Boolean)
    If Trim$(CStr(cell.Value2 & "")) <> neu Then
        cell.Value2 = neu
        changed = True
    End If
End Sub

Private Function LastSurveyBitMd(ByVal ss As Worksheet) As Double
    Dim r As Long
    Dim vD As Variant
    LastSurveyBitMd = 0#
    r = MDL_ContDI.LastSlidesheetSurveyRow()
    If r < Y_DATA_FIRST Then Exit Function
    vD = ss.Cells(r, "D").Value2
    If IsNumeric(vD) Then LastSurveyBitMd = CDbl(vD)
End Function

Private Function LoadNamedPlanTargets(ByVal ps As Worksheet, _
        ByRef md() As Double, ByRef inc() As Double, ByRef azm() As Double, _
        ByRef tvd() As Double, ByRef nm() As String) As Long

    Dim lastR As Long
    Dim r As Long
    Dim n As Long
    Dim autoNm As String, userNm As String, showNm As String
    Dim vMd As Variant

    lastR = ps.Cells(ps.rows.Count, 1).End(xlUp).row
    ReDim md(1 To 80): ReDim inc(1 To 80): ReDim azm(1 To 80)
    ReDim tvd(1 To 80): ReDim nm(1 To 80)
    n = 0
    For r = 3 To lastR
        vMd = ps.Cells(r, 1).Value2
        If Not IsNumeric(vMd) Then GoTo NextPs
        autoNm = UCase$(Trim$(CStr(ps.Cells(r, 11).Value2 & "")))
        userNm = UCase$(Trim$(CStr(ps.Cells(r, 12).Value2 & "")))
        If userNm <> "" Then
            showNm = userNm
        Else
            showNm = autoNm
        End If
        If showNm = "" Then
            showNm = SeedNameFromText(CStr(ps.Cells(r, 10).Value2 & ""))
        End If
        If Not IsKnownTargetName(showNm) Then GoTo NextPs
        n = n + 1
        If n > 80 Then Exit For
        md(n) = PickedNum(ps.Cells(r, OVR_MD).Value2, CDbl(vMd))
        inc(n) = PickedNum(ps.Cells(r, OVR_INC).Value2, val(ps.Cells(r, 2).Value2 & ""))
        azm(n) = PickedNum(ps.Cells(r, OVR_AZM).Value2, val(ps.Cells(r, 3).Value2 & ""))
        tvd(n) = PickedNum(ps.Cells(r, OVR_TVD).Value2, val(ps.Cells(r, 4).Value2 & ""))
        nm(n) = showNm
NextPs:
    Next r
    If n > 0 Then
        ReDim Preserve md(1 To n)
        ReDim Preserve inc(1 To n)
        ReDim Preserve azm(1 To n)
        ReDim Preserve tvd(1 To n)
        ReDim Preserve nm(1 To n)
    End If
    LoadNamedPlanTargets = n
End Function

Private Function IsKnownTargetName(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    IsKnownTargetName = (s = "KOP" Or s = "TANGENT" Or s = "HEEL" _
                      Or s = "SOT" Or s = "EOT" Or s = "TD" _
                      Or s = "NUDGE" Or s = "VERTICAL" _
                      Or s = "ICP" Or s = "BUILD")
End Function

Private Function PlanSecSheetExists() As Boolean
    On Error Resume Next
    PlanSecSheetExists = Not (ThisWorkbook.Worksheets(PLANSEC_SHEET) Is Nothing)
    On Error GoTo 0
End Function

Private Sub ApplyTargetNameDropdown(ByVal ss As Worksheet)
    Dim rng As Range
    On Error Resume Next
    Set rng = ss.Range(ss.Cells(TGT_FIRST, "Y"), ss.Cells(TGT_LAST, "Y"))
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=TGT_NAMES
        .InCellDropdown = True
        .ShowError = False
        .ShowInput = False
    End With
    rng.Locked = False
    On Error GoTo 0
End Sub

' Persist a T2:Y5 edit onto _OC_PlanSec. The name goes to USERNAME.
' MD / INC / AZM / TVD go to the override columns, which the next sync reads
' back, so the geologist's HEEL TVD is not replaced by the imported plan.
Public Sub RememberPlanTargetOverride(ByVal Target As Range)
    Dim ss As Worksheet
    Dim ps As Worksheet
    Dim cell As Range
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Dim nm As String
    Dim md As Double
    Dim haveMd As Boolean
    Dim row As Long
    Dim ovrCol As Long
    Dim planCol As Long
    Dim entered As Double

    On Error GoTo Fail
    If m_syncingTgt Then Exit Sub
    If Target Is Nothing Then Exit Sub
    Set ss = Target.Worksheet
    If Not PlanSecSheetExists() Then Exit Sub
    Set ps = ThisWorkbook.Worksheets(PLANSEC_SHEET)

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ps)

    For Each cell In Target.Cells
        If cell.row < TGT_FIRST Or cell.row > TGT_LAST Then GoTo NextOv
        If cell.Column < 21 Or cell.Column > 25 Then GoTo NextOv
        nm = UCase$(Trim$(CStr(ss.Cells(cell.row, "Y").Value2 & "")))
        haveMd = IsNumeric(ss.Cells(cell.row, "U").Value2)
        If haveMd Then md = CDbl(ss.Cells(cell.row, "U").Value2)
        ' The MD cell itself now holds the new number, so match on the name.
        If cell.Column = 21 Then haveMd = False
        row = FindNamedTargetRow(ps, nm, md, haveMd)
        If row = 0 Then GoTo NextOv
        If cell.Column = 25 Then
            If nm = "" Or Not IsKnownTargetName(nm) Then
                ps.Cells(row, 12).ClearContents
            Else
                ps.Cells(row, 12).value = nm
            End If
            GoTo NextOv
        End If
        If Not IsNumeric(cell.Value2) Then GoTo NextOv
        entered = CDbl(cell.Value2)
        Select Case cell.Column
            Case 21: ovrCol = OVR_MD: planCol = 1
            Case 22: ovrCol = OVR_INC: planCol = 2
            Case 23: ovrCol = OVR_AZM: planCol = 3
            Case 24: ovrCol = OVR_TVD: planCol = 4
        End Select
        WriteTargetOverride ps, row, ovrCol, planCol, entered
NextOv:
    Next cell

    SheetReprotectAfterVba ps, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ps, wasProt
    Application.EnableEvents = prevEvents
End Sub

Private Function PickedNum(ByVal ovr As Variant, ByVal planVal As Double) As Double
    If IsNumeric(ovr) Then
        If Len(Trim$(CStr(ovr & ""))) > 0 Then
            PickedNum = CDbl(ovr)
            Exit Function
        End If
    End If
    PickedNum = planVal
End Function

Private Sub WriteTargetOverride(ByVal ps As Worksheet, ByVal row As Long, _
        ByVal ovrCol As Long, ByVal planCol As Long, ByVal entered As Double)
    Dim planVal As Double
    If IsNumeric(ps.Cells(row, planCol).Value2) Then
        planVal = CDbl(ps.Cells(row, planCol).Value2)
        If Abs(entered - planVal) < 0.005 Then
            ps.Cells(row, ovrCol).ClearContents
            Exit Sub
        End If
    End If
    ps.Cells(row, ovrCol).Value2 = entered
End Sub

' haveMd: match name and the current target MD. Otherwise the name alone,
' and only when that name is on one station.
Private Function FindNamedTargetRow(ByVal ps As Worksheet, ByVal nm As String, _
        ByVal md As Double, ByVal haveMd As Boolean) As Long
    Dim lastR As Long
    Dim r As Long
    Dim showNm As String
    Dim rowMd As Double
    Dim hits As Long
    Dim hit As Long
    Dim v As Variant

    FindNamedTargetRow = 0
    lastR = ps.Cells(ps.rows.Count, 1).End(xlUp).row
    For r = 3 To lastR
        If Not IsNumeric(ps.Cells(r, 1).Value2) Then GoTo NextFind
        showNm = UCase$(Trim$(CStr(ps.Cells(r, 12).Value2 & "")))
        If Len(showNm) = 0 Then showNm = UCase$(Trim$(CStr(ps.Cells(r, 11).Value2 & "")))
        If Len(nm) > 0 And showNm <> nm Then GoTo NextFind
        v = ps.Cells(r, OVR_MD).Value2
        If IsNumeric(v) And Len(Trim$(CStr(v & ""))) > 0 Then
            rowMd = CDbl(v)
        Else
            rowMd = CDbl(ps.Cells(r, 1).Value2)
        End If
        If haveMd Then
            If Abs(rowMd - md) < 0.05 Then
                FindNamedTargetRow = r
                Exit Function
            End If
        Else
            hits = hits + 1
            hit = r
        End If
NextFind:
    Next r
    If Not haveMd And hits = 1 Then FindNamedTargetRow = hit
End Function

' --- Targets form -----------------------------------------------------------

Private Sub LoadTargetEditor()
    Dim ps As Worksheet
    Dim lastR As Long
    Dim r As Long
    Dim n As Long
    Dim autoNm As String
    Dim userNm As String
    Dim showNm As String
    Dim vMd As Variant

    m_teN = 0
    If Not PlanSecSheetExists() Then Exit Sub
    Set ps = ThisWorkbook.Worksheets(PLANSEC_SHEET)
    lastR = ps.Cells(ps.rows.Count, 1).End(xlUp).row
    ReDim m_teRow(1 To 40)
    ReDim m_teNm(1 To 40)
    ReDim m_teMd(1 To 40)
    ReDim m_teInc(1 To 40)
    ReDim m_teAzm(1 To 40)
    ReDim m_teTvd(1 To 40)
    For r = 3 To lastR
        vMd = ps.Cells(r, 1).Value2
        If Not IsNumeric(vMd) Then GoTo NextEd
        autoNm = UCase$(Trim$(CStr(ps.Cells(r, 11).Value2 & "")))
        userNm = UCase$(Trim$(CStr(ps.Cells(r, 12).Value2 & "")))
        If Len(userNm) > 0 Then showNm = userNm Else showNm = autoNm
        If Len(showNm) = 0 Then showNm = SeedNameFromText(CStr(ps.Cells(r, 10).Value2 & ""))
        If Not IsKnownTargetName(showNm) Then GoTo NextEd
        n = n + 1
        If n > 40 Then Exit For
        m_teRow(n) = r
        m_teNm(n) = showNm
        m_teMd(n) = PickedNum(ps.Cells(r, OVR_MD).Value2, CDbl(vMd))
        m_teInc(n) = PickedNum(ps.Cells(r, OVR_INC).Value2, val(ps.Cells(r, 2).Value2 & ""))
        m_teAzm(n) = PickedNum(ps.Cells(r, OVR_AZM).Value2, val(ps.Cells(r, 3).Value2 & ""))
        m_teTvd(n) = PickedNum(ps.Cells(r, OVR_TVD).Value2, val(ps.Cells(r, 4).Value2 & ""))
NextEd:
    Next r
    m_teN = n
End Sub

Public Function TargetEditorCount() As Long
    LoadTargetEditor
    TargetEditorCount = m_teN
End Function

Public Sub TargetEditorItem(ByVal i As Long, ByRef planRow As Long, ByRef nm As String, _
        ByRef md As Double, ByRef incD As Double, ByRef azmD As Double, ByRef tvdD As Double)
    planRow = m_teRow(i)
    nm = m_teNm(i)
    md = m_teMd(i)
    incD = m_teInc(i)
    azmD = m_teAzm(i)
    tvdD = m_teTvd(i)
End Sub

Public Function TargetNameOk(ByVal nm As String) As Boolean
    TargetNameOk = IsKnownTargetName(UCase$(Trim$(nm)))
End Function

' Writes one named station. Returns a message when the row cannot be saved.
Public Function TargetEditorStore(ByVal planRow As Long, ByVal nm As String, _
        ByVal md As Double, ByVal incD As Double, ByVal azmD As Double, _
        ByVal tvdD As Double) As String
    Dim ps As Worksheet
    Dim wasProt As Boolean
    Dim showNm As String

    TargetEditorStore = ""
    showNm = UCase$(Trim$(nm))
    If Not IsKnownTargetName(showNm) Then
        TargetEditorStore = "Name must be one of: " & TGT_NAMES
        Exit Function
    End If
    If azmD >= 360# And azmD < 360.001 Then azmD = 0#
    If md <= 0# Or incD < 0# Or incD > 180# Or azmD < 0# Or azmD >= 360# Or tvdD < 0# Then
        TargetEditorStore = showNm & ": MD, INC (0-180), AZM (0-360) and TVD have to be in range."
        Exit Function
    End If
    If Not PlanSecSheetExists() Then
        TargetEditorStore = "Plan sections are not loaded."
        Exit Function
    End If
    Set ps = ThisWorkbook.Worksheets(PLANSEC_SHEET)
    If planRow < 3 Then
        TargetEditorStore = "That target row is not on the plan."
        Exit Function
    End If
    wasProt = SheetUnprotectForVba(ps)
    On Error GoTo Fail
    ps.Cells(planRow, 12).value = showNm
    WriteTargetOverride ps, planRow, OVR_MD, 1, md
    WriteTargetOverride ps, planRow, OVR_INC, 2, incD
    WriteTargetOverride ps, planRow, OVR_AZM, 3, azmD
    WriteTargetOverride ps, planRow, OVR_TVD, 4, tvdD
    SheetReprotectAfterVba ps, wasProt
    Exit Function
Fail:
    TargetEditorStore = "Could not save " & showNm & "."
    On Error Resume Next
    SheetReprotectAfterVba ps, wasProt
End Function

' Rebuild the Slidesheet window and the names Cont DI reads, from the saved rows.
Public Sub TargetEditorCommit()
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail
    SyncPlanTargetWindow
    On Error Resume Next
    ThisWorkbook.Worksheets(MDL_ContDI.SH_CONTDI).Calculate
    On Error GoTo Fail
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    Application.EnableEvents = prevEvents
End Sub

' ================================================================================
'  SURVEY ROW FORMULA REPAIR
'
'  Every survey row (14:505) carries the same calculation formulas down the
'  columns (TVD, DLS, build/turn, toolface, cumulative N/E, bit projection ...).
'  Cutting, dragging or overwriting F:G on a row makes Excel re-point those
'  formulas: they end up reading G for Inc (or #REF!), and TVD / N / E, and the
'  proximity gauge that reads them, are wrong from that row down.
'
'  Repair = for each calculation column, take the formula the column uses on
'  >= 90 % of its rows and put it back on the rows that differ. Only cells that
'  already hold a formula are written; typed values (D, F, G, T, U ...) are never
'  touched.
' ================================================================================

Private Function RepairColumns() As Variant
    RepairColumns = Array("B", "C", "E", "H", "I", "J", "K", "L", "N", "O", "P", "Q", _
                          "R", "S", "V", "W", "X", "AA", "AB", "AI", "AJ", "AK", "AL", _
                          "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT")
End Function

' Returns the number of formula cells put back. Interactive unless silent.
Public Function RepairSurveyRowFormulas(Optional ByVal silent As Boolean = False) As Long
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean, prevCalc As XlCalculation, prevScreen As Boolean
    Dim cols As Variant, col As Variant
    Dim arr As Variant
    Dim r As Long, c As Long, n As Long, fixed As Long
    Dim counts As Object, modeF As String, modeN As Long, total As Long
    Dim k As Variant, f As String
    Dim rowsFixed As Object
    Dim sample As String

    If m_repairBusy Then Exit Function
    m_repairBusy = True
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    On Error GoTo Clean

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    wasProt = SheetUnprotectForVba(ss)

    Set rowsFixed = CreateObject("Scripting.Dictionary")
    cols = RepairColumns()
    For Each col In cols
        c = ss.Range(CStr(col) & "1").Column
        arr = ss.Range(ss.Cells(Y_DATA_FIRST + 1, c), ss.Cells(Y_LAST, c)).formula
        Set counts = CreateObject("Scripting.Dictionary")
        total = 0
        For n = 1 To UBound(arr, 1)
            If Left$(CStr(arr(n, 1)), 1) = "=" Then
                f = ss.Cells(Y_DATA_FIRST + n, c).FormulaR1C1
                counts(f) = CLng(counts(f)) + 1
                total = total + 1
            End If
        Next n
        modeF = "": modeN = 0
        For Each k In counts.keys
            If counts(k) > modeN Then modeN = counts(k): modeF = CStr(k)
        Next k
        If total >= 20 And modeN >= 0.9 * total Then
            For n = 1 To UBound(arr, 1)
                If Left$(CStr(arr(n, 1)), 1) = "=" Then
                    r = Y_DATA_FIRST + n
                    If ss.Cells(r, c).FormulaR1C1 <> modeF Then
                        If Len(sample) < 120 Then sample = sample & CStr(col) & r & " "
                        ss.Cells(r, c).FormulaR1C1 = modeF
                        fixed = fixed + 1
                        rowsFixed(r) = True
                    End If
                End If
            Next n
        End If
    Next col

Clean:
    Dim errTxt As String
    If Err.Number <> 0 Then errTxt = Err.Number & " " & Err.Description
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.Calculation = prevCalc          ' settle recalc of the repaired cells
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    m_repairBusy = False
    RepairSurveyRowFormulas = fixed
    If Len(errTxt) > 0 Then
        If Not silent Then MsgBox "Survey formula repair stopped: " & errTxt, vbExclamation, "Repair survey formulas"
        Exit Function
    End If
    If fixed > 0 Then
        RefreshSlideComments
        MDL_PlanGauge.RenderPlanGauge
    End If
    If Not silent Then
        If fixed = 0 Then
            MsgBox "No damaged survey formulas found.", vbInformation, "Repair survey formulas"
        Else
            Dim rl As String
            For Each k In rowsFixed.keys
                If Len(rl) < 90 Then rl = rl & k & " "
            Next k
            MsgBox "Put back " & fixed & " formula cells on " & rowsFixed.Count & _
                   " rows (" & Trim$(rl) & IIf(Len(rl) >= 90, " ...", "") & ")." & vbCrLf & _
                   "TVD, N/E and the gauge are recalculated.", vbInformation, "Repair survey formulas"
        End If
    End If
End Function

' True when row r's TVD does not follow from its own Inc and the row above
' (a shifted reference) or any calculation cell on it holds an error. Reads
' values only, so it works on the protected sheet.
Private Function SurveyRowDamaged(ByVal ss As Worksheet, ByVal r As Long) As Boolean
    Dim e0 As Variant, e1 As Variant, i0 As Variant, i1 As Variant
    Dim h0 As Variant, h1 As Variant, c As Variant
    SurveyRowDamaged = False
    If r <= Y_DATA_FIRST Or r > Y_LAST Then Exit Function
    e1 = ss.Cells(r, "E").Value2: e0 = ss.Cells(r - 1, "E").Value2
    i1 = ss.Cells(r, "F").Value2: i0 = ss.Cells(r - 1, "F").Value2
    h1 = ss.Cells(r, "H").Value2: h0 = ss.Cells(r - 1, "H").Value2
    For Each c In Array("H", "J", "K", "AL", "AM")
        If isError(ss.Cells(r, CStr(c)).Value2) Then
            If Not isError(i1) Then
                If IsNumeric(i1) And Len(CStr(i1 & "")) > 0 Then SurveyRowDamaged = True: Exit Function
            End If
        End If
    Next c
    If isError(e1) Or isError(e0) Or isError(i1) Or isError(i0) Or isError(h1) Or isError(h0) Then Exit Function
    If Not (IsNumeric(e1) And IsNumeric(e0) And IsNumeric(i1) And IsNumeric(i0) _
            And IsNumeric(h1) And IsNumeric(h0)) Then Exit Function
    If Len(CStr(e1 & "")) = 0 Or Len(CStr(e0 & "")) = 0 Or Len(CStr(i1 & "")) = 0 _
       Or Len(CStr(i0 & "")) = 0 Or Len(CStr(h1 & "")) = 0 Or Len(CStr(h0 & "")) = 0 Then Exit Function
    If CDbl(e1) <= CDbl(e0) Then Exit Function
    Dim expect As Double
    expect = Cos(((CDbl(i1) + CDbl(i0)) / 2#) * 3.14159265358979 / 180#) * (CDbl(e1) - CDbl(e0)) + CDbl(h0)
    If Abs(CDbl(h1) - expect) > 0.05 Then SurveyRowDamaged = True
End Function

' Called after Inc / Azm are typed or pasted. Damage shows on the entered rows and
' the next one down (that row reads this row's F:G). Offers the repair once per
' session; "No" keeps it quiet until the workbook is reopened.
Public Sub CheckSurveyRowsAfterEntry(ByVal Target As Range)
    Dim ss As Worksheet
    Dim r As Long, r1 As Long, r2 As Long
    Dim Bad As Long
    If m_repairDeclined Or m_repairBusy Then Exit Sub
    On Error GoTo Done
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    r1 = Target.row
    r2 = Target.row + Target.rows.Count - 1
    If r2 - r1 > 40 Then r2 = r1 + 40
    For r = r1 To r2 + 1
        If SurveyRowDamaged(ss, r) Then Bad = r: Exit For
    Next r
    If Bad = 0 Then GoTo Done
    If MsgBox("The survey formulas on row " & Bad & " are damaged (a cell was moved, cut or" & _
              " dragged), so TVD, N/E and the proximity gauge are wrong from there down." & _
              vbCrLf & vbCrLf & "Put the formulas back now? Your typed surveys are not changed.", _
              vbYesNo + vbExclamation, "Slidesheet survey formulas") = vbYes Then
        RepairSurveyRowFormulas
    Else
        m_repairDeclined = True
    End If
Done:
End Sub






