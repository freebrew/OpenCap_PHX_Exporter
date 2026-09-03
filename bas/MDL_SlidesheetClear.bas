Attribute VB_Name = "MDL_SlidesheetClear"
Option Explicit

Private Const SS_SHEET As String = "Slidesheet"
Private Const BTN_NAME As String = "btnClearSlidesheetRanges"
Private Const BTN_CAPTION As String = "Clear Ranges"
' Merged Z1:Z3 - left of the plan proximity gauge which owns merged AA1:AC7
Private Const BTN_ANCHOR As String = "Z1:Z3"
Private Const Y_FIRST As Long = 12
Private Const Y_LAST As Long = 305
Private Const Y_DATA_FIRST As Long = 13
Private Const Y_LABEL As String = "RKB"
' Sheet comment yellow (same as Module16 pale yellow). Not Excel idx-6 65535.
Private Const COMMENT_YELLOW As Long = 13434879  ' RGB(255,255,204)
Private Const TGT_FIRST As Long = 2
Private Const TGT_LAST As Long = 5
Private Const PLANSEC_SHEET As String = "_OC_PlanSec"
Private Const TGT_NAMES As String = "NUDGE,VERTICAL,KOP,TANGENT,SOT,EOT,HEEL,TD"
' Full named-target list for BURR / slide UDFs (T2:Y5 is display-only).
Private Const FULL_TGT_MD_COL As Long = 14
Private Const FULL_TGT_FIRST As Long = 3
Private Const FULL_TGT_ROWS As Long = 80
Private m_updatingY As Boolean
Private m_hlTgtRow As Long   ' last highlighted target row (0 = none)
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
'   D,F,G,T,U rows 13:305
'   Sail waypoints AC14:AD33 (display AA=Inc Next, AB=Geo Window; helper AE)
'   Targets U2:X5
'   Comments Y12:Y305 (then restore RKB + auto slide comments)
Public Sub ClearSlidesheetRanges()
    Dim ss As Worksheet

    On Error GoTo Fail

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ss.Range("D13:D305").ClearContents
    ss.Range("F13:F305").ClearContents
    ss.Range("G13:G305").ClearContents
    ss.Range("T13:T305").ClearContents
    ss.Range("U13:U305").ClearContents
    ss.Range("AC14:AD33").ClearContents
    ss.Range("U2:X5").ClearContents
    ss.Range("Y12:Y305").ClearContents
    ss.Range("Z13:Z305").ClearContents
    SyncPlanTargetWindowOnSheet ss
    ss.Range("Y12").Value2 = Y_LABEL

    Application.EnableEvents = True
    RefreshSlideComments forceAll:=True
    HighlightActiveTarget

    Application.ScreenUpdating = True

    MsgBox "Cleared:" & vbCrLf & _
           "  D/F/G/T/U 13:305" & vbCrLf & _
           "  Sail waypoints AC14:AD33" & vbCrLf & _
           "  Targets U2:X5" & vbCrLf & _
           "  Comments Y12:Y305 / ROT Z13:Z305 (auto text restored where inputs exist)", _
           vbInformation
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
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

    Set area = Intersect(Target, ss.Range("D13:G305,T13:U305"))
    If area Is Nothing Then Exit Function

    Set seen = CreateObject("Scripting.Dictionary")
    For Each cell In area.Cells
        r = cell.Row
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

' Rewrite Z on every stand row (C − T). Used after the formula change so
' existing leftover-rotate values catch up without a survey keystroke.
Public Sub RefreshAllRotateMetres()
    Dim ss As Worksheet
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    RefreshRotateMetres ss.Range("C13:C305")
End Sub

' Z = C − T on the same row. C/D/T edits call this without a full Y rebuild.
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
    Set area = Intersect(Target, ss.Range("C13:D305,T13:T305"))
    If area Is Nothing Then Exit Sub

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    Set seen = CreateObject("Scripting.Dictionary")
    For Each cell In area.Cells
        r = cell.Row
        If r >= Y_DATA_FIRST And r <= Y_LAST Then
            If Not seen.Exists(r) Then
                seen.Add r, True
                WriteRotateMetresOnSheet ss, r
            End If
        End If
    Next cell

Clean:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

' Same-row leftover rotate: course (C) minus already-slid (T). Blank T = 0.
' Leaves a hand-typed Z alone. Clamps at 0 when T > C.
Private Sub WriteRotateMetresOnSheet(ByVal ss As Worksheet, ByVal r As Long)
    Dim course As Double
    Dim slid As Double
    Dim remain As Double
    Dim rotTxt As String
    Dim zCell As Range
    Dim hasC As Boolean

    Set zCell = ss.Cells(r, "Z")
    If Not IsAutoRotText(Trim$(CStr(zCell.Value2 & ""))) Then Exit Sub

    On Error Resume Next
    ss.Cells(r, "C").Calculate
    On Error GoTo 0

    hasC = IsNumberValue(ss.Cells(r, "C").Value2)
    If Not hasC Then
        If Len(Trim$(CStr(zCell.Value2 & ""))) > 0 Then zCell.ClearContents
        Exit Sub
    End If

    course = CDbl(ss.Cells(r, "C").Value2)
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

        If IsNumberValue(arArr(i, 1)) Then
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

        ' Z: leftover rotate = this stand's C minus T (already slid).
        WriteRotateMetresOnSheet ss, r

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

    m_lastRefreshAt = Timer          ' stamp only on a completed refresh

Clean:
    On Error Resume Next
    EndTextMeasure                   ' delete the reusable textbox (needs sheet still unprotected)
    SheetReprotectAfterVba ss, wasProt
    ' Restore calc while events are still off so the settle recalc cannot
    ' re-enter via Worksheet_Calculate.
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    m_updatingY = False
End Sub

' Hidden AR (BURR) / AS (Metres To Slide) must start from the PROJECTED bit
' (W = Inc@Bit, X = Azm@Bit) — the same attitude AT (required TF) uses — not
' the survey station (F / G). Older sheets carry
'   ProjBurr(D13,F13,G13,...) / ProjMetersToSlide(D13,F13,G13,...)
' which sized the slide from the survey while the TF was aimed from the bit,
' so Y paired metres and toolface that belonged to two different manoeuvres.
' Rewrites only the F{r},G{r} argument pair; every other token is preserved.
'
' W / X (Inc@Bit / Azm@Bit) used to pass IF(AK="",0,AK): a blank dial became
' TF 0 = high side, so a rotary stand was projected as a full build with the
' demonstrated N/O walk. Pass AK through so ProjIncAtBit / ProjAzmAtBit see
' the blank and project pure rotary (bit attitude = survey attitude).
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

    On Error GoTo Done
    For Each col In Array("AR", "AS", "W", "X")
        Set rng = ss.Range(ss.Cells(Y_DATA_FIRST, col), ss.Cells(Y_LAST, col))
        arr = rng.Formula
        dirty = False
        For i = 1 To UBound(arr, 1)
            r = Y_DATA_FIRST + i - 1
            f = CStr(arr(i, 1) & "")
            If Len(f) > 1 And Left$(f, 1) = "=" Then
                If col = "W" Or col = "X" Then
                    oldTok = "IF(AK" & r & "="""",0,AK" & r & ")"
                    newTok = "AK" & r
                Else
                    oldTok = "(D" & r & ",F" & r & ",G" & r & ","
                    newTok = "(D" & r & ",W" & r & ",X" & r & ","
                End If
                If InStr(1, f, oldTok, vbBinaryCompare) > 0 Then
                    arr(i, 1) = Replace(f, oldTok, newTok)
                    dirty = True
                End If
            End If
        Next i
        If dirty Then rng.Formula = arr
    Next col
Done:
End Sub

' Pad left/right text with spaces so the right part sits on the cell's right edge.
' Uses point widths (font-accurate), not ColumnWidth character units.
Private Function FitLeftRightInCell(ByVal leftPart As String, ByVal rightPart As String, _
                                    ByVal cell As Range, ByVal spacePts As Double) As String
    Dim avail As Double
    Dim leftPts As Double
    Dim rightPts As Double
    Dim PAD As Long
    ' Excel keeps a small built-in cell pad; 2pt leaves BURR flush to the right grid.
    Const CELL_MARGIN_PTS As Double = 2#

    leftPts = MeasureTextWidthPts(leftPart, cell)
    rightPts = MeasureTextWidthPts(rightPart, cell)
    avail = cell.Width - CELL_MARGIN_PTS
    If spacePts <= 0# Then spacePts = 1#

    PAD = CLng(Application.WorksheetFunction.Round((avail - leftPts - rightPts) / spacePts, 0))
    If PAD < 2 Then PAD = 2
    FitLeftRightInCell = leftPart & Space$(PAD) & rightPart
End Function

' Sliding left, BURR, then leftover-rotate flush to the right grid.
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
    avail = cell.Width - CELL_MARGIN_PTS
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
    sz = proto.Font.Size
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
        .Font.Size = sz
        .Font.bold = bld
    End With
    MeasureTextWidthPts = shp.Width
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
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
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

    ' Use the last surveyed row (F filled), not the last bit-only tally row.
    ' Bit-only depths past a target MD would otherwise advance the highlight
    ' while BURR is still computed against the prior survey's aim (e.g. TAR2 / 8.37).
    lastSurvRow = 0
    bitMd = 0#
    aimMd = 0#
    For r = 12 To 305
        vF = ss.Cells(r, "F").Value2
        If IsNumeric(vF) Then
            If Len(Trim$(CStr(vF & ""))) > 0 Then
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
    ' Old pipe format or new padded "Sliding ... BURR 0.00" format.
    If s Like "Sliding *| BURR *" Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    If (Left$(s, 8) = "Sliding ") And (InStr(1, s, "BURR ", vbTextCompare) > 0) Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    ' Leftover "Nm rot" that used to live in Y — clear those back to slide/BURR only.
    IsAutoSlideCommentText = (InStr(1, s, "m rot", vbTextCompare) > 0)
End Function

Private Function IsAutoRotText(ByVal t As String) As Boolean
    Dim s As String
    s = Trim$(t)
    If Len(s) = 0 Then
        IsAutoRotText = True
        Exit Function
    End If
    If (right$(s, 5) = "m ROT") Or (right$(s, 5) = "m rot") Then
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

    Set btn = ss.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    btn.name = BTN_NAME
    btn.caption = BTN_CAPTION
    btn.OnAction = "'" & ThisWorkbook.name & "'!ClearSlidesheetRanges"

    Set shp = ss.Shapes(BTN_NAME)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.Size = 10
    On Error GoTo 0
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

    changed = False
    For vis = 0 To 3
        r = TGT_FIRST + vis
        If startI + vis <= n Then
            i = startI + vis
            newMd = md(i): newInc = inc(i): newAzm = azm(i)
            newTvd = tvd(i): newNm = nm(i)
        Else
            newMd = "": newInc = "": newAzm = ""
            newTvd = "": newNm = ""
        End If
        If Not SameTgtCell(ss.Cells(r, "U").Value2, newMd) Then
            ss.Cells(r, "U").Value2 = newMd
            changed = True
        End If
        If Not SameTgtCell(ss.Cells(r, "V").Value2, newInc) Then
            ss.Cells(r, "V").Value2 = newInc
            changed = True
        End If
        If Not SameTgtCell(ss.Cells(r, "W").Value2, newAzm) Then
            ss.Cells(r, "W").Value2 = newAzm
            changed = True
        End If
        If Not SameTgtCell(ss.Cells(r, "X").Value2, newTvd) Then
            ss.Cells(r, "X").Value2 = newTvd
            changed = True
        End If
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
    Dim startRef As String
    Dim lastR As Long

    lastR = FULL_TGT_FIRST + FULL_TGT_ROWS - 1
    mdRef = "='" & PLANSEC_SHEET & "'!$N$" & FULL_TGT_FIRST & ":$N$" & lastR
    incRef = "='" & PLANSEC_SHEET & "'!$O$" & FULL_TGT_FIRST & ":$O$" & lastR
    azmRef = "='" & PLANSEC_SHEET & "'!$P$" & FULL_TGT_FIRST & ":$P$" & lastR
    tvdRef = "='" & PLANSEC_SHEET & "'!$Q$" & FULL_TGT_FIRST & ":$Q$" & lastR
    startRef = "='" & PLANSEC_SHEET & "'!$N$1"

    SetWbName "ProjTargets_MD", mdRef
    SetWbName "ProjTargets_INC", incRef
    SetWbName "ProjTargets_AZM", azmRef
    SetWbName "ProjTargets_TVD", tvdRef
    SetWbName "ProjBuildStartMd", startRef
End Sub

Private Sub SetWbName(ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    ThisWorkbook.names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.names.Add name:=nm, refersTo:=refersTo
End Sub

' N1 = first KOP/TANGENT MD (BURR start). N3:Q = every named plan station.
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

    lastR = FULL_TGT_FIRST + FULL_TGT_ROWS - 1
    ps.Range(ps.Cells(FULL_TGT_FIRST, 14), ps.Cells(lastR, 17)).ClearContents
    For i = 1 To n
        If i > FULL_TGT_ROWS Then Exit For
        ps.Cells(FULL_TGT_FIRST + i - 1, 14).Value2 = md(i)
        ps.Cells(FULL_TGT_FIRST + i - 1, 15).Value2 = inc(i)
        ps.Cells(FULL_TGT_FIRST + i - 1, 16).Value2 = azm(i)
        ps.Cells(FULL_TGT_FIRST + i - 1, 17).Value2 = tvd(i)
    Next i

    SheetReprotectAfterVba ps, wasProt
End Sub

Private Function SameTgtCell(ByVal cur As Variant, ByVal neu As Variant) As Boolean
    If IsNumeric(cur) And IsNumeric(neu) Then
        SameTgtCell = (Abs(CDbl(cur) - CDbl(neu)) < 0.005)
        Exit Function
    End If
    SameTgtCell = (Trim$(CStr(cur & "")) = Trim$(CStr(neu & "")))
End Function

Private Function LastSurveyBitMd(ByVal ss As Worksheet) As Double
    Dim r As Long
    Dim vF As Variant, vD As Variant
    LastSurveyBitMd = 0#
    For r = 12 To 305
        vF = ss.Cells(r, "F").Value2
        If IsNumeric(vF) Then
            If Len(Trim$(CStr(vF & ""))) > 0 Then
                vD = ss.Cells(r, "D").Value2
                If IsNumeric(vD) Then LastSurveyBitMd = CDbl(vD)
            End If
        End If
    Next r
End Function

Private Function LoadNamedPlanTargets(ByVal ps As Worksheet, _
        ByRef md() As Double, ByRef inc() As Double, ByRef azm() As Double, _
        ByRef tvd() As Double, ByRef nm() As String) As Long

    Dim lastR As Long
    Dim r As Long
    Dim n As Long
    Dim autoNm As String, userNm As String, showNm As String
    Dim vMd As Variant

    lastR = ps.Cells(ps.Rows.Count, 1).End(xlUp).Row
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
        If Not IsKnownTargetName(showNm) Then GoTo NextPs
        n = n + 1
        If n > 80 Then Exit For
        md(n) = CDbl(vMd)
        inc(n) = val(ps.Cells(r, 2).Value2 & "")
        azm(n) = val(ps.Cells(r, 3).Value2 & "")
        tvd(n) = val(ps.Cells(r, 4).Value2 & "")
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
                      Or s = "NUDGE" Or s = "VERTICAL")
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

' Persist a Y2:Y5 dropdown change onto _OC_PlanSec USERNAME (matched by MD).
Public Sub RememberPlanTargetOverride(ByVal Target As Range)
    Dim ss As Worksheet
    Dim ps As Worksheet
    Dim cell As Range
    Dim tgtMd As Variant
    Dim lastR As Long
    Dim r As Long
    Dim nm As String

    On Error GoTo Fail
    If m_syncingTgt Then Exit Sub
    If Target Is Nothing Then Exit Sub
    Set ss = Target.Worksheet
    If Not PlanSecSheetExists() Then Exit Sub
    Set ps = ThisWorkbook.Worksheets(PLANSEC_SHEET)
    lastR = ps.Cells(ps.Rows.Count, 1).End(xlUp).Row

    For Each cell In Target.Cells
        If cell.Column = 25 And cell.Row >= TGT_FIRST And cell.Row <= TGT_LAST Then
            tgtMd = ss.Cells(cell.Row, "U").Value2
            nm = UCase$(Trim$(CStr(cell.Value2 & "")))
            If Not IsNumeric(tgtMd) Then GoTo NextOv
            For r = 3 To lastR
                If IsNumeric(ps.Cells(r, 1).Value2) Then
                    If Abs(CDbl(ps.Cells(r, 1).Value2) - CDbl(tgtMd)) < 0.05 Then
                        If nm = "" Then
                            ps.Cells(r, 12).ClearContents
                        Else
                            ps.Cells(r, 12).Value = nm
                        End If
                        Exit For
                    End If
                End If
            Next r
        End If
NextOv:
    Next cell
    Exit Sub
Fail:
End Sub














