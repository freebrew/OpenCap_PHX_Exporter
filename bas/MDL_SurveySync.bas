Attribute VB_Name = "MDL_SurveySync"
Option Explicit

' ================================================================================
'  MDL_SurveySync
'
'  Audit & sync Slidesheet surveys (E/F/G: MD, Inc, Azi) against a Well Seeker X
'  Survey Report export (.xlsx / .xlsm / .xls / .csv).
'
'  Safety rules:
'    - Audit-first: shows every differing row and asks Yes/No before writing.
'    - Wrong-well guard: refuses to sync if < 50% of Slidesheet survey MDs match.
'    - Interpolated formation/casing markers in Well Seeker X (blank DLS) and
'      tie-on rows (<= E12) / trailing PTB rows are skipped automatically.
'    - Never inserts/deletes Slidesheet rows; updates ONLY E (if MD differs),
'      F (Inc), and G (Azi) on matched existing rows. Slides, bit depths,
'      formulas (H/AL/AM), colors, and notes remain untouched.
' ================================================================================

Private Const SS_SHEET As String = "Slidesheet"
Private Const SETUP_SHEET As String = "Setup"
Private Const DATA_SHEET As String = "Data"
Private Const SURV_FIRST As Long = 13
Private Const SURV_LAST As Long = 520
Private Const TOL_EXACT_MD As Double = 0.05
Private Const TOL_NEAR_MD As Double = 0.5
Private Const TOL_DIFF As Double = 0.005

' Button entry point on Setup ("Sync Surveys").
Public Sub SyncWellSeekerSurveys()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select Well Seeker X Survey Report (.xlsx or .csv)"
    fd.Filters.Clear
    fd.Filters.Add "Well Seeker X Survey Report (*.xlsx;*.xlsm;*.xls;*.csv)", "*.xlsx;*.xlsm;*.xls;*.csv", 1
    fd.Filters.Add "Excel Files (*.xlsx;*.xlsm;*.xls)", "*.xlsx;*.xlsm;*.xls", 2
    fd.Filters.Add "CSV Files (*.csv)", "*.csv", 3
    fd.Filters.Add "All Files (*.*)", "*.*", 4
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub

    AuditOrSyncWellSeekerFile fd.SelectedItems(1), True
End Sub

' Testable entry point. When promptApply = False, audits only and returns a summary
' without displaying a MsgBox or modifying any cells.
Public Function AuditOrSyncWellSeekerFile(ByVal fPath As String, _
        Optional ByVal promptApply As Boolean = True) As String
    Dim wsMD() As Double, wsInc() As Double, wsAzi() As Double
    Dim nWs As Long, nInterpSkipped As Long
    Dim tieMdWs As Double, hasTieWs As Boolean
    Dim errMsg As String
    Dim fileName As String

    AuditOrSyncWellSeekerFile = ""
    fPath = Trim$(fPath & "")
    If Len(fPath) = 0 Then Exit Function
    fileName = BaseName(fPath)

    Dim ss As Worksheet
    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    On Error GoTo 0
    If ss Is Nothing Then
        AuditOrSyncWellSeekerFile = "ERROR: Slidesheet not found."
        If promptApply Then MsgBox AuditOrSyncWellSeekerFile, vbExclamation, "Sync Surveys"
        Exit Function
    End If

    ' Tie-on threshold from Row 12 (E12, e.g. 5.55 m KBE); default 10 m if blank.
    Dim tieThreshold As Double
    tieThreshold = 10#
    If IsNumeric(ss.Cells(SURV_FIRST - 1, 5).Value2) And Len(CStr(ss.Cells(SURV_FIRST - 1, 5).Value2 & "")) > 0 Then
        tieThreshold = CDbl(ss.Cells(SURV_FIRST - 1, 5).Value2) + 0.05
    End If

    Dim ext As String
    ext = LCase$(mid$(fPath, InStrRev(fPath, ".") + 1))
    If ext = "csv" Or ext = "txt" Then
        nWs = ReadWellSeekerCsv(fPath, tieThreshold, wsMD, wsInc, wsAzi, nInterpSkipped, tieMdWs, hasTieWs, errMsg)
    Else
        nWs = ReadWellSeekerXlsx(fPath, tieThreshold, wsMD, wsInc, wsAzi, nInterpSkipped, tieMdWs, hasTieWs, errMsg)
    End If

    If Len(errMsg) > 0 Or nWs < 1 Then
        If Len(errMsg) = 0 Then errMsg = "No survey rows found in " & fileName & " (expected MD, INC, AZI columns)."
        AuditOrSyncWellSeekerFile = "ERROR: " & errMsg
        If promptApply Then MsgBox errMsg, vbExclamation, "Sync Surveys"
        Exit Function
    End If

    ' Load Slidesheet survey rows (13..520, skipping LOOKUP summary rows)
    Dim ssRow() As Long, ssMD() As Double, ssInc() As Double, ssAzi() As Double
    Dim nSs As Long
    nSs = LoadSlidesheetSurveys(ss, ssRow, ssMD, ssInc, ssAzi)
    If nSs < 1 Then
        AuditOrSyncWellSeekerFile = "Slidesheet has no entered surveys yet (" & nWs & " stations in " & fileName & ")."
        If promptApply Then MsgBox AuditOrSyncWellSeekerFile, vbInformation, "Sync Surveys"
        Exit Function
    End If

    ' Match Slidesheet rows 1-to-1 against Well Seeker stations:
    '   Pass 1: exact MD within +-0.05 m
    '   Pass 2: nearest unmatched MD within +-0.50 m (catches 0.01-0.50 m MD typos)
    Dim ssMatch() As Long, wsUsed() As Boolean
    ReDim ssMatch(0 To nSs - 1)
    ReDim wsUsed(0 To nWs - 1)
    Dim i As Long, j As Long, bestJ As Long
    Dim d As Double, bestD As Double, nMatched As Long
    For i = 0 To nSs - 1: ssMatch(i) = -1: Next i

    For i = 0 To nSs - 1
        bestJ = -1: bestD = TOL_EXACT_MD + 0.0001
        For j = 0 To nWs - 1
            If Not wsUsed(j) Then
                d = Abs(ssMD(i) - wsMD(j))
                If d < bestD Then bestD = d: bestJ = j
            End If
        Next j
        If bestJ >= 0 Then
            ssMatch(i) = bestJ: wsUsed(bestJ) = True: nMatched = nMatched + 1
        End If
    Next i

    For i = 0 To nSs - 1
        If ssMatch(i) < 0 Then
            bestJ = -1: bestD = TOL_NEAR_MD + 0.0001
            For j = 0 To nWs - 1
                If Not wsUsed(j) Then
                    d = Abs(ssMD(i) - wsMD(j))
                    If d < bestD Then bestD = d: bestJ = j
                End If
            Next j
            If bestJ >= 0 Then
                ssMatch(i) = bestJ: wsUsed(bestJ) = True: nMatched = nMatched + 1
            End If
        End If
    Next i

    ' Wrong-well safety guard: require at least 50% of Slidesheet stations to match.
    If nMatched * 2 < nSs Then
        errMsg = "Only " & nMatched & " of " & nSs & " Slidesheet survey depths matched " & fileName & "." & vbCrLf & _
                 "Check that you selected the right well's Well Seeker X export. No cells were changed."
        AuditOrSyncWellSeekerFile = "BLOCKED: " & nMatched & "/" & nSs & " matched"
        If promptApply Then MsgBox errMsg, vbExclamation, "Sync Surveys — Well Mismatch"
        Exit Function
    End If

    ' Collect diffs + detect paste shifts (Inc/Azi matching adjacent Well Seeker row)
    Dim diffIdx() As Long, nDiff As Long, nShiftHits As Long
    Dim shiftFirstRow As Long, shiftLastRow As Long
    ReDim diffIdx(0 To nSs - 1)
    nDiff = 0: nShiftHits = 0: shiftFirstRow = 0: shiftLastRow = 0

    For i = 0 To nSs - 1
        j = ssMatch(i)
        If j >= 0 Then
            If Abs(ssMD(i) - wsMD(j)) > TOL_DIFF _
                    Or Abs(ssInc(i) - wsInc(j)) > TOL_DIFF _
                    Or Abs(ssAzi(i) - wsAzi(j)) > TOL_DIFF Then
                diffIdx(nDiff) = i
                nDiff = nDiff + 1
                If IsShiftedFromNeighbor(ssInc(i), ssAzi(i), j, nWs, wsInc, wsAzi) Then
                    nShiftHits = nShiftHits + 1
                    If shiftFirstRow = 0 Then shiftFirstRow = ssRow(i)
                    shiftLastRow = ssRow(i)
                End If
            End If
        End If
    Next i

    ' Count deeper Well Seeker stations not yet on Slidesheet (reported only)
    Dim maxSsMd As Double: maxSsMd = ssMD(nSs - 1)
    Dim nDeeperWs As Long: nDeeperWs = 0
    For j = 0 To nWs - 1
        If Not wsUsed(j) And wsMD(j) > maxSsMd + TOL_EXACT_MD Then nDeeperWs = nDeeperWs + 1
    Next j

    ' Check Row 12 tie-on note (informational)
    Dim tieNote As String: tieNote = ""
    If hasTieWs Then
        Dim ssTie As Double: ssTie = 0#
        If IsNumeric(ss.Cells(SURV_FIRST - 1, 5).Value2) And Len(CStr(ss.Cells(SURV_FIRST - 1, 5).Value2 & "")) > 0 Then
            ssTie = CDbl(ss.Cells(SURV_FIRST - 1, 5).Value2)
        End If
        If Abs(ssTie - tieMdWs) > 0.05 Then
            tieNote = vbCrLf & "Note: Tie-on Row 12 (E12=" & Format$(ssTie, "0.00") & "m) differs from Well Seeker KBE (" & Format$(tieMdWs, "0.00") & "m)."
        End If
    End If

    If nDiff = 0 Then
        AuditOrSyncWellSeekerFile = "OK: 0 diffs across " & nMatched & "/" & nSs & " surveys (" & fileName & ")"
        If nDeeperWs > 0 Then AuditOrSyncWellSeekerFile = AuditOrSyncWellSeekerFile & "; " & nDeeperWs & " deeper in WS"
        If promptApply Then
            WriteSyncStatusToSetup fileName & "  [OK: 0 diffs @ " & Format$(Now, "dd-mmm hh:nn") & "]"
            Dim okMsg As String
            okMsg = "Slidesheet matches Well Seeker X across all " & nMatched & " matched surveys (max delta <= 0.005)."
            If nDeeperWs > 0 Then
                okMsg = okMsg & vbCrLf & vbCrLf & "Well Seeker X has " & nDeeperWs & " deeper survey(s) past " & Format$(maxSsMd, "0.00") & " m not yet on Slidesheet."
            End If
            If Len(tieNote) > 0 Then okMsg = okMsg & vbCrLf & tieNote
            MsgBox okMsg, vbInformation, "Sync Surveys — Clean Match"
        End If
        Exit Function
    End If

    ' Build preview text for audit summary / confirmation dialog
    Dim msg As String
    msg = "File: " & fileName & vbCrLf & _
          "Matched: " & nMatched & " / " & nSs & " Slidesheet surveys   |   Differences: " & nDiff & " row(s)"
    If nShiftHits >= 2 Then
        msg = msg & vbCrLf & "WARNING: Paste-shift detected around rows " & shiftFirstRow & "-" & shiftLastRow & " (" & nShiftHits & " shifted rows)!"
    End If
    msg = msg & vbCrLf & vbCrLf

    Dim maxShow As Long: maxShow = 16
    Dim k As Long, idx As Long, wIdx As Long
    For k = 0 To nDiff - 1
        If k >= maxShow Then
            msg = msg & "... and " & (nDiff - maxShow) & " more row(s)." & vbCrLf
            Exit For
        End If
        idx = diffIdx(k): wIdx = ssMatch(idx)
        msg = msg & "Row " & ssRow(idx) & " (" & Format$(wsMD(wIdx), "0.00") & "m): " & _
              Format$(ssInc(idx), "0.00") & "/" & Format$(ssAzi(idx), "0.00") & " -> " & _
              Format$(wsInc(wIdx), "0.00") & "/" & Format$(wsAzi(wIdx), "0.00")
        If Abs(ssMD(idx) - wsMD(wIdx)) > TOL_DIFF Then
            msg = msg & "  [MD " & Format$(ssMD(idx), "0.00") & "->" & Format$(wsMD(wIdx), "0.00") & "]"
        End If
        msg = msg & vbCrLf
    Next k

    If nDeeperWs > 0 Then
        msg = msg & vbCrLf & "Note: " & nDeeperWs & " deeper survey(s) in Well Seeker X past " & Format$(maxSsMd, "0.00") & "m (not auto-appended)."
    End If
    If Len(tieNote) > 0 Then msg = msg & tieNote

    AuditOrSyncWellSeekerFile = "DIFFS: " & nDiff & " row(s) across " & nMatched & "/" & nSs & _
                                " surveys (shifted=" & nShiftHits & ", interpSkipped=" & nInterpSkipped & ")"
    If Not promptApply Then Exit Function

    msg = msg & vbCrLf & vbCrLf & _
          "Only columns E/F/G (MD/Inc/Azi) on these " & nDiff & " matched row(s) will be updated." & vbCrLf & _
          "Slides, bit depths, notes, and all formulas (H/AL/AM) stay untouched." & vbCrLf & vbCrLf & _
          "Apply these updates and refresh the proximity gauge now?"

    If MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "Sync Surveys from Well Seeker X") <> vbYes Then
        WriteSyncStatusToSetup fileName & "  [Audit: " & nDiff & " diffs, not applied]"
        Exit Function
    End If

    ' Apply ONLY E (if MD differs), F, G on matched rows with events off.
    Dim prevEvents As Boolean: prevEvents = Application.EnableEvents
    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Dim wasProt As Boolean
    On Error GoTo ApplyFail
    ScreenBeginBusy "Syncing " & nDiff & " survey row(s) from Well Seeker X..."
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)

    For k = 0 To nDiff - 1
        idx = diffIdx(k): wIdx = ssMatch(idx)
        If Abs(ssMD(idx) - wsMD(wIdx)) > TOL_DIFF Then
            ss.Cells(ssRow(idx), 5).value = wsMD(wIdx)
        End If
        ss.Cells(ssRow(idx), 6).value = wsInc(wIdx)
        ss.Cells(ssRow(idx), 7).value = wsAzi(wIdx)
    Next k

    SheetReprotectAfterVba ss, wasProt
    Application.Calculation = prevCalc
    Application.Calculate

    On Error Resume Next
    MDL_PlanGauge.RenderPlanGaugeStrict
    On Error GoTo ApplyFail

    Application.EnableEvents = prevEvents
    ScreenEndBusy

    WriteSyncStatusToSetup fileName & "  [Synced " & nDiff & " rows @ " & Format$(Now, "dd-mmm hh:nn") & "]"

    Dim posSummary As String: posSummary = ReadDataWellboreSummary()
    AuditOrSyncWellSeekerFile = "SYNCED: " & nDiff & " row(s) from " & fileName
    MsgBox "Updated " & nDiff & " survey row(s) on Slidesheet and refreshed the proximity gauge." & _
           IIf(Len(posSummary) > 0, vbCrLf & vbCrLf & posSummary, ""), _
           vbInformation, "Sync Surveys Complete"
    Exit Function

ApplyFail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    ScreenEndBusy
    AuditOrSyncWellSeekerFile = "ERROR: " & Err.Description
    MsgBox "Survey sync failed: " & Err.Description, vbExclamation, "Sync Surveys"
End Function

' True when a Slidesheet row's Inc/Azi differs materially from its own station
' and matches the adjacent Well Seeker station (classic copy/paste row-shift error).
Private Function IsShiftedFromNeighbor(ByVal sInc As Double, ByVal sAzi As Double, _
        ByVal wIdx As Long, ByVal nWs As Long, _
        wsInc() As Double, wsAzi() As Double) As Boolean
    IsShiftedFromNeighbor = False
    If Abs(sInc - wsInc(wIdx)) <= 0.1 And Abs(sAzi - wsAzi(wIdx)) <= 0.2 Then Exit Function
    If wIdx > 0 Then
        If Abs(sInc - wsInc(wIdx - 1)) <= 0.02 And Abs(sAzi - wsAzi(wIdx - 1)) <= 0.05 Then
            IsShiftedFromNeighbor = True
            Exit Function
        End If
    End If
    If wIdx + 1 < nWs Then
        If Abs(sInc - wsInc(wIdx + 1)) <= 0.02 And Abs(sAzi - wsAzi(wIdx + 1)) <= 0.05 Then
            IsShiftedFromNeighbor = True
        End If
    End If
End Function

' Read Slidesheet survey rows 13..520, skipping LOOKUP summary rows.
Private Function LoadSlidesheetSurveys(ws As Worksheet, _
        ByRef rows() As Long, ByRef mds() As Double, _
        ByRef incs() As Double, ByRef azis() As Double) As Long
    Dim cap As Long: cap = SURV_LAST - SURV_FIRST + 1
    ReDim rows(0 To cap): ReDim mds(0 To cap)
    ReDim incs(0 To cap): ReDim azis(0 To cap)
    Dim r As Long, n As Long
    Dim vM As Variant, vI As Variant, vA As Variant
    n = 0
    For r = SURV_FIRST To SURV_LAST
        If IsSummaryRow(ws, r) Then GoTo NextR
        vM = ws.Cells(r, 5).Value2
        vI = ws.Cells(r, 6).Value2
        vA = ws.Cells(r, 7).Value2
        If IsNumeric(vM) And IsNumeric(vI) And IsNumeric(vA) Then
            If Len(CStr(vM & "")) > 0 And Len(CStr(vI & "")) > 0 And Len(CStr(vA & "")) > 0 Then
                rows(n) = r
                mds(n) = CDbl(vM)
                incs(n) = CDbl(vI)
                azis(n) = CDbl(vA)
                n = n + 1
            End If
        End If
NextR:
    Next r
    LoadSlidesheetSurveys = n
End Function

Private Function IsSummaryRow(ws As Worksheet, ByVal r As Long) As Boolean
    Dim c As Long, f As String
    For c = 4 To 6
        f = ""
        On Error Resume Next
        f = CStr(ws.Cells(r, c).formula)
        On Error GoTo 0
        If InStr(1, f, "LOOKUP", vbTextCompare) > 0 Then
            IsSummaryRow = True
            Exit Function
        End If
    Next c
    IsSummaryRow = False
End Function

' Read Well Seeker X survey stations from an .xlsx / .xlsm / .xls export.
' Reuses an already-open workbook without closing it; otherwise opens read-only and closes.
Private Function ReadWellSeekerXlsx(ByVal fPath As String, ByVal tieThreshold As Double, _
        ByRef outMD() As Double, ByRef outInc() As Double, ByRef outAzi() As Double, _
        ByRef nInterpSkipped As Long, ByRef tieMdWs As Double, ByRef hasTieWs As Boolean, _
        ByRef errMsg As String) As Long
    Dim wb As Workbook, srcWb As Workbook, openedHere As Boolean
    Dim prevEvents As Boolean, prevAlerts As Boolean
    Dim ws As Worksheet, data As Variant
    Dim n As Long

    ReadWellSeekerXlsx = 0: nInterpSkipped = 0: hasTieWs = False: errMsg = ""
    openedHere = False
    For Each wb In Application.Workbooks
        If LCase$(wb.FullName) = LCase$(fPath) Then
            Set srcWb = wb
            Exit For
        End If
    Next wb

    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    On Error GoTo OpenFail
    If srcWb Is Nothing Then
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Set srcWb = Application.Workbooks.Open(fileName:=fPath, UpdateLinks:=0, ReadOnly:=True, Notify:=False, AddToMru:=False)
        openedHere = True
    End If

    ' Prefer a sheet whose name contains "Survey", else first worksheet.
    For Each ws In srcWb.Worksheets
        If InStr(1, ws.name, "Survey", vbTextCompare) > 0 Then
            data = ws.UsedRange.Value2
            n = ParseWellSeekerGrid(data, tieThreshold, outMD, outInc, outAzi, nInterpSkipped, tieMdWs, hasTieWs)
            If n > 0 Then Exit For
        End If
    Next ws
    If n = 0 Then
        For Each ws In srcWb.Worksheets
            data = ws.UsedRange.Value2
            n = ParseWellSeekerGrid(data, tieThreshold, outMD, outInc, outAzi, nInterpSkipped, tieMdWs, hasTieWs)
            If n > 0 Then Exit For
        Next ws
    End If

Cleanup:
    On Error Resume Next
    If openedHere And Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Application.DisplayAlerts = prevAlerts
    Application.EnableEvents = prevEvents
    On Error GoTo 0
    ReadWellSeekerXlsx = n
    Exit Function

OpenFail:
    errMsg = "Could not open " & BaseName(fPath) & ": " & Err.Description
    Resume Cleanup
End Function

' Parse a 2D cell array from a Well Seeker X Survey Report worksheet.
Private Function ParseWellSeekerGrid(ByVal data As Variant, ByVal tieThreshold As Double, _
        ByRef outMD() As Double, ByRef outInc() As Double, ByRef outAzi() As Double, _
        ByRef nInterpSkipped As Long, ByRef tieMdWs As Double, ByRef hasTieWs As Boolean) As Long
    ParseWellSeekerGrid = 0
    If Not IsArray(data) Then Exit Function
    Dim rLo As Long, rHi As Long, cLo As Long, cHi As Long
    rLo = LBound(data, 1): rHi = UBound(data, 1)
    cLo = LBound(data, 2): cHi = UBound(data, 2)

    ' Scan first 150 rows for header containing MD + INC + AZI
    Dim hdrRow As Long, colMd As Long, colInc As Long, colAzi As Long, colDls As Long
    Dim r As Long, c As Long, hdrLim As Long, tok As String
    hdrRow = 0: colMd = 0: colInc = 0: colAzi = 0: colDls = 0
    hdrLim = rLo + 150: If hdrLim > rHi Then hdrLim = rHi

    For r = rLo To hdrLim
        colMd = 0: colInc = 0: colAzi = 0: colDls = 0
        For c = cLo To cHi
            tok = NormalizeHeader(data(r, c))
            If colMd = 0 And (tok = "MD" Or tok = "MEASUREDDEPTH") Then colMd = c
            If colInc = 0 And (tok = "INC" Or tok = "INCLINATION") Then colInc = c
            If colAzi = 0 And (tok = "AZI" Or tok = "AZM" Or tok = "AZIMUTH") Then colAzi = c
            If colDls = 0 And (tok = "DLS" Or tok = "DOGLEGRATE") Then colDls = c
        Next c
        If colMd > 0 And colInc > 0 And colAzi > 0 Then
            hdrRow = r
            Exit For
        End If
    Next r
    If hdrRow = 0 Then Exit Function

    ReDim outMD(0 To rHi - hdrRow)
    ReDim outInc(0 To rHi - hdrRow)
    ReDim outAzi(0 To rHi - hdrRow)

    Dim n As Long: n = 0
    Dim vM As Variant, vI As Variant, vA As Variant, vD As Variant
    Dim mdVal As Double, incVal As Double, aziVal As Double
    Dim rowHasPtb As Boolean

    For r = hdrRow + 1 To rHi
        vM = data(r, colMd)
        vI = data(r, colInc)
        vA = data(r, colAzi)
        If Not (IsNumeric(vM) And IsNumeric(vI) And IsNumeric(vA)) Then GoTo NextGridRow
        If Len(Trim$(CStr(vM & ""))) = 0 Or Len(Trim$(CStr(vI & ""))) = 0 Or Len(Trim$(CStr(vA & ""))) = 0 Then GoTo NextGridRow

        mdVal = CDbl(vM): incVal = CDbl(vI): aziVal = CDbl(vA)
        If incVal < -0.01 Or incVal > 180.01 Then GoTo NextGridRow

        ' Skip trailing PTB / Projected to Bit annotation rows if present
        rowHasPtb = False
        For c = cLo To cHi
            If VarType(data(r, c)) = vbString Then
                Dim sCell As String: sCell = UCase$(Trim$(CStr(data(r, c))))
                If InStr(1, sCell, "PTB") > 0 Or InStr(1, sCell, "PROJECTED") > 0 Or InStr(1, sCell, "PROJECTION") > 0 Then
                    rowHasPtb = True
                    Exit For
                End If
            End If
        Next c
        If rowHasPtb Then GoTo NextGridRow

        ' Track KBE tie-on row (e.g. 5.55 m, Inc 0, Azi 0) and skip rows <= tieThreshold
        If mdVal > 0.01 And mdVal <= 15# And incVal <= 0.01 Then
            tieMdWs = mdVal: hasTieWs = True
        End If
        If mdVal <= tieThreshold Then GoTo NextGridRow

        ' Well Seeker X interpolated formation/casing markers have blank DLS while real surveys have numeric DLS
        If colDls > 0 Then
            vD = data(r, colDls)
            If Not IsNumeric(vD) Or Len(Trim$(CStr(vD & ""))) = 0 Then
                nInterpSkipped = nInterpSkipped + 1
                GoTo NextGridRow
            End If
        End If

        outMD(n) = mdVal: outInc(n) = incVal: outAzi(n) = aziVal
        n = n + 1
NextGridRow:
    Next r
    ParseWellSeekerGrid = n
End Function

' Fallback reader for plain CSV exports (MD, INC, AZI, optional DLS).
Private Function ReadWellSeekerCsv(ByVal fPath As String, ByVal tieThreshold As Double, _
        ByRef outMD() As Double, ByRef outInc() As Double, ByRef outAzi() As Double, _
        ByRef nInterpSkipped As Long, ByRef tieMdWs As Double, ByRef hasTieWs As Boolean, _
        ByRef errMsg As String) As Long
    Dim fNum As Integer, Ln As String
    Dim lines() As String, nLines As Long, i As Long, c As Long
    ReadWellSeekerCsv = 0: nInterpSkipped = 0: hasTieWs = False: errMsg = ""

    On Error GoTo CsvFail
    fNum = FreeFile
    Open fPath For Input As #fNum
    ReDim lines(0 To 500)
    nLines = 0
    Do While Not EOF(fNum)
        Line Input #fNum, Ln
        If nLines > UBound(lines) Then ReDim Preserve lines(0 To UBound(lines) * 2)
        lines(nLines) = Ln
        nLines = nLines + 1
    Loop
    Close #fNum
    If nLines < 2 Then Exit Function

    Dim maxCols As Long: maxCols = 20
    Dim grid() As Variant
    ReDim grid(1 To nLines, 1 To maxCols)
    Dim parts() As String
    For i = 0 To nLines - 1
        parts = SplitCsvRow(lines(i))
        For c = LBound(parts) To UBound(parts)
            If c + 1 <= maxCols Then
                Dim sVal As String: sVal = Trim$(parts(c))
                If IsNumeric(sVal) And Len(sVal) > 0 Then
                    grid(i + 1, c + 1) = CDbl(sVal)
                Else
                    grid(i + 1, c + 1) = sVal
                End If
            End If
        Next c
    Next i
    ReadWellSeekerCsv = ParseWellSeekerGrid(grid, tieThreshold, outMD, outInc, outAzi, nInterpSkipped, tieMdWs, hasTieWs)
    Exit Function

CsvFail:
    On Error Resume Next
    If fNum > 0 Then Close #fNum
    On Error GoTo 0
    errMsg = "Could not read " & BaseName(fPath) & ": " & Err.Description
End Function

Private Function SplitCsvRow(ByVal Ln As String) As String()
    Dim out() As String, n As Long, i As Long, ch As String
    Dim inQ As Boolean, cur As String
    ReDim out(0 To 30)
    n = 0: inQ = False: cur = ""
    For i = 1 To Len(Ln)
        ch = mid$(Ln, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf ch = "," And Not inQ Then
            If n > UBound(out) Then ReDim Preserve out(0 To UBound(out) * 2)
            out(n) = cur: n = n + 1: cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    If n > UBound(out) Then ReDim Preserve out(0 To UBound(out) + 1)
    out(n) = cur
    ReDim Preserve out(0 To n)
    SplitCsvRow = out
End Function

Private Function NormalizeHeader(ByVal v As Variant) As String
    Dim s As String, i As Long, ch As String, res As String
    s = UCase$(Trim$(CStr(v & "")))
    res = ""
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If ch = "(" Then Exit For
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then res = res & ch
    Next i
    NormalizeHeader = res
End Function

Private Function BaseName(ByVal p As String) As String
    Dim k As Long: k = InStrRev(p, "\")
    Dim k2 As Long: k2 = InStrRev(p, "/")
    If k2 > k Then k = k2
    If k > 0 Then BaseName = mid$(p, k + 1) Else BaseName = p
End Function

Private Sub WriteSyncStatusToSetup(ByVal text As String)
    Dim ws As Worksheet, wasProt As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SETUP_SHEET)
    If ws Is Nothing Then Exit Sub
    wasProt = SheetUnprotectForVba(ws)
    Dim anchor As Range
    Set anchor = ws.Range("OC_ImpSurvRow")
    If anchor Is Nothing Then
        ws.Range("M31").MergeArea.Cells(1, 1).value = text
    Else
        Dim btnArea As Range: Set btnArea = anchor.MergeArea
        btnArea.Cells(1, btnArea.Columns.Count).Offset(0, 1).MergeArea.Cells(1, 1).value = text
    End If
    SheetReprotectAfterVba ws, wasProt
    On Error GoTo 0
End Sub

Private Function ReadDataWellboreSummary() As String
    Dim wd As Worksheet, r As Long, lbl As String, v1 As String, v2 As String
    Dim lines As String
    On Error Resume Next
    Set wd = ThisWorkbook.Worksheets(DATA_SHEET)
    If wd Is Nothing Then Exit Function
    For r = 10 To 25
        lbl = Trim$(CStr(wd.Cells(r, 4).text))
        If InStr(1, lbl, "R/L from", vbTextCompare) > 0 Or InStr(1, lbl, "Above/Below", vbTextCompare) > 0 Then
            v1 = Trim$(CStr(wd.Cells(r, 5).text))
            v2 = Trim$(CStr(wd.Cells(r, 6).text))
            If Len(lines) > 0 Then lines = lines & vbCrLf
            lines = lines & lbl & ": " & v1 & " " & v2
        End If
    Next r
    On Error GoTo 0
    ReadDataWellboreSummary = lines
End Function


