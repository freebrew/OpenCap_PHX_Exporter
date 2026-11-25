Attribute VB_Name = "MDL_BrowserHandoff"
Option Explicit

' Applies a browser visible-tab handoff onto this workbook.
' The bundle is <workbook folder>\.phx\handoff\pending\<id>\ with READY written last.
' Scalar cells only. A formula is never overwritten. Numbers compare by value
' (relative 1e-9) so a 17-digit browser serial and an Excel double agree.
' "00" and "2-4" stay text.

Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
    ByRef phAlgorithm As LongPtr, _
    ByVal pszAlgId As LongPtr, _
    ByVal pszImplementation As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
    ByVal hObject As LongPtr, _
    ByVal pszProperty As LongPtr, _
    ByRef pbOutput As Any, _
    ByVal cbOutput As Long, _
    ByRef pcbResult As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
    ByVal hAlgorithm As LongPtr, _
    ByRef phHash As LongPtr, _
    ByRef pbHashObject As Byte, _
    ByVal cbHashObject As Long, _
    ByVal pbSecret As LongPtr, _
    ByVal cbSecret As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbInput As Byte, _
    ByVal cbInput As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr, _
    ByRef pbOutput As Byte, _
    ByVal cbOutput As Long, _
    ByVal dwFlags As Long) As Long

Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _
    ByVal hHash As LongPtr) As Long

Private Type HoRow
    sheetName As String
    a1 As String
    BaseKind As String
    BaseValue As String
    NextKind As String
    NextValue As String
    Decision As String
    ConflictNo As Long
End Type

Private mRows() As HoRow
Private mCount As Long
Private mMessage As String
Private mIncomingBha As Long
Private mExcelBha As Long

' The review form reads these while it is open.
Public gHO_Result As String
Public gHO_Summary As String
Public gHO_N As Long
Public gHO_Addr() As String
Public gHO_Base() As String
Public gHO_Excel() As String
Public gHO_Browser() As String
Public gHO_Pick() As String

Public Function BrowserHandoff_LastMessage() As String
    BrowserHandoff_LastMessage = mMessage
End Function

' Smoke check: SHA-256 of the raw file bytes, lower-case hex.
Public Function BrowserHandoff_Sha256File(ByVal path As String) As String
    BrowserHandoff_Sha256File = Sha256File(path)
End Function

Public Sub BrowserHandoff_CheckOnOpen()
    On Error GoTo Fail
    Dim pending As String
    pending = OldestReady()
    If Len(pending) = 0 Then Exit Sub
    If MsgBox("A browser handoff is waiting." & vbCrLf & vbCrLf & _
              "Review and apply it now?", vbYesNo + vbQuestion, "Browser handoff") = vbYes Then
        ApplyCore "ask"
    End If
    Exit Sub
Fail:
End Sub

Public Sub BrowserHandoff_ApplyPending()
    ApplyCore "ask"
End Sub

' choice is ask, excel, browser, or reject. ask is the button. The others are for the smoke test.
Public Sub BrowserHandoff_ApplyHeadless(ByVal choice As String)
    ApplyCore LCase$(Trim$(choice))
End Sub

Private Sub ApplyCore(ByVal choice As String)
    Dim pending As String
    Dim id As String
    Dim problem As String
    Dim i As Long
    Dim applyCount As Long
    Dim conflictCount As Long
    Dim skippedFormulas As Long
    Dim written As Long
    Dim resolved As Long
    Dim pick As String
    Dim prevEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim touchedApp As Boolean
    Dim opened As Object
    Dim touched As Object
    Dim fso As Object

    On Error GoTo Fail
    mMessage = ""
    If choice <> "ask" And choice <> "excel" And choice <> "browser" And choice <> "reject" Then choice = "ask"
    Set fso = CreateObject("Scripting.FileSystemObject")
    pending = OldestReady()
    If Len(pending) = 0 Then
        mMessage = "No browser handoff is waiting."
        If choice = "ask" Then MsgBox mMessage, vbInformation, "Browser handoff"
        Exit Sub
    End If
    id = fso.GetFileName(pending)

    If choice = "reject" Then
        FinishReject pending, id, "Rejected before apply."
        Exit Sub
    End If

    problem = VerifyBundle(pending)
    If Len(problem) > 0 Then
        mMessage = problem
        If choice = "ask" Then
            If MsgBox(problem & vbCrLf & vbCrLf & "Reject this handoff so the browser can export again?", _
                      vbYesNo + vbExclamation, "Browser handoff") = vbYes Then
                FinishReject pending, id, problem
            End If
        End If
        Exit Sub
    End If

    LoadRows pending & "\cells.csv"
    applyCount = 0
    conflictCount = 0
    skippedFormulas = 0
    For i = 1 To mCount
        If mRows(i).Decision = "apply" Then applyCount = applyCount + 1
        If mRows(i).Decision = "conflict" Then
            conflictCount = conflictCount + 1
            mRows(i).ConflictNo = conflictCount
        End If
        If mRows(i).Decision = "formula" Then skippedFormulas = skippedFormulas + 1
    Next i

    If conflictCount > 0 And choice = "ask" Then
        PrepareReview applyCount, conflictCount, skippedFormulas
        gHO_Result = ""
        frmHandoffReview.Show vbModal
        If gHO_Result <> "apply" Then
            FinishReject pending, id, "Cancelled. No cells were written."
            Exit Sub
        End If
    ElseIf conflictCount = 0 And choice = "ask" Then
        If applyCount = 0 Then
            FinishApplied pending, id, 0, 0, skippedFormulas, "Nothing to apply."
            Exit Sub
        End If
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Apply " & CStr(applyCount) & " cell(s) from the browser?" & vbCrLf & _
                     CStr(skippedFormulas) & " formula cell(s) will be left alone.", _
                     vbYesNoCancel + vbQuestion, "Browser handoff")
        If ans = vbCancel Then
            FinishReject pending, id, "Cancelled. No cells were written."
            Exit Sub
        End If
        If ans = vbNo Then
            mMessage = "Left the handoff waiting."
            Exit Sub
        End If
    End If

    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    touchedApp = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    written = 0
    resolved = 0
    Set opened = CreateObject("Scripting.Dictionary")
    Set touched = CreateObject("Scripting.Dictionary")

    If PendingWrites(choice) > 0 Then BackupWorkbook

    For i = 1 To mCount
        If mRows(i).Decision <> "apply" And mRows(i).Decision <> "conflict" Then GoTo NextRow
        pick = "excel"
        If mRows(i).Decision = "apply" Then
            pick = "browser"
        ElseIf choice = "browser" Then
            pick = "browser"
        ElseIf choice = "ask" Then
            If mRows(i).ConflictNo >= 1 And mRows(i).ConflictNo <= gHO_N Then
                pick = LCase$(gHO_Pick(mRows(i).ConflictNo))
            End If
        End If
        If mRows(i).Decision = "conflict" Then resolved = resolved + 1
        If pick <> "browser" Then GoTo NextRow
        If WriteOne(mRows(i), opened, touched) Then written = written + 1
NextRow:
    Next i

    ReprotectOpened opened
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail
    RunHooks touched
    If written > 0 Then ThisWorkbook.Save
    FinishApplied pending, id, written, resolved, skippedFormulas, _
        "Applied " & CStr(written) & " cell(s)."
    RestoreApp touchedApp, prevEvents, prevCalc, prevScreen
    touchedApp = False
    Erase mRows
    Exit Sub
Fail:
    Dim errDesc As String
    errDesc = Err.Description
    On Error Resume Next
    If Not opened Is Nothing Then ReprotectOpened opened
    RestoreApp touchedApp, prevEvents, prevCalc, prevScreen
    mMessage = "Browser handoff stopped: " & errDesc
    If choice = "ask" Then MsgBox mMessage, vbExclamation, "Browser handoff"
End Sub

Private Function PendingWrites(ByVal choice As String) As Long
    Dim i As Long
    Dim n As Long
    Dim pick As String
    For i = 1 To mCount
        If mRows(i).Decision = "apply" Then
            n = n + 1
        ElseIf mRows(i).Decision = "conflict" Then
            pick = "excel"
            If choice = "browser" Then pick = "browser"
            If choice = "ask" Then
                If mRows(i).ConflictNo >= 1 And mRows(i).ConflictNo <= gHO_N Then pick = LCase$(gHO_Pick(mRows(i).ConflictNo))
            End If
            If pick = "browser" Then n = n + 1
        End If
    Next i
    PendingWrites = n
End Function

Private Function WriteOne(ByRef row As HoRow, ByVal opened As Object, ByVal touched As Object) As Boolean
    Dim ws As Worksheet
    Dim cell As Range
    Dim wasProt As Boolean
    On Error GoTo GiveUp
    Set ws = ThisWorkbook.Worksheets(row.sheetName)
    If Not opened.Exists(ws.name) Then
        wasProt = SheetUnprotectForVba(ws)
        opened.Add ws.name, wasProt
    End If
    Set cell = ws.Range(row.a1)
    If cell.MergeCells Then Set cell = cell.MergeArea.Cells(1, 1)
    If cell.HasFormula Then GoTo GiveUp
    WriteCell cell, row.NextKind, row.NextValue
    If Not touched.Exists(ws.name) Then touched.Add ws.name, True
    WriteOne = True
    Exit Function
GiveUp:
    WriteOne = False
End Function

Private Sub ReprotectOpened(ByVal opened As Object)
    Dim key As Variant
    Dim ws As Worksheet
    If opened Is Nothing Then Exit Sub
    On Error Resume Next
    For Each key In opened.keys
        Set ws = ThisWorkbook.Worksheets(CStr(key))
        SheetReprotectAfterVba ws, CBool(opened(key))
    Next key
End Sub

Private Sub RunHooks(ByVal touched As Object)
    On Error Resume Next
    If touched.Exists("Setup") Or touched.Exists("Data") Then
        MDL_Setup.SyncOpenCapCsvsInPlace
        MDL_DayRoll.DayRoll_EnsureReportHourFormulas
        MDL_TDCalc.RefreshTdPlannedFromPlan
        MDL_ToolHours.ToolHours_Sync
    End If
    If touched.Exists("Slidesheet") Then
        MDL_SlidesheetClear.RefreshSlideComments False
        MDL_SlidesheetClear.RefreshRotateMetres ThisWorkbook.Worksheets("Slidesheet").Range("C13:D505")
        MDL_PlanGauge.RenderPlanGaugeStrict
    End If
    If touched.Exists("Cont DI") Then MDL_ContDI.ContDI_RefreshFormulas
    ' DL Tracker I/J are trusted, including a typed HIA seed. Do not recompute them.
End Sub

Private Sub RestoreApp(ByVal touchedApp As Boolean, ByVal prevEvents As Boolean, ByVal prevCalc As XlCalculation, ByVal prevScreen As Boolean)
    If Not touchedApp Then Exit Sub
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Sub

Private Sub PrepareReview(ByVal applyCount As Long, ByVal conflictCount As Long, ByVal skippedFormulas As Long)
    Dim i As Long
    Dim n As Long
    gHO_N = conflictCount
    ReDim gHO_Addr(1 To conflictCount)
    ReDim gHO_Base(1 To conflictCount)
    ReDim gHO_Excel(1 To conflictCount)
    ReDim gHO_Browser(1 To conflictCount)
    ReDim gHO_Pick(1 To conflictCount)
    n = 0
    For i = 1 To mCount
        If mRows(i).Decision <> "conflict" Then GoTo NextConflict
        n = n + 1
        gHO_Addr(n) = mRows(i).sheetName & "!" & mRows(i).a1
        gHO_Base(n) = ShortText(mRows(i).BaseKind, mRows(i).BaseValue)
        gHO_Browser(n) = ShortText(mRows(i).NextKind, mRows(i).NextValue)
        gHO_Excel(n) = CurrentText(mRows(i).sheetName, mRows(i).a1)
        gHO_Pick(n) = "excel"
NextConflict:
    Next i
    gHO_Summary = CStr(applyCount) & " cell(s) will be written from the browser. " & _
        CStr(conflictCount) & " conflict(s) need a choice. " & _
        CStr(skippedFormulas) & " formula cell(s) stay as they are."
End Sub

Private Function CurrentText(ByVal sheetName As String, ByVal a1 As String) As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim kind As String
    Dim value As String
    On Error GoTo Blank
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Set cell = ws.Range(a1)
    If cell.MergeCells Then Set cell = cell.MergeArea.Cells(1, 1)
    ReadCellKind cell, kind, value
    CurrentText = ShortText(kind, value)
    Exit Function
Blank:
    CurrentText = "(unreadable)"
End Function

Private Function ShortText(ByVal kind As String, ByVal value As String) As String
    Dim shown As String
    If StrComp(kind, "blank", vbTextCompare) = 0 Or Len(value) = 0 Then
        ShortText = "(blank)"
        Exit Function
    End If
    shown = value
    If Len(shown) > 40 Then shown = Left$(shown, 37) & "..."
    ShortText = kind & " " & shown
End Function

Private Sub LoadRows(ByVal path As String)
    Dim rows As Collection
    Dim header As Collection
    Dim row As Collection
    Dim cSheet As Long
    Dim cA1 As Long
    Dim cBaseKind As Long
    Dim cBaseValue As Long
    Dim cNextKind As Long
    Dim cNextValue As Long
    Dim r As Long
    Dim n As Long

    Set rows = ParseCsv(ReadUtf8(path))
    mCount = 0
    Erase mRows
    If rows.Count < 2 Then Exit Sub
    Set header = rows(1)
    cSheet = ColIndex(header, "sheet")
    cA1 = ColIndex(header, "a1")
    cBaseKind = ColIndex(header, "baseKind")
    cBaseValue = ColIndex(header, "baseValue")
    cNextKind = ColIndex(header, "nextKind")
    cNextValue = ColIndex(header, "nextValue")
    If cSheet = 0 Or cA1 = 0 Or cBaseKind = 0 Or cBaseValue = 0 Or cNextKind = 0 Or cNextValue = 0 Then
        Err.Raise vbObjectError + 711, "MDL_BrowserHandoff", "cells.csv is missing a column."
    End If
    n = rows.Count - 1
    If n < 1 Then Exit Sub
    ReDim mRows(1 To n)
    mIncomingBha = 0
    mExcelBha = CurrentExcelBha()
    For r = 2 To rows.Count
        Set row = rows(r)
        mCount = mCount + 1
        mRows(mCount).sheetName = FieldAt(row, cSheet)
        mRows(mCount).a1 = FieldAt(row, cA1)
        mRows(mCount).BaseKind = FieldAt(row, cBaseKind)
        mRows(mCount).BaseValue = FieldAt(row, cBaseValue)
        mRows(mCount).NextKind = FieldAt(row, cNextKind)
        mRows(mCount).NextValue = FieldAt(row, cNextValue)
        If StrComp(mRows(mCount).sheetName, "Data", vbTextCompare) = 0 _
           And StrComp(mRows(mCount).a1, "H3", vbTextCompare) = 0 Then
            mIncomingBha = ParseBhaLong(mRows(mCount).NextKind, mRows(mCount).NextValue)
        End If
    Next r
    For r = 1 To mCount
        mRows(r).Decision = ClassifyRow(mRows(r))
    Next r
End Sub

Private Function ParseBhaLong(ByVal kind As String, ByVal value As String) As Long
    Dim d As Double
    If Not TryPlainNumber(kind, value, d) Then Exit Function
    If d > 0 Then ParseBhaLong = CLng(Round(d, 0))
End Function

Private Function CurrentExcelBha() As Long
    Dim ws As Worksheet
    Dim cell As Range
    Dim kind As String
    Dim value As String
    On Error GoTo Done
    Set ws = ThisWorkbook.Worksheets("Data")
    Set cell = ws.Range("H3")
    If cell.HasFormula Then GoTo Done
    ReadCellKind cell, kind, value
    CurrentExcelBha = ParseBhaLong(kind, value)
Done:
End Function

Private Function IsBrowserAuthoritativeSheet(ByVal sheetName As String) As Boolean
    IsBrowserAuthoritativeSheet = (StrComp(sheetName, "Setup", vbTextCompare) = 0 _
        Or StrComp(sheetName, "Slidesheet", vbTextCompare) = 0)
End Function

' Code-painted / computed Data tables rebuilt from _OC_* / _TH_Hours / _OC_Survey / _OC_AC / Slidesheet.
' Never written by handoff; hooks rebuild them cleanly on both sides.
Private Function IsDerivedCell(ByVal sheetName As String, ByVal cell As Range) As Boolean
    Dim r As Long
    Dim c As Long
    If StrComp(sheetName, "Data", vbTextCompare) <> 0 Then Exit Function
    r = cell.row
    c = cell.Column
    If r >= 13 And r <= 15 And (c = 5 Or c = 6) Then IsDerivedCell = True: Exit Function
    If r >= 35 And r <= 42 And c >= 2 And c <= 6 Then IsDerivedCell = True: Exit Function
    If r >= 45 And r <= 55 And c = 2 Then IsDerivedCell = True: Exit Function
    If r >= 30 And r <= 41 And (c = 15 Or c = 16) Then IsDerivedCell = True: Exit Function
    If r >= 31 And r <= 41 And (c = 18 Or c = 19) Then IsDerivedCell = True: Exit Function
    If r >= 44 And r <= 55 And (c = 15 Or c = 16) Then IsDerivedCell = True: Exit Function
    If r >= 45 And r <= 55 And (c = 18 Or c = 19) Then IsDerivedCell = True: Exit Function
    If r = 54 And c >= 9 And c <= 13 Then IsDerivedCell = True: Exit Function
    If r >= 2 And r <= 400 And c >= 24 And c <= 28 Then IsDerivedCell = True: Exit Function
End Function

' Active-BHA-scoped cells on Data (H3, C22, H4:Q22). Governed atomically by the higher BHA #.
Private Function IsActiveBhaCell(ByVal sheetName As String, ByVal cell As Range) As Boolean
    Dim r As Long
    Dim c As Long
    If StrComp(sheetName, "Data", vbTextCompare) <> 0 Then Exit Function
    r = cell.row
    c = cell.Column
    If r = 3 And c = 8 Then IsActiveBhaCell = True: Exit Function
    If r = 22 And c = 3 Then IsActiveBhaCell = True: Exit Function
    If r >= 4 And r <= 22 And c >= 8 And c <= 17 Then IsActiveBhaCell = True: Exit Function
End Function

Private Function ClassifyRow(ByRef row As HoRow) As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim kind As String
    Dim value As String
    On Error GoTo skip
    If Len(row.sheetName) = 0 Or Len(row.a1) = 0 Then GoTo skip
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(row.sheetName)
    On Error GoTo skip
    If ws Is Nothing Then GoTo skip
    If ws.Visible <> xlSheetVisible Then GoTo skip
    If StrComp(row.NextKind, "formula", vbTextCompare) = 0 Then
        ClassifyRow = "formula"
        Exit Function
    End If
    Set cell = ws.Range(row.a1)
    If cell.MergeCells Then Set cell = cell.MergeArea.Cells(1, 1)
    If IsDerivedCell(ws.Name, cell) Then
        ClassifyRow = "skip"
        Exit Function
    End If
    If cell.HasFormula Then
        ClassifyRow = "formula"
        Exit Function
    End If
    ReadCellKind cell, kind, value
    If ValuesEqual(kind, value, row.NextKind, row.NextValue) Then
        ClassifyRow = "skip"
    ElseIf IsBrowserAuthoritativeSheet(ws.Name) Then
        ' Setup and Slidesheet are one-way Browser -> Excel. No conflict prompt.
        ClassifyRow = "apply"
    ElseIf IsActiveBhaCell(ws.Name, cell) And mIncomingBha > 0 And mExcelBha > 0 And mIncomingBha <> mExcelBha Then
        ' Newest BHA # owns Data!H3, C22, and H4:Q22 atomically.
        If mIncomingBha > mExcelBha Then
            ClassifyRow = "apply"
        Else
            ClassifyRow = "skip"
        End If
    ElseIf ValuesEqual(kind, value, row.BaseKind, row.BaseValue) Then
        ClassifyRow = "apply"
    Else
        ClassifyRow = "conflict"
    End If
    Exit Function
skip:
    ClassifyRow = "skip"
End Function

Private Sub ReadCellKind(ByVal cell As Range, ByRef kind As String, ByRef value As String)
    Dim v As Variant
    If cell.HasFormula Then
        kind = "formula"
        value = CStr(cell.formula)
        Exit Sub
    End If
    v = cell.Value2
    If IsEmpty(v) Or IsNull(v) Then
        kind = "blank"
        value = ""
        Exit Sub
    End If
    If VarType(v) = vbBoolean Then
        kind = "bool"
        If CBool(v) Then value = "TRUE" Else value = "FALSE"
        Exit Sub
    End If
    If isError(v) Then
        kind = "text"
        value = CStr(cell.text)
        Exit Sub
    End If
    If VarType(v) = vbString Then
        kind = "text"
        value = CStr(v)
        Exit Sub
    End If
    If IsNumeric(v) Then
        kind = "number"
        value = Trim$(Str$(CDbl(v)))
        Exit Sub
    End If
    kind = "text"
    value = CStr(v)
End Sub

Private Sub WriteCell(ByVal cell As Range, ByVal kind As String, ByVal value As String)
    If StrComp(kind, "blank", vbTextCompare) = 0 Or (StrComp(kind, "text", vbTextCompare) = 0 And Len(value) = 0) Then
        cell.ClearContents
        Exit Sub
    End If
    If StrComp(kind, "bool", vbTextCompare) = 0 Then
        cell.value = (StrComp(value, "TRUE", vbTextCompare) = 0)
        Exit Sub
    End If
    If StrComp(kind, "number", vbTextCompare) = 0 Then
        cell.value = val(value)
        Exit Sub
    End If
    If NeedTextFormat(value) Then cell.numberFormat = "@"
    cell.value = value
End Sub

Private Function NeedTextFormat(ByVal value As String) As Boolean
    Dim re As Object
    If Len(value) = 0 Then Exit Function
    If IsNumeric(value) Then
        NeedTextFormat = True
        Exit Function
    End If
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[0-9]+-[0-9]+$"
    NeedTextFormat = re.Test(value)
End Function

Private Function ValuesEqual(ByVal k1 As String, ByVal v1 As String, ByVal k2 As String, ByVal v2 As String) As Boolean
    Dim n1 As Double
    Dim n2 As Double
    Dim has1 As Boolean
    Dim has2 As Boolean
    has1 = TryPlainNumber(k1, v1, n1)
    has2 = TryPlainNumber(k2, v2, n2)
    If has1 And has2 Then
        ValuesEqual = NumbersClose(n1, n2)
        Exit Function
    End If
    If StrComp(NormKind(k1, v1), NormKind(k2, v2), vbTextCompare) <> 0 Then Exit Function
    If StrComp(NormKind(k1, v1), "blank", vbTextCompare) = 0 Then
        ValuesEqual = True
        Exit Function
    End If
    If StrComp(k1, "bool", vbTextCompare) = 0 Then
        ValuesEqual = (StrComp(v1, v2, vbTextCompare) = 0)
        Exit Function
    End If
    ValuesEqual = (StrComp(v1, v2, vbBinaryCompare) = 0)
End Function

Private Function NormKind(ByVal kind As String, ByVal value As String) As String
    If StrComp(kind, "text", vbTextCompare) = 0 And Len(value) = 0 Then
        NormKind = "blank"
    Else
        NormKind = LCase$(kind)
    End If
End Function

Private Function TryPlainNumber(ByVal kind As String, ByVal value As String, ByRef n As Double) As Boolean
    Dim re As Object
    On Error GoTo GiveUp
    If StrComp(kind, "number", vbTextCompare) = 0 Then
        If Len(Trim$(value)) = 0 Then GoTo GiveUp
        n = val(value)
        TryPlainNumber = True
        Exit Function
    End If
    If StrComp(kind, "text", vbTextCompare) <> 0 Then GoTo GiveUp
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^-?(0|[1-9][0-9]*)(\.[0-9]+)?$"
    If Not re.Test(value) Then GoTo GiveUp
    n = val(value)
    TryPlainNumber = True
    Exit Function
GiveUp:
    TryPlainNumber = False
End Function

Private Function NumbersClose(ByVal leftVal As Double, ByVal rightVal As Double) As Boolean
    Dim span As Double
    If leftVal = 0 And rightVal = 0 Then
        NumbersClose = True
        Exit Function
    End If
    span = Abs(leftVal)
    If Abs(rightVal) > span Then span = Abs(rightVal)
    If span < 1 Then span = 1
    NumbersClose = (Abs(leftVal - rightVal) <= 0.000000001 * span)
End Function

Private Function VerifyBundle(ByVal folder As String) As String
    Dim manifest As Object
    Dim token As String
    Dim mode As String
    Dim name As Variant
    Dim expected As String
    Dim actual As String
    Dim names As Variant
    On Error GoTo Bad
    If Not FileExists(folder & "\manifest.csv") Then
        VerifyBundle = "manifest.csv is missing."
        Exit Function
    End If
    If Not FileExists(folder & "\cells.csv") Then
        VerifyBundle = "cells.csv is missing."
        Exit Function
    End If
    Set manifest = ManifestMap(ReadUtf8(folder & "\manifest.csv"))
    If Not manifest.Exists("schema") Or CStr(manifest("schema")) <> "1" Then
        VerifyBundle = "This handoff schema is not supported."
        Exit Function
    End If
    token = JobToken(ThisWorkbook.name)
    If Not manifest.Exists("jobNumber") Or Len(token) = 0 Or StrComp(CStr(manifest("jobNumber")), token, vbTextCompare) <> 0 Then
        VerifyBundle = "This handoff is for a different job than " & token & "."
        Exit Function
    End If
    mode = ""
    If manifest.Exists("mode") Then mode = CStr(manifest("mode"))
    If Len(mode) > 0 And StrComp(mode, "visible-data", vbTextCompare) <> 0 Then
        VerifyBundle = "This handoff is not a visible-tab data bundle."
        Exit Function
    End If
    names = Array("cells.csv", "styles.csv", "merges.csv", "geometry.csv", "sheets.csv", "names.csv")
    For Each name In names
        If Not manifest.Exists(CStr(name)) Then GoTo NextName
        If Not FileExists(folder & "\" & CStr(name)) Then
            VerifyBundle = CStr(name) & " is listed but missing."
            Exit Function
        End If
        expected = LCase$(CStr(manifest(CStr(name))))
        actual = Sha256File(folder & "\" & CStr(name))
        If StrComp(expected, actual, vbTextCompare) <> 0 Then
            VerifyBundle = CStr(name) & " does not match the manifest."
            Exit Function
        End If
NextName:
    Next name
    VerifyBundle = ""
    Exit Function
Bad:
    VerifyBundle = "The handoff could not be read. " & Err.Description
End Function

Private Function ManifestMap(ByVal text As String) As Object
    Dim map As Object
    Dim rows As Collection
    Dim row As Collection
    Dim r As Long
    Set map = CreateObject("Scripting.Dictionary")
    Set rows = ParseCsv(text)
    For r = 2 To rows.Count
        Set row = rows(r)
        If row.Count >= 2 Then map(CStr(row(1))) = CStr(row(2))
    Next r
    Set ManifestMap = map
End Function

Private Function OldestReady() As String
    Dim fso As Object
    Dim pending As String
    Dim folder As Object
    Dim best As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    pending = ThisWorkbook.path & "\.phx\handoff\pending"
    If Not fso.FolderExists(pending) Then Exit Function
    best = ""
    For Each folder In fso.GetFolder(pending).SubFolders
        If fso.FileExists(folder.path & "\READY") Then
            If Len(best) = 0 Or StrComp(folder.name, best, vbTextCompare) < 0 Then best = folder.name
        End If
    Next folder
    If Len(best) > 0 Then OldestReady = pending & "\" & best
End Function

' Same rule as the browser workbookToken: a run of 4 or more digits, otherwise the last word.
Private Function JobToken(ByVal fileName As String) As String
    Dim re As Object
    Dim matches As Object
    Dim BaseName As String
    Dim i As Long
    Dim ch As String
    Dim word As String
    Dim lastWord As String
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\d{4,})"
    re.Global = False
    If re.Test(fileName) Then
        Set matches = re.Execute(fileName)
        JobToken = CStr(matches(0).SubMatches(0))
        Exit Function
    End If
    BaseName = fileName
    If Len(BaseName) > 5 Then
        If StrComp(Right$(BaseName, 5), ".xlsm", vbTextCompare) = 0 Then BaseName = Left$(BaseName, Len(BaseName) - 5)
    End If
    lastWord = ""
    word = ""
    For i = 1 To Len(BaseName)
        ch = mid$(BaseName, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            word = word & ch
        Else
            If Len(word) > 0 Then lastWord = word
            word = ""
        End If
    Next i
    If Len(word) > 0 Then lastWord = word
    JobToken = lastWord
End Function

Private Sub FinishApplied(ByVal pending As String, ByVal id As String, ByVal written As Long, ByVal resolved As Long, ByVal skippedFormulas As Long, ByVal message As String)
    WriteReceipt id, "applied", CStr(written), CStr(resolved), CStr(skippedFormulas), message
    MoveBundle pending, "applied", id
    mMessage = message
End Sub

Private Sub FinishReject(ByVal pending As String, ByVal id As String, ByVal message As String)
    WriteReceipt id, "rejected", "0", "0", "0", message
    MoveBundle pending, "rejected", id
    mMessage = message
End Sub

Private Sub MoveBundle(ByVal src As String, ByVal bucket As String, ByVal id As String)
    Dim fso As Object
    Dim destRoot As String
    Dim dest As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    destRoot = ThisWorkbook.path & "\.phx\handoff\" & bucket
    EnsureFolder destRoot
    dest = destRoot & "\" & id
    If fso.FolderExists(dest) Then fso.DeleteFolder dest, True
    fso.MoveFolder src, dest
End Sub

Private Sub WriteReceipt(ByVal id As String, ByVal status As String, ByVal operations As String, ByVal conflictsResolved As String, ByVal skippedFormulas As String, ByVal message As String)
    Dim folder As String
    Dim path As String
    Dim stm As Object
    Dim body As String
    folder = ThisWorkbook.path & "\.phx\handoff\receipts"
    EnsureFolder folder
    path = folder & "\" & id & ".csv"
    body = "key,value" & vbLf & _
        "status," & CsvField(status) & vbLf & _
        "handoffId," & CsvField(id) & vbLf & _
        "operations," & CsvField(operations) & vbLf & _
        "conflictsResolved," & CsvField(conflictsResolved) & vbLf & _
        "skippedFormulas," & CsvField(skippedFormulas) & vbLf & _
        "message," & CsvField(message) & vbLf
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText body
    stm.SaveToFile path, 2
    stm.Close
End Sub

Private Sub BackupWorkbook()
    Dim folder As String
    Dim dest As String
    folder = ThisWorkbook.path & "\backups"
    EnsureFolder folder
    dest = folder & "\" & ThisWorkbook.name & ".bak_handoff_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs dest
End Sub

Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Dim parent As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then EnsureFolder parent
    End If
    fso.CreateFolder path
End Sub

Private Function FileExists(ByVal path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

Private Function CsvField(ByVal value As String) As String
    If InStr(1, value, ",", vbBinaryCompare) > 0 Or InStr(1, value, """", vbBinaryCompare) > 0 _
        Or InStr(1, value, vbLf, vbBinaryCompare) > 0 Or InStr(1, value, vbCr, vbBinaryCompare) > 0 Then
        CsvField = """" & Replace(value, """", """""") & """"
    Else
        CsvField = value
    End If
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadUtf8 = stm.ReadText(-1)
    stm.Close
End Function

Private Function Sha256File(ByVal path As String) As String
    Dim hAlg As LongPtr
    Dim hHash As LongPtr
    Dim status As Long
    Dim cbData As Long
    Dim cbHashObject As Long
    Dim cbHash As Long
    Dim hashObject() As Byte
    Dim digest() As Byte
    Dim n As Long
    Dim fn As Integer
    Dim bytes() As Byte
    Dim i As Long
    Dim out As String
    Dim errNum As Long
    Dim errText As String

    status = BCryptOpenAlgorithmProvider(hAlg, StrPtr("SHA256"), 0, 0)
    If status <> 0 Then
        errNum = vbObjectError + 710
        errText = "SHA-256 could not start (" & Hex$(status) & ")."
        GoTo Clean
    End If
    status = BCryptGetProperty(hAlg, StrPtr("ObjectLength"), cbHashObject, 4, cbData, 0)
    If status <> 0 Then
        errNum = vbObjectError + 710
        errText = "SHA-256 could not size the hash (" & Hex$(status) & ")."
        GoTo Clean
    End If
    status = BCryptGetProperty(hAlg, StrPtr("HashDigestLength"), cbHash, 4, cbData, 0)
    If status <> 0 Then
        errNum = vbObjectError + 710
        errText = "SHA-256 could not size the digest (" & Hex$(status) & ")."
        GoTo Clean
    End If
    ReDim hashObject(0 To cbHashObject - 1)
    ReDim digest(0 To cbHash - 1)
    status = BCryptCreateHash(hAlg, hHash, hashObject(0), cbHashObject, 0, 0, 0)
    If status <> 0 Then
        errNum = vbObjectError + 710
        errText = "SHA-256 could not hash (" & Hex$(status) & ")."
        GoTo Clean
    End If
    On Error GoTo Blow

    n = FileLen(path)
    If n > 0 Then
        ReDim bytes(0 To n - 1)
        fn = FreeFile
        Open path For Binary Access Read As #fn
        Get #fn, , bytes
        Close #fn
        fn = 0
        status = BCryptHashData(hHash, bytes(0), n, 0)
        If status <> 0 Then Err.Raise vbObjectError + 710, "MDL_BrowserHandoff", "SHA-256 failed (" & Hex$(status) & ")."
    End If
    status = BCryptFinishHash(hHash, digest(0), cbHash, 0)
    If status <> 0 Then Err.Raise vbObjectError + 710, "MDL_BrowserHandoff", "SHA-256 failed (" & Hex$(status) & ")."
    For i = 0 To cbHash - 1
        out = out & Right$("0" & LCase$(Hex$(digest(i))), 2)
    Next i
    Sha256File = out
    GoTo Clean
Blow:
    errNum = Err.Number
    errText = Err.Description
Clean:
    On Error Resume Next
    If fn <> 0 Then Close #fn
    If hHash <> 0 Then BCryptDestroyHash hHash
    If hAlg <> 0 Then BCryptCloseAlgorithmProvider hAlg, 0
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "MDL_BrowserHandoff", errText
End Function

Private Function ParseCsv(ByVal text As String) As Collection
    Dim rows As New Collection
    Dim fields As Collection
    Dim field As String
    Dim quoted As Boolean
    Dim i As Long
    Dim ch As String
    Dim nxt As String
    Set fields = New Collection
    If Len(text) > 0 Then
        If AscW(mid$(text, 1, 1)) = &HFEFF Then text = mid$(text, 2)
    End If
    i = 1
    Do While i <= Len(text)
        ch = mid$(text, i, 1)
        If quoted Then
            If ch = """" Then
                If i < Len(text) Then nxt = mid$(text, i + 1, 1) Else nxt = ""
                If nxt = """" Then
                    field = field & """"
                    i = i + 2
                Else
                    quoted = False
                    i = i + 1
                End If
            Else
                field = field & ch
                i = i + 1
            End If
        ElseIf ch = """" Then
            quoted = True
            i = i + 1
        ElseIf ch = "," Then
            fields.Add field
            field = ""
            i = i + 1
        ElseIf ch = vbLf Then
            fields.Add field
            AddCsvRow rows, fields
            Set fields = New Collection
            field = ""
            i = i + 1
        ElseIf ch = vbCr Then
            i = i + 1
        Else
            field = field & ch
            i = i + 1
        End If
    Loop
    If Len(field) > 0 Or fields.Count > 0 Then
        fields.Add field
        AddCsvRow rows, fields
    End If
    Set ParseCsv = rows
End Function

Private Sub AddCsvRow(ByVal rows As Collection, ByVal fields As Collection)
    If fields.Count = 1 Then
        If CStr(fields(1)) = "" Then Exit Sub
    End If
    rows.Add fields
End Sub

Private Function ColIndex(ByVal header As Collection, ByVal name As String) As Long
    Dim i As Long
    For i = 1 To header.Count
        If StrComp(CStr(header(i)), name, vbTextCompare) = 0 Then
            ColIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function FieldAt(ByVal row As Collection, ByVal index As Long) As String
    If index < 1 Or index > row.Count Then
        FieldAt = ""
    Else
        FieldAt = CStr(row(index))
    End If
End Function

