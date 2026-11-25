Attribute VB_Name = "MDL_DownlinkTracker"
Option Explicit

' DL Tracker tab. ADD ENTRY (A1:B1) opens frmDLEntry prefilled from the last
' entry; CLEAR (C1:E1) empties typed log cells only. CLIENT sits in F1:G1.
' Entry writes touch only unlocked, non-formula cells, so D/L Setting (H),
' Dnlk # (B), hidden helper columns and the command tables (AA:AR) stay.
' The tab is protected with a blank password; buttons and the Dnlk # repair
' unprotect briefly and restore the same protection options.

Public Const SH_DLT As String = "DL Tracker"
Private Const SH_DLT_OLD As String = "Downlink Tracker V4.6"
Private Const DLT_LOG As String = "A3:U674"
Public Const DLT_FIRST As Long = 3
Public Const DLT_LAST As Long = 674
Public Const DLT_BTN As String = "DLT_ClearBtn"
Public Const DLT_ADD_BTN As String = "DLT_AddBtn"

' Log columns.
Public Const DLC_RUN As Long = 1
Public Const DLC_NUM As Long = 2
Public Const DLC_PAGE As Long = 3
Public Const DLC_CMD As Long = 4
Public Const DLC_PERIOD As Long = 5
Public Const DLC_TYPE As Long = 6
Public Const DLC_BOTTOM As Long = 7
Public Const DLC_SET As Long = 8
Public Const DLC_I As Long = 9
Public Const DLC_J As Long = 10
Public Const DLC_DEPTH As Long = 11
Public Const DLC_DATE As Long = 12
Public Const DLC_TIME As Long = 13
Public Const DLC_CODE As Long = 14
Public Const DLC_MON As Long = 15
Public Const DLC_PASS As Long = 16
Public Const DLC_REF As Long = 17
Public Const DLC_DROP As Long = 18
Public Const DLC_LOW As Long = 19
Public Const DLC_RPM As Long = 20
Public Const DLC_NOTE As Long = 21

' Command tables: one key column per DL page, setting text one column right.
Private Const TBL_FIRST As Long = 3
Private Const TBL_LAST As Long = 66

' ComputeIJ results.
Public Const DLT_IJ_KEEP As Long = 0     ' nothing to work from - leave I/J alone
Public Const DLT_IJ_SET As Long = 1      ' outI / outJ are the new values
Public Const DLT_IJ_BLANK As Long = 2    ' page change / unsupported - blank I/J

' Handed to frmDLEntry by DLT_AddEntry.
Public gDLT_TargetRow As Long
Public gDLT_SourceRow As Long

Private mPDraw As Boolean, mPScen As Boolean
Private mPFmtCells As Boolean, mPFmtCols As Boolean, mPFmtRows As Boolean
Private mPInsCols As Boolean, mPInsRows As Boolean, mPInsLinks As Boolean
Private mPDelCols As Boolean, mPDelRows As Boolean
Private mPSort As Boolean, mPFilter As Boolean, mPPivot As Boolean

' Mouse wheel on an open form dropdown. Excel's ComboBox list ignores the
' wheel; while frmDLEntry is open a low-level hook scrolls that list.
Private mWheelHook As LongPtr
Private mWheelBusy As Boolean
Private mWheelComboName As String

#If VBA7 Then
Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetClassNameW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)
#End If

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A

' ================================================================================
'  BUTTONS
' ================================================================================

Public Sub DLT_AddEntry()
    Dim ws As Worksheet
    Dim lastR As Long

    Set ws = DltSheet()
    If ws Is Nothing Then
        MsgBox "The '" & SH_DLT & "' tab is not in this workbook.", vbInformation, "Add Downlink Entry"
        Exit Sub
    End If
    lastR = DLT_LastEntryRow(ws)
    If lastR >= DLT_LAST Then
        MsgBox "The downlink log is full (row " & DLT_LAST & ").", vbExclamation, "Add Downlink Entry"
        Exit Sub
    End If
    gDLT_SourceRow = lastR
    gDLT_TargetRow = lastR + 1
    frmDLEntry.Show
End Sub

Public Sub DLT_ClearEntries()
    Dim ws As Worksheet
    Dim n As Long

    Set ws = DltSheet()
    If ws Is Nothing Then
        MsgBox "The '" & SH_DLT & "' tab is not in this workbook.", vbInformation, "Clear Downlink Tracker"
        Exit Sub
    End If
    If MsgBox("Clear every downlink entry in rows 3-674?" & vbCrLf & vbCrLf & _
              "Client / Rig / Well / Job # and all formulas stay." & vbCrLf & _
              "This cannot be undone.", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Clear Downlink Tracker") <> vbYes Then Exit Sub

    n = DLT_ClearLog(ws)
    If n < 0 Then
        MsgBox "Clear stopped: " & Err.Description, vbCritical, "Clear Downlink Tracker"
    Else
        MsgBox n & " entries cleared.", vbInformation, "Clear Downlink Tracker"
    End If
End Sub

' Returns cells cleared, or -1 on error (Err still set for the caller).
Public Function DLT_ClearLog(ByVal ws As Worksheet) As Long
    Dim c As Range
    Dim n As Long
    Dim r As Long
    Dim wasProt As Boolean
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail
    wasProt = DltUnprotect(ws)
    For Each c In ws.Range(DLT_LOG).Cells
        If Not c.Locked Then
            If Not c.HasFormula Then
                If Not IsEmpty(c.Value2) Then
                    c.ClearContents
                    n = n + 1
                End If
            End If
        End If
    Next c
    For r = DLT_FIRST To DLT_LAST
        PutSettingFormula ws, r
    Next r
    DLT_ClearLog = n
    GoTo Restore

Fail:
    DLT_ClearLog = -1

Restore:
    DltReprotect ws, wasProt
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Function

' ADD ENTRY on A1:B1, CLEAR on C1:E1, CLIENT on F1:G1. Frozen header row.
Public Sub DLT_EnsureButtons()
    Dim ws As Worksheet
    Dim wasProt As Boolean

    Set ws = DltSheet()
    If ws Is Nothing Then Exit Sub
    wasProt = DltUnprotect(ws)
    On Error GoTo Restore
    PaintHeaderChrome ws
    PlaceButton ws, DLT_ADD_BTN, ws.Range("A1:B1"), "ADD ENTRY", "DLT_AddEntry"
    PlaceButton ws, DLT_BTN, ws.Range("C1:E1"), "CLEAR", "DLT_ClearEntries"
Restore:
    DltReprotect ws, wasProt
End Sub

Public Sub DLT_EnsureClearButton()
    DLT_EnsureButtons
End Sub

' Called from Workbook_SheetActivate: paint the header and buttons when the
' layout is still the old A1:G1 CLIENT bar / CLEAR-on-U1, or a button is gone.
Public Sub DLT_EnsureButtonsIfMissing()
    Dim ws As Worksheet
    Set ws = DltSheet()
    If ws Is Nothing Then Exit Sub
    EnsureSettingOnLog ws
    If HeaderChromeOk(ws) And HasShape(ws, DLT_ADD_BTN) And HasShape(ws, DLT_BTN) Then Exit Sub
    DLT_EnsureButtons
End Sub

Private Function HeaderChromeOk(ByVal ws As Worksheet) As Boolean
    On Error GoTo Done
    If ws.Range("A1").MergeArea.Address(False, False) <> "A1:B1" Then Exit Function
    If ws.Range("C1").MergeArea.Address(False, False) <> "C1:E1" Then Exit Function
    If ws.Range("F1").MergeArea.Address(False, False) <> "F1:G1" Then Exit Function
    If UCase$(Trim$(CStr(ws.Range("F1").Value2 & ""))) <> "CLIENT" Then Exit Function
    HeaderChromeOk = True
Done:
End Function

' A1:B1 ADD ENTRY, C1:E1 CLEAR, F1:G1 CLIENT. Does not touch H1 (client name)
' or the rest of the identity row (RIG / well / job #).
Private Sub PaintHeaderChrome(ByVal ws As Worksheet)
    Dim gray As Long
    Dim fSize As Double
    Dim fName As String

    gray = ws.Range("A1").Interior.Color
    If gray = 0 Then gray = RGB(191, 191, 191)
    fSize = ws.Range("A1").Font.size
    If fSize < 8# Then fSize = 11#
    fName = ws.Range("A1").Font.name
    If Len(fName) = 0 Then fName = "Calibri"

    On Error Resume Next
    ws.Range("A1:G1").UnMerge
    On Error GoTo 0
    ws.Range("A1:E1").ClearContents
    On Error Resume Next
    ws.Range("A1:B1").Merge
    ws.Range("C1:E1").Merge
    ws.Range("F1:G1").Merge
    On Error GoTo 0

    With ws.Range("A1:G1")
        .Interior.Color = gray
        .Font.bold = True
        .Font.size = fSize
        .Font.name = fName
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Locked = True
    End With
    ws.Range("F1").value = "CLIENT"
End Sub

' Cheap self-heal: if the first or last entry row lost its H formula, put it
' back on every log row. The template still has H further down after a clear.
Private Sub EnsureSettingOnLog(ByVal ws As Worksheet)
    Dim lastR As Long
    Dim wasProt As Boolean
    Dim need As Boolean
    Dim r As Long
    wasProt = DltUnprotect(ws)
    On Error GoTo Restore
    lastR = DLT_LastEntryRow(ws)
    If lastR < DLT_FIRST Then lastR = DLT_FIRST
    need = (CStr(ws.Cells(DLT_FIRST, DLC_SET).formula) <> SettingFormula(DLT_FIRST)) _
        Or (CStr(ws.Cells(lastR, DLC_SET).formula) <> SettingFormula(lastR))
    If need Then
        For r = DLT_FIRST To DLT_LAST
            PutSettingFormula ws, r
        Next r
    End If
Restore:
    DltReprotect ws, wasProt
End Sub

Private Sub PlaceButton(ByVal ws As Worksheet, ByVal nm As String, ByVal anchor As Range, _
                        ByVal caption As String, ByVal action As String)
    Dim btn As Object
    Dim w As Double, h As Double

    On Error Resume Next
    ws.Buttons(nm).Delete
    On Error GoTo 0
    w = anchor.width - 6
    h = anchor.Height - 4
    If w < 24 Then w = 24
    If h < 12 Then h = 12
    Set btn = ws.Buttons.Add(anchor.Left + 3, anchor.Top + 2, w, h)
    btn.name = nm
    btn.caption = caption
    btn.OnAction = action
    btn.Placement = xlFreeFloating
    btn.PrintObject = False
    btn.Font.bold = True
End Sub

Private Function HasShape(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = ws.Shapes(nm)
    HasShape = Not s Is Nothing
End Function

' Puts the Dnlk # counter formula back where a number was typed over it.
Public Function DLT_RepairDnlkCounter() As Long
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim r As Long, n As Long

    Set ws = DltSheet()
    If ws Is Nothing Then Exit Function
    wasProt = DltUnprotect(ws)
    On Error GoTo Restore
    For r = DLT_FIRST To DLT_LAST
        If Not ws.Cells(r, DLC_NUM).HasFormula Then
            If Len(Trim$(CStr(ws.Cells(r, DLC_NUM).Value2 & ""))) > 0 Then
                ws.Cells(r, DLC_NUM).formula = CounterFormula(r)
                n = n + 1
            End If
        End If
    Next r
Restore:
    DltReprotect ws, wasProt
    DLT_RepairDnlkCounter = n
End Function

' The template carries the counter only part way down the log; a new entry
' below that gets it on its own row.
Private Sub EnsureCounter(ByVal ws As Worksheet, ByVal r As Long)
    Dim wasProt As Boolean
    If ws.Cells(r, DLC_NUM).HasFormula Then Exit Sub
    wasProt = DltUnprotect(ws)
    On Error Resume Next
    ws.Cells(r, DLC_NUM).formula = CounterFormula(r)
    On Error GoTo 0
    DltReprotect ws, wasProt
End Sub

Private Function CounterFormula(ByVal r As Long) As String
    If r = DLT_FIRST Then
        CounterFormula = "=IF(D" & r & "="""","""",1)"
    Else
        CounterFormula = "=IF(D" & r & "="""","""",B" & (r - 1) & "+1)"
    End If
End Function

' The template's H formula: VLOOKUP the command on that DL page. A half-cleared
' log (or an older write that stored the setting as a value) leaves H blank;
' WriteEntry never fills H, so the formula has to be on the row.
Private Function SettingFormula(ByVal r As Long) As String
    Dim c As String, d As String, inner As String
    c = "C" & r
    d = "D" & r
    inner = "IF(" & c & "=0,VLOOKUP(" & d & ",$AA$3:$AB$66,2,FALSE)," & _
            "IF(" & c & "=1,VLOOKUP(" & d & ",$AC$3:$AD$66,2,FALSE)," & _
            "IF(" & c & "=2,VLOOKUP(" & d & ",$AE$3:$AF$66,2,FALSE)," & _
            "IF(" & c & "=3,VLOOKUP(" & d & ",$AG$3:$AH$66,2,FALSE)," & _
            "IF(" & c & "=7,VLOOKUP(" & d & ",$AI$3:$AJ$66,2,FALSE)," & _
            "IF(" & c & "=""S"",VLOOKUP(" & d & ",$AK$3:$AL$66,2,FALSE),""""))))))"
    SettingFormula = "=IF(" & c & "="""","""",IF(IFERROR(" & inner & ","""")=0,"""",IFERROR(" & inner & ","""")))"
End Function

Private Sub PutSettingFormula(ByVal ws As Worksheet, ByVal r As Long)
    Dim cell As Range
    Dim f As String
    f = SettingFormula(r)
    Set cell = ws.Cells(r, DLC_SET)
    If CStr(cell.formula) = f Then
        cell.Locked = True
        cell.FormulaHidden = True
        Exit Sub
    End If
    cell.formula = f
    cell.Locked = True
    cell.FormulaHidden = True
End Sub

Private Sub EnsureSettingFormula(ByVal ws As Worksheet, ByVal r As Long)
    Dim wasProt As Boolean
    If CStr(ws.Cells(r, DLC_SET).formula) = SettingFormula(r) Then Exit Sub
    wasProt = DltUnprotect(ws)
    On Error Resume Next
    PutSettingFormula ws, r
    On Error GoTo 0
    DltReprotect ws, wasProt
End Sub

' Restore D/L Setting formulas on every log row. Typed columns are not touched.
Public Function DLT_RepairSettingFormulas() As Long
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim r As Long, n As Long
    Dim f As String

    Set ws = DltSheet()
    If ws Is Nothing Then Exit Function
    wasProt = DltUnprotect(ws)
    On Error GoTo Restore
    For r = DLT_FIRST To DLT_LAST
        f = SettingFormula(r)
        If CStr(ws.Cells(r, DLC_SET).formula) <> f Then
            PutSettingFormula ws, r
            n = n + 1
        End If
    Next r
Restore:
    DltReprotect ws, wasProt
    DLT_RepairSettingFormulas = n
End Function

' ================================================================================
'  LOG LOOKUPS
' ================================================================================

Public Function DltSheet() As Worksheet
    On Error Resume Next
    Set DltSheet = ThisWorkbook.Worksheets(SH_DLT)
    If DltSheet Is Nothing Then Set DltSheet = ThisWorkbook.Worksheets(SH_DLT_OLD)
    On Error GoTo 0
End Function

Public Function DLT_IsTracker(ByVal sh As Object) As Boolean
    Dim nm As String
    On Error Resume Next
    nm = sh.name
    On Error GoTo 0
    DLT_IsTracker = (StrComp(nm, SH_DLT, vbTextCompare) = 0) Or _
                    (StrComp(nm, SH_DLT_OLD, vbTextCompare) = 0)
End Function

' Last log row with a command in D; DLT_FIRST - 1 when the log is empty.
' A (Run#) and Q:T are template-filled down the log, so they cannot mark it.
Public Function DLT_LastEntryRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(DLT_LAST + 1, DLC_CMD).End(xlUp).row
    If r > DLT_LAST Then r = DLT_LAST
    Do While r >= DLT_FIRST
        If Len(Trim$(CStr(ws.Cells(r, DLC_CMD).Value2 & ""))) > 0 Then Exit Do
        r = r - 1
    Loop
    If r < DLT_FIRST Then r = DLT_FIRST - 1
    DLT_LastEntryRow = r
End Function

' "0", "3", "S" ... (C holds numbers for numbered pages, text for S).
Public Function DLT_PageKey(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(Trim$(CStr(v & "")))
    If Len(s) > 0 And IsNumeric(s) Then s = CStr(CLng(CDbl(s)))
    DLT_PageKey = s
End Function

Private Function PageTableCol(ByVal PageKey As String) As Long
    Select Case PageKey
        Case "0": PageTableCol = 27   ' AA:AB Page 0 Manual
        Case "1": PageTableCol = 29   ' AC:AD Page 1 Vertical
        Case "2": PageTableCol = 31   ' AE:AF Page 2 IH
        Case "3": PageTableCol = 33   ' AG:AH Page 3 HIA
        Case "7": PageTableCol = 35   ' AI:AJ Page 7 Engineering
        Case "S": PageTableCol = 37   ' AK:AL Single Page Mode
        Case Else: PageTableCol = 0
    End Select
End Function

' D/L Setting text for a page + command, as the tab's H column shows it.
Public Function DLT_LookupSetting(ByVal ws As Worksheet, ByVal PageKey As String, _
                                  ByVal cmd As String) As String
    Dim c As Long, r As Long
    c = PageTableCol(PageKey)
    cmd = Trim$(cmd)
    If c = 0 Or Len(cmd) = 0 Then Exit Function
    For r = TBL_FIRST To TBL_LAST
        If StrComp(Trim$(ws.Cells(r, c).text), cmd, vbTextCompare) = 0 Then
            DLT_LookupSetting = Trim$(ws.Cells(r, c + 1).text)
            Exit Function
        End If
    Next r
End Function

' Supported commands for a page as a 2-column list (key, setting).
Public Function DLT_CommandList(ByVal ws As Worksheet, ByVal PageKey As String) As Variant
    Dim c As Long, r As Long, n As Long
    Dim k As String, s As String
    Dim tmp() As String, out() As String

    c = PageTableCol(PageKey)
    If c = 0 Then Exit Function
    ReDim tmp(1 To TBL_LAST - TBL_FIRST + 1, 1 To 2)
    For r = TBL_FIRST To TBL_LAST
        k = Trim$(ws.Cells(r, c).text)
        s = Trim$(ws.Cells(r, c + 1).text)
        If Len(k) > 0 And Len(s) > 0 And InStr(1, s, "ERROR", vbTextCompare) = 0 Then
            n = n + 1
            tmp(n, 1) = k
            tmp(n, 2) = s
        End If
    Next r
    If n = 0 Then Exit Function
    ReDim out(0 To n - 1, 0 To 1)
    For r = 1 To n
        out(r - 1, 0) = tmp(r, 1)
        out(r - 1, 1) = tmp(r, 2)
    Next r
    DLT_CommandList = out
End Function

' Values from a cell's list validation (range or comma list), blanks skipped.
Public Function DLT_ValidationList(ByVal c As Range) As Variant
    Dim f As String
    Dim src As Range, cell As Range
    Dim items() As String, parts() As String
    Dim n As Long, i As Long
    Dim t As String

    On Error GoTo Done
    f = c.Validation.Formula1
    If Len(f) = 0 Then GoTo Done
    ReDim items(0 To 0)
    If Left$(f, 1) = "=" Then
        Set src = c.Worksheet.Range(mid$(f, 2))
        ReDim items(0 To src.Cells.Count - 1)
        For Each cell In src.Cells
            t = Trim$(CStr(cell.Value2 & ""))
            If Len(t) > 0 Then items(n) = t: n = n + 1
        Next cell
    Else
        parts = Split(f, ",")
        ReDim items(0 To UBound(parts))
        For i = 0 To UBound(parts)
            t = Trim$(parts(i))
            If Len(t) > 0 Then items(n) = t: n = n + 1
        Next i
    End If
    If n = 0 Then GoTo Done
    ReDim Preserve items(0 To n - 1)
    DLT_ValidationList = items
Done:
End Function

' Two success codes per command, 64 apart (flow vs RPM detection; the tool
' cannot always tell them apart, so either one is a pass). 1-x -> 256+x or
' 320+x; 2-x -> 288+x or 352+x.
Public Function DLT_SuccessCodes(ByVal cmd As String, ByRef lo As Long, ByRef hi As Long) As Boolean
    Dim s As String, a As String, b As String
    Dim p As Long, g As Long, x As Long

    s = Trim$(cmd)
    p = InStr(s, "-")
    If p < 2 Then Exit Function
    a = Left$(s, p - 1)
    b = mid$(s, p + 1)
    If Not IsNumeric(a) Or Not IsNumeric(b) Then Exit Function
    g = CLng(a)
    x = CLng(b)
    If g < 1 Or g > 2 Or x < 0 Or x > 31 Then Exit Function
    lo = 256 + (g - 1) * 32 + x
    hi = lo + 64
    DLT_SuccessCodes = True
End Function

' Bit depth (Slidesheet D) on the deepest good survey row.
Public Function DLT_LastSurveyBitDepth(ByRef depth As Double) As Boolean
    Dim r As Long
    Dim v As Variant
    On Error GoTo Done
    r = MDL_ContDI.LastSlidesheetSurveyRow()
    If r <= 0 Then Exit Function
    v = ThisWorkbook.Worksheets("Slidesheet").Cells(r, "D").Value2
    If IsNumber(v) Then
        depth = CDbl(v)
        DLT_LastSurveyBitDepth = True
    End If
Done:
End Function

' ================================================================================
'  TF/INC (I) AND SR%/AZI (J) FROM THE D/L SETTING
' ================================================================================

' Baseline I/J for an increment: last same-page row, or INC/AZM typed on the
' inbound "Go to Page N" row. Stops at a page change so a return to HIA does
' not reuse the previous visit.
Public Function DLT_PrevIJ(ByVal ws As Worksheet, ByVal beforeRow As Long, ByVal PageKey As String, _
                           ByRef prevI As Double, ByRef prevJ As Double) As Boolean
    Dim r As Long
    Dim dest As String, rowPage As String
    For r = beforeRow - 1 To DLT_FIRST Step -1
        dest = DLT_GoToDest(RowSetting(ws, r))
        If Len(dest) > 0 Then
            If dest = PageKey Then
                If IsNumber(ws.Cells(r, DLC_I).Value2) And IsNumber(ws.Cells(r, DLC_J).Value2) Then
                    prevI = CDbl(ws.Cells(r, DLC_I).Value2)
                    prevJ = CDbl(ws.Cells(r, DLC_J).Value2)
                    DLT_PrevIJ = True
                End If
            End If
            Exit Function
        End If
        rowPage = DLT_PageKey(ws.Cells(r, DLC_PAGE).Value2)
        If rowPage = PageKey Then
            If IsNumber(ws.Cells(r, DLC_I).Value2) And IsNumber(ws.Cells(r, DLC_J).Value2) Then
                prevI = CDbl(ws.Cells(r, DLC_I).Value2)
                prevJ = CDbl(ws.Cells(r, DLC_J).Value2)
                DLT_PrevIJ = True
                Exit Function
            End If
        End If
    Next r
End Function

' D/L Setting text for log row r (table lookup, then the H formula result).
Private Function RowSetting(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim cmd As String
    cmd = Trim$(CStr(ws.Cells(r, DLC_CMD).Value2 & ""))
    RowSetting = DLT_LookupSetting(ws, DLT_PageKey(ws.Cells(r, DLC_PAGE).Value2), cmd)
    If Len(RowSetting) = 0 Then RowSetting = Trim$(CStr(ws.Cells(r, DLC_SET).Value2 & ""))
End Function

' "Go to Page 3: HIA Mode" -> "3". Blank when this row is not a page change.
Private Function DLT_GoToDest(ByVal setting As String) As String
    Dim ch As String
    If Left$(Trim$(setting), 11) <> "Go to Page " Then Exit Function
    ch = mid$(Trim$(setting), 12, 1)
    If ch Like "[0-9]" Then DLT_GoToDest = ch
End Function

' "Set TF = 324 degrees, SR = 100%"         -> I 324, J 100
' "Increase Target INC by 0.5 deg ..."       -> I prev + 0.5, J prev
' "Decrease Target AZI by 1.0 deg ..."       -> I prev, J prev - 1
' "Increase SR by 10%" / "... TF by 12 ..."  -> step the matching value
' "Go to Page 3: HIA Mode" / ERROR           -> blank
' anything else (bit period, gain, mode)     -> carry the previous values
Public Function DLT_ComputeIJ(ByVal setting As String, ByVal hasPrev As Boolean, _
                              ByVal prevI As Double, ByVal prevJ As Double, _
                              ByRef outI As Double, ByRef outJ As Double) As Long
    Dim s As String
    Dim dirn As Double, amt As Double

    s = Trim$(setting)
    DLT_ComputeIJ = DLT_IJ_KEEP
    If Len(s) = 0 Then Exit Function
    If InStr(1, s, "ERROR", vbTextCompare) > 0 Then DLT_ComputeIJ = DLT_IJ_BLANK: Exit Function

    If InStr(1, s, "TF =", vbTextCompare) > 0 And InStr(1, s, "SR =", vbTextCompare) > 0 Then
        If NumAfter(s, "TF =", outI) And NumAfter(s, "SR =", outJ) Then DLT_ComputeIJ = DLT_IJ_SET
        Exit Function
    End If
    If Left$(s, 10) = "Go to Page" Or Left$(s, 12) = "Go to Manual" Then
        DLT_ComputeIJ = DLT_IJ_BLANK
        Exit Function
    End If
    If Not hasPrev Then Exit Function

    outI = prevI
    outJ = prevJ
    DLT_ComputeIJ = DLT_IJ_SET
    dirn = 1#
    If InStr(1, s, "Decrease", vbTextCompare) > 0 Or InStr(1, s, "Nudge Down", vbTextCompare) > 0 Then dirn = -1#

    If NumAfter(s, "Target INC by", amt) Then
        outI = Round(prevI + dirn * amt, 2)
    ElseIf NumAfter(s, "Inclination Target", amt) Then
        outI = Round(prevI + dirn * amt, 2)
    ElseIf NumAfter(s, "Target AZI by", amt) Then
        outJ = Wrap360(prevJ + dirn * amt)
    ElseIf NumAfter(s, "SR by", amt) Then
        outJ = prevJ + dirn * amt
        If outJ < 0# Then outJ = 0#
        If outJ > 100# Then outJ = 100#
    ElseIf NumAfter(s, "TF by", amt) Then
        outI = Wrap360(prevI + dirn * amt)
    End If
End Function

' Recompute I/J on one log row from its page, command and Pass/Fail.
' A failed downlink leaves the tool where it was, so it carries the prior values.
Public Sub DLT_RefreshIJ(ByVal ws As Worksheet, ByVal r As Long)
    Dim cmd As String, PageKey As String, setting As String
    Dim pI As Double, pJ As Double, oI As Double, oJ As Double
    Dim hasPrev As Boolean

    cmd = Trim$(CStr(ws.Cells(r, DLC_CMD).Value2 & ""))
    If Len(cmd) = 0 Then Exit Sub
    PageKey = DLT_PageKey(ws.Cells(r, DLC_PAGE).Value2)
    hasPrev = DLT_PrevIJ(ws, r, PageKey, pI, pJ)
    If UCase$(Trim$(CStr(ws.Cells(r, DLC_PASS).Value2 & ""))) = "FAIL" Then
        If hasPrev Then WriteIJ ws, r, pI, pJ
        Exit Sub
    End If
    ' First command after a blank go-to: INC/AZM typed on this row is the start.
    If Not hasPrev Then
        If IsNumber(ws.Cells(r, DLC_I).Value2) And IsNumber(ws.Cells(r, DLC_J).Value2) Then
            pI = CDbl(ws.Cells(r, DLC_I).Value2)
            pJ = CDbl(ws.Cells(r, DLC_J).Value2)
            hasPrev = True
        End If
    End If
    setting = DLT_LookupSetting(ws, PageKey, cmd)
    Select Case DLT_ComputeIJ(setting, hasPrev, pI, pJ, oI, oJ)
        Case DLT_IJ_SET: WriteIJ ws, r, oI, oJ
        Case DLT_IJ_BLANK: ws.Range(ws.Cells(r, DLC_I), ws.Cells(r, DLC_J)).ClearContents
    End Select
End Sub

' Workbook_SheetChange hook: C (page), D (command) or P (Pass/Fail) typed on
' the log -> recompute that row's I/J. Bulk pastes / clears are left alone.
Public Sub DLT_OnSheetChange(ByVal ws As Worksheet, ByVal Target As Range)
    Dim hit As Range, c As Range
    Dim ev As Boolean
    Dim lastRow As Long

    Set hit = Application.Intersect(Target, ws.Range( _
        "C" & DLT_FIRST & ":D" & DLT_LAST & ",P" & DLT_FIRST & ":P" & DLT_LAST))
    If hit Is Nothing Then Exit Sub
    If hit.Cells.Count > 60 Then Exit Sub
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done
    lastRow = 0
    For Each c In hit.Cells
        If c.row <> lastRow Then
            EnsureSettingFormula ws, c.row
            DLT_RefreshIJ ws, c.row
            lastRow = c.row
        End If
    Next c
Done:
    Application.EnableEvents = ev
End Sub

' ================================================================================
'  ENTRY WRITE
' ================================================================================

' vals(1 To 21) by log column; Empty clears. B and H (formulas) are never
' written, nor any locked cell.
Public Sub DLT_WriteEntry(ByVal ws As Worksheet, ByVal r As Long, ByRef vals() As Variant)
    Dim c As Long
    Dim cell As Range
    Dim ev As Boolean

    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done
    EnsureCounter ws, r
    EnsureSettingFormula ws, r
    For c = DLC_RUN To DLC_NOTE
        If c <> DLC_NUM And c <> DLC_SET Then
            Set cell = ws.Cells(r, c)
            If Not cell.Locked And Not cell.HasFormula Then
                If IsEmpty(vals(c)) Then
                    cell.ClearContents
                Else
                    cell.Value2 = vals(c)
                End If
            End If
        End If
    Next c
Done:
    Application.EnableEvents = ev
    If Err.Number <> 0 Then Err.Raise Err.Number, "DLT_WriteEntry", Err.Description
End Sub

' ================================================================================
'  HELPERS
' ================================================================================

Private Sub WriteIJ(ByVal ws As Worksheet, ByVal r As Long, ByVal vI As Double, ByVal vJ As Double)
    ws.Cells(r, DLC_I).Value2 = vI
    ws.Cells(r, DLC_J).Value2 = vJ
End Sub

' First number after key in s ("TF = 324 degrees" -> 324).
Private Function NumAfter(ByVal s As String, ByVal key As String, ByRef v As Double) As Boolean
    Dim p As Long, i As Long
    Dim ch As String, num As String

    p = InStr(1, s, key, vbTextCompare)
    If p = 0 Then Exit Function
    i = p + Len(key)
    Do While i <= Len(s)
        ch = mid$(s, i, 1)
        If ch Like "[0-9.]" Then Exit Do
        If ch <> " " Then Exit Function
        i = i + 1
    Loop
    Do While i <= Len(s)
        ch = mid$(s, i, 1)
        If Not ch Like "[0-9.]" Then Exit Do
        num = num & ch
        i = i + 1
    Loop
    If Len(num) = 0 Or Not IsNumeric(num) Then Exit Function
    v = val(num)
    NumAfter = True
End Function

Private Function Wrap360(ByVal a As Double) As Double
    a = a - 360# * Int(a / 360#)
    Wrap360 = Round(a, 2)
End Function

Private Function IsNumber(ByVal v As Variant) As Boolean
    If isError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    IsNumber = IsNumeric(v)
End Function

Private Function DltUnprotect(ByVal ws As Worksheet) As Boolean
    If Not ws.ProtectContents Then Exit Function
    mPDraw = ws.ProtectDrawingObjects
    mPScen = ws.ProtectScenarios
    With ws.Protection
        mPFmtCells = .AllowFormattingCells
        mPFmtCols = .AllowFormattingColumns
        mPFmtRows = .AllowFormattingRows
        mPInsCols = .AllowInsertingColumns
        mPInsRows = .AllowInsertingRows
        mPInsLinks = .AllowInsertingHyperlinks
        mPDelCols = .AllowDeletingColumns
        mPDelRows = .AllowDeletingRows
        mPSort = .AllowSorting
        mPFilter = .AllowFiltering
        mPPivot = .AllowUsingPivotTables
    End With
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PROTECT_PWD
    If ws.ProtectContents Then ws.Unprotect Password:=""
    On Error GoTo 0
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 701, "MDL_DownlinkTracker", "Could not unprotect '" & ws.name & "'."
    End If
    DltUnprotect = True
End Function

' ================================================================================
'  DROPDOWN MOUSE WHEEL (frmDLEntry)
' ================================================================================

' The combo whose list is opening. DropButtonClick on the form reports it,
' because ActiveControl is not always that combo while the list is open.
Public Sub DLT_WheelCombo(ByVal controlName As String)
    mWheelComboName = controlName
End Sub

Public Function DLT_WheelReady() As Boolean
    DLT_WheelReady = (mWheelHook <> 0)
End Function

Public Sub DLT_WheelOn()
    If mWheelHook <> 0 Then Exit Sub
    mWheelHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf DLT_WheelProc, 0, 0)
End Sub

Public Sub DLT_WheelOff()
    If mWheelHook = 0 Then Exit Sub
    UnhookWindowsHookEx mWheelHook
    mWheelHook = 0
End Sub

' Every mouse event while the form is open comes through here. Only a wheel
' notch is handled, and only when a dropdown list is actually open.
Public Function DLT_WheelProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    On Error GoTo Pass
    If nCode < 0 Or mWheelBusy Then GoTo Pass
    If wParam = WM_MOUSEWHEEL Then
        mWheelBusy = True
        If ScrollOpenCombo(lParam) Then
            mWheelBusy = False
            DLT_WheelProc = 1
            Exit Function
        End If
        mWheelBusy = False
    End If
Pass:
    DLT_WheelProc = CallNextHookEx(mWheelHook, nCode, wParam, lParam)
End Function

' The popup ignores wheel and key messages. Moving the open combo's ListIndex
' keeps the list open and scrolls the highlight. Positive wheel goes up the list.
Private Function ScrollOpenCombo(ByVal lParam As LongPtr) As Boolean
    Dim w As Integer
    Dim lines As Long
    Dim cbo As Object
    Dim i As Long

    If FindPopup() = 0 Then Exit Function
    CopyMemory w, ByVal (lParam + 10), 2
    If w = 0 Then Exit Function
    lines = CLng(w) \ 120
    If lines = 0 Then
        If w > 0 Then lines = 1 Else lines = -1
    Else
        lines = lines * 3
    End If
    If lines > 12 Then lines = 12
    If lines < -12 Then lines = -12
    On Error GoTo Done
    If Len(mWheelComboName) > 0 Then Set cbo = frmDLEntry.Controls(mWheelComboName)
    If cbo Is Nothing Then Set cbo = frmDLEntry.ActiveControl
    If cbo Is Nothing Then Exit Function
    If TypeName(cbo) <> "ComboBox" Then Exit Function
    If cbo.ListCount < 1 Then Exit Function
    i = cbo.listIndex - lines
    If i < 0 Then i = 0
    If i > cbo.ListCount - 1 Then i = cbo.ListCount - 1
    If i <> cbo.listIndex Then cbo.listIndex = i
    ScrollOpenCombo = True
Done:
End Function

' The open MSForms list is a top-level window whose class starts with
' "F3 MdcPopup". The rest of the class name changes every Excel session.
Private Function FindPopup() As LongPtr
    Dim h As LongPtr
    Dim buf As String
    Dim n As Long
    Dim pid As Long
    Const GW_CHILD As Long = 5
    Const GW_HWNDNEXT As Long = 2

    h = GetWindow(GetDesktopWindow(), GW_CHILD)
    Do While h <> 0
        pid = 0
        GetWindowThreadProcessId h, pid
        If pid = GetCurrentProcessId() And IsWindowVisible(h) <> 0 Then
            buf = String$(160, 0)
            n = GetClassNameW(h, StrPtr(buf), 160)
            If n >= 11 Then
                If Left$(buf, 11) = "F3 MdcPopup" Then
                    FindPopup = h
                    Exit Function
                End If
            End If
        End If
        h = GetWindow(h, GW_HWNDNEXT)
    Loop
End Function

Private Sub DltReprotect(ByVal ws As Worksheet, ByVal wasProt As Boolean)
    If Not wasProt Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=SHEET_PROTECT_PWD, DrawingObjects:=mPDraw, Contents:=True, _
               Scenarios:=mPScen, AllowFormattingCells:=mPFmtCells, _
               AllowFormattingColumns:=mPFmtCols, AllowFormattingRows:=mPFmtRows, _
               AllowInsertingColumns:=mPInsCols, AllowInsertingRows:=mPInsRows, _
               AllowInsertingHyperlinks:=mPInsLinks, AllowDeletingColumns:=mPDelCols, _
               AllowDeletingRows:=mPDelRows, AllowSorting:=mPSort, _
               AllowFiltering:=mPFilter, AllowUsingPivotTables:=mPPivot
End Sub








