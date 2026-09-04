Attribute VB_Name = "MDL_ToolHours"
Option Explicit

' ================================================================================
'  MDL_ToolHours - Data tab: 3rd-party tool + motor "Current Hours" from Q24
' ================================================================================
' Data layout this module relies on:
'   H3              selected BHA#
'   Q24             BHA TOTALS circ hours  (=SUM(Q4:Q22))  -> hours on this BHA
'   C22             selected motor S/N (dropdown fed by B45:B55)
'   D24:D32         3rd-party tools picked for this BHA (dropdowns fed by O31:O38;
'                   D29 is the "Activated Agitator Hours" header and never matches)
'   O28:S30 title   "Enter any 3rd Party Tools/hours below"
'   O33:P38         3rd-party serial | Q previous | R current | S total (=Q+R)
'   O39:S41 title   "Enter Motors & hours below"
'   O42:P55         motor serial     | Q previous | R current | S total (=Q+R)
'
' Rule: a serial is "on the selected BHA" when it is picked in D24:D32, or it
' appears in _OC_BHA (Serial #) for BHA# = H3. OpenCap's export does not carry
' FieldCap "Other Tools" (third-party) components - those BHA rows come through
' with no serial - so the D-column picks are the authority for rentals OpenCap
' does not know about. Matched rows get R = Q24. Unmatched rows are left alone
' (their R keeps whatever it held), and Q "previous hours" is never touched.
'
' Motor: the row in O42:O55 whose serial = C22 gets R = Q24. If C22 is blank the
' motor is taken from _OC_BHA (Description contains "Mud Motor") for BHA# = H3.
'
' Entry points:
'   ToolHours_Sync          full recompute (RefreshData, Pason form, manual)
'   ToolHours_OnDataChange  Data Worksheet_Change hook (cheap range gate)
' ================================================================================

Private Const SH_DATA As String = "Data"
Private Const SH_BHA As String = "_OC_BHA"

Private Const CELL_BHA As String = "H3"
Private Const CELL_HOURS As String = "Q24"
Private Const CELL_MOTOR As String = "C22"
Private Const RNG_PICKS As String = "D24:D32"
Private Const RNG_HOURS_SRC As String = "P4:Q22"

Private Const COL_SERIAL As Long = 15      ' O (merged O:P)
Private Const COL_CURRENT As Long = 18     ' R
Private Const TOOLS_TITLE_TOKEN As String = "3rd Party"
Private Const TOOLS_TITLE_TO_DATA As Long = 5   ' O28 title -> O33 first row
Private Const MOTOR_TITLE_TOKEN As String = "Enter Motors"
Private Const MOTOR_TITLE_TO_DATA As Long = 3   ' O39 title -> O42 first row
Private Const MOTOR_ROWS As Long = 14           ' O42:O55

Private Const TITLE_TYPO As String = "Toosl"
Private Const TITLE_FIX As String = "Tools"

Private mBusy As Boolean

' Data Worksheet_Change hook. Only recomputes when an input cell moved.
Public Sub ToolHours_OnDataChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim watch As Range
    Dim toolsFirst As Long, toolsLast As Long
    Dim motorFirst As Long, motorLast As Long

    On Error GoTo Quiet
    If mBusy Then Exit Sub
    If Target Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    If Not Target.Worksheet Is ws Then Exit Sub

    LocateTables ws, toolsFirst, toolsLast, motorFirst, motorLast

    Set watch = Union(ws.Range(CELL_BHA), ws.Range(CELL_MOTOR), _
                      ws.Range(RNG_PICKS), ws.Range(RNG_HOURS_SRC), _
                      ws.Range(ws.Cells(toolsFirst, COL_SERIAL), ws.Cells(motorLast, COL_SERIAL)))
    If Intersect(Target, watch) Is Nothing Then Exit Sub

    ToolHours_Sync
    Exit Sub
Quiet:
    ' Never raise out of Worksheet_Change
End Sub

' Full recompute of R (Current Hours) for the selected BHA's tools + motor.
Public Sub ToolHours_Sync()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Dim hrs As Double
    Dim bha As Long
    Dim motorSn As String
    Dim picks As Collection
    Dim bhaSerials As Collection
    Dim toolsFirst As Long, toolsLast As Long
    Dim motorFirst As Long, motorLast As Long
    Dim r As Long
    Dim sn As String

    If mBusy Then Exit Sub
    If Not SheetExistsTH(SH_DATA) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_DATA)

    mBusy = True
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail

    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail

    hrs = NumOrZero(ws.Range(CELL_HOURS).Value)
    bha = BhaNumber(ws.Range(CELL_BHA).Value)
    motorSn = Trim$(CStr(ws.Range(CELL_MOTOR).Value & ""))

    Set picks = CollectPicks(ws)
    Set bhaSerials = CollectBhaSerials(bha)
    If Len(motorSn) = 0 Then motorSn = BhaMotorSerial(bha)

    LocateTables ws, toolsFirst, toolsLast, motorFirst, motorLast

    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo FailProt

    FixTitleSpelling ws

    ' 3rd-party tools: on this BHA if picked in D24:D32 or listed in _OC_BHA
    For r = toolsFirst To toolsLast
        sn = SerialAt(ws, r)
        If Len(sn) > 0 Then
            If InSet(picks, sn) Or InSet(bhaSerials, sn) Then
                WriteCurrent ws, r, hrs
            End If
        End If
    Next r

    ' Motor: the selected motor S/N only
    If Len(motorSn) > 0 Then
        For r = motorFirst To motorLast
            sn = SerialAt(ws, r)
            If Len(sn) > 0 Then
                If StrComp(sn, motorSn, vbTextCompare) = 0 Then WriteCurrent ws, r, hrs
            End If
        Next r
    End If

    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
    mBusy = False
    Exit Sub

FailProt:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
Fail:
    On Error Resume Next
    Application.EnableEvents = prevEvents
    mBusy = False
End Sub

' ------------------------------------------------------------------------------
'  Internals
' ------------------------------------------------------------------------------

' Find both tables from their title text in column O so a row shift does not
' silently point us at the wrong cells. Falls back to the documented layout.
Private Sub LocateTables(ByVal ws As Worksheet, ByRef toolsFirst As Long, _
                         ByRef toolsLast As Long, ByRef motorFirst As Long, _
                         ByRef motorLast As Long)
    Dim tTitle As Range, mTitle As Range

    Set tTitle = FindInColumn(ws, COL_SERIAL, TOOLS_TITLE_TOKEN)
    Set mTitle = FindInColumn(ws, COL_SERIAL, MOTOR_TITLE_TOKEN)

    If tTitle Is Nothing Then toolsFirst = 33 Else toolsFirst = tTitle.Row + TOOLS_TITLE_TO_DATA
    If mTitle Is Nothing Then
        motorFirst = 42
    Else
        motorFirst = mTitle.Row + MOTOR_TITLE_TO_DATA
    End If
    motorLast = motorFirst + MOTOR_ROWS - 1

    If mTitle Is Nothing Then
        toolsLast = 38
    Else
        toolsLast = mTitle.Row - 1
    End If
    If toolsLast < toolsFirst Then toolsLast = toolsFirst
End Sub

Private Function FindInColumn(ByVal ws As Worksheet, ByVal col As Long, _
                              ByVal token As String) As Range
    On Error Resume Next
    Set FindInColumn = ws.Columns(col).Find(What:=token, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)
    On Error GoTo 0
End Function

' "Enter any 3rd Party Toosl/hours below" -> "... Tools/hours below"
Private Sub FixTitleSpelling(ByVal ws As Worksheet)
    Dim c As Range
    Dim txt As String
    Set c = FindInColumn(ws, COL_SERIAL, TITLE_TYPO)
    If c Is Nothing Then Exit Sub
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    txt = CStr(c.Value & "")
    If InStr(1, txt, TITLE_TYPO, vbTextCompare) = 0 Then Exit Sub
    c.Value = Replace(txt, TITLE_TYPO, TITLE_FIX, 1, -1, vbTextCompare)
End Sub

Private Function SerialAt(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim c As Range
    Set c = ws.Cells(r, COL_SERIAL)
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    SerialAt = Trim$(CStr(c.Value & ""))
End Function

Private Sub WriteCurrent(ByVal ws As Worksheet, ByVal r As Long, ByVal hrs As Double)
    Dim c As Range
    Set c = ws.Cells(r, COL_CURRENT)
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    ' IsNumeric(Empty) is True, so test blank first - a blank cell must become 0.
    If Not IsEmpty(c.Value) Then
        If IsNumeric(c.Value) Then
            If Abs(CDbl(c.Value) - hrs) < 0.000001 Then Exit Sub
        End If
    End If
    c.Value = hrs
End Sub

' Serials picked in the "3rd Party Tools" / "Activated Agitator" dropdowns.
Private Function CollectPicks(ByVal ws As Worksheet) As Collection
    Dim col As New Collection
    Dim c As Range
    Dim sn As String
    For Each c In ws.Range(RNG_PICKS).Cells
        sn = Trim$(CStr(c.Value & ""))
        If Len(sn) > 0 Then AddUnique col, sn
    Next c
    Set CollectPicks = col
End Function

' Every Serial # that _OC_BHA lists for this BHA#.
Private Function CollectBhaSerials(ByVal bha As Long) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    Dim cNum As Long, cSn As Long
    Dim lastR As Long, r As Long
    Dim sn As String

    Set CollectBhaSerials = col
    If bha <= 0 Then Exit Function
    If Not SheetExistsTH(SH_BHA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    cNum = HeaderCol(ws, "BHA #")
    cSn = HeaderCol(ws, "Serial #")
    If cNum = 0 Or cSn = 0 Then Exit Function

    lastR = ws.Cells(ws.Rows.Count, cNum).End(xlUp).Row
    For r = 2 To lastR
        If BhaNumber(ws.Cells(r, cNum).Value) = bha Then
            sn = Trim$(CStr(ws.Cells(r, cSn).Value & ""))
            If Len(sn) > 0 Then AddUnique col, sn
        End If
    Next r
End Function

' Motor serial for this BHA# from _OC_BHA (Description contains "Mud Motor").
Private Function BhaMotorSerial(ByVal bha As Long) As String
    Dim ws As Worksheet
    Dim cNum As Long, cSn As Long, cDesc As Long
    Dim lastR As Long, r As Long
    Dim sn As String

    BhaMotorSerial = ""
    If bha <= 0 Then Exit Function
    If Not SheetExistsTH(SH_BHA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    cNum = HeaderCol(ws, "BHA #")
    cSn = HeaderCol(ws, "Serial #")
    cDesc = HeaderCol(ws, "Description")
    If cNum = 0 Or cSn = 0 Or cDesc = 0 Then Exit Function

    lastR = ws.Cells(ws.Rows.Count, cNum).End(xlUp).Row
    For r = 2 To lastR
        If BhaNumber(ws.Cells(r, cNum).Value) = bha Then
            If InStr(1, CStr(ws.Cells(r, cDesc).Value & ""), "Mud Motor", vbTextCompare) > 0 Then
                sn = Trim$(CStr(ws.Cells(r, cSn).Value & ""))
                If Len(sn) > 0 Then
                    BhaMotorSerial = sn
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function HeaderCol(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim lastC As Long, c As Long
    HeaderCol = 0
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value & "")), header, vbTextCompare) = 0 Then
            HeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Sub AddUnique(ByVal col As Collection, ByVal sn As String)
    On Error Resume Next
    col.Add sn, UCase$(sn)
    On Error GoTo 0
End Sub

Private Function InSet(ByVal col As Collection, ByVal sn As String) As Boolean
    Dim v As Variant
    InSet = False
    If col Is Nothing Then Exit Function
    On Error Resume Next
    v = col.Item(UCase$(sn))
    InSet = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function BhaNumber(ByVal v As Variant) As Long
    BhaNumber = 0
    On Error Resume Next
    If IsNumeric(v) Then BhaNumber = CLng(v)
    On Error GoTo 0
End Function

Private Function NumOrZero(ByVal v As Variant) As Double
    NumOrZero = 0#
    On Error Resume Next
    If IsNumeric(v) Then NumOrZero = CDbl(v)
    On Error GoTo 0
End Function

Private Function SheetExistsTH(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    SheetExistsTH = Not ws Is Nothing
End Function
