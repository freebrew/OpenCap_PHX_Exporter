Attribute VB_Name = "MDL_CostsForm"
Option Explicit

' Calendar / edit UI for the Costs sheet (sheet stays hidden like Pipe Tally).
' REFRESH / SyncCostsFromOpenCap still owns CSV merge (CSV wins on same date).

Private Const SH_COSTS_TAB As String = "Costs"
Private Const DATA_SHEET As String = "Data"
Private Const COSTS_FIRST_ROW As Long = 3
Private Const COSTS_LAST_ROW As Long = 72
Private Const COSTS_COL_DATE As Long = 3   ' C
Private Const COSTS_COL_DAILY As Long = 5  ' E
Private Const COSTS_COL_TOTAL As Long = 7  ' G
Private Const COSTS_COL_BATCH As Long = 8  ' H
Private Const BTN_COSTS As String = "btnCostsForm"
Private Const BTN_FULL As String = "Button 73"
Private Const BTN_REFRESH As String = "BtnImport"

Public Sub ShowCostsForm()
    On Error GoTo Fail
    HideCostsSheet
    If CostsSheet() Is Nothing Then
        MsgBox "Costs sheet not found in this workbook.", vbExclamation, "Costs"
        Exit Sub
    End If
    frmCosts.Show vbModeless
    Exit Sub
Fail:
    MsgBox "Could not open Costs form." & vbCrLf & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Costs"
End Sub

Public Sub CloseCostsFormIfOpen()
    On Error Resume Next
    Unload frmCosts
    On Error GoTo 0
End Sub

Public Sub HideCostsSheet()
    Dim ws As Worksheet
    Set ws = CostsSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Visible = xlSheetHidden
    On Error GoTo 0
End Sub

Public Sub EnsureCostsUi()
    On Error Resume Next
    HideCostsSheet
    On Error GoTo 0
    EnsureCostsFormButton
End Sub

' Place Costs between Full Report and REFRESH on the Data toolbar.
Public Sub EnsureCostsFormButton()
    Dim ws As Worksheet
    Dim fullBtn As Button
    Dim refBtn As Button
    Dim btn As Button
    Dim costW As Double
    Dim topPt As Double
    Dim hPt As Double
    Dim fullRight As Double
    Dim wasProt As Boolean
    Dim errN As Long
    Dim errD As String

    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)

    On Error Resume Next
    Set fullBtn = ws.Buttons(BTN_FULL)
    Set refBtn = ws.Buttons(BTN_REFRESH)
    On Error GoTo 0
    If fullBtn Is Nothing Or refBtn Is Nothing Then
        Err.Raise vbObjectError + 810, "EnsureCostsFormButton", _
            "Data toolbar anchors missing (need '" & BTN_FULL & "' and '" & BTN_REFRESH & "')."
    End If

    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo FailBtn

    costW = 55#
    topPt = fullBtn.Top
    hPt = fullBtn.Height
    fullRight = fullBtn.Left + fullBtn.Width

    ' Shift REFRESH right to make room
    refBtn.Left = fullRight + costW
    refBtn.Top = topPt
    refBtn.Height = hPt

    On Error Resume Next
    ws.Buttons(BTN_COSTS).Delete
    On Error GoTo FailBtn

    Set btn = ws.Buttons.Add(fullRight, topPt, costW, hPt)
    btn.name = BTN_COSTS
    btn.caption = "Costs"
    btn.OnAction = "'" & ThisWorkbook.name & "'!ShowCostsForm"
    btn.Placement = xlMoveAndSize

    On Error Resume Next
    btn.Font.bold = True
    btn.Font.Size = 9
    btn.Font.name = "Consolas"
    On Error GoTo 0

    SheetReprotectAfterVba ws, wasProt
    Exit Sub

FailBtn:
    errN = Err.Number
    errD = Err.Description
    SheetReprotectAfterVba ws, wasProt
    Err.Raise errN, "EnsureCostsFormButton", errD
End Sub

Public Function CostsSheet() As Worksheet
    On Error Resume Next
    Set CostsSheet = ThisWorkbook.Worksheets(SH_COSTS_TAB)
    On Error GoTo 0
End Function

' Returns Dictionary dateSerial(String) -> daily cost (Double)
Public Function CostsDateMap() As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim dSerial As Long
    Dim seed As Double

    Set dict = CreateObject("Scripting.Dictionary")
    Set CostsDateMap = dict
    Set ws = CostsSheet()
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    ws.Calculate
    On Error GoTo 0

    seed = 0#
    If IsNumeric(ws.Cells(COSTS_FIRST_ROW, COSTS_COL_DATE).Value2) Then
        seed = CDbl(ws.Cells(COSTS_FIRST_ROW, COSTS_COL_DATE).Value2)
    End If

    For r = COSTS_FIRST_ROW To COSTS_LAST_ROW
        If Len(Trim$(CStr(ws.Cells(r, COSTS_COL_DAILY).Value & ""))) = 0 Then GoTo NextRow
        dSerial = ParseCostDateLocal(ws.Cells(r, COSTS_COL_DATE).Value)
        If dSerial = 0 Then dSerial = ParseCostDateLocal(ws.Cells(r, COSTS_COL_DATE).Value2)
        If dSerial = 0 And seed > 0 Then dSerial = CLng(seed) + (r - COSTS_FIRST_ROW)
        If dSerial > 0 Then
            If IsNumeric(ws.Cells(r, COSTS_COL_DAILY).Value) Then
                dict(CStr(dSerial)) = CDbl(ws.Cells(r, COSTS_COL_DAILY).Value)
            End If
        End If
NextRow:
    Next r
End Function

Public Function CostsFindRow(ByVal dateSerial As Long) As Long
    Dim ws As Worksheet
    Dim r As Long
    Dim dSerial As Long
    Dim seed As Double

    CostsFindRow = 0
    Set ws = CostsSheet()
    If ws Is Nothing Or dateSerial <= 0 Then Exit Function

    On Error Resume Next
    ws.Calculate
    On Error GoTo 0

    seed = 0#
    If IsNumeric(ws.Cells(COSTS_FIRST_ROW, COSTS_COL_DATE).Value2) Then
        seed = CDbl(ws.Cells(COSTS_FIRST_ROW, COSTS_COL_DATE).Value2)
    End If

    For r = COSTS_FIRST_ROW To COSTS_LAST_ROW
        dSerial = ParseCostDateLocal(ws.Cells(r, COSTS_COL_DATE).Value)
        If dSerial = 0 Then dSerial = ParseCostDateLocal(ws.Cells(r, COSTS_COL_DATE).Value2)
        If dSerial = 0 And seed > 0 And Len(Trim$(CStr(ws.Cells(r, COSTS_COL_DAILY).Value & ""))) > 0 Then
            dSerial = CLng(seed) + (r - COSTS_FIRST_ROW)
        End If
        If dSerial = dateSerial Then
            CostsFindRow = r
            Exit Function
        End If
    Next r
End Function

Public Sub CostsGetDay(ByVal dateSerial As Long, _
        ByRef daily As Variant, ByRef total As Variant, ByRef batch As Variant, _
        ByRef found As Boolean)
    Dim ws As Worksheet
    Dim r As Long

    daily = Empty: total = Empty: batch = Empty
    found = False
    r = CostsFindRow(dateSerial)
    If r = 0 Then Exit Sub

    Set ws = CostsSheet()
    On Error Resume Next
    ws.Calculate
    On Error GoTo 0

    found = True
    daily = ws.Cells(r, COSTS_COL_DAILY).Value
    total = ws.Cells(r, COSTS_COL_TOTAL).Value
    batch = ws.Cells(r, COSTS_COL_BATCH).Value
End Sub

' Write daily cost (and optional batch) for an existing or new trailing day row.
Public Function CostsSetDay(ByVal dateSerial As Long, ByVal daily As Double, _
                            ByVal batchText As String) As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim wasProt As Boolean
    Dim firstEmpty As Long
    Dim isNew As Boolean
    Dim prevSerial As Long

    CostsSetDay = False
    Set ws = CostsSheet()
    If ws Is Nothing Or dateSerial <= 0 Then Exit Function

    r = CostsFindRow(dateSerial)
    isNew = False
    If r = 0 Then
        ' Append on first empty daily cell in the day block
        firstEmpty = COSTS_FIRST_ROW
        Do While firstEmpty <= COSTS_LAST_ROW
            If Len(Trim$(CStr(ws.Cells(firstEmpty, COSTS_COL_DAILY).Value & ""))) = 0 Then Exit Do
            firstEmpty = firstEmpty + 1
        Loop
        If firstEmpty > COSTS_LAST_ROW Then Exit Function
        r = firstEmpty
        isNew = True
    End If

    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo FailWrite

    ws.Cells(r, COSTS_COL_DAILY).numberFormat = "$#,##0.00"
    ws.Cells(r, COSTS_COL_DAILY).Value = daily

    If isNew Then
        prevSerial = 0
        If r > COSTS_FIRST_ROW Then
            prevSerial = ParseCostDateLocal(ws.Cells(r - 1, COSTS_COL_DATE).Value)
            If prevSerial = 0 Then prevSerial = ParseCostDateLocal(ws.Cells(r - 1, COSTS_COL_DATE).Value2)
        End If
        With ws.Cells(r, COSTS_COL_DATE).MergeArea.Cells(1, 1)
            .numberFormat = "[$-F800]dddd, mmmm dd, yyyy"
            If r = COSTS_FIRST_ROW Or prevSerial = 0 Or dateSerial <> prevSerial + 1 Then
                .Value = dateSerial
            Else
                .Formula = "=IF(E" & r & "<>"""",C" & (r - 1) & "+1,"""")"
            End If
        End With

        ws.Cells(r, COSTS_COL_TOTAL).numberFormat = "$#,##0.00"
        If r = COSTS_FIRST_ROW Then
            ws.Cells(r, COSTS_COL_TOTAL).Formula = _
                "=IF(ISBLANK(E" & r & "),"""",E" & r & ")"
        Else
            ws.Cells(r, COSTS_COL_TOTAL).Formula = _
                "=IF(ISBLANK(E" & r & "),"""",SUM($E$" & COSTS_FIRST_ROW & ":E" & r & "))"
        End If
    End If

    ws.Cells(r, COSTS_COL_BATCH).Value = Trim$(batchText)

    On Error Resume Next
    ws.Calculate
    On Error GoTo 0
    SheetReprotectAfterVba ws, wasProt
    CostsSetDay = True
    Exit Function

FailWrite:
    SheetReprotectAfterVba ws, wasProt
    CostsSetDay = False
End Function

Public Function ParseCostDateLocal(ByVal v As Variant) As Long
    ParseCostDateLocal = 0
    On Error GoTo Fail
    If isError(v) Then Exit Function
    If IsDate(v) Then
        ParseCostDateLocal = CLng(CDate(v))
        Exit Function
    End If
    If IsNumeric(v) Then
        Dim n As Double
        n = CDbl(v)
        If n >= 30000 And n < 100000 Then
            ParseCostDateLocal = CLng(n)
            Exit Function
        End If
    End If
    If Len(Trim$(CStr(v & ""))) > 0 Then
        If IsDate(CStr(v)) Then ParseCostDateLocal = CLng(CDate(CStr(v)))
    End If
Fail:
End Function



