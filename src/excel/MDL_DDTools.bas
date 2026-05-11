Option Explicit

' ================================================================================
'  MODULE: MDL_DDTools
'  FieldCap DD Tools - Excel Dashboard
'
'  QUICK START:
'   1. Alt+F11 > Insert > Module > paste this file
'   2. Run InitDDTools once from the Immediate window:  InitDDTools
'   3. Use the REFRESH button on the DD Tools sheet
'
'  ARCHITECTURE:
'   - "DD Tools"  = visible dashboard sheet (created automatically)
'   - "_FC_BHA"   = hidden raw BHA CSV data
'   - "_FC_Crew"  = hidden raw Crew CSV data
'   - "_FC_Job"   = hidden raw Job Details CSV data
'  All hidden sheets use xlSheetVeryHidden
'  All rendering uses ScreenUpdating=False + Calculation=Manual
' ================================================================================

' Sheet names
Private Const SH_UI   As String = "DD Tools"
Private Const SH_BHA  As String = "_FC_BHA"
Private Const SH_CREW As String = "_FC_Crew"
Private Const SH_JOB  As String = "_FC_Job"

' Layout: row anchors (tight, no wasted space)
Private Const R_TOOLBAR As Long = 1   ' title + all buttons in one row
Private Const R_JOB     As Long = 2   ' job info subtitle
Private Const R_KPI     As Long = 3   ' KPI stat cards
Private Const R_DIV     As Long = 4   ' thin accent divider / freeze line
Private Const R_DETAIL  As Long = 5   ' BHA detail starts here

' Layout: column span
Private Const C_LAST As Long = 14
' Main inventory/cumulative tables width — trim section bars to this many columns (A:H)
Private Const COL_DATA_LAST As Long = 8

' Right-side summary block: four columns J:M (gap column I after main table A:H)
Private Const COL_SUMMARY_L As Long = 10    ' J:K = KPI-style metrics
Private Const COL_SUMMARY_R As Long = 12    ' L:M = BHA detail metrics
Private Const COL_SUMMARY_END As Long = 13   ' M
' Raw BHA summary — right of KPI island J:M with gap col N; table starts at O
Private Const COL_RAW_SUMMARY_START As Long = 15  ' O (14 = spacer after M)

' Colors: clean, professional, minimal
Private Function cWhite() As Long: cWhite = RGB(255, 255, 255): End Function
Private Function cGrayBg() As Long: cGrayBg = RGB(245, 245, 245): End Function
Private Function cGrayMed() As Long: cGrayMed = RGB(200, 200, 200): End Function
Private Function cGrayDk() As Long: cGrayDk = RGB(120, 120, 120): End Function
Private Function cBlack() As Long: cBlack = RGB(30, 30, 30): End Function
Private Function cRed() As Long: cRed = RGB(180, 40, 40): End Function
Private Function cTeal() As Long: cTeal = RGB(0, 79, 79): End Function
Private Function cTealLt() As Long: cTealLt = RGB(0, 110, 110): End Function
Private Function cAccentLine() As Long: cAccentLine = RGB(0, 60, 60): End Function

' ================================================================================
'  PUBLIC ENTRY POINTS
' ================================================================================

Public Sub InitDDTools()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    SetupDataSheets

    Dim ws As Worksheet
    If Not SheetExists(SH_UI) Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SH_UI
    End If

    SetupUISheet Worksheets(SH_UI)
    AddControlButtons Worksheets(SH_UI)

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "DD Tools sheet created." & vbLf & _
           "Place your three FieldCap CSV exports in the same folder as this workbook," & vbLf & _
           "then click REFRESH on the sheet.", _
           vbInformation, "DD Tools Ready"
End Sub

Public Sub RefreshData()
    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    If wbPath = "" Then
        MsgBox "Save the workbook to a folder first." & vbLf & _
               "The Refresh button scans the same folder as this file.", _
               vbInformation, "DD Tools"
        Exit Sub
    End If

    Dim fJob As String, fCrew As String, fBHA As String
    fJob = FindLatestCsvPathByToken(wbPath, "job-details")
    fBHA = FindLatestCsvPathByToken(wbPath, "bha-equipment")
    fCrew = FindLatestCsvPathByToken(wbPath, "crew")

    Dim missing As String
    If fJob = "" Then missing = missing & Chr(10) & "  *job-details*.csv"
    If fCrew = "" Then missing = missing & Chr(10) & "  *crew*.csv"
    If fBHA = "" Then missing = missing & Chr(10) & "  *bha-equipment*.csv"

    If missing <> "" Then
        MsgBox "Missing CSV export(s) in:" & Chr(10) & wbPath & Chr(10) & missing & Chr(10) & Chr(10) & _
               "Export them from the FieldCap Chrome plugin, then click Refresh again.", _
               vbExclamation, "DD Tools - Files Not Found"
        Exit Sub
    End If

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Found (newest matches):" & Chr(10) & _
                     "  " & Mid(fJob, InStrRev(fJob, Application.PathSeparator) + 1) & _
                     "  [" & Format(FileDateTime(fJob), "yyyy-mm-dd hh:nn:ss") & "]" & Chr(10) & _
                     "  " & Mid(fCrew, InStrRev(fCrew, Application.PathSeparator) + 1) & _
                     "  [" & Format(FileDateTime(fCrew), "yyyy-mm-dd hh:nn:ss") & "]" & Chr(10) & _
                     "  " & Mid(fBHA, InStrRev(fBHA, Application.PathSeparator) + 1) & _
                     "  [" & Format(FileDateTime(fBHA), "yyyy-mm-dd hh:nn:ss") & "]" & Chr(10) & Chr(10) & _
                     "Import and rebuild DD Tools?", _
                     vbQuestion + vbYesNo, "DD Tools Refresh")
    If confirm = vbNo Then Exit Sub

    Dim prevSheet As Worksheet
    Dim prevCellAddr As String
    Set prevSheet = ActiveSheet
    prevCellAddr = ActiveCell.Address(False, False)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "DD Tools: importing CSVs..."

    On Error GoTo ErrHandler

    SetupDataSheets
    ImportCSVToSheet fJob, Worksheets(SH_JOB)
    ImportCSVToSheet fCrew, Worksheets(SH_CREW)
    ImportCSVToSheet fBHA, Worksheets(SH_BHA)

    Application.StatusBar = "DD Tools: building dashboard..."
    BuildUI

    On Error Resume Next
    prevSheet.Activate
    prevSheet.Range(prevCellAddr).Select
    On Error GoTo 0

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "DD Tools"
End Sub

Public Sub RebuildDashboard()
    If Not SheetExists(SH_BHA) Then
        MsgBox "No data found. Click REFRESH first.", vbInformation, "DD Tools"
        Exit Sub
    End If

    Dim prevSheet As Worksheet
    Dim prevCellAddr As String
    Set prevSheet = ActiveSheet
    prevCellAddr = ActiveCell.Address(False, False)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "DD Tools: rebuilding..."

    On Error GoTo ErrHandler
    BuildUI

    On Error Resume Next
    prevSheet.Activate
    prevSheet.Range(prevCellAddr).Select
    On Error GoTo 0

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "DD Tools"
End Sub

Public Sub SelectBHA()
    Dim callerName As String
    On Error Resume Next
    callerName = Application.Caller
    On Error GoTo 0
    If InStr(callerName, "_") = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(callerName, "_")
    Dim bhaNum As Long
    bhaNum = CLng(parts(UBound(parts)))

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SH_UI)

    HighlightBHAButton ws, bhaNum
    DrawBHADetail ws, bhaNum
    DrawCumulativeTable ws
    DrawRawBHASummaryTable ws

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error in SelectBHA: " & Err.Description, vbCritical, "DD Tools"
End Sub

' ================================================================================
'  DATA SHEET SETUP & CSV IMPORT
' ================================================================================

Private Sub SetupDataSheets()
    Dim names(2) As String
    names(0) = SH_JOB
    names(1) = SH_CREW
    names(2) = SH_BHA

    Dim i As Integer
    For i = 0 To 2
        If Not SheetExists(names(i)) Then
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = names(i)
        End If
        Worksheets(names(i)).Visible = xlSheetVeryHidden
    Next i
End Sub

Private Sub ImportCSVToSheet(filePath As String, ws As Worksheet)
    ws.Cells.Clear

    Dim fileNum As Integer
    fileNum = FreeFile

    Open filePath For Input As #fileNum

    Dim rowIdx As Long
    rowIdx = 1

    Do While Not EOF(fileNum)
        Dim rawLine As String
        Line Input #fileNum, rawLine

        Dim fields() As String
        fields = ParseCSVLine(rawLine)

        Dim c As Long
        For c = 0 To UBound(fields)
            ws.Cells(rowIdx, c + 1).Value = SmartConvert(fields(c))
        Next c

        rowIdx = rowIdx + 1
    Loop

    Close #fileNum
End Sub

Private Function FindLatestCsvPathByToken(folderPath As String, token As String) As String
    FindLatestCsvPathByToken = ""
    Dim newestStamp As Date
    newestStamp = 0

    Dim fname As String
    fname = Dir(folderPath & Application.PathSeparator & "*.csv")
    Do While fname <> ""
        Dim lname As String
        lname = LCase(fname)
        If InStr(lname, LCase(token)) > 0 Then
            Dim full As String
            full = folderPath & Application.PathSeparator & fname
            On Error Resume Next
            Dim stamp As Date
            stamp = FileDateTime(full)
            If Err.Number = 0 Then
                If stamp >= newestStamp Then
                    newestStamp = stamp
                    FindLatestCsvPathByToken = full
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
        fname = Dir()
    Loop
End Function

Private Function ParseCSVLine(ByVal line As String) As String()
    Dim result(200) As String
    Dim idx As Long: idx = 0
    Dim pos As Long: pos = 1
    Dim inQ As Boolean: inQ = False
    Dim token As String: token = ""

    Do While pos <= Len(line)
        Dim ch As String
        ch = Mid(line, pos, 1)

        If ch = Chr(34) Then
            If inQ And Mid(line, pos + 1, 1) = Chr(34) Then
                token = token & Chr(34)
                pos = pos + 1
            Else
                inQ = Not inQ
            End If
        ElseIf ch = "," And Not inQ Then
            result(idx) = token
            idx = idx + 1
            token = ""
        Else
            token = token & ch
        End If

        pos = pos + 1
    Loop

    result(idx) = token

    Dim out() As String
    ReDim out(idx)
    Dim k As Long
    For k = 0 To idx
        out(k) = result(k)
    Next k
    ParseCSVLine = out
End Function

Private Function SmartConvert(s As String) As Variant
    s = Trim(s)
    If s = "" Then
        SmartConvert = ""
    ElseIf IsNumeric(s) Then
        SmartConvert = CDbl(s)
    Else
        SmartConvert = s
    End If
End Function

' ================================================================================
'  MAIN UI BUILD ORCHESTRATOR
' ================================================================================

Private Sub BuildUI()
    Dim uiStep As String
    On Error GoTo ErrHandler

    Dim ws As Worksheet

    uiStep = "get/create UI sheet"
    If Not SheetExists(SH_UI) Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SH_UI
    Else
        Set ws = Worksheets(SH_UI)
    End If

    uiStep = "clear sheet"
    ws.Cells.UnMerge
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    Dim btnDel As Button
    For Each btnDel In ws.Buttons
        btnDel.Delete
    Next btnDel

    uiStep = "SetupUISheet": SetupUISheet ws
    uiStep = "DrawToolbar": DrawToolbar ws

    uiStep = "GetBHAList"
    Dim bhaList As Variant
    bhaList = GetBHAList()
    Dim bub As Long: bub = SafeUBound(bhaList)

    uiStep = "DrawBHAButtons": DrawBHAButtons ws, bhaList
    uiStep = "AddControlButtons": AddControlButtons ws

    If bub >= 0 Then
        Dim defaultBHA As Long
        defaultBHA = CLng(bhaList(bub))
        uiStep = "HighlightBHAButton": HighlightBHAButton ws, defaultBHA
        uiStep = "DrawBHADetail": DrawBHADetail ws, defaultBHA
        uiStep = "DrawCumulativeTable": DrawCumulativeTable ws
        uiStep = "DrawRawBHASummaryTable": DrawRawBHASummaryTable ws
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & " in BuildUI [" & uiStep & "]:" & Chr(10) & Err.Description, _
           vbCritical, "DD Tools - BuildUI"
End Sub

' ================================================================================
'  SHEET SETUP & LAYOUT
' ================================================================================

Private Sub SetupUISheet(ws As Worksheet)
    With ws
        .Cells.Interior.Color = cWhite()
        .Cells.Font.Color = cBlack()
        .Cells.Font.Name = "Consolas"
        .Cells.Font.Size = 9
        .Tab.Color = cTeal()
    End With

    If ActiveSheet.Name = SH_UI Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If

    ' Column widths - tight, purposeful
    ws.Columns("A").ColumnWidth = 4     ' row # / index
    ws.Columns("B").ColumnWidth = 14    ' serial #
    ws.Columns("C").ColumnWidth = 12    ' item code
    ws.Columns("D").ColumnWidth = 30    ' description
    ws.Columns("E:W").ColumnWidth = 11  ' hours + KPI island + gap + RAW BHA SUMMARY

    ' Row heights
    ws.Rows(R_TOOLBAR).RowHeight = 26
    ws.Rows(R_JOB).RowHeight = 15
    ws.Rows(R_KPI).RowHeight = 14
    ws.Rows(R_DIV).RowHeight = 2
End Sub

' ================================================================================
'  TOOLBAR ROW (Title + Buttons consolidated in Row 1)
' ================================================================================

Private Sub DrawToolbar(ws As Worksheet)
    ' Title in A1:C1
    With ws.Range("A1:C1")
        .Merge
        .Value = "FIELDCAP DD TOOLS"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = cBlack()
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
    End With

    ' Job info in row 2
    Dim jobLabel As String
    jobLabel = GetJobHeaderLabel()
    With ws.Cells(R_JOB, 1)
        .Value = jobLabel
        .Font.Size = 8
        .Font.Color = cGrayDk()
        .Font.Bold = False
    End With

    ' Freeze line row: no colored banner (keep neutral)
    With ws.Range(ws.Cells(R_DIV, 1), ws.Cells(R_DIV, COL_DATA_LAST))
        .Interior.Color = cWhite()
        .Borders.LineStyle = xlNone
    End With
End Sub

' ================================================================================
'  FOUR-COLUMN SUMMARY ISLAND (J:M) — merged headers; KPI left, BHA detail right;
'  rows aligned horizontally (same row index for line 1, 2, …).
' ================================================================================

Private Sub IslandSummaryPair(ws As Worksheet, r As Long, labelCol As Long, _
                              label As String, val As String, rowBg As Long)
    With ws.Cells(r, labelCol)
        .Value = label
        .Interior.Color = rowBg
        .Font.Color = cGrayDk()
        .Font.Size = 8
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignRight
    End With
    With ws.Cells(r, labelCol + 1)
        .Value = val
        .Interior.Color = rowBg
        .Font.Color = cBlack()
        .Font.Size = 9
        .HorizontalAlignment = xlHAlignLeft
    End With
End Sub

Private Sub DrawSummaryIslandFourColumns(ws As Worksheet, bhaNum As Long, _
    totMetres As Double, totHrs As Double, totSlid As Double, totRot As Double, totCirc As Double, _
    slideFrac As Double, crewCount As Long, _
    section As String, status As String, motor As String, guid As String, _
    metres As Double, sldHrs As Double, rotHrs As Double, crcHrs As Double, blwHrs As Double, _
    actOn As String, cmpOn As String)

    ' Legacy horizontal KPI strip + old island area
    With ws.Range(ws.Cells(R_KPI, 1), ws.Cells(R_KPI, C_LAST))
        .UnMerge
        .ClearContents
        .ClearFormats
        .Interior.Color = cWhite()
        .Borders.LineStyle = xlNone
    End With
    ' Current island (J:M) + legacy position (Z:AC) so old layouts do not leave ghosts
    With ws.Range(ws.Cells(R_DETAIL, COL_SUMMARY_L), ws.Cells(22, COL_SUMMARY_END))
        .UnMerge
        .ClearContents
        .ClearFormats
        .Interior.Color = cWhite()
        .Borders.LineStyle = xlNone
    End With
    With ws.Range(ws.Cells(R_DETAIL, 26), ws.Cells(22, 29))
        .UnMerge
        .ClearContents
        .ClearFormats
        .Interior.Color = cWhite()
        .Borders.LineStyle = xlNone
    End With

    Dim hdr As Long
    hdr = R_DETAIL
    ws.Rows(hdr).RowHeight = 15

    With ws.Range(ws.Cells(hdr, COL_SUMMARY_L), ws.Cells(hdr, COL_SUMMARY_L + 1))
        .Merge
        .Value = "SELECTED BHA"
        .Interior.Color = cGrayBg()
        .Font.Bold = True
        .Font.Size = 8
        .Font.Color = cGrayDk()
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With ws.Range(ws.Cells(hdr, COL_SUMMARY_R), ws.Cells(hdr, COL_SUMMARY_END))
        .Merge
        .Value = "BHA " & CStr(bhaNum)
        .Interior.Color = cGrayBg()
        .Font.Bold = True
        .Font.Size = 8
        .Font.Color = cGrayDk()
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    With ws.Range(ws.Cells(hdr, COL_SUMMARY_L), ws.Cells(hdr, COL_SUMMARY_L + 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cGrayMed()
        .Weight = xlHairline
    End With
    With ws.Range(ws.Cells(hdr, COL_SUMMARY_R), ws.Cells(hdr, COL_SUMMARY_END)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cGrayMed()
        .Weight = xlHairline
    End With

    Dim i As Long
    Dim r As Long
    Dim maxRows As Long
    maxRows = 13

    Dim rowBg As Long

    For i = 1 To maxRows
        r = hdr + i
        ws.Rows(r).RowHeight = 15
        ' Match component table stripe: row 1 white, row 2 gray, ...
        If i Mod 2 = 1 Then rowBg = cWhite() Else rowBg = cGrayBg()

        Select Case i
            Case 1
                IslandSummaryPair ws, r, COL_SUMMARY_L, "METRES", Format(totMetres, "#,##0.0") & " m", rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Section", UCase(section), rowBg
            Case 2
                IslandSummaryPair ws, r, COL_SUMMARY_L, "TOTAL HRS", Format(totHrs, "0.00") & " h", rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Status", UCase(status), rowBg
            Case 3
                IslandSummaryPair ws, r, COL_SUMMARY_L, "SLIDE", Format(totSlid, "0.00") & " h", rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Motor", motor, rowBg
            Case 4
                IslandSummaryPair ws, r, COL_SUMMARY_L, "ROTATE", Format(totRot, "0.00") & " h", rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Guidance", guid, rowBg
            Case 5
                IslandSummaryPair ws, r, COL_SUMMARY_L, "CIRC", Format(totCirc, "0.00") & " h", rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Metres", Format(metres, "#,##0.00"), rowBg
            Case 6
                IslandSummaryPair ws, r, COL_SUMMARY_L, "SLD %", Format(slideFrac, "0.00"), rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Hrs Slid", Format(sldHrs, "0.00"), rowBg
            Case 7
                IslandSummaryPair ws, r, COL_SUMMARY_L, "CREW", CStr(crewCount), rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Hrs Rot", Format(rotHrs, "0.00"), rowBg
            Case 8
                ws.Cells(r, COL_SUMMARY_L).ClearContents
                ws.Cells(r, COL_SUMMARY_L + 1).ClearContents
                ws.Cells(r, COL_SUMMARY_L).Interior.Color = rowBg
                ws.Cells(r, COL_SUMMARY_L + 1).Interior.Color = rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Hrs Circ", Format(crcHrs, "0.00"), rowBg
            Case 9
                ws.Cells(r, COL_SUMMARY_L).ClearContents
                ws.Cells(r, COL_SUMMARY_L + 1).ClearContents
                ws.Cells(r, COL_SUMMARY_L).Interior.Color = rowBg
                ws.Cells(r, COL_SUMMARY_L + 1).Interior.Color = rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Total Hrs", Format(totHrs, "0.00"), rowBg
            Case 10
                ws.Cells(r, COL_SUMMARY_L).ClearContents
                ws.Cells(r, COL_SUMMARY_L + 1).ClearContents
                ws.Cells(r, COL_SUMMARY_L).Interior.Color = rowBg
                ws.Cells(r, COL_SUMMARY_L + 1).Interior.Color = rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Below Rot", Format(blwHrs, "0.00"), rowBg
            Case 11
                ws.Cells(r, COL_SUMMARY_L).ClearContents
                ws.Cells(r, COL_SUMMARY_L + 1).ClearContents
                ws.Cells(r, COL_SUMMARY_L).Interior.Color = rowBg
                ws.Cells(r, COL_SUMMARY_L + 1).Interior.Color = rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Activated", actOn, rowBg
            Case 12
                ws.Cells(r, COL_SUMMARY_L).ClearContents
                ws.Cells(r, COL_SUMMARY_L + 1).ClearContents
                ws.Cells(r, COL_SUMMARY_L).Interior.Color = rowBg
                ws.Cells(r, COL_SUMMARY_L + 1).Interior.Color = rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Completed", cmpOn, rowBg
            Case 13
                ws.Cells(r, COL_SUMMARY_L).ClearContents
                ws.Cells(r, COL_SUMMARY_L + 1).ClearContents
                ws.Cells(r, COL_SUMMARY_L).Interior.Color = rowBg
                ws.Cells(r, COL_SUMMARY_L + 1).Interior.Color = rowBg
                IslandSummaryPair ws, r, COL_SUMMARY_R, "Slide %", Format(slideFrac, "0.000"), rowBg
        End Select
    Next i

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(hdr, COL_SUMMARY_L), ws.Cells(hdr + maxRows, COL_SUMMARY_END))
    With rng
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = cGrayMed()
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = cGrayMed()
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = cGrayMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = cGrayMed()
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = cGrayMed()
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = cGrayMed()
    End With
End Sub

' ================================================================================
'  BHA SELECTOR BUTTONS (in toolbar row, starting col F)
' ================================================================================

Private Sub DrawBHAButtons(ws As Worksheet, bhaList As Variant)
    ' Delete old BHA buttons
    Dim btnDel As Button
    For Each btnDel In ws.Buttons
        If Left(btnDel.Name, 7) = "BHABtn_" Then btnDel.Delete
    Next btnDel

    If Not IsArray(bhaList) Then Exit Sub
    Dim ub As Long
    On Error Resume Next
    ub = UBound(bhaList)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    If ub < 0 Then Exit Sub

    ' BHA buttons start at column 6 (F) in toolbar row
    Dim startCol As Long: startCol = 6
    Dim i As Long
    For i = 0 To ub
        Dim bNum As Long
        bNum = CLng(bhaList(i))

        Dim targetCol As Long
        targetCol = startCol + i

        ' Ensure column exists with usable width
        If ws.Columns(targetCol).ColumnWidth < 8 Then
            ws.Columns(targetCol).ColumnWidth = 8
        End If

        Dim targetCell As Range
        Set targetCell = ws.Cells(R_TOOLBAR, targetCol)

        Dim btn As Button
        Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)
        btn.Name = "BHABtn_" & bNum
        btn.OnAction = "'" & ThisWorkbook.Name & "'!SelectBHA"
        btn.Placement = xlMoveAndSize
        StyleBtn btn, "BHA" & bNum, cTeal(), cAccentLine(), cBlack(), False
    Next i
End Sub

Private Sub HighlightBHAButton(ws As Worksheet, activeBHA As Long)
    Dim btn As Button
    For Each btn In ws.Buttons
        If Left(btn.Name, 7) = "BHABtn_" Then
            Dim n As Long
            n = CLng(Split(btn.Name, "_")(1))
            If n = activeBHA Then
                StyleBtn btn, "BHA" & CStr(n), cTealLt(), cAccentLine(), cBlack(), True
            Else
                StyleBtn btn, "BHA" & CStr(n), cTeal(), cAccentLine(), cBlack(), False
            End If
        End If
    Next btn

    On Error Resume Next
    ws.Names.Add "DD_ActiveBHA", "=" & Chr(34) & CStr(activeBHA) & Chr(34)
    On Error GoTo 0
End Sub

' ================================================================================
'  CONTROL BUTTONS (REFRESH in D1, REBUILD in E1)
' ================================================================================

Private Sub AddControlButtons(ws As Worksheet)
    Dim btnDel As Button
    For Each btnDel In ws.Buttons
        If btnDel.Name = "BtnImport" Or btnDel.Name = "BtnRebuild" Then btnDel.Delete
    Next btnDel

    ' REFRESH in D1
    Dim cellRef As Range
    Set cellRef = ws.Cells(R_TOOLBAR, 4)
    Dim bImport As Button
    Set bImport = ws.Buttons.Add(cellRef.Left, cellRef.Top, cellRef.Width, cellRef.Height)
    bImport.Name = "BtnImport"
    bImport.OnAction = "'" & ThisWorkbook.Name & "'!RefreshData"
    bImport.Placement = xlMoveAndSize
    StyleBtn bImport, "REFRESH", cTeal(), cAccentLine(), cBlack(), True

    ' REBUILD in E1
    Dim cellReb As Range
    Set cellReb = ws.Cells(R_TOOLBAR, 5)
    Dim bRebuild As Button
    Set bRebuild = ws.Buttons.Add(cellReb.Left, cellReb.Top, cellReb.Width, cellReb.Height)
    bRebuild.Name = "BtnRebuild"
    bRebuild.OnAction = "'" & ThisWorkbook.Name & "'!RebuildDashboard"
    bRebuild.Placement = xlMoveAndSize
    StyleBtn bRebuild, "REBUILD", cTeal(), cAccentLine(), cBlack(), True
End Sub

Private Sub StyleBtn(btn As Button, caption As String, fillColor As Long, lineColor As Long, textColor As Long, isBold As Boolean)
    On Error Resume Next
    btn.Caption = caption
    btn.Font.Name = "Consolas"
    btn.Font.Size = 9
    btn.Font.Bold = isBold
    btn.Font.Color = textColor
    btn.ShapeRange.Fill.ForeColor.RGB = fillColor
    btn.ShapeRange.Line.ForeColor.RGB = lineColor
    btn.ShapeRange.Line.Weight = 0.75
    On Error GoTo 0
End Sub

' ================================================================================
'  BHA DETAIL SECTION (Row 5+)
' ================================================================================

Private Sub DrawBHADetail(ws As Worksheet, bhaNum As Long)
    ' Clear detail area generously (include far-right island columns)
    Dim clearTo As Long: clearTo = R_DETAIL + 120
    With ws.Range(ws.Cells(R_DETAIL, 1), ws.Cells(clearTo, 30))
        .UnMerge
        .ClearContents
        .ClearFormats
        .Interior.Color = cWhite()
        .Font.Color = cBlack()
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Font.Bold = False
        .Borders.LineStyle = xlNone
    End With

    If Not SheetExists(SH_BHA) Then Exit Sub

    Dim bhaWs As Worksheet
    Set bhaWs = Worksheets(SH_BHA)

    Dim fr As Long: fr = GetFirstBHARow(bhaNum)
    If fr = 0 Then Exit Sub

    ' Column lookups
    Dim cSec As Long: cSec = FindCol(bhaWs, "Section")
    Dim cStat As Long: cStat = FindCol(bhaWs, "Status")
    Dim cMot As Long: cMot = FindCol(bhaWs, "Motor")
    Dim cGui As Long: cGui = FindCol(bhaWs, "Guidance")
    Dim cMtr As Long: cMtr = FindCol(bhaWs, "Metres Drilled")
    Dim cTot As Long: cTot = FindCol(bhaWs, "BHA Total Hrs")
    Dim cSld As Long: cSld = FindCol(bhaWs, "BHA Hrs Slid")
    Dim cRot As Long: cRot = FindCol(bhaWs, "BHA Hrs Rot")
    Dim cCrc As Long: cCrc = FindCol(bhaWs, "BHA Hrs Circ")
    Dim cBlw As Long: cBlw = FindCol(bhaWs, "BHA Below Rot")
    Dim cAct As Long: cAct = FindCol(bhaWs, "Activated On")
    Dim cCmp As Long: cCmp = FindCol(bhaWs, "Completed On")
    Dim cBNum As Long: cBNum = FindCol(bhaWs, "BHA #")
    Dim cSer As Long: cSer = FindCol(bhaWs, "Serial #")
    Dim cCod As Long: cCod = FindCol(bhaWs, "Item Code")
    Dim cDes As Long: cDes = FindCol(bhaWs, "Description")
    Dim cMSld As Long: cMSld = FindCol(bhaWs, "BHA Mtrs Slid")
    Dim cMRot As Long: cMRot = FindCol(bhaWs, "BHA Mtrs Rot")

    ' BHA-level values
    Dim section As String: section = SafeStr(bhaWs.Cells(fr, cSec))
    Dim status As String: status = SafeStr(bhaWs.Cells(fr, cStat))
    Dim motor As String: motor = SafeStr(bhaWs.Cells(fr, cMot))
    Dim guid As String: guid = SafeStr(bhaWs.Cells(fr, cGui))
    Dim metres As Double: metres = GetBHAMetricValue(bhaWs, cBNum, cMtr, bhaNum)
    Dim totHrs As Double: totHrs = GetNum(bhaWs.Cells(fr, cTot))
    Dim sldHrs As Double: sldHrs = GetNum(bhaWs.Cells(fr, cSld))
    Dim rotHrs As Double: rotHrs = GetNum(bhaWs.Cells(fr, cRot))
    Dim crcHrs As Double: crcHrs = GetNum(bhaWs.Cells(fr, cCrc))
    Dim blwHrs As Double: blwHrs = GetNum(bhaWs.Cells(fr, cBlw))
    Dim actOn As String: actOn = SafeStr(bhaWs.Cells(fr, cAct))
    Dim cmpOn As String: cmpOn = SafeStr(bhaWs.Cells(fr, cCmp))

    Dim crewCount As Long
    crewCount = 0
    If SheetExists(SH_CREW) Then
        On Error Resume Next
        crewCount = Application.WorksheetFunction.CountA(Worksheets(SH_CREW).Columns(1)) - 1
        If crewCount < 0 Then crewCount = 0
        On Error GoTo 0
    End If

    Dim slideFrac As Double
    Dim slideM As Double
    Dim rotM As Double
    slideFrac = 0
    slideM = 0
    rotM = 0
    If cMSld > 0 Then slideM = GetNum(bhaWs.Cells(fr, cMSld))
    If cMRot > 0 Then rotM = GetNum(bhaWs.Cells(fr, cMRot))
    ' Match FieldCap BHA slide share: metres slid / (slid + rotate metres). Fallback: slid hrs / total hrs.
    If slideM + rotM > 0 Then
        slideFrac = slideM / (slideM + rotM)
    ElseIf totHrs > 0 Then
        slideFrac = sldHrs / totHrs
    End If

    DrawSummaryIslandFourColumns ws, bhaNum, metres, totHrs, sldHrs, rotHrs, crcHrs, slideFrac, crewCount, _
        section, status, motor, guid, metres, sldHrs, rotHrs, crcHrs, blwHrs, actOn, cmpOn

    Dim r As Long: r = R_DETAIL

    ' ========================================================================
    ' COMPONENT TABLE (columns A–H, same count as cumulative table)
    ' #, SERIAL #, ITEM CODE, DESCRIPTION, HRS SLD, HRS ROT, TOTAL HRS, BHAs
    ' ========================================================================
    ' Column headers
    ws.Rows(r).RowHeight = 15
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST))
        .Interior.Color = cGrayBg()
        .Font.Color = cGrayDk()
        .Font.Bold = True
        .Font.Size = 8
    End With
    ws.Cells(r, 1).Value = "#"
    ws.Cells(r, 1).HorizontalAlignment = xlHAlignCenter
    ws.Cells(r, 2).Value = "SERIAL #"
    ws.Cells(r, 3).Value = "ITEM CODE"
    ws.Cells(r, 4).Value = "DESCRIPTION"
    ws.Cells(r, 5).Value = "HRS SLD"
    ws.Cells(r, 5).HorizontalAlignment = xlHAlignRight
    ws.Cells(r, 6).Value = "HRS ROT"
    ws.Cells(r, 6).HorizontalAlignment = xlHAlignRight
    ws.Cells(r, 7).Value = "TOTAL HRS"
    ws.Cells(r, 7).HorizontalAlignment = xlHAlignRight
    ws.Cells(r, COL_DATA_LAST).Value = "BHAs"
    ws.Cells(r, COL_DATA_LAST).HorizontalAlignment = xlHAlignRight
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cGrayMed()
        .Weight = xlHairline
    End With
    r = r + 1

    ' -- Component data rows --
    Dim lastRow As Long
    lastRow = bhaWs.Cells(bhaWs.Rows.Count, cBNum).End(xlUp).Row

    Dim compIdx As Long: compIdx = 1
    Dim dr As Long
    For dr = 2 To lastRow
        Dim rawBNum As Variant
        rawBNum = bhaWs.Cells(dr, cBNum).Value
        If IsNumeric(rawBNum) Then
            If CLng(rawBNum) = bhaNum Then
                ' Skip blank rows (no serial AND no item code AND no description)
                Dim sSer As String: sSer = SafeStr(bhaWs.Cells(dr, cSer))
                Dim sCod As String: sCod = SafeStr(bhaWs.Cells(dr, cCod))
                Dim sDes As String: sDes = SafeStr(bhaWs.Cells(dr, cDes))
                If sSer = "" And sCod = "" And sDes = "" Then GoTo SkipComp

                ws.Rows(r).RowHeight = 15

                Dim bg As Long
                If compIdx Mod 2 = 0 Then bg = cGrayBg() Else bg = cWhite()

                ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Interior.Color = bg

                PutCell ws, r, 1, CStr(compIdx), bg, cGrayDk(), False, xlHAlignCenter
                PutCell ws, r, 2, sSer, bg, cBlack(), False, xlHAlignLeft
                PutCell ws, r, 3, sCod, bg, cGrayDk(), False, xlHAlignLeft
                PutCell ws, r, 4, sDes, bg, cBlack(), False, xlHAlignLeft

                ' Hours columns align with the island metrics
                With ws.Cells(r, 5)
                    .Value = Format(sldHrs, "0.00")
                    .Interior.Color = bg
                    .Font.Color = cBlack()
                    .Font.Size = 9
                    .HorizontalAlignment = xlHAlignRight
                End With
                With ws.Cells(r, 6)
                    .Value = Format(rotHrs, "0.00")
                    .Interior.Color = bg
                    .Font.Color = cBlack()
                    .Font.Size = 9
                    .HorizontalAlignment = xlHAlignRight
                End With
                With ws.Cells(r, 7)
                    .Value = Format(totHrs, "0.00")
                    .Interior.Color = bg
                    .Font.Color = cBlack()
                    .Font.Bold = True
                    .Font.Size = 9
                    .HorizontalAlignment = xlHAlignRight
                End With
                With ws.Cells(r, COL_DATA_LAST)
                    .Value = bhaNum
                    .Interior.Color = bg
                    .Font.Color = cGrayDk()
                    .Font.Size = 9
                    .HorizontalAlignment = xlHAlignRight
                End With

                compIdx = compIdx + 1
                r = r + 1
SkipComp:
            End If
        End If
    Next dr

    ' Footer border
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = cGrayMed()
        .Weight = xlHairline
    End With
    r = r + 1

    ' Store cumulative table anchor
    On Error Resume Next
    ws.Names.Add "DD_CumulRow", "='" & SH_UI & "'!$A$" & r
    On Error GoTo 0
End Sub

Private Sub MetricRow(ws As Worksheet, r As Long, c As Long, label As String, val As String)
    ws.Cells(r, c).Value = label
    ws.Cells(r, c).Font.Color = cGrayDk()
    ws.Cells(r, c).Font.Size = 8
    ws.Cells(r, c).Font.Bold = True
    ws.Cells(r, c).HorizontalAlignment = xlHAlignRight
    ws.Cells(r, c + 1).Value = val
    ws.Cells(r, c + 1).Font.Color = cBlack()
    ws.Cells(r, c + 1).Font.Size = 9
    ws.Cells(r, c + 1).HorizontalAlignment = xlHAlignLeft
End Sub

' ================================================================================
'  CUMULATIVE COMPONENT HOURS TABLE
' ================================================================================

Private Sub DrawCumulativeTable(ws As Worksheet)
    Dim cStart As Long: cStart = 0
    On Error Resume Next
    cStart = ws.Range("DD_CumulRow").Row
    On Error GoTo 0
    If cStart = 0 Then cStart = 60

    ' Clear downward
    With ws.Range(ws.Cells(cStart, 1), ws.Cells(cStart + 200, C_LAST + 2))
        .UnMerge
        .ClearContents
        .ClearFormats
        .Interior.Color = cWhite()
        .Font.Color = cBlack()
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Borders.LineStyle = xlNone
    End With

    If Not SheetExists(SH_BHA) Then Exit Sub

    Dim bhaWs As Worksheet
    Set bhaWs = Worksheets(SH_BHA)

    Dim cBNum As Long: cBNum = FindCol(bhaWs, "BHA #")
    Dim cSer As Long: cSer = FindCol(bhaWs, "Serial #")
    Dim cCod As Long: cCod = FindCol(bhaWs, "Item Code")
    Dim cDes As Long: cDes = FindCol(bhaWs, "Description")
    Dim cTot As Long: cTot = FindCol(bhaWs, "BHA Total Hrs")
    Dim cSld As Long: cSld = FindCol(bhaWs, "BHA Hrs Slid")
    Dim cRot As Long: cRot = FindCol(bhaWs, "BHA Hrs Rot")
    Dim cMtr As Long: cMtr = FindCol(bhaWs, "Metres Drilled")

    Dim lastRow As Long
    lastRow = bhaWs.Cells(bhaWs.Rows.Count, cBNum).End(xlUp).Row

    ' Rollup arrays (serial-first fatigue model)
    Dim rSerial() As String
    Dim rDesc() As String
    Dim rCode() As String
    Dim rTot() As Double
    Dim rSld() As Double
    Dim rRot() As Double
    Dim rMtr() As Double
    Dim rCnt() As Long
    Dim rSz As Long: rSz = 0

    ReDim rSerial(0): ReDim rDesc(0): ReDim rCode(0)
    ReDim rTot(0): ReDim rSld(0)
    ReDim rRot(0): ReDim rMtr(0): ReDim rCnt(0)

    If cSer = 0 Then Exit Sub

    Dim seenSerialBha As New Collection
    Dim dr As Long
    For dr = 2 To lastRow
        If Not IsNumeric(bhaWs.Cells(dr, cBNum).Value) Then GoTo SkipRow
        Dim bNum As Long: bNum = CLng(bhaWs.Cells(dr, cBNum).Value)

        Dim srl As String
        srl = Trim(SafeStr(bhaWs.Cells(dr, cSer)))
        If srl = "" Then GoTo SkipRow

        Dim sbKey As String
        sbKey = srl & "|" & CStr(bNum)
        On Error Resume Next
        seenSerialBha.Add 1, sbKey
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo SkipRow
        End If
        On Error GoTo 0

        Dim idx As Long: idx = -1
        Dim k As Long
        For k = 0 To rSz - 1
            If rSerial(k) = srl Then
                idx = k
                Exit For
            End If
        Next k

        Dim codPart As String
        codPart = ""
        If cCod > 0 Then codPart = SafeStr(bhaWs.Cells(dr, cCod))

        If idx = -1 Then
            idx = rSz
            rSz = rSz + 1
            ReDim Preserve rSerial(rSz - 1): ReDim Preserve rDesc(rSz - 1): ReDim Preserve rCode(rSz - 1)
            ReDim Preserve rTot(rSz - 1): ReDim Preserve rSld(rSz - 1)
            ReDim Preserve rRot(rSz - 1): ReDim Preserve rMtr(rSz - 1)
            ReDim Preserve rCnt(rSz - 1)
            rSerial(idx) = srl
            rDesc(idx) = SafeStr(bhaWs.Cells(dr, cDes))
            rCode(idx) = codPart
        Else
            ' Pick up item code from any row for this serial if still blank
            If rCode(idx) = "" And codPart <> "" Then rCode(idx) = codPart
        End If

        rCnt(idx) = rCnt(idx) + 1

        Dim bfr As Long: bfr = GetFirstBHARow(bNum)
        If bfr > 0 Then
            rTot(idx) = rTot(idx) + GetNum(bhaWs.Cells(bfr, cTot))
            rSld(idx) = rSld(idx) + GetNum(bhaWs.Cells(bfr, cSld))
            rRot(idx) = rRot(idx) + GetNum(bhaWs.Cells(bfr, cRot))
            rMtr(idx) = rMtr(idx) + GetBHAMetricValue(bhaWs, cBNum, cMtr, bNum)
        End If
SkipRow:
    Next dr

    If rSz = 0 Then Exit Sub

    ' Sort descending by total hours
    Dim swapped As Boolean
    Do
        swapped = False
        For k = 0 To rSz - 2
            If rTot(k) < rTot(k + 1) Then
                SwapS rSerial, k, k + 1: SwapS rDesc, k, k + 1: SwapS rCode, k, k + 1
                SwapD rTot, k, k + 1: SwapD rSld, k, k + 1
                SwapD rRot, k, k + 1: SwapD rMtr, k, k + 1
                SwapL rCnt, k, k + 1
                swapped = True
            End If
        Next k
    Loop While swapped

    ' Render cumulative section
    Dim r As Long: r = cStart

    ' Section header (same width as data columns)
    ws.Rows(r).RowHeight = 18
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST))
        .Interior.Color = cGrayBg()
        .Font.Color = cBlack()
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Cells(r, 1).Value = "CUMULATIVE SERIAL HOURS (All BHAs)"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cTeal()
        .Weight = xlThin
    End With
    r = r + 1

    ' Column headers — match top table: #, SERIAL, ITEM CODE, DESCRIPTION, HRS SLD, HRS ROT, TOTAL HRS, BHAs
    ws.Rows(r).RowHeight = 15
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST))
        .Interior.Color = cGrayBg()
        .Font.Color = cGrayDk()
        .Font.Bold = True
        .Font.Size = 8
    End With
    ws.Cells(r, 1).Value = "#"
    ws.Cells(r, 1).HorizontalAlignment = xlHAlignCenter
    ws.Cells(r, 2).Value = "SERIAL #"
    ws.Cells(r, 3).Value = "ITEM CODE"
    ws.Cells(r, 4).Value = "DESCRIPTION"
    ws.Cells(r, 5).Value = "HRS SLD"
    ws.Cells(r, 5).HorizontalAlignment = xlHAlignRight
    ws.Cells(r, 6).Value = "HRS ROT"
    ws.Cells(r, 6).HorizontalAlignment = xlHAlignRight
    ws.Cells(r, 7).Value = "TOTAL HRS"
    ws.Cells(r, 7).HorizontalAlignment = xlHAlignRight
    ws.Cells(r, COL_DATA_LAST).Value = "BHAs"
    ws.Cells(r, COL_DATA_LAST).HorizontalAlignment = xlHAlignRight
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cGrayMed()
        .Weight = xlHairline
    End With
    r = r + 1

    ' Data rows
    Dim ri As Long
    For ri = 0 To rSz - 1
        ws.Rows(r).RowHeight = 15

        Dim bg As Long
        If (ri + 1) Mod 2 = 0 Then bg = cGrayBg() Else bg = cWhite()
        ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Interior.Color = bg

        PutCell ws, r, 1, CStr(ri + 1), bg, cGrayDk(), False, xlHAlignCenter
        PutCell ws, r, 2, rSerial(ri), bg, cBlack(), False, xlHAlignLeft
        PutCell ws, r, 3, rCode(ri), bg, cGrayDk(), False, xlHAlignLeft
        PutCell ws, r, 4, rDesc(ri), bg, cBlack(), False, xlHAlignLeft

        PutNum ws, r, 5, rSld(ri), bg
        PutNum ws, r, 6, rRot(ri), bg

        ' Total hours - bold if > 0, red if > 300
        With ws.Cells(r, 7)
            .Value = Format(rTot(ri), "0.00")
            .Interior.Color = bg
            .HorizontalAlignment = xlHAlignRight
            .Font.Size = 9
            If rTot(ri) > 300 Then
                .Font.Color = cRed()
                .Font.Bold = True
            ElseIf rTot(ri) > 0 Then
                .Font.Color = cBlack()
                .Font.Bold = True
            Else
                .Font.Color = cGrayDk()
                .Font.Bold = False
            End If
        End With

        ' BHAs count (last column)
        With ws.Cells(r, COL_DATA_LAST)
            .Value = rCnt(ri)
            .Interior.Color = bg
            .Font.Color = cGrayDk()
            .HorizontalAlignment = xlHAlignRight
        End With

        r = r + 1
    Next ri

    ' Bottom border
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = cGrayMed()
        .Weight = xlThin
    End With
    r = r + 1

    ' ========================================================================
    ' OVER-LIMIT COMPONENTS (>300h) - fatigue warning
    ' ========================================================================
    Dim oSerial() As String, oDesc() As String, oCode() As String
    Dim oTot() As Double, oBhaCnt() As Long
    Dim oSz As Long: oSz = 0
    ReDim oSerial(0): ReDim oDesc(0): ReDim oCode(0)
    ReDim oTot(0): ReDim oBhaCnt(0)

    For k = 0 To rSz - 1
        If rTot(k) > 300# Then
            oSz = oSz + 1
            ReDim Preserve oSerial(oSz - 1)
            ReDim Preserve oDesc(oSz - 1)
            ReDim Preserve oCode(oSz - 1)
            ReDim Preserve oTot(oSz - 1)
            ReDim Preserve oBhaCnt(oSz - 1)
            oSerial(oSz - 1) = rSerial(k)
            oDesc(oSz - 1) = rDesc(k)
            oCode(oSz - 1) = rCode(k)
            oTot(oSz - 1) = rTot(k)
            oBhaCnt(oSz - 1) = rCnt(k)
        End If
    Next k

    ' Sort over-limit descending
    If oSz > 1 Then
        Do
            swapped = False
            For k = 0 To oSz - 2
                If oTot(k) < oTot(k + 1) Then
                    SwapS oSerial, k, k + 1
                    SwapS oDesc, k, k + 1
                    SwapS oCode, k, k + 1
                    SwapD oTot, k, k + 1
                    SwapL oBhaCnt, k, k + 1
                    swapped = True
                End If
            Next k
        Loop While swapped
    End If

    ' Section header
    ws.Rows(r).RowHeight = 18
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST))
        .Interior.Color = cWhite()
        .Font.Color = cRed()
        .Font.Bold = True
        .Font.Size = 10
    End With
    ws.Cells(r, 1).Value = "OVER-LIMIT (>300h Serial Hrs)"
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cRed()
        .Weight = xlThin
    End With
    r = r + 1

    If oSz = 0 Then
        ws.Cells(r, 1).Value = "No components over 300h."
        ws.Cells(r, 1).Font.Color = cGrayDk()
        r = r + 1
    Else
        ' Column headers — aligned with tables above
        ws.Rows(r).RowHeight = 14
        ws.Cells(r, 1).Value = "#"
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 8
        ws.Cells(r, 1).Font.Color = cGrayDk()
        ws.Cells(r, 1).HorizontalAlignment = xlHAlignCenter
        ws.Cells(r, 2).Value = "SERIAL #"
        ws.Cells(r, 2).Font.Bold = True
        ws.Cells(r, 2).Font.Size = 8
        ws.Cells(r, 2).Font.Color = cGrayDk()
        ws.Cells(r, 3).Value = "ITEM CODE"
        ws.Cells(r, 3).Font.Bold = True
        ws.Cells(r, 3).Font.Size = 8
        ws.Cells(r, 3).Font.Color = cGrayDk()
        ws.Cells(r, 4).Value = "DESCRIPTION"
        ws.Cells(r, 4).Font.Bold = True
        ws.Cells(r, 4).Font.Size = 8
        ws.Cells(r, 4).Font.Color = cGrayDk()
        ws.Cells(r, 7).Value = "SERIAL HRS"
        ws.Cells(r, 7).Font.Bold = True
        ws.Cells(r, 7).Font.Size = 8
        ws.Cells(r, 7).Font.Color = cGrayDk()
        ws.Cells(r, 7).HorizontalAlignment = xlHAlignRight
        ws.Cells(r, COL_DATA_LAST).Value = "BHAs"
        ws.Cells(r, COL_DATA_LAST).Font.Bold = True
        ws.Cells(r, COL_DATA_LAST).Font.Size = 8
        ws.Cells(r, COL_DATA_LAST).Font.Color = cGrayDk()
        ws.Cells(r, COL_DATA_LAST).HorizontalAlignment = xlHAlignRight
        r = r + 1

        ' Data rows
        For k = 0 To oSz - 1
            ws.Rows(r).RowHeight = 15
            PutCell ws, r, 1, CStr(k + 1), cWhite(), cGrayDk(), False, xlHAlignCenter
            PutCell ws, r, 2, oSerial(k), cWhite(), cRed(), True, xlHAlignLeft
            PutCell ws, r, 3, oCode(k), cWhite(), cGrayDk(), False, xlHAlignLeft
            PutCell ws, r, 4, oDesc(k), cWhite(), cBlack(), False, xlHAlignLeft
            With ws.Cells(r, 7)
                .Value = Format(oTot(k), "0.00")
                .HorizontalAlignment = xlHAlignRight
                .Font.Color = cRed()
                .Font.Bold = True
            End With
            With ws.Cells(r, COL_DATA_LAST)
                .Value = oBhaCnt(k)
                .HorizontalAlignment = xlHAlignRight
                .Font.Color = cGrayDk()
            End With
            r = r + 1
        Next k
    End If

    ' Final border
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, COL_DATA_LAST)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = cGrayMed()
        .Weight = xlHairline
    End With
End Sub

' ================================================================================
'  RAW BHA SUMMARY (columns O:W — right of KPI J:M, gap col N; header row = R_DETAIL)
' ================================================================================

Private Sub DrawRawBHASummaryTable(ws As Worksheet)
    If Not SheetExists(SH_BHA) Then Exit Sub

    Dim bhaWs As Worksheet
    Set bhaWs = Worksheets(SH_BHA)

    Dim cBNum As Long: cBNum = FindCol(bhaWs, "BHA #")
    Dim cMtr As Long: cMtr = FindCol(bhaWs, "Metres Drilled")
    Dim cMSld As Long: cMSld = FindCol(bhaWs, "BHA Mtrs Slid")
    Dim cMRot As Long: cMRot = FindCol(bhaWs, "BHA Mtrs Rot")
    Dim cSld As Long: cSld = FindCol(bhaWs, "BHA Hrs Slid")
    Dim cRot As Long: cRot = FindCol(bhaWs, "BHA Hrs Rot")
    Dim cCrc As Long: cCrc = FindCol(bhaWs, "BHA Hrs Circ")
    Dim cTot As Long: cTot = FindCol(bhaWs, "BHA Total Hrs")
    Dim cBlw As Long: cBlw = FindCol(bhaWs, "BHA Below Rot")
    If cBNum = 0 Then Exit Sub

    Dim startCol As Long: startCol = COL_RAW_SUMMARY_START
    Dim endCol As Long: endCol = startCol + 8
    Dim hdrRow As Long: hdrRow = R_DETAIL
    Dim titleRow As Long: titleRow = R_KPI

    ' Clear title band, raw block, and legacy N-started tables
    With ws.Range(ws.Cells(titleRow, 14), ws.Cells(hdrRow + 45, endCol))
        On Error Resume Next
        .UnMerge
        On Error GoTo 0
        .ClearContents
        .ClearFormats
        .Interior.Color = cWhite()
        .Font.Color = cBlack()
        .Font.Name = "Consolas"
        .Font.Size = 9
        .Borders.LineStyle = xlNone
    End With

    ' Section title (row 3 — same row band cleared only through col N elsewhere)
    With ws.Range(ws.Cells(titleRow, startCol), ws.Cells(titleRow, endCol))
        .Merge
        .Value = "RAW BHA SUMMARY"
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = cGrayDk()
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
        .Interior.Color = cWhite()
    End With

    ' Header row — match component / KPI header on row 5
    Dim hdr() As String
    hdr = Split("BHA,Meters,Mtrs Sld,Mtrs Rot,Hrs Sld,Hrs Rot,Hrs Crc,Total Hrs,Below Rot", ",")
    Dim i As Long
    For i = 0 To UBound(hdr)
        ws.Cells(hdrRow, startCol + i).Value = hdr(i)
        ws.Cells(hdrRow, startCol + i).HorizontalAlignment = xlHAlignCenter
    Next i
    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(hdrRow, endCol))
        .Interior.Color = cGrayBg()
        .Font.Color = cGrayDk()
        .Font.Bold = True
        .Font.Size = 8
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = cGrayMed()
            .Weight = xlHairline
        End With
    End With
    ws.Rows(hdrRow).RowHeight = 15

    ' Data rows — same stripe pattern as component table (row 1 white, 2 gray, …)
    Dim bhaList As Variant
    bhaList = GetBHAList()
    Dim ub As Long: ub = SafeUBound(bhaList)
    Dim r As Long: r = hdrRow + 1
    Dim j As Long
    Dim stripe As Long
    stripe = 0
    For i = 0 To ub
        Dim fr As Long
        fr = GetFirstBHARow(CLng(bhaList(i)))
        If fr > 0 Then
            stripe = stripe + 1
            Dim rawBg As Long
            If stripe Mod 2 = 1 Then rawBg = cWhite() Else rawBg = cGrayBg()

            ws.Rows(r).Hidden = False
            If ws.Rows(r).RowHeight < 14 Then ws.Rows(r).RowHeight = 15

            ws.Cells(r, startCol + 0).Value = CLng(bhaList(i))
            Dim totalM As Double: totalM = GetBHAMetricValue(bhaWs, cBNum, cMtr, CLng(bhaList(i)))
            Dim slideM As Double: slideM = 0
            Dim rotM As Double: rotM = 0
            If cMSld > 0 Then slideM = GetNum(bhaWs.Cells(fr, cMSld))
            If cMRot > 0 Then rotM = GetNum(bhaWs.Cells(fr, cMRot))
            If rotM = 0 And totalM > 0 And slideM > 0 And totalM >= slideM Then rotM = totalM - slideM
            If slideM = 0 And totalM > 0 And rotM > 0 And totalM >= rotM Then slideM = totalM - rotM

            ws.Cells(r, startCol + 1).Value = totalM
            ws.Cells(r, startCol + 2).Value = slideM
            ws.Cells(r, startCol + 3).Value = rotM
            ws.Cells(r, startCol + 4).Value = GetNum(bhaWs.Cells(fr, cSld))
            ws.Cells(r, startCol + 5).Value = GetNum(bhaWs.Cells(fr, cRot))
            ws.Cells(r, startCol + 6).Value = GetNum(bhaWs.Cells(fr, cCrc))
            ws.Cells(r, startCol + 7).Value = GetNum(bhaWs.Cells(fr, cTot))
            ws.Cells(r, startCol + 8).Value = GetNum(bhaWs.Cells(fr, cBlw))

            With ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
                .Interior.Color = rawBg
            End With
            ws.Cells(r, startCol + 0).Font.Color = cGrayDk()
            For j = 1 To 8
                ws.Cells(r, startCol + j).HorizontalAlignment = xlHAlignRight
                ws.Cells(r, startCol + j).Font.Color = cBlack()
                ws.Cells(r, startCol + j).Font.Size = 9
            Next j
            ws.Cells(r, startCol + 0).HorizontalAlignment = xlHAlignCenter
            ws.Cells(r, startCol + 0).Font.Size = 9

            r = r + 1
        End If
    Next i

    Dim tblEnd As Long
    If r > hdrRow + 1 Then
        tblEnd = r - 1
    Else
        tblEnd = hdrRow
    End If

    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(tblEnd, endCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = cGrayMed()
        .Borders.Weight = xlHairline
    End With

    On Error Resume Next
    ws.Names.Add "DD_RawBHASummary", "='" & SH_UI & "'!" & _
        ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(tblEnd, endCol)).Address
    On Error GoTo 0
End Sub

' ================================================================================
'  DATA HELPERS
' ================================================================================

Private Function GetBHAList() As Variant
    GetBHAList = Array()

    If Not SheetExists(SH_BHA) Then Exit Function

    Dim ws As Worksheet
    Set ws = Worksheets(SH_BHA)

    Dim col As Long: col = FindCol(ws, "BHA #")
    If col = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Dim seen As New Collection
    Dim r As Long
    For r = 2 To lastRow
        Dim v As Variant: v = ws.Cells(r, col).Value
        If IsNumeric(v) And v <> "" Then
            On Error Resume Next
            seen.Add CLng(v), CStr(CLng(v))
            On Error GoTo 0
        End If
    Next r

    If seen.Count = 0 Then Exit Function

    Dim result() As Variant
    ReDim result(seen.Count - 1)

    Dim i As Long
    For i = 1 To seen.Count
        result(i - 1) = seen(i)
    Next i

    ' Insertion sort ascending
    Dim j As Long, tmp As Variant
    For i = 1 To UBound(result)
        tmp = result(i)
        j = i - 1
        Do While j >= 0
            If result(j) <= tmp Then Exit Do
            result(j + 1) = result(j)
            j = j - 1
        Loop
        result(j + 1) = tmp
    Next i

    GetBHAList = result
End Function

Private Function GetFirstBHARow(bhaNum As Long) As Long
    GetFirstBHARow = 0
    If Not SheetExists(SH_BHA) Then Exit Function

    Dim ws As Worksheet: Set ws = Worksheets(SH_BHA)
    Dim col As Long: col = FindCol(ws, "BHA #")
    If col = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, col).Value) Then
            If CLng(ws.Cells(r, col).Value) = bhaNum Then
                GetFirstBHARow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function GetBHAMetricValue(ws As Worksheet, bhaCol As Long, metricCol As Long, bhaNum As Long) As Double
    GetBHAMetricValue = 0
    If bhaCol = 0 Or metricCol = 0 Then Exit Function

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, bhaCol).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, bhaCol).Value) Then
            If CLng(ws.Cells(r, bhaCol).Value) = bhaNum Then
                Dim v As Double
                v = GetNum(ws.Cells(r, metricCol))
                If v <> 0 Then
                    GetBHAMetricValue = v
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function GetJobHeaderLabel() As String
    GetJobHeaderLabel = "No job data imported"
    If Not SheetExists(SH_JOB) Then Exit Function

    Dim ws As Worksheet: Set ws = Worksheets(SH_JOB)

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim label As String: label = ""
    Dim c As Long
    For c = 1 To lastCol
        Dim hd As String: hd = Trim(LCase(SafeStr(ws.Cells(1, c))))
        Dim val As String: val = SafeStr(ws.Cells(2, c))
        If val = "" Then GoTo NextCol

        Select Case True
            Case InStr(hd, "job") > 0 And InStr(hd, "id") > 0
                label = label & "Job " & val & "  |  "
            Case InStr(hd, "client") > 0 And InStr(hd, "job") > 0
                label = label & val & "  |  "
            Case InStr(hd, "well") > 0
                label = label & val & "  |  "
            Case InStr(hd, "status") > 0
                label = label & UCase(val)
        End Select
NextCol:
    Next c

    If Trim(label) <> "" Then GetJobHeaderLabel = Trim(label)
End Function

' ================================================================================
'  UTILITY FUNCTIONS
' ================================================================================

Private Function SafeUBound(arr As Variant) As Long
    SafeUBound = -1
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    SafeUBound = UBound(arr)
    On Error GoTo 0
End Function

Private Function SheetExists(name As String) As Boolean
    On Error Resume Next
    SheetExists = Not (ThisWorkbook.Sheets(name) Is Nothing)
    On Error GoTo 0
End Function

Private Function FindCol(ws As Worksheet, header As String) As Long
    FindCol = 0
    Dim last As Long
    last = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To last
        If Trim(SafeStr(ws.Cells(1, c))) = header Then
            FindCol = c
            Exit Function
        End If
    Next c
End Function

Private Function SafeStr(cell As Range) As String
    On Error Resume Next
    SafeStr = Trim(CStr(cell.Value))
    On Error GoTo 0
End Function

Private Function GetNum(cell As Range) As Double
    GetNum = 0
    On Error Resume Next
    If IsNumeric(cell.Value) Then GetNum = CDbl(cell.Value)
    On Error GoTo 0
End Function

Private Sub PutCell(ws As Worksheet, r As Long, c As Long, val As String, _
                    bg As Long, fg As Long, bold As Boolean, align As Long)
    With ws.Cells(r, c)
        .NumberFormat = "@"
        .Value = val
        .Interior.Color = bg
        .Font.Color = fg
        .Font.Bold = bold
        .Font.Size = 9
        .Font.Name = "Consolas"
        .HorizontalAlignment = align
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub PutNum(ws As Worksheet, r As Long, c As Long, val As Double, bg As Long)
    With ws.Cells(r, c)
        .Value = val
        .Interior.Color = bg
        .Font.Color = cBlack()
        .Font.Size = 9
        .Font.Name = "Consolas"
        .HorizontalAlignment = xlHAlignRight
        .VerticalAlignment = xlVAlignCenter
        .NumberFormat = "0.00"
    End With
End Sub

Private Sub SwapS(arr() As String, i As Long, j As Long)
    Dim t As String: t = arr(i): arr(i) = arr(j): arr(j) = t
End Sub
Private Sub SwapD(arr() As Double, i As Long, j As Long)
    Dim t As Double: t = arr(i): arr(i) = arr(j): arr(j) = t
End Sub
Private Sub SwapL(arr() As Long, i As Long, j As Long)
    Dim t As Long: t = arr(i): arr(i) = arr(j): arr(j) = t
End Sub
