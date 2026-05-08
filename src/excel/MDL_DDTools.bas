Option Explicit

' ================================================================================
'  MODULE: MDL_DDTools
'  FieldCap DD Tools ? Excel Dashboard
'
'  QUICK START:
'   1. Alt+F11 ? Insert ? Module ? paste this file
'   2. Run InitDDTools once from the Immediate window:  ?InitDDTools
'   3. Use the "Import & Build" button on the DD Tools sheet
'
'  ARCHITECTURE:
'   ? "DD Tools"  ? visible dashboard sheet (created automatically)
'   ? "_FC_BHA"   ? hidden raw BHA CSV data
'   ? "_FC_Crew"  ? hidden raw Crew CSV data
'   ? "_FC_Job"   ? hidden raw Job Details CSV data
'  All hidden sheets use xlSheetVeryHidden (do not appear in tab bar)
'  All rendering uses ScreenUpdating=False + Calculation=Manual to prevent flicker
' ================================================================================

' ?? Sheet names ??????????????????????????????????????????????????????????????
Private Const SH_UI   As String = "DD Tools"
Private Const SH_BHA  As String = "_FC_BHA"
Private Const SH_CREW As String = "_FC_Crew"
Private Const SH_JOB  As String = "_FC_Job"

' ?? Layout: row anchors ???????????????????????????????????????????????????????
Private Const R_TITLE  As Long = 1
Private Const R_SUB    As Long = 2
Private Const R_DIV1   As Long = 3
Private Const R_STATS  As Long = 4
Private Const R_DIV2   As Long = 5
Private Const R_BTNS   As Long = 6
Private Const R_DIV3   As Long = 7
Private Const R_DETAIL As Long = 9   ' BHA detail begins here

' ?? Layout: column span ???????????????????????????????????????????????????????
Private Const C_LAST As Long = 15    ' rightmost column used

' Colors: high-contrast neutral (no dark filled backgrounds).
Private Function cBg()       As Long: cBg       = RGB(255, 255, 255): End Function
Private Function cCard()     As Long: cCard     = RGB(255, 255, 255): End Function
Private Function cHeader()   As Long: cHeader   = RGB(255, 255, 255): End Function
Private Function cAccent()   As Long: cAccent   = RGB(0,   0,   0):   End Function
Private Function cAccentH()  As Long: cAccentH  = RGB(0,   0,   0):   End Function
Private Function cBorder()   As Long: cBorder   = RGB(180, 180, 180): End Function
Private Function cText()     As Long: cText     = RGB(0,   0,   0):   End Function
Private Function cTextSec()  As Long: cTextSec  = RGB(40,  40,  40):  End Function
Private Function cAmber()    As Long: cAmber    = RGB(0,   0,   0):   End Function
Private Function cGreen()    As Long: cGreen    = RGB(30,  90,  30):  End Function  ' muted green
Private Function cRed()      As Long: cRed      = RGB(160, 30,  30):  End Function  ' muted red
Private Function cBtnFill()  As Long: cBtnFill  = RGB(0,   79,  79):  End Function
Private Function cBtnFillH() As Long: cBtnFillH = RGB(0,   105, 105): End Function
Private Function cBtnLine()  As Long: cBtnLine  = RGB(0,   60,  60):  End Function
Private Function cBtnText()  As Long: cBtnText  = RGB(0,   0,   0):   End Function


' ================================================================================
'  PUBLIC ENTRY POINTS
' ================================================================================

' Run once from the Immediate Window to bootstrap the DD Tools sheet
Public Sub InitDDTools()
    Application.ScreenUpdating = False
    Application.EnableEvents   = False

    SetupDataSheets

    Dim ws As Worksheet
    If Not SheetExists(SH_UI) Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SH_UI
    End If

    SetupUISheet Worksheets(SH_UI)
    AddControlButtons Worksheets(SH_UI)

    Application.ScreenUpdating = True
    Application.EnableEvents   = True

    MsgBox "DD Tools sheet created." & vbLf & _
           "Place your three FieldCap CSV exports in the same folder as this workbook," & vbLf & _
           "then click REFRESH on the sheet.", _
           vbInformation, "DD Tools Ready"
End Sub

' Scan the workbook's own folder for the three plugin CSV exports and import them.
' Matching is case-insensitive InStr on filename:
'   contains "job-details"    ? Job Details CSV
'   contains "crew"           ? Crew CSV
'   contains "bha-equipment"  ? BHA Equipment CSV
Public Sub RefreshData()
    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    If wbPath = "" Then
        MsgBox "Save the workbook to a folder first." & vbLf & _
               "The Refresh button scans the same folder as this file.", _
               vbInformation, "DD Tools"
        Exit Sub
    End If

    ' ?? Scan folder for matching CSVs (pick newest match per type) ???????????
    Dim fJob As String, fCrew As String, fBHA As String
    fJob = FindLatestCsvPathByToken(wbPath, "job-details")
    fBHA = FindLatestCsvPathByToken(wbPath, "bha-equipment")
    fCrew = FindLatestCsvPathByToken(wbPath, "crew")

    ' ?? Report any missing files ??????????????????????????????????????????????
    Dim missing As String
    If fJob  = "" Then missing = missing & Chr(10) & "  *  *job-details*.csv"
    If fCrew = "" Then missing = missing & Chr(10) & "  *  *crew*.csv"
    If fBHA  = "" Then missing = missing & Chr(10) & "  *  *bha-equipment*.csv"

    If missing <> "" Then
        MsgBox "Missing CSV export(s) in:" & Chr(10) & wbPath & Chr(10) & missing & Chr(10) & Chr(10) & _
               "Export them from the FieldCap Chrome plugin, then click Refresh again.", _
               vbExclamation, "DD Tools - Files Not Found"
        Exit Sub
    End If

    ' ?? Show what was found, confirm ?????????????????????????????????????????
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Found (newest matches):" & Chr(10) & _
                     "  " & Mid(fJob,  InStrRev(fJob,  Application.PathSeparator) + 1) & _
                     "  [" & Format(FileDateTime(fJob), "yyyy-mm-dd hh:nn:ss") & "]" & Chr(10) & _
                     "  " & Mid(fCrew, InStrRev(fCrew, Application.PathSeparator) + 1) & _
                     "  [" & Format(FileDateTime(fCrew), "yyyy-mm-dd hh:nn:ss") & "]" & Chr(10) & _
                     "  " & Mid(fBHA,  InStrRev(fBHA,  Application.PathSeparator) + 1) & _
                     "  [" & Format(FileDateTime(fBHA), "yyyy-mm-dd hh:nn:ss") & "]" & Chr(10) & Chr(10) & _
                     "Import and rebuild DD Tools?", _
                     vbQuestion + vbYesNo, "DD Tools Refresh")
    If confirm = vbNo Then Exit Sub

    ' Preserve user's current focus so refresh from other sheets does not redirect.
    Dim prevSheet As Worksheet
    Dim prevCellAddr As String
    Set prevSheet = ActiveSheet
    prevCellAddr = ActiveCell.Address(False, False)

    ' ?? Import & build ????????????????????????????????????????????????????????
    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual
    Application.StatusBar = "DD Tools: importing CSVs..."

    On Error GoTo ErrHandler

    SetupDataSheets
    ImportCSVToSheet fJob,  Worksheets(SH_JOB)
    ImportCSVToSheet fCrew, Worksheets(SH_CREW)
    ImportCSVToSheet fBHA,  Worksheets(SH_BHA)

    Application.StatusBar = "DD Tools: building dashboard..."
    BuildUI

    On Error Resume Next
    prevSheet.Activate
    prevSheet.Range(prevCellAddr).Select
    On Error GoTo 0

    Application.StatusBar      = False
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Application.StatusBar      = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "DD Tools"
End Sub

' Rebuild dashboard from already-imported hidden sheet data (no file dialogs)
Public Sub RebuildDashboard()
    If Not SheetExists(SH_BHA) Then
        MsgBox "No data found. Click REFRESH first (CSVs must be in the same folder as this workbook).", vbInformation, "DD Tools"
        Exit Sub
    End If

    Dim prevSheet As Worksheet
    Dim prevCellAddr As String
    Set prevSheet = ActiveSheet
    prevCellAddr = ActiveCell.Address(False, False)

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual
    Application.StatusBar = "DD Tools: rebuilding..."

    On Error GoTo ErrHandler
    BuildUI

    On Error Resume Next
    prevSheet.Activate
    prevSheet.Range(prevCellAddr).Select
    On Error GoTo 0

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Application.StatusBar      = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "DD Tools"
End Sub

' BHA button click handler ? reads Application.Caller to identify which BHA
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
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = Worksheets(SH_UI)

    HighlightBHAButton ws, bhaNum
    DrawStats           ws, bhaNum
    DrawBHADetail       ws, bhaNum
    DrawCumulativeTable ws
    DrawRawBHASummaryTable ws

    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
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
        ' Very hidden ? does not appear in sheet tab bar at all
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

' Handles quoted fields with embedded commas/newlines
Private Function ParseCSVLine(ByVal line As String) As String()
    Dim result(200) As String
    Dim idx    As Long: idx     = 0
    Dim pos    As Long: pos     = 1
    Dim inQ    As Boolean: inQ  = False
    Dim token  As String: token = ""

    Do While pos <= Len(line)
        Dim ch As String
        ch = Mid(line, pos, 1)

        If ch = Chr(34) Then                          ' double-quote
            If inQ And Mid(line, pos + 1, 1) = Chr(34) Then  ' escaped quote ""
                token = token & Chr(34)
                pos = pos + 1
            Else
                inQ = Not inQ
            End If
        ElseIf ch = "," And Not inQ Then
            result(idx) = token
            idx   = idx + 1
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

    uiStep = "SetupUISheet":  SetupUISheet  ws
    uiStep = "DrawHeader":    DrawHeader    ws
    uiStep = "DrawStats":     DrawStats     ws

    uiStep = "GetBHAList"
    Dim bhaList As Variant
    bhaList = GetBHAList()
    Dim bub As Long: bub = SafeUBound(bhaList)

    uiStep = "DrawBHAButtons (" & (bub + 1) & " BHAs)": DrawBHAButtons ws, bhaList
    uiStep = "AddControlButtons":                        AddControlButtons ws

    If bub >= 0 Then
        Dim defaultBHA As Long
        defaultBHA = CLng(bhaList(bub))
        uiStep = "HighlightBHAButton BHA=" & defaultBHA: HighlightBHAButton ws, defaultBHA
        uiStep = "DrawStats BHA=" & defaultBHA:         DrawStats           ws, defaultBHA
        uiStep = "DrawBHADetail BHA=" & defaultBHA:      DrawBHADetail       ws, defaultBHA
        uiStep = "DrawCumulativeTable":                  DrawCumulativeTable ws
        uiStep = "DrawRawBHASummaryTable":               DrawRawBHASummaryTable ws
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & " in BuildUI [" & uiStep & "]:" & Chr(10) & Err.Description, _
           vbCritical, "DD Tools - BuildUI"
End Sub

' ================================================================================
'  SHEET CHROME & LAYOUT
' ================================================================================

Private Sub SetupUISheet(ws As Worksheet)
    With ws
        .Cells.Interior.Color = cBg()
        .Cells.Font.Color     = cText()
        .Cells.Font.Name      = "Consolas"
        .Cells.Font.Size      = 9
        .Tab.Color            = cAccent()
    End With

    ' DisplayGridlines / DisplayHeadings are Window-level properties.
    ' Do not force sheet activation during refresh to preserve user focus.
    If ActiveSheet.Name = SH_UI Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If

    ' Column widths (characters at default font)
    ws.Columns("A").ColumnWidth = 1.5    ' left margin
    ws.Columns("B").ColumnWidth = 5      ' index / BHA#
    ws.Columns("C").ColumnWidth = 14     ' serial#
    ws.Columns("D").ColumnWidth = 12     ' item code
    ws.Columns("E").ColumnWidth = 32     ' description (wide)
    ws.Columns("F").ColumnWidth = 14     ' sub-description
    ws.Columns("G").ColumnWidth = 8      ' status / gap
    ws.Columns("H").ColumnWidth = 9      ' length
    ws.Columns("I").ColumnWidth = 9      ' accum length
    ws.Columns("J").ColumnWidth = 8      ' max OD
    ws.Columns("K").ColumnWidth = 8      ' min ID
    ws.Columns("L").ColumnWidth = 10     ' top conn
    ws.Columns("M").ColumnWidth = 10     ' bot conn
    ws.Columns("N").ColumnWidth = 10     ' total hrs
    ws.Columns("O").ColumnWidth = 10     ' extra

    ' Fixed row heights for the chrome
    ws.Rows(R_TITLE).RowHeight = 28
    ws.Rows(R_SUB).RowHeight   = 14
    ws.Rows(R_DIV1).RowHeight  = 2
    ws.Rows(R_STATS).RowHeight = 28
    ws.Rows(R_DIV2).RowHeight  = 2
    ws.Rows(R_BTNS).RowHeight  = 34
    ws.Rows(R_DIV3).RowHeight  = 2
    ws.Rows(8).RowHeight       = 4   ' spacer before detail
End Sub

Private Sub DrawHeader(ws As Worksheet)
    ' ?? Title ????????????????????????????????????????????????????????????????
    With ws.Range("A1:D1")
        .Merge
        .Value              = "  FIELDCAP  DD TOOLS"
        .Interior.Color     = cHeader()
        .Font.Color         = cAmber()
        .Font.Size          = 14
        .Font.Bold          = True
        .Font.Name          = "Consolas"
        .VerticalAlignment  = xlVAlignCenter
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color     = cAccent()
            .Weight    = xlMedium
        End With
    End With

    ' ?? Sub-title: job info ???????????????????????????????????????????????????
    Dim jobLabel As String
    jobLabel = GetJobHeaderLabel()
    With ws.Range(ws.Cells(R_SUB, 1), ws.Cells(R_SUB, C_LAST))
        .UnMerge
        .Value             = ""
        .Interior.Color    = cHeader()
        .Font.Color        = cTextSec()
        .Font.Size         = 9
        .VerticalAlignment = xlVAlignCenter
    End With
    ws.Cells(R_SUB, 1).Value = "  " & jobLabel

    ' ?? Accent divider ????????????????????????????????????????????????????????
    With ws.Range(ws.Cells(R_DIV1, 1), ws.Cells(R_DIV1, C_LAST))
        .Interior.Color = cAccent()
    End With
End Sub

' ================================================================================
'  STATS STRIP
' ================================================================================

Private Sub DrawStats(ws As Worksheet, Optional selectedBHA As Variant)
    Dim step As String
    On Error GoTo ErrHandler

    step = "SheetExists SH_BHA"
    If Not SheetExists(SH_BHA) Then Exit Sub

    step = "get bhaWs"
    Dim bhaWs As Worksheet
    Set bhaWs = Worksheets(SH_BHA)

    step = "FindCol BHA#":            Dim colBNum As Long: colBNum = FindCol(bhaWs, "BHA #")
    step = "FindCol Metres Drilled":  Dim colMtr  As Long: colMtr  = FindCol(bhaWs, "Metres Drilled")
    step = "FindCol BHA Total Hrs":   Dim colTot  As Long: colTot  = FindCol(bhaWs, "BHA Total Hrs")
    step = "FindCol BHA Hrs Slid":    Dim colSld  As Long: colSld  = FindCol(bhaWs, "BHA Hrs Slid")
    step = "FindCol BHA Hrs Rot":     Dim colRot  As Long: colRot  = FindCol(bhaWs, "BHA Hrs Rot")

    If colBNum = 0 Or colMtr = 0 Or colTot = 0 Or colSld = 0 Or colRot = 0 Then
        MsgBox "Column not found in _FC_BHA:" & Chr(10) & _
               "BHA#=" & colBNum & "  Metres=" & colMtr & _
               "  TotHrs=" & colTot & "  Slid=" & colSld & "  Rot=" & colRot & Chr(10) & _
               "Row1: " & bhaWs.Cells(1,1) & " | " & bhaWs.Cells(1,2) & " | " & bhaWs.Cells(1,7) & _
               " | " & bhaWs.Cells(1,8), vbExclamation, "DD Tools"
        Exit Sub
    End If

    step = "GetBHAList"
    Dim bhaList As Variant
    bhaList = GetBHAList()

    step = "SafeUBound"
    Dim dsub As Long: dsub = SafeUBound(bhaList)

    Dim activeBhaNum As Long
    If IsMissing(selectedBHA) Or IsEmpty(selectedBHA) Then
        If dsub >= 0 Then activeBhaNum = CLng(bhaList(dsub))
    Else
        activeBhaNum = CLng(selectedBHA)
    End If

    step = "GetFirstBHARow activeBHA=" & activeBhaNum
    Dim fr As Long
    fr = GetFirstBHARow(activeBhaNum)

    step = "totals active row"
    Dim totMetres As Double, totHrs As Double
    Dim totSlid   As Double, totRot As Double
    If fr > 0 Then
        step = "GetBHAMetricValue Metres"
        totMetres = GetBHAMetricValue(bhaWs, colBNum, colMtr, activeBhaNum)
        step = "GetNum TotHrs fr=" & fr & " col=" & colTot
        totHrs    = GetNum(bhaWs.Cells(fr, colTot))
        step = "GetNum Slid fr=" & fr & " col=" & colSld
        totSlid   = GetNum(bhaWs.Cells(fr, colSld))
        step = "GetNum Rot fr=" & fr & " col=" & colRot
        totRot    = GetNum(bhaWs.Cells(fr, colRot))
    End If

    step = "slidePct"
    Dim slidePct As Double
    If totHrs > 0 Then slidePct = totSlid / totHrs * 100

    step = "crewCount"
    Dim crewCount As Long
    If SheetExists(SH_CREW) Then
        On Error Resume Next
        crewCount = Application.WorksheetFunction.CountA( _
                        Worksheets(SH_CREW).Columns(1)) - 1
        On Error GoTo ErrHandler
        If crewCount < 0 Then crewCount = 0
    End If

    step = "StatBox Metres":   Dim r As Long: r = R_STATS
    StatBox ws, r, 2,  3,  "METRES DRILLED", Format(totMetres, "#,##0.00") & " m"
    step = "StatBox TotalHrs": StatBox ws, r, 4,  5,  "TOTAL HRS",  Format(totHrs,   "0.00") & " h"
    step = "StatBox SlideHrs": StatBox ws, r, 6,  7,  "SLIDE HRS",  Format(totSlid,  "0.00") & " h"
    step = "StatBox RotHrs":   StatBox ws, r, 8,  9,  "ROTATE HRS", Format(totRot,   "0.00") & " h"
    step = "StatBox SlidePct": StatBox ws, r, 10, 11, "SLIDE %",     Format(slidePct,"0.0") & "%"
    step = "StatBox BHA":      StatBox ws, r, 12, 13, "BHA",         CStr(IIf(activeBhaNum > 0, activeBhaNum, dsub + 1))
    step = "StatBox Crew":     StatBox ws, r, 14, 15, "CREW",        CStr(crewCount)

    step = "divider"
    With ws.Range(ws.Cells(R_DIV2, 1), ws.Cells(R_DIV2, C_LAST))
        .Interior.Color = cBorder()
    End With
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & " in DrawStats [" & step & "]:" & Chr(10) & Err.Description, _
           vbCritical, "DD Tools - DrawStats"
End Sub

Private Sub StatBox(ws As Worksheet, r As Long, _
                    c1 As Long, c2 As Long, _
                    label As String, val As String)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Merge
        .Value             = label & Chr(10) & val
        .WrapText          = True
        .Interior.Color    = cCard()
        .Font.Color        = cText()
        .Font.Size         = 9
        .Font.Name         = "Consolas"
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment   = xlVAlignCenter
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color     = cAccent()
            .Weight    = xlHairline
        End With
    End With
End Sub

' ================================================================================
'  BHA SELECTOR BUTTONS  (dynamic ? one button per BHA, auto-created on import)
' ================================================================================

Private Sub DrawBHAButtons(ws As Worksheet, bhaList As Variant)
    Dim step As String
    On Error GoTo ErrHandler

    step = "delete old BHABtn_ controls"
    Dim btnDel As Button
    For Each btnDel In ws.Buttons
        If Left(btnDel.Name, 7) = "BHABtn_" Then btnDel.Delete
    Next btnDel
    Dim shpDel As Shape
    For Each shpDel In ws.Shapes
        If Left(shpDel.Name, 7) = "BHABtn_" Then shpDel.Delete
    Next shpDel

    If Not IsArray(bhaList) Then Exit Sub
    Dim ub As Long
    On Error Resume Next
    ub = UBound(bhaList)
    If Err.Number <> 0 Then Exit Sub
    On Error GoTo ErrHandler
    If ub < 0 Then Exit Sub

    ' Place BHA buttons from G1 and flow right one cell each.
    Dim btnRow As Long: btnRow = 1
    Dim startColRight As Long: startColRight = 7   ' G

    Dim i As Long
    For i = 0 To ub
        Dim bNum As Long
        bNum = CLng(bhaList(i))

        Dim label As String
        label = "BHA" & bNum

        Dim targetCol As Long
        targetCol = startColRight + i

        Dim targetCell As Range
        Set targetCell = ws.Cells(btnRow, targetCol)

        ' Keep each button in equal-width cells for reliable alignment.
        ws.Columns(targetCol).ColumnWidth = 9

        step = "AddButton i=" & i
        Dim btn As Button
        Set btn = ws.Buttons.Add(targetCell.Left, targetCell.Top, targetCell.Width, targetCell.Height)

        step = "btn.Name"
        btn.Name = "BHABtn_" & bNum
        step = "btn.OnAction"
        btn.OnAction = "'" & ThisWorkbook.Name & "'!SelectBHA"
        step = "btn.Placement"
        btn.Placement = xlMoveAndSize

        step = "btn.Caption"
        StyleCellButton btn, label, cBtnFill(), cBtnLine(), cBtnText(), False
    Next i

    step = "divider row"
    With ws.Range(ws.Cells(R_DIV3, 1), ws.Cells(R_DIV3, C_LAST))
        .Interior.Color = cBorder()
    End With
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & " in DrawBHAButtons [" & step & "]:" & Chr(10) & Err.Description, _
           vbCritical, "DD Tools - DrawBHAButtons"
End Sub

Private Sub HighlightBHAButton(ws As Worksheet, activeBHA As Long)
    Dim btn As Button
    For Each btn In ws.Buttons
        If Left(btn.Name, 7) = "BHABtn_" Then
            Dim n As Long
            n = CLng(Split(btn.Name, "_")(1))
            If n = activeBHA Then
                StyleCellButton btn, "BHA" & CStr(n), cBtnFillH(), cBtnLine(), cBtnText(), True
            Else
                StyleCellButton btn, "BHA" & CStr(n), cBtnFill(), cBtnLine(), cBtnText(), False
            End If
        End If
    Next btn

    ' Store active BHA in a named cell for later reference
    On Error Resume Next
    ws.Names.Add "DD_ActiveBHA", "=" & Chr(34) & CStr(activeBHA) & Chr(34)
    On Error GoTo 0
End Sub

' ================================================================================
'  BHA DETAIL TABLE
' ================================================================================

Private Sub DrawBHADetail(ws As Worksheet, bhaNum As Long)
    ' ?? Clear detail area ????????????????????????????????????????????????????
    Dim clearTo As Long: clearTo = R_DETAIL + 100
    With ws.Range(ws.Cells(R_DETAIL, 1), ws.Cells(clearTo, C_LAST + 2))
        .ClearContents
        .ClearFormats
        .Interior.Color      = cBg()
        .Font.Color          = cText()
        .Font.Name           = "Consolas"
        .Font.Size           = 9
        .Font.Bold           = False
        .Borders.LineStyle   = xlNone
    End With

    If Not SheetExists(SH_BHA) Then Exit Sub

    Dim bhaWs As Worksheet
    Set bhaWs = Worksheets(SH_BHA)

    Dim fr As Long: fr = GetFirstBHARow(bhaNum)
    If fr = 0 Then Exit Sub

    ' Column lookups (from CSV header row 1)
    Dim cSec  As Long: cSec  = FindCol(bhaWs, "Section")
    Dim cStat As Long: cStat = FindCol(bhaWs, "Status")
    Dim cMot  As Long: cMot  = FindCol(bhaWs, "Motor")
    Dim cGui  As Long: cGui  = FindCol(bhaWs, "Guidance")
    Dim cMtr  As Long: cMtr  = FindCol(bhaWs, "Metres Drilled")
    Dim cTot  As Long: cTot  = FindCol(bhaWs, "BHA Total Hrs")
    Dim cSld  As Long: cSld  = FindCol(bhaWs, "BHA Hrs Slid")
    Dim cRot  As Long: cRot  = FindCol(bhaWs, "BHA Hrs Rot")
    Dim cCrc  As Long: cCrc  = FindCol(bhaWs, "BHA Hrs Circ")
    Dim cBlw  As Long: cBlw  = FindCol(bhaWs, "BHA Below Rot")
    Dim cAct  As Long: cAct  = FindCol(bhaWs, "Activated On")
    Dim cCmp  As Long: cCmp  = FindCol(bhaWs, "Completed On")
    Dim cBNum As Long: cBNum = FindCol(bhaWs, "BHA #")
    Dim cSer  As Long: cSer  = FindCol(bhaWs, "Serial #")
    Dim cCod  As Long: cCod  = FindCol(bhaWs, "Item Code")
    Dim cDes  As Long: cDes  = FindCol(bhaWs, "Description")
    Dim cSub  As Long: cSub  = FindCol(bhaWs, "Sub Description")

    ' BHA-level values (all components share these)
    Dim section As String: section = SafeStr(bhaWs.Cells(fr, cSec))
    Dim status  As String: status  = SafeStr(bhaWs.Cells(fr, cStat))
    Dim motor   As String: motor   = SafeStr(bhaWs.Cells(fr, cMot))
    Dim guid    As String: guid    = SafeStr(bhaWs.Cells(fr, cGui))
    Dim metres  As Double: metres  = GetBHAMetricValue(bhaWs, cBNum, cMtr, bhaNum)
    Dim totHrs  As Double: totHrs  = GetNum(bhaWs.Cells(fr, cTot))
    Dim sldHrs  As Double: sldHrs  = GetNum(bhaWs.Cells(fr, cSld))
    Dim rotHrs  As Double: rotHrs  = GetNum(bhaWs.Cells(fr, cRot))
    Dim crcHrs  As Double: crcHrs  = GetNum(bhaWs.Cells(fr, cCrc))
    Dim blwHrs  As Double: blwHrs  = GetNum(bhaWs.Cells(fr, cBlw))
    Dim actOn   As String: actOn   = SafeStr(bhaWs.Cells(fr, cAct))
    Dim cmpOn   As String: cmpOn   = SafeStr(bhaWs.Cells(fr, cCmp))

    Dim r As Long: r = R_DETAIL

    ' ?? BHA title bar ?????????????????????????????????????????????????????????
    ws.Rows(r).RowHeight = 20
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST))
        .Merge
        .Value             = "  BHA " & bhaNum & "  |  " & UCase(section) & _
                             "  |  STATUS: " & UCase(status) & _
                             "  |  MOTOR: " & motor & _
                             "  |  GUIDANCE: " & guid
        .Interior.Color    = cHeader()
        .Font.Color        = cAmber()
        .Font.Bold         = True
        .Font.Size         = 10
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color     = cAccent()
            .Weight    = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color     = cAccent()
            .Weight    = xlMedium
        End With
    End With
    r = r + 1

    ' ?? Hours strip ???????????????????????????????????????????????????????????
    ws.Rows(r).RowHeight = 28

    Dim hLabels(7) As String
    Dim hValues(7) As String
    hLabels(0) = "METRES":   hValues(0) = Format(metres,  "#,##0.00") & " m"
    hLabels(1) = "HRS SLID": hValues(1) = Format(sldHrs,  "0.00") & " h"
    hLabels(2) = "HRS ROT":  hValues(2) = Format(rotHrs,  "0.00") & " h"
    hLabels(3) = "HRS CIRC": hValues(3) = Format(crcHrs,  "0.00") & " h"
    hLabels(4) = "TOTAL HRS":hValues(4) = Format(totHrs,  "0.00") & " h"
    hLabels(5) = "BELOW ROT":hValues(5) = Format(blwHrs,  "0.00") & " h"
    hLabels(6) = "ACTIVATED":hValues(6) = actOn
    hLabels(7) = "COMPLETED":hValues(7) = cmpOn

    ' Keep top strip synchronized with selected BHA every time detail renders.
    DrawStats ws, bhaNum

    Dim col As Long: col = 2
    Dim hi As Integer
    For hi = 0 To 7
        With ws.Cells(r, col)
            .Value             = hLabels(hi) & Chr(10) & hValues(hi)
            .WrapText          = True
            .Interior.Color    = cCard()
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment   = xlVAlignCenter
            .Font.Size         = 8
            .Font.Name         = "Consolas"
            If hi = 4 Then
                .Font.Color = cAmber()
                .Font.Bold  = True
            Else
                .Font.Color = cText()
            End If
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color     = cBorder()
                .Weight    = xlHairline
            End With
        End With
        col = col + 1
    Next hi
    r = r + 1

    ' Thin divider
    ws.Rows(r).RowHeight = 2
    ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST)).Interior.Color = cBorder()
    r = r + 1

    ' ?? Component table column headers ????????????????????????????????????????
    ws.Rows(r).RowHeight = 14

    Dim hdr(4) As String: Dim hdrC(4) As Long
    hdr(0) = "#":         hdrC(0) = 2
    hdr(1) = "SERIAL #":  hdrC(1) = 3
    hdr(2) = "ITEM CODE": hdrC(2) = 4
    hdr(3) = "DESCRIPTION": hdrC(3) = 5
    hdr(4) = "TOTAL HRS": hdrC(4) = 8

    Dim h As Integer
    For h = 0 To 4
        With ws.Cells(r, hdrC(h))
            .Value             = hdr(h)
            .Interior.Color    = cHeader()
            .Font.Color        = cAmber()
            .Font.Bold         = True
            .Font.Size         = 8
            .HorizontalAlignment = xlHAlignCenter
        End With
    Next h
    r = r + 1

    ' Accent underline under headers
    ws.Rows(r).RowHeight = 1
    ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST)).Interior.Color = cAccent()
    r = r + 1

    ' ?? Component data rows ???????????????????????????????????????????????????
    Dim lastRow As Long
    lastRow = bhaWs.Cells(bhaWs.Rows.Count, cBNum).End(xlUp).Row

    Dim compIdx As Long: compIdx = 1
    Dim dr As Long
    For dr = 2 To lastRow
        Dim rawBNum As Variant
        rawBNum = bhaWs.Cells(dr, cBNum).Value
        If IsNumeric(rawBNum) Then
            If CLng(rawBNum) = bhaNum Then
                ws.Rows(r).RowHeight = 14

                Dim bg As Long
                If compIdx Mod 2 = 0 Then bg = cCard() Else bg = cBg()

                ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST)).Interior.Color = bg

                PutCell ws, r, 2,  CStr(compIdx),                    bg, cTextSec(), False, xlRight
                PutCell ws, r, 3,  SafeStr(bhaWs.Cells(dr, cSer)),   bg, cText(),    False, xlLeft
                PutCell ws, r, 4,  SafeStr(bhaWs.Cells(dr, cCod)),   bg, cTextSec(), False, xlLeft
                PutCell ws, r, 5,  SafeStr(bhaWs.Cells(dr, cDes)),   bg, cText(),    False, xlLeft
                ' Total Hrs column ? amber highlight, same value for all in BHA
                With ws.Cells(r, 8)
                    .Value             = Format(totHrs, "0.00")
                    .Interior.Color    = bg
                    .Font.Color        = cAmber()
                    .Font.Bold         = True
                    .Font.Size         = 9
                    .HorizontalAlignment = xlRight
                End With

                compIdx = compIdx + 1
                r = r + 1
            End If
        End If
    Next dr

    ' Footer rule
    ws.Rows(r).RowHeight = 2
    ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST)).Interior.Color = cBorder()
    r = r + 3

    ' Store cumulative table anchor (sheet-scoped named range)
    On Error Resume Next
    ws.Names.Add "DD_CumulRow", "='" & SH_UI & "'!$A$" & r
    On Error GoTo 0
End Sub

' ================================================================================
'  RIGHT-SIDE RAW BHA SUMMARY TABLE (VLOOKUP-friendly)
' ================================================================================
Private Sub DrawRawBHASummaryTable(ws As Worksheet)
    If Not SheetExists(SH_BHA) Then Exit Sub

    Dim bhaWs As Worksheet
    Set bhaWs = Worksheets(SH_BHA)

    Dim cBNum As Long: cBNum = FindCol(bhaWs, "BHA #")
    Dim cMtr  As Long: cMtr  = FindCol(bhaWs, "Metres Drilled")
    Dim cMSld As Long: cMSld = FindCol(bhaWs, "BHA Mtrs Slid")
    Dim cMRot As Long: cMRot = FindCol(bhaWs, "BHA Mtrs Rot")
    Dim cSld  As Long: cSld  = FindCol(bhaWs, "BHA Hrs Slid")
    Dim cRot  As Long: cRot  = FindCol(bhaWs, "BHA Hrs Rot")
    Dim cCrc  As Long: cCrc  = FindCol(bhaWs, "BHA Hrs Circ")
    Dim cTot  As Long: cTot  = FindCol(bhaWs, "BHA Total Hrs")
    Dim cBlw  As Long: cBlw  = FindCol(bhaWs, "BHA Below Rot")
    If cBNum = 0 Then Exit Sub

    Dim startCol As Long: startCol = 17 ' Q
    ' Avoid DD detail divider rows (height 1-2) so summary rows don't appear hidden.
    Dim startRow As Long: startRow = R_DETAIL + 5
    Dim endCol As Long: endCol = startCol + 8

    ' Keep this region isolated and unmerged so it never collides with merged UI areas.
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 20, endCol))
        On Error Resume Next
        .UnMerge
        On Error GoTo 0
        .ClearContents
        .ClearFormats
        .Interior.Color = cBg()
        .Font.Color = cText()
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With

    ws.Cells(startRow, startCol).Value = "RAW BHA SUMMARY (for formulas)"
    ws.Cells(startRow, startCol).Font.Bold = True
    ws.Cells(startRow, startCol).Font.Size = 10
    ws.Cells(startRow, startCol).Font.Color = cText()
    ws.Rows(startRow).RowHeight = 16

    Dim hdr() As String
    hdr = Split("BHA,Meters Drilled,Mtrs Slid,Mtrs Rot,Hrs Slid,Hrs Rot,Hrs Circ,Total Hrs,Below Rot", ",")
    Dim i As Long
    For i = 0 To UBound(hdr)
        With ws.Cells(startRow + 1, startCol + i)
            .Value = hdr(i)
            .Font.Bold = True
            .Interior.Color = cBg()
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = cBorder()
        End With
    Next i
    ws.Rows(startRow + 1).RowHeight = 16

    Dim bhaList As Variant
    bhaList = GetBHAList()
    Dim ub As Long: ub = SafeUBound(bhaList)
    Dim r As Long: r = startRow + 2
    For i = 0 To ub
        Dim fr As Long
        fr = GetFirstBHARow(CLng(bhaList(i)))
        If fr > 0 Then
            ' Ensure summary rows remain visible even if left-side dashboard uses
            ' compact divider row heights.
            ws.Rows(r).Hidden = False
            If ws.Rows(r).RowHeight < 14 Then ws.Rows(r).RowHeight = 14
            ws.Cells(r, startCol + 0).Value = CLng(bhaList(i))
            ' Pull metres as the first non-zero metric across all rows for this BHA
            ' (first row can be blank/0 in some exports).
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
            r = r + 1
        End If
    Next i

    ' Light grid for readability
    With ws.Range(ws.Cells(startRow + 1, startCol), ws.Cells(r - 1, endCol))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = cBorder()
    End With
    ws.Range(ws.Cells(startRow + 2, startCol), ws.Cells(r - 1, startCol)).HorizontalAlignment = xlHAlignCenter

    ' Helpful named range for downstream formulas.
    On Error Resume Next
    ws.Names.Add "DD_RawBHASummary", "='" & SH_UI & "'!" & _
        ws.Range(ws.Cells(startRow + 1, startCol), ws.Cells(r - 1, endCol)).Address
    On Error GoTo 0
End Sub

' ================================================================================
'  CUMULATIVE COMPONENT HOURS TABLE
' ================================================================================

Private Sub DrawCumulativeTable(ws As Worksheet)
    ' Find cumulative anchor row
    Dim cStart As Long: cStart = 0
    On Error Resume Next
    cStart = ws.Range("DD_CumulRow").Row
    On Error GoTo 0
    If cStart = 0 Then cStart = 80

    ' Clear downward
    With ws.Range(ws.Cells(cStart, 1), ws.Cells(cStart + 200, C_LAST + 2))
        .ClearContents
        .ClearFormats
        .Interior.Color    = cBg()
        .Font.Color        = cText()
        .Font.Name         = "Consolas"
        .Font.Size         = 9
        .Font.Bold         = False
        .Borders.LineStyle = xlNone
    End With

    If Not SheetExists(SH_BHA) Then Exit Sub

    Dim bhaWs As Worksheet
    Set bhaWs = Worksheets(SH_BHA)

    Dim cBNum As Long: cBNum = FindCol(bhaWs, "BHA #")
    Dim cSer  As Long: cSer  = FindCol(bhaWs, "Serial #")
    Dim cCod  As Long: cCod  = FindCol(bhaWs, "Item Code")
    Dim cDes  As Long: cDes  = FindCol(bhaWs, "Description")
    Dim cTot  As Long: cTot  = FindCol(bhaWs, "BHA Total Hrs")
    Dim cSld  As Long: cSld  = FindCol(bhaWs, "BHA Hrs Slid")
    Dim cRot  As Long: cRot  = FindCol(bhaWs, "BHA Hrs Rot")
    Dim cMtr  As Long: cMtr  = FindCol(bhaWs, "Metres Drilled")

    Dim lastRow As Long
    lastRow = bhaWs.Cells(bhaWs.Rows.Count, cBNum).End(xlUp).Row

    ' ?? Rollup arrays (SERIAL-FIRST fatigue model) ????????????????????????????
    Dim rSerial() As String
    Dim rDesc()   As String
    Dim rTot()    As Double
    Dim rSld()    As Double
    Dim rRot()    As Double
    Dim rMtr()    As Double
    Dim rCnt()    As Long
    Dim rSz       As Long: rSz = 0

    ReDim rSerial(0): ReDim rDesc(0)
    ReDim rTot(0):    ReDim rSld(0)
    ReDim rRot(0):    ReDim rMtr(0): ReDim rCnt(0)

    If cSer = 0 Then Exit Sub

    Dim seenSerialBha As New Collection
    Dim dr As Long
    For dr = 2 To lastRow
        If Not IsNumeric(bhaWs.Cells(dr, cBNum).Value) Then GoTo SkipRow
        Dim bNum As Long: bNum = CLng(bhaWs.Cells(dr, cBNum).Value)

        Dim srl As String
        srl = Trim(SafeStr(bhaWs.Cells(dr, cSer)))
        If srl = "" Then GoTo SkipRow

        ' De-dupe repeated rows of same serial inside same BHA.
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

        ' Find in serial rollup
        Dim idx As Long: idx = -1
        Dim k As Long
        For k = 0 To rSz - 1
            If rSerial(k) = srl Then
                idx = k
                Exit For
            End If
        Next k

        If idx = -1 Then
            idx = rSz
            rSz = rSz + 1
            ReDim Preserve rSerial(rSz - 1): ReDim Preserve rDesc(rSz - 1)
            ReDim Preserve rTot(rSz - 1):    ReDim Preserve rSld(rSz - 1)
            ReDim Preserve rRot(rSz - 1):    ReDim Preserve rMtr(rSz - 1)
            ReDim Preserve rCnt(rSz - 1)
            rSerial(idx) = srl
            rDesc(idx) = SafeStr(bhaWs.Cells(dr, cDes))
        End If

        rCnt(idx) = rCnt(idx) + 1 ' number of BHAs this serial appears in

        ' Accumulate BHA-level hours once per serial-per-BHA
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

    ' ?? Sort descending by total hours ?????????????????????????????????????????
    Dim swapped As Boolean
    Do
        swapped = False
        For k = 0 To rSz - 2
            If rTot(k) < rTot(k + 1) Then
                SwapS rSerial, k, k + 1: SwapS rDesc, k, k + 1
                SwapD rTot,  k, k + 1: SwapD rSld,  k, k + 1
                SwapD rRot,  k, k + 1: SwapD rMtr,  k, k + 1
                SwapL rCnt,  k, k + 1
                swapped = True
            End If
        Next k
    Loop While swapped

    ' ?? Render cumulative section ??????????????????????????????????????????????
    Dim r As Long: r = cStart

    ws.Rows(r).RowHeight = 8
    r = r + 1

    ' Section title
    ws.Rows(r).RowHeight = 18
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST))
        .Merge
        .Value             = "  CUMULATIVE COMPONENT HOURS  |  ALL BHAs"
        .Interior.Color    = cHeader()
        .Font.Color        = cAmber()
        .Font.Bold         = True
        .Font.Size         = 10
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color     = cAccent()
            .Weight    = xlThin
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color     = cAccent()
            .Weight    = xlMedium
        End With
    End With
    r = r + 1

    ' Column headers
    ws.Rows(r).RowHeight = 14
    Dim ch(6) As String: Dim cc(6) As Long
    ch(0) = "SERIAL #":     cc(0) = 2
    ch(1) = "DESCRIPTION":  cc(1) = 3
    ch(2) = "COUNT":        cc(2) = 8
    ch(3) = "METRES":       cc(3) = 9
    ch(4) = "HRS SLID":     cc(4) = 10
    ch(5) = "HRS ROT":      cc(5) = 11
    ch(6) = "TOTAL HRS":    cc(6) = 14

    Dim ci As Integer
    For ci = 0 To 6
        With ws.Cells(r, cc(ci))
            .Value             = ch(ci)
            .Interior.Color    = cHeader()
            .Font.Color        = cAmber()
            .Font.Bold         = True
            .Font.Size         = 8
            .HorizontalAlignment = xlHAlignCenter
        End With
    Next ci
    r = r + 1

    ws.Rows(r).RowHeight = 1
    ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST)).Interior.Color = cAccent()
    r = r + 1

    ' Data rows
    Dim ri As Long
    For ri = 0 To rSz - 1
        ws.Rows(r).RowHeight = 14

        Dim bg As Long
        If (ri + 1) Mod 2 = 0 Then bg = cCard() Else bg = cBg()
        ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST)).Interior.Color = bg

        PutCell ws, r, 2,  rSerial(ri), bg, cTextSec(), False, xlLeft
        PutCell ws, r, 3,  rDesc(ri), bg, cText(),    False, xlLeft

        With ws.Cells(r, 8)
            .Value = rCnt(ri)
            .Interior.Color = bg
            .Font.Color = cTextSec()
            .HorizontalAlignment = xlRight
        End With

        PutNum ws, r, 9,  rMtr(ri), bg
        PutNum ws, r, 10, rSld(ri), bg
        PutNum ws, r, 11, rRot(ri), bg

        ' Total hours ? highlight if > 0
        With ws.Cells(r, 14)
            .Value = Format(rTot(ri), "0.00")
            .Interior.Color = bg
            If rTot(ri) > 0 Then
                .Font.Color = cAmber()
                .Font.Bold  = True
            Else
                .Font.Color = cTextSec()
                .Font.Bold  = False
            End If
            .HorizontalAlignment = xlRight
        End With

        r = r + 1
    Next ri

    ' Footer rule
    ws.Rows(r).RowHeight = 2
    ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST)).Interior.Color = cBorder()
    r = r + 2

    ' ------------------------------------------------------------------------
    ' OVER-LIMIT COMPONENTS (SERIAL HOURS > 300) - fatigue tracking
    ' ------------------------------------------------------------------------
    Dim oSerial() As String, oDesc() As String
    Dim oTot() As Double, oBhaCnt() As Long
    Dim oSz As Long: oSz = 0
    ReDim oSerial(0): ReDim oDesc(0)
    ReDim oTot(0): ReDim oBhaCnt(0)

    For k = 0 To rSz - 1
        If rTot(k) > 300# Then
            oSz = oSz + 1
            ReDim Preserve oSerial(oSz - 1)
            ReDim Preserve oDesc(oSz - 1)
            ReDim Preserve oTot(oSz - 1)
            ReDim Preserve oBhaCnt(oSz - 1)
            oSerial(oSz - 1) = rSerial(k)
            oDesc(oSz - 1) = rDesc(k)
            oTot(oSz - 1) = rTot(k)
            oBhaCnt(oSz - 1) = rCnt(k)
        End If
    Next k

    ' Sort over-limit rows descending by serial hours
    If oSz > 1 Then
        Do
            swapped = False
            For k = 0 To oSz - 2
                If oTot(k) < oTot(k + 1) Then
                    SwapS oSerial, k, k + 1
                    SwapS oDesc, k, k + 1
                    SwapD oTot, k, k + 1
                    SwapL oBhaCnt, k, k + 1
                    swapped = True
                End If
            Next k
        Loop While swapped
    End If

    ' Section header
    ws.Rows(r).RowHeight = 18
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, C_LAST))
        .Merge
        .Value = "  OVER-LIMIT COMPONENTS  (>300 h)  |  SERIAL-LEVEL"
        .Interior.Color = cHeader()
        .Font.Color = cText()
        .Font.Bold = True
        .Font.Size = 10
    End With
    r = r + 1

    ' Column headers
    ws.Rows(r).RowHeight = 14
    Dim oh(3) As String: Dim oc(3) As Long
    oh(0) = "SERIAL #":  oc(0) = 2
    oh(1) = "DESCRIPTION": oc(1) = 4
    oh(2) = "BHAs":      oc(2) = 10
    oh(3) = "SERIAL HRS": oc(3) = 14

    For ci = 0 To 3
        With ws.Cells(r, oc(ci))
            .Value = oh(ci)
            .Interior.Color = cBg()
            .Font.Color = cText()
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
        End With
    Next ci
    r = r + 1

    ' Data rows
    If oSz = 0 Then
        ws.Cells(r, 2).Value = "No components over 300h."
        ws.Cells(r, 2).Font.Color = cTextSec()
        r = r + 1
    Else
        For k = 0 To oSz - 1
            ws.Rows(r).RowHeight = 14
            PutCell ws, r, 2, oSerial(k), cBg(), cText(), True, xlLeft
            PutCell ws, r, 4, oDesc(k), cBg(), cText(), False, xlLeft
            With ws.Cells(r, 10)
                .Value = oBhaCnt(k)
                .HorizontalAlignment = xlRight
                .Font.Color = cTextSec()
            End With
            With ws.Cells(r, 14)
                .Value = Format(oTot(k), "0.00")
                .HorizontalAlignment = xlRight
                .Font.Color = cRed()
                .Font.Bold = True
            End With
            r = r + 1
        Next k
    End If

    ' Grid lines for over-limit section
    With ws.Range(ws.Cells(r - IIf(oSz = 0, 2, oSz + 1), 2), ws.Cells(r - 1, C_LAST))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = cBorder()
    End With
End Sub

' ================================================================================
'  CONTROL BUTTONS  (Import & Build, Rebuild)
' ================================================================================

Private Sub AddControlButtons(ws As Worksheet)
    ' Remove existing control buttons
    Dim btnDel As Button
    For Each btnDel In ws.Buttons
        If btnDel.Name = "BtnImport" Or btnDel.Name = "BtnRebuild" Then btnDel.Delete
    Next btnDel
    Dim shpDel As Shape
    For Each shpDel In ws.Shapes
        If shpDel.Name = "BtnImport" Or shpDel.Name = "BtnRebuild" Then shpDel.Delete
    Next shpDel

    ' Place control buttons in E1 and F1 (embedded, same style).
    Dim colRefresh As Long: colRefresh = 5
    Dim colRebuild As Long: colRebuild = 6
    ws.Columns(colRefresh).ColumnWidth = 10
    ws.Columns(colRebuild).ColumnWidth = 10

    Dim cellRefresh As Range
    Dim cellRebuild As Range
    Set cellRefresh = ws.Cells(R_TITLE, colRefresh)
    Set cellRebuild = ws.Cells(R_TITLE, colRebuild)

    Dim bImport As Button
    Set bImport = ws.Buttons.Add(cellRefresh.Left, cellRefresh.Top, cellRefresh.Width, cellRefresh.Height)
    With bImport
        .Name = "BtnImport"
        .OnAction = "'" & ThisWorkbook.Name & "'!RefreshData"
        .Placement = xlMoveAndSize
    End With
    StyleCellButton bImport, "REFRESH", cBtnFill(), cBtnLine(), cBtnText(), True

    Dim bRebuild As Button
    Set bRebuild = ws.Buttons.Add(cellRebuild.Left, cellRebuild.Top, cellRebuild.Width, cellRebuild.Height)
    With bRebuild
        .Name = "BtnRebuild"
        .OnAction = "'" & ThisWorkbook.Name & "'!RebuildDashboard"
        .Placement = xlMoveAndSize
    End With
    StyleCellButton bRebuild, "REBUILD", cBtnFill(), cBtnLine(), cBtnText(), True
End Sub

Private Sub StyleCellButton(btn As Button, caption As String, fillColor As Long, lineColor As Long, textColor As Long, isBold As Boolean)
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
'  DATA HELPERS
' ================================================================================

Private Function GetBHAList() As Variant
    GetBHAList = Array()

    Dim step As String
    On Error GoTo ErrHandler

    step = "SheetExists SH_BHA"
    If Not SheetExists(SH_BHA) Then Exit Function

    step = "Worksheets(SH_BHA)"
    Dim ws As Worksheet
    Set ws = Worksheets(SH_BHA)

    step = "FindCol BHA#"
    Dim col As Long: col = FindCol(ws, "BHA #")
    If col = 0 Then
        MsgBox "GetBHAList: 'BHA #' column not found." & Chr(10) & _
               "Row 1, col 1 = [" & ws.Cells(1,1).Value & "]  col 2 = [" & ws.Cells(1,2).Value & "]", _
               vbExclamation, "DD Tools"
        Exit Function
    End If

    step = "lastRow"
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row

    step = "build seen collection lastRow=" & lastRow
    Dim seen As New Collection
    Dim r As Long
    For r = 2 To lastRow
        Dim v As Variant: v = ws.Cells(r, col).Value
        If IsNumeric(v) And v <> "" Then
            On Error Resume Next
            seen.Add CLng(v), CStr(CLng(v))
            On Error GoTo ErrHandler
        End If
    Next r

    step = "seen.Count=" & seen.Count
    If seen.Count = 0 Then Exit Function

    step = "ReDim result"
    Dim result() As Variant
    ReDim result(seen.Count - 1)

    step = "fill result from seen"
    Dim i As Long
    For i = 1 To seen.Count
        result(i - 1) = seen(i)
    Next i

    step = "insertion sort UBound=" & UBound(result)
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
    Exit Function

ErrHandler:
    MsgBox "Error " & Err.Number & " in GetBHAList [" & step & "]:" & Chr(10) & Err.Description, _
           vbCritical, "DD Tools - GetBHAList"
End Function

Private Function GetFirstBHARow(bhaNum As Long) As Long
    GetFirstBHARow = 0
    If Not SheetExists(SH_BHA) Then Exit Function

    Dim ws As Worksheet: Set ws = Worksheets(SH_BHA)
    Dim col As Long:     col = FindCol(ws, "BHA #")
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

' Return first non-zero metric for a BHA by scanning all rows for that BHA.
' This avoids false 0 values when the first BHA row is sparse.
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

Private Function GetBHASection(bhaNum As Long) As String
    GetBHASection = ""
    Dim fr As Long: fr = GetFirstBHARow(bhaNum)
    If fr = 0 Then Exit Function

    Dim ws As Worksheet: Set ws = Worksheets(SH_BHA)
    Dim col As Long: col = FindCol(ws, "Section")
    If col = 0 Then Exit Function

    GetBHASection = SafeStr(ws.Cells(fr, col))
End Function

Private Function GetJobHeaderLabel() As String
    GetJobHeaderLabel = "No job data imported"
    If Not SheetExists(SH_JOB) Then Exit Function

    Dim ws As Worksheet: Set ws = Worksheets(SH_JOB)

    ' Job Details CSV has headers in row 1, values in row 2
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
                label = label & "ID: " & val & "  "
            Case InStr(hd, "client") > 0 And InStr(hd, "job") > 0
                label = label & val & "  "
            Case InStr(hd, "well") > 0
                label = label & "WELL: " & val & "  "
            Case InStr(hd, "status") > 0
                label = label & "[" & UCase(val) & "]"
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
        .NumberFormat        = "@"   ' force text so numeric-looking codes don't show as 1E+05
        .Value               = val
        .Interior.Color      = bg
        .Font.Color          = fg
        .Font.Bold           = bold
        .Font.Size           = 9
        .Font.Name           = "Consolas"
        .HorizontalAlignment = align
        .VerticalAlignment   = xlVAlignCenter
    End With
End Sub

Private Sub PutNum(ws As Worksheet, r As Long, c As Long, val As Double, bg As Long)
    With ws.Cells(r, c)
        .Value               = val
        .Interior.Color      = bg
        .Font.Color          = cText()
        .Font.Size           = 9
        .Font.Name           = "Consolas"
        .HorizontalAlignment = xlRight
        .VerticalAlignment   = xlVAlignCenter
        .NumberFormat        = "0.00"
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
