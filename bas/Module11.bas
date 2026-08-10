Attribute VB_Name = "Module11"
Option Explicit

Public Sub RTP_printW_Email()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim c As Range
    Dim filePath As String
    Dim attachedCount As Long
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim wsData As Worksheet
    Dim r As Long
    Dim reportPng As String
    Dim prevInteractive As Boolean
    Dim prevCursor As XlMousePointer
    Dim tmpDir As String

    On Error GoTo CleanFail

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    prevInteractive = Application.Interactive
    prevCursor = Application.Cursor

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Interactive = False
    Application.Cursor = xlWait
    Application.Calculation = xlCalculationManual

    Set wsData = ThisWorkbook.Worksheets("Data")
    On Error Resume Next
    wsData.Range("B2:F55").Calculate
    On Error GoTo CleanFail

    ' Corridor + ops PNG is an attachment only; the mail body is the B2:F55 HTML table.
    tmpDir = Environ$("TEMP")
    reportPng = ""
    On Error Resume Next
    reportPng = MDL_CorridorImage.RenderCorridorPng(tmpDir & "\daily_report.png", "all")
    On Error GoTo CleanFail

    Set OutApp = GetOutlookAppSafe()
    If OutApp Is Nothing Then
        MsgBox "EMAIL cancelled or Outlook desktop is not available." & vbCrLf & vbCrLf & _
               "Open classic Outlook first, then click EMAIL again." & vbCrLf & _
               "(The new Outlook / Mail app does not support this macro.)", _
               vbExclamation, "EMAIL"
        GoTo CleanExit
    End If

    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = "company.man@northwind.example"
        .CC = "field@northwind.example;ops@demo-dd.example"
        .BCC = ""
        .Subject = "Daily Drilling Summary - Apex-214 - demo-rig"

        attachedCount = 0
        For Each c In wsData.Range("I29:I33")
            filePath = Trim$(CStr(c.Value & ""))
            If Len(filePath) > 0 Then
                If FileExistsFast(filePath) Then
                    .Attachments.Add filePath
                    attachedCount = attachedCount + 1
                End If
            End If
        Next c

        If Len(reportPng) > 0 Then
            If FileExistsFast(reportPng) Then
                .Attachments.Add reportPng
                attachedCount = attachedCount + 1
            End If
        End If

        .HTMLBody = RangeToHTML(wsData.Range("B2:F55"))
    End With

    ' I29:I32 are merged across columns; ClearContents on I29 alone fails.
    ' Clear each MergeArea. Keep Attachement 5 (I33).
    For r = 29 To 32
        wsData.Range("I" & r).MergeArea.ClearContents
    Next r

    ' Restore Excel UI before showing the mail so Outlook does not fight a locked
    ' Excel session for focus.
    Application.Cursor = prevCursor
    Application.Interactive = prevInteractive
    Application.ScreenUpdating = prevScreenUpdating

    OutMail.Display

CleanExit:
    On Error Resume Next
    Application.Cursor = prevCursor
    Application.Interactive = prevInteractive
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Set OutMail = Nothing
    Set OutApp = Nothing
    Set wsData = Nothing
    Exit Sub

CleanFail:
    Dim msg As String
    msg = Err.Description
    Resume CleanExitMessage

CleanExitMessage:
    On Error Resume Next
    Application.Cursor = prevCursor
    Application.Interactive = prevInteractive
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Set OutMail = Nothing
    Set OutApp = Nothing
    Set wsData = Nothing
    MsgBox "EMAIL failed:" & vbCrLf & vbCrLf & msg, vbExclamation, "EMAIL"
End Sub


Private Function GetOutlookAppSafe() As Object
    Dim app As Object

    ' Prefer an already-running Outlook instance (avoids CreateObject hangs).
    On Error Resume Next
    Set app = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If Not app Is Nothing Then
        Set GetOutlookAppSafe = app
        Exit Function
    End If

    If Not IsOutlookDesktopRegistered() Then
        Set GetOutlookAppSafe = Nothing
        Exit Function
    End If

    If MsgBox( _
        "Outlook is not running." & vbCrLf & vbCrLf & _
        "Click Yes to start Outlook (may take a few seconds)." & vbCrLf & _
        "Click No to cancel.", _
        vbYesNo + vbQuestion, "EMAIL") <> vbYes Then
        Set GetOutlookAppSafe = Nothing
        Exit Function
    End If

    On Error Resume Next
    Set app = CreateObject("Outlook.Application")
    If Err.Number <> 0 Or app Is Nothing Then
        Err.Clear
        Set GetOutlookAppSafe = Nothing
        Exit Function
    End If
    On Error GoTo 0

    Set GetOutlookAppSafe = app
End Function

Private Function IsOutlookDesktopRegistered() As Boolean
    Dim wsh As Object
    Dim progId As String

    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    progId = CStr(wsh.RegRead("HKLM\SOFTWARE\Classes\Outlook.Application\CLSID\"))
    IsOutlookDesktopRegistered = (Len(progId) > 0)
    On Error GoTo 0
End Function

Private Function FileExistsFast(ByVal filePath As String) As Boolean
    Dim fso As Object

    ' Skip UNC / mapped-drive probes that can hang Excel for a long time.
    If Left$(filePath, 2) = "\\" Then
        FileExistsFast = False
        Exit Function
    End If

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsFast = fso.FileExists(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsFast = False
    End If
    On Error GoTo 0
End Function

Public Function RangeToHTML(rng As Range, Optional TempFile As String = "") As String
    ' Build HTML that preserves Excel column widths / row heights / light styling
    ' so the Outlook body matches the workbook aspect ratio.
    Dim r As Long
    Dim c As Long
    Dim nRows As Long
    Dim nCols As Long
    Dim html As String
    Dim cellText As String
    Dim cell As Range
    Dim bg As Long
    Dim fg As Long
    Dim bgCss As String
    Dim fgCss As String
    Dim boldCss As String
    Dim alignCss As String
    Dim heightCss As String
    Dim skipCell As Boolean
    Dim colPx() As Long
    Dim totalPx As Long
    Dim rowPx As Long
    Dim mergeCols As Long
    Dim mergeRows As Long
    Dim style As String

    nRows = rng.Rows.Count
    nCols = rng.Columns.Count
    ReDim colPx(1 To nCols)

    totalPx = 0
    For c = 1 To nCols
        colPx(c) = ColWidthToPx(rng.Columns(c).ColumnWidth)
        totalPx = totalPx + colPx(c)
    Next c

    ' color-scheme:light stops Outlook dark mode from inverting black text onto
    ' white cells (which reads as a blank table).
    html = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
    html = html & "<meta name='color-scheme' content='light only'>"
    html = html & "<meta name='supported-color-schemes' content='light'>"
    html = html & "<style type='text/css'>"
    html = html & ":root{color-scheme:light only;}"
    html = html & "body{margin:0;padding:8px;background:#ffffff;color:#000000;}"
    html = html & "table.xlmail{border-collapse:collapse;table-layout:fixed;"
    html = html & "width:" & CStr(totalPx) & "px;max-width:100%;"
    html = html & "font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;background:#ffffff;}"
    html = html & "table.xlmail td{border:1px solid #000000;padding:2px 4px;vertical-align:middle;"
    html = html & "overflow:hidden;word-wrap:break-word;color:#000000;}"
    html = html & "</style></head><body>"
    html = html & "<table class='xlmail' cellspacing='0' cellpadding='0' width='" & CStr(totalPx) & _
           "' bgcolor='#ffffff' style='color:#000000;background:#ffffff;'>"

    ' Explicit column widths keep Excel proportions in Outlook.
    html = html & "<colgroup>"
    For c = 1 To nCols
        html = html & "<col style='width:" & CStr(colPx(c)) & "px;' width='" & CStr(colPx(c)) & "'>"
    Next c
    html = html & "</colgroup>"

    For r = 1 To nRows
        rowPx = RowHeightToPx(rng.Rows(r).rowHeight)
        html = html & "<tr style='height:" & CStr(rowPx) & "px;'>"

        For c = 1 To nCols
            Set cell = rng.Cells(r, c)
            skipCell = False

            If cell.MergeCells Then
                If cell.Address(False, False) <> cell.MergeArea.Cells(1, 1).Address(False, False) Then
                    skipCell = True
                End If
            End If

            If Not skipCell Then
                cellText = Trim$(CStr(cell.text))
                cellText = HtmlEncode(cellText)
                If Len(cellText) = 0 Then cellText = "&nbsp;"

                bgCss = "background-color:#ffffff;"
                fgCss = "color:#000000;"
                boldCss = ""
                alignCss = "text-align:left;"
                heightCss = "height:" & CStr(rowPx) & "px;"

                On Error Resume Next
                If cell.DisplayFormat.Interior.Pattern <> xlNone Then
                    bg = cell.DisplayFormat.Interior.Color
                    bgCss = "background-color:#" & NormalizeBgColor(bg) & ";"
                End If
                Err.Clear

                fg = cell.DisplayFormat.Font.Color
                If Err.Number = 0 Then
                    fgCss = "color:#" & NormalizeFgColor(fg, bgCss) & ";"
                End If
                Err.Clear

                If cell.Font.bold Then boldCss = "font-weight:bold;"
                Err.Clear

                Select Case cell.HorizontalAlignment
                    Case xlCenter
                        alignCss = "text-align:center;"
                    Case xlRight
                        alignCss = "text-align:right;"
                    Case xlGeneral
                        If IsNumeric(cell.Value2) Then
                            alignCss = "text-align:right;"
                        Else
                            alignCss = "text-align:left;"
                        End If
                    Case Else
                        alignCss = "text-align:left;"
                End Select
                Err.Clear
                On Error GoTo 0

                ' Always emit black text for the mail body. Theme / DisplayFormat
                ' colours are kept only for the fill; Outlook dark mode otherwise
                ' inverts light fills and drops the glyphs.
                fgCss = "color:#000000;"
                style = bgCss & fgCss & boldCss & alignCss & heightCss & "border:1px solid #000000;"

                If cell.MergeCells Then
                    mergeCols = cell.MergeArea.Columns.Count
                    mergeRows = cell.MergeArea.Rows.Count
                    html = html & "<td colspan='" & mergeCols & "' rowspan='" & mergeRows & _
                           "' bgcolor='#" & mid$(bgCss, InStr(bgCss, "#") + 1, 6) & "' style='" & style & "'>" & _
                           "<font color='#000000'>" & cellText & "</font></td>"
                Else
                    html = html & "<td bgcolor='#" & mid$(bgCss, InStr(bgCss, "#") + 1, 6) & _
                           "' style='" & style & "'><font color='#000000'>" & cellText & "</font></td>"
                End If
            End If
        Next c
        html = html & "</tr>"
    Next r

    html = html & "</table></body></html>"
    RangeToHTML = html
End Function

Private Function ColWidthToPx(ByVal colWidth As Double) As Long
    ' Approx Excel column-width units -> CSS pixels (Calibri 11).
    If colWidth <= 0 Then
        ColWidthToPx = 20
    Else
        ColWidthToPx = CLng(colWidth * 7# + 5#)
        If ColWidthToPx < 20 Then ColWidthToPx = 20
    End If
End Function

Private Function RowHeightToPx(ByVal rowHeight As Double) As Long
    ' Excel row height is in points; convert to pixels (96 DPI).
    If rowHeight <= 0 Then rowHeight = 15
    RowHeightToPx = CLng(rowHeight * 96# / 72#)
    If RowHeightToPx < 16 Then RowHeightToPx = 16
End Function

Private Function NormalizeBgColor(ByVal oleColor As Long) As String
    Dim r As Long, g As Long, b As Long, lum As Long
    r = oleColor And &HFF&
    g = (oleColor \ &H100&) And &HFF&
    b = (oleColor \ &H10000) And &HFF&
    lum = (r + g + b) \ 3

    ' Theme/misread dark fills collapse the email; map to Excel-like light gray/white.
    If lum < 90 Then
        NormalizeBgColor = "F2F2F2"
    ElseIf r > 245 And g > 245 And b > 245 Then
        NormalizeBgColor = "FFFFFF"
    Else
        NormalizeBgColor = right$("0" & Hex$(r), 2) & right$("0" & Hex$(g), 2) & right$("0" & Hex$(b), 2)
    End If
End Function

Private Function NormalizeFgColor(ByVal oleColor As Long, ByVal bgCss As String) As String
    Dim r As Long, g As Long, b As Long, lum As Long
    r = oleColor And &HFF&
    g = (oleColor \ &H100&) And &HFF&
    b = (oleColor \ &H10000) And &HFF&
    lum = (r + g + b) \ 3

    ' White/near-white text on light cells is unreadable in Outlook.
    If lum > 200 Then
        NormalizeFgColor = "000000"
    ElseIf lum < 30 Then
        NormalizeFgColor = "000000"
    Else
        NormalizeFgColor = right$("0" & Hex$(r), 2) & right$("0" & Hex$(g), 2) & right$("0" & Hex$(b), 2)
    End If
End Function

Private Function HtmlEncode(ByVal s As String) As String
    Dim out As String
    out = s
    out = Replace(out, "&", "&amp;")
    out = Replace(out, "<", "&lt;")
    out = Replace(out, ">", "&gt;")
    out = Replace(out, """", "&quot;")
    out = Replace(out, vbCrLf, "<br>")
    out = Replace(out, vbLf, "<br>")
    out = Replace(out, vbCr, "<br>")
    HtmlEncode = out
End Function




