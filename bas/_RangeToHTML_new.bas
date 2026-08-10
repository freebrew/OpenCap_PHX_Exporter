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

    html = "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
    html = html & "<style type='text/css'>"
    html = html & "body{margin:0;padding:8px;background:#ffffff;}"
    html = html & "table.xlmail{border-collapse:collapse;table-layout:fixed;"
    html = html & "width:" & CStr(totalPx) & "px;max-width:100%;"
    html = html & "font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;background:#ffffff;}"
    html = html & "table.xlmail td{border:1px solid #000000;padding:2px 4px;vertical-align:middle;"
    html = html & "overflow:hidden;word-wrap:break-word;}"
    html = html & "</style></head><body>"
    html = html & "<table class='xlmail' cellspacing='0' cellpadding='0' width='" & CStr(totalPx) & "'>"

    ' Explicit column widths keep Excel proportions in Outlook.
    html = html & "<colgroup>"
    For c = 1 To nCols
        html = html & "<col style='width:" & CStr(colPx(c)) & "px;' width='" & CStr(colPx(c)) & "'>"
    Next c
    html = html & "</colgroup>"

    For r = 1 To nRows
        rowPx = RowHeightToPx(rng.Rows(r).RowHeight)
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
                cellText = Trim$(CStr(cell.Text))
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

                If cell.Font.Bold Then boldCss = "font-weight:bold;"
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

                style = bgCss & fgCss & boldCss & alignCss & heightCss & "border:1px solid #000000;"

                If cell.MergeCells Then
                    mergeCols = cell.MergeArea.Columns.Count
                    mergeRows = cell.MergeArea.Rows.Count
                    html = html & "<td colspan='" & mergeCols & "' rowspan='" & mergeRows & _
                           "' style='" & style & "'>" & cellText & "</td>"
                Else
                    html = html & "<td style='" & style & "'>" & cellText & "</td>"
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
        NormalizeBgColor = Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
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
        NormalizeFgColor = Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
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
