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
    Dim picOk As Boolean
    Dim prevInteractive As Boolean
    Dim prevCursor As XlMousePointer
    Dim tmpDir As String
    Dim wasProt As Boolean

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

    ' Daily plot PNG: corridor band only (wellbore room, vertical section and
    ' the survey / position / requirement cards). The ops tables are already in
    ' the B2:F55 HTML table, so they are not repeated in the picture. The PNG is
    ' attached as a visible file AND inlined (CID) as the last row of that
    ' table, so table + picture read as one block. Outlook hides CID-referenced
    ' images from the paperclip list, so the visible attach must NOT get a
    ' Content-ID. No daily PDF.
    tmpDir = Environ$("TEMP")
    reportPng = ""
    On Error Resume Next
    reportPng = MDL_CorridorImage.RenderCorridorPng(tmpDir & "\daily_report.png", "corridor")
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
        .To = MailToFromSheet(wsData)
        .CC = MailCcFromSheet(wsData)
        .BCC = ""
        .Subject = DailyMailSubject()

        attachedCount = 0
        For Each c In wsData.Range("I29:I33")
            filePath = StoredAttachPath(c)
            If Len(filePath) > 0 Then
                If FileExistsFast(filePath) Then
                    .Attachments.Add filePath
                    attachedCount = attachedCount + 1
                End If
            End If
        Next c

        picOk = False
        If Len(reportPng) > 0 Then picOk = FileExistsFast(reportPng)
        If picOk Then
            .Attachments.Add reportPng
            attachedCount = attachedCount + 1
            AttachDailyPngInline OutMail, reportPng, RangePxWidth(wsData.Range("B2:F55")) - 2
        End If
        ' Picture is the last row of the table (one unit), not a block below it.
        .HTMLBody = RangeToHTML(wsData.Range("B2:F55"), "", picOk)
    End With

    ' I29:I31 are one-shot. Keep Attachement 4 (I32) and 5 (I33) for every email.
    ' Clear each MergeArea — ClearContents on I alone fails when I:J is merged.
    ' Data is usually protected in the field; StoreAttachPath already unprotects
    ' when a file is picked, but this send-path write did not.
    wasProt = SheetUnprotectForVba(wsData)
    For r = 29 To 31
        ClearAttachCell wsData.Range("I" & r)
    Next r
    SheetReprotectAfterVba wsData, wasProt
    wasProt = False

    ' Restore Excel UI before showing the mail so Outlook does not fight a locked
    ' Excel session for focus.
    Application.Cursor = prevCursor
    Application.Interactive = prevInteractive
    Application.ScreenUpdating = prevScreenUpdating

    OutMail.Display

CleanExit:
    On Error Resume Next
    If Not wsData Is Nothing Then SheetReprotectAfterVba wsData, wasProt
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
    If Not wsData Is Nothing Then SheetReprotectAfterVba wsData, wasProt
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

    ' Outlook is installed but not running: start it automatically (no prompt).
    Application.StatusBar = "EMAIL: starting Outlook..."
    On Error Resume Next
    Set app = CreateObject("Outlook.Application")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    If app Is Nothing Then
        ' CreateObject can fail while Outlook is still loading its profile:
        ' launch it, then wait up to ~30 s for it to register.
        Dim wsh As Object, t As Single
        On Error Resume Next
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run "outlook.exe", 6, False
        If Err.Number <> 0 Then Err.Clear
        t = Timer
        Do While app Is Nothing And Abs(Timer - t) < 30!
            DoEvents
            Set app = GetObject(, "Outlook.Application")
            If Err.Number <> 0 Then Err.Clear
        Loop
        On Error GoTo 0
    End If
    Application.StatusBar = False

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

' Second copy of the PNG, Content-ID only. Outlook hides this one from the
' attachment well because the HTML body references cid:daily_report.png.
' The visible paperclip file is the plain .Attachments.Add above.
Private Sub AttachDailyPngInline(ByVal mail As Object, ByVal pngPath As String, _
                                 ByVal widthPx As Long)
    Dim att As Object
    Dim inlinePath As String
    Dim fso As Object

    inlinePath = Environ$("TEMP") & "\daily_report_inline.png"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(inlinePath) Then fso.DeleteFile inlinePath, True
    fso.CopyFile pngPath, inlinePath, True
    On Error GoTo 0
    If Not FileExistsFast(inlinePath) Then
        inlinePath = pngPath
    Else
        ' The inline copy is resampled (anti-aliased) to the width it is shown
        ' at, so Outlook never rescales it. The attached file stays full size.
        On Error Resume Next
        MDL_CorridorImage.PngDownscale inlinePath, widthPx
        On Error GoTo 0
    End If

    Set att = mail.Attachments.Add(inlinePath)
    On Error Resume Next
    att.PropertyAccessor.SetProperty _
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "daily_report.png"
    att.PropertyAccessor.SetProperty _
        "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True   ' PR_ATTACHMENT_HIDDEN
    On Error GoTo 0
End Sub

' Table width in CSS pixels (same sum RangeToHTML uses).
Private Function RangePxWidth(ByVal rng As Range) As Long
    Dim c As Long
    For c = 1 To rng.Columns.Count
        RangePxWidth = RangePxWidth + ColWidthToPx(rng.Columns(c).ColumnWidth)
    Next c
End Function

' The picture sized to the table's inner width so it fills the last row.
Private Function DailyPngHtml(ByVal widthPx As Long) As String
    If widthPx < 100 Then widthPx = 100
    DailyPngHtml = "<img src='cid:daily_report.png' width='" & CStr(widthPx) & _
        "' alt='Daily wellbore plot' " & _
        "style='display:block;width:" & CStr(widthPx) & "px;max-width:100%;height:auto;border:0;'>"
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

Public Function RangeToHTML(rng As Range, Optional TempFile As String = "", _
                            Optional ByVal withPicRow As Boolean = False) As String
    ' Build HTML that preserves Excel column widths / row heights / light styling
    ' so the Outlook body matches the workbook aspect ratio.
    Dim r As Long
    Dim c As Long
    Dim nRows As Long
    Dim nCols As Long
    Dim html As String
    Dim CellText As String
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

    nRows = rng.rows.Count
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
        rowPx = RowHeightToPx(rng.rows(r).rowHeight)
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
                CellText = Trim$(CStr(cell.text))
                CellText = HtmlEncode(CellText)
                If Len(CellText) = 0 Then CellText = "&nbsp;"

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
                    mergeRows = cell.MergeArea.rows.Count
                    html = html & "<td colspan='" & mergeCols & "' rowspan='" & mergeRows & _
                           "' bgcolor='#" & mid$(bgCss, InStr(bgCss, "#") + 1, 6) & "' style='" & style & "'>" & _
                           "<font color='#000000'>" & CellText & "</font></td>"
                Else
                    html = html & "<td bgcolor='#" & mid$(bgCss, InStr(bgCss, "#") + 1, 6) & _
                           "' style='" & style & "'><font color='#000000'>" & CellText & "</font></td>"
                End If
            End If
        Next c
        html = html & "</tr>"
    Next r

    ' Daily picture = last row, full table width, same black frame as the cells.
    If withPicRow Then
        html = html & "<tr><td colspan='" & CStr(nCols) & "' bgcolor='#262626' " & _
               "style='padding:0;background-color:#262626;border:1px solid #000000;'>" & _
               DailyPngHtml(totalPx - 2) & "</td></tr>"
    End If

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

Private Function DailyMailSubject() As String
    Dim code As String, nm As String
    On Error Resume Next
    code = Trim$(MDL_Setup.OcJobField("Job Code", "Job ID"))
    nm = Trim$(MDL_Setup.OcJobField("Job Name"))
    On Error GoTo 0
    If code = "" Then code = "Job"
    If nm = "" Then
        DailyMailSubject = "Daily Drilling Summary - " & code
    Else
        DailyMailSubject = "Daily Drilling Summary - " & code & " - " & nm
    End If
End Function

Private Function MailToFromSheet(ByVal ws As Worksheet) As String
    Dim r As Long, role As String, addr As String
    MailToFromSheet = ""
    For r = 29 To 33
        role = UCase$(Trim$(CStr(ws.Cells(r, "K").Value2 & "")))
        addr = Trim$(CStr(ws.Cells(r, "L").Value2 & ""))
        If addr = "" Then GoTo NextTo
        If role = "TO" Or role = "" Then
            MailToFromSheet = addr
            Exit Function
        End If
NextTo:
    Next r
    MailToFromSheet = Trim$(CStr(ws.Cells(29, "L").Value2 & ""))
End Function

Private Function MailCcFromSheet(ByVal ws As Worksheet) As String
    Dim r As Long, role As String, addr As String
    Dim acc As String
    acc = ""
    For r = 29 To 33
        role = UCase$(Trim$(CStr(ws.Cells(r, "K").Value2 & "")))
        addr = Trim$(CStr(ws.Cells(r, "L").Value2 & ""))
        If addr = "" Then GoTo NextCc
        If r = 29 And (role = "TO" Or role = "") Then GoTo NextCc
        If role = "TO" Then GoTo NextCc
        If acc <> "" Then acc = acc & "; "
        acc = acc & addr
NextCc:
    Next r
    MailCcFromSheet = acc
End Function

' The visible filename decides what is attached. The hidden note / hyperlink
' only supply the full path for that same filename, so a blank cell (or a
' stale note left behind after the text was cleared) attaches nothing.
Public Function StoredAttachPath(ByVal cell As Range) As String
    Dim shown As String
    Dim s As String
    StoredAttachPath = ""
    If cell Is Nothing Then Exit Function
    Set cell = cell.MergeArea.Cells(1, 1)
    shown = Trim$(CStr(cell.Value2 & ""))
    If Len(shown) = 0 Then Exit Function
    If InStr(shown, "\") > 0 Or InStr(shown, "/") > 0 Then
        StoredAttachPath = shown
        Exit Function
    End If

    On Error Resume Next
    If Not cell.Comment Is Nothing Then s = Trim$(cell.Comment.text)
    On Error GoTo 0
    If PathLeafIs(s, shown) Then
        StoredAttachPath = s
        Exit Function
    End If

    s = ""
    On Error Resume Next
    If cell.Hyperlinks.Count > 0 Then s = Trim$(cell.Hyperlinks(1).Address)
    On Error GoTo 0
    If PathLeafIs(s, shown) Then StoredAttachPath = s
End Function

Private Function PathLeafIs(ByVal fullPath As String, ByVal leaf As String) As Boolean
    Dim p As Long
    PathLeafIs = False
    If Len(fullPath) = 0 Then Exit Function
    p = InStrRev(Replace(fullPath, "/", "\"), "\")
    PathLeafIs = (StrComp(mid$(fullPath, p + 1), leaf, vbTextCompare) = 0)
End Function

Private Sub ClearAttachCell(ByVal cell As Range)
    On Error Resume Next
    With cell.MergeArea
        .ClearContents
        .ClearComments
        .ClearHyperlinks
    End With
    On Error GoTo 0
End Sub

Public Sub StoreAttachPath(ByVal cell As Range, ByVal fullPath As String)
    Dim leaf As String
    Dim ws As Worksheet
    Dim dest As Range
    Dim wasProt As Boolean

    If cell Is Nothing Then Exit Sub
    Set dest = cell.MergeArea.Cells(1, 1)
    Set ws = dest.Worksheet
    wasProt = SheetUnprotectForVba(ws)

    fullPath = Trim$(fullPath)
    If InStrRev(fullPath, "\") > 0 Then
        leaf = mid$(fullPath, InStrRev(fullPath, "\") + 1)
    ElseIf InStrRev(fullPath, "/") > 0 Then
        leaf = mid$(fullPath, InStrRev(fullPath, "/") + 1)
    Else
        leaf = fullPath
    End If

    On Error Resume Next
    dest.ClearComments
    dest.Hyperlinks.Delete
    On Error GoTo 0

    dest.Value = leaf
    If Len(fullPath) > 0 Then
        On Error Resume Next
        dest.AddComment fullPath
        If Not dest.Comment Is Nothing Then
            dest.Comment.Visible = False
            dest.Comment.Shape.TextFrame.AutoSize = True
        End If
        ' '#' in a filename is a hyperlink fragment and raises 1004 (e.g. BHA # 2.pdf).
        If InStr(1, fullPath, "#", vbBinaryCompare) = 0 Then
            dest.Hyperlinks.Add anchor:=dest, Address:=fullPath, TextToDisplay:=leaf
        End If
        On Error GoTo 0
    End If

    If IsKeepAttachRow(dest.row) Then
        dest.MergeArea.Interior.Color = KeepAttachFill()
    Else
        dest.MergeArea.Interior.Color = RGB(255, 255, 255)
    End If
    dest.Font.Color = RGB(0, 0, 0)
    dest.Font.size = 8
    dest.Font.name = "Calibri"
    dest.HorizontalAlignment = xlLeft

    SheetReprotectAfterVba ws, wasProt
End Sub

Private Function IsKeepAttachRow(ByVal r As Long) As Boolean
    IsKeepAttachRow = (r = 32 Or r = 33)
End Function

Private Function KeepAttachFill() As Long
    ' Amber pin: last two attachment slots survive EMAIL.
    KeepAttachFill = RGB(255, 242, 204)
End Function

Private Sub ApplyKeepAttachStyle(ByVal ws As Worksheet)
    ws.Range("I32").MergeArea.Interior.Color = KeepAttachFill()
    ws.Range("I33").MergeArea.Interior.Color = KeepAttachFill()
    ws.Range("H28:M33").BorderAround LineStyle:=xlContinuous, Weight:=xlThick, Color:=RGB(0, 0, 0)
End Sub

' Format only — does not change attachment paths or mail addresses.
Public Sub StylePersistentAttachCells()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim prevSU As Boolean

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets("Data")
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    wasProt = SheetUnprotectForVba(ws)
    ApplyKeepAttachStyle ws
    SheetReprotectAfterVba ws, wasProt
    Application.ScreenUpdating = prevSU
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.ScreenUpdating = True
End Sub

' Unmerge I28:M33, show filenames, hold full paths in comments, add To/CC emails.
' Print-friendly: light gray headers, white body, no dark fills.
Public Sub ApplyDataMailPanel()
    Dim ws As Worksheet
    Dim r As Long
    Dim saved(29 To 33) As String
    Dim hdr As Long, body As Long, toFill As Long
    Dim wasProt As Boolean

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets("Data")
    wasProt = False
    On Error Resume Next
    wasProt = ws.ProtectContents
    ws.Unprotect
    On Error GoTo Fail

    For r = 29 To 33
        saved(r) = StoredAttachPath(ws.Cells(r, "I"))
        If saved(r) = "" Then saved(r) = Trim$(CStr(ws.Cells(r, "I").MergeArea.Cells(1, 1).Value2 & ""))
    Next r

    On Error Resume Next
    ws.Range("I28:M33").UnMerge
    On Error GoTo Fail

    hdr = RGB(242, 242, 242)
    body = RGB(255, 255, 255)
    toFill = RGB(232, 244, 248)

    ws.Range("H28").Value = "File"
    ws.Range("I28").Value = "Attachment"
    ws.Range("K28").Value = "Send"
    ws.Range("L28").Value = "Email"
    StyleMailHeader ws.Range("H28:I28"), hdr
    StyleMailHeader ws.Range("K28:M28"), hdr

    For r = 29 To 33
        ws.Cells(r, "H").Value = "Attach " & CStr(r - 28)
        ws.Cells(r, "H").Font.size = 8
        ws.Cells(r, "H").Font.bold = True
        ws.Cells(r, "H").Interior.Color = hdr
        ws.Cells(r, "H").Font.Color = RGB(0, 0, 0)

        If r = 29 Then
            ws.Cells(r, "K").Value = "To"
            ws.Cells(r, "K").Interior.Color = toFill
            ws.Cells(r, "L").Interior.Color = toFill
            ws.Cells(r, "M").Interior.Color = toFill
        Else
            ws.Cells(r, "K").Value = "CC"
            ws.Cells(r, "K").Interior.Color = body
            ws.Cells(r, "L").Interior.Color = body
            ws.Cells(r, "M").Interior.Color = body
        End If
        ws.Cells(r, "K").Font.size = 8
        ws.Cells(r, "K").Font.bold = True
        ws.Cells(r, "K").HorizontalAlignment = xlCenter
        ws.Cells(r, "L").Font.size = 8
        ws.Cells(r, "L").HorizontalAlignment = xlLeft
        StoreAttachPath ws.Cells(r, "I"), saved(r)
        On Error Resume Next
        ws.Range(ws.Cells(r, "I"), ws.Cells(r, "J")).Merge
        ws.Range(ws.Cells(r, "L"), ws.Cells(r, "M")).Merge
        On Error GoTo Fail
    Next r
    On Error Resume Next
    ws.Range("I28:J28").Merge
    ws.Range("L28:M28").Merge
    On Error GoTo Fail

    SeedMailAddressesIfEmpty ws
    ApplyKeepAttachStyle ws

    If wasProt Then
        On Error Resume Next
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If
    Exit Sub
Fail:
    On Error Resume Next
    If wasProt Then ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
End Sub

Private Sub StyleMailHeader(ByVal rng As Range, ByVal fill As Long)
    With rng
        .Interior.Color = fill
        .Font.bold = True
        .Font.size = 8
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub SeedMailAddressesIfEmpty(ByVal ws As Worksheet)
    Dim raw As String, parts() As String, i As Long, em As String
    Dim dest As Long
    If Trim$(CStr(ws.Cells(29, "L").Value2 & "")) <> "" Then Exit Sub
    On Error Resume Next
    ws.Cells(29, "L").Value = Trim$(MDL_Setup.OcJobField("CompanyManEmail"))
    raw = MDL_Setup.OcJobField("ClientEmail")
    On Error GoTo 0
    If Len(raw) = 0 Then Exit Sub
    raw = Replace(Replace(raw, ",", ";"), " ", "")
    parts = Split(raw, ";")
    dest = 30
    For i = LBound(parts) To UBound(parts)
        If dest > 33 Then Exit For
        em = LCase$(Trim$(parts(i)))
        If em = "" Then GoTo NextSeed
        If InStr(em, "@") < 2 Then GoTo NextSeed
        If InStr(em, "phxtech.com") > 0 Then GoTo NextSeed
        ws.Cells(dest, "L").Value = Trim$(parts(i))
        dest = dest + 1
NextSeed:
    Next i
End Sub





















