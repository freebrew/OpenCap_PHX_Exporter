Attribute VB_Name = "Module11"
Option Explicit

Public Sub RTP_printW_Email()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim rng As Range
    Dim HTMLBodyText As String
    Dim c As Range
    Dim filePath As String
    Dim attachedCount As Long
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation

    On Error GoTo CleanFail

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set rng = Sheets("Data").Range("R5:V61")
    HTMLBodyText = RangeToHTML(rng)

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
        .To = "paraakita520@gmail.com"
        .CC = "paramount.d-c@paramountres.com;logan.mercer@paramountres.com"
        .BCC = ""
        .Subject = "Daily Drilling Summary - AKITA520 - sinclairdrill"
        .HTMLBody = HTMLBodyText

        attachedCount = 0
        For Each c In Sheets("Data").Range("Y11:Y15")
            filePath = Trim$(CStr(c.Value & ""))
            If Len(filePath) > 0 Then
                If FileExistsFast(filePath) Then
                    .Attachments.Add filePath
                    attachedCount = attachedCount + 1
                End If
            End If
        Next c

        .Display
    End With

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Set OutMail = Nothing
    Set OutApp = Nothing
    Exit Sub

CleanFail:
    Dim msg As String
    msg = Err.Description
    Resume CleanExitMessage

CleanExitMessage:
    On Error Resume Next
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
    Set OutMail = Nothing
    Set OutApp = Nothing
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
    ' Build HTML directly from visible cell text.
    ' Avoids Workbooks.Add / PublishObjects, which commonly freeze Excel.
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
    Dim skipCell As Boolean

    nRows = rng.Rows.Count
    nCols = rng.Columns.Count

    html = "<html><body>"
    html = html & "<table border='1' cellspacing='0' cellpadding='3' style='border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>"

    For r = 1 To nRows
        html = html & "<tr>"
        For c = 1 To nCols
            Set cell = rng.Cells(r, c)
            skipCell = False

            ' Skip cells hidden by merges (keep only the top-left merge cell).
            If cell.MergeCells Then
                If cell.Address(False, False) <> cell.MergeArea.Cells(1, 1).Address(False, False) Then
                    skipCell = True
                End If
            End If

            If Not skipCell Then
                cellText = Trim$(cell.Text)
                cellText = HtmlEncode(cellText)
                If Len(cellText) = 0 Then cellText = "&nbsp;"

                bgCss = ""
                fgCss = ""
                boldCss = ""
                alignCss = ""

                On Error Resume Next
                If cell.Interior.Pattern <> xlNone Then
                    bg = cell.Interior.Color
                    If bg <> 16777215 Then
                        bgCss = "background-color:#" & HtmlColor(bg) & ";"
                    End If
                End If
                Err.Clear

                fg = cell.Font.Color
                If Err.Number = 0 Then
                    If fg <> 0 Then
                        fgCss = "color:#" & HtmlColor(fg) & ";"
                    End If
                End If
                Err.Clear

                If cell.Font.Bold Then boldCss = "font-weight:bold;"
                Err.Clear

                Select Case cell.HorizontalAlignment
                    Case xlCenter
                        alignCss = "text-align:center;"
                    Case xlRight
                        alignCss = "text-align:right;"
                    Case Else
                        alignCss = "text-align:left;"
                End Select
                Err.Clear
                On Error GoTo 0

                If cell.MergeCells Then
                    html = html & "<td colspan='" & cell.MergeArea.Columns.Count & _
                           "' rowspan='" & cell.MergeArea.Rows.Count & _
                           "' style='" & bgCss & fgCss & boldCss & alignCss & "'>" & cellText & "</td>"
                Else
                    html = html & "<td style='" & bgCss & fgCss & boldCss & alignCss & "'>" & cellText & "</td>"
                End If
            End If
        Next c
        html = html & "</tr>"
    Next r

    html = html & "</table></body></html>"
    RangeToHTML = html
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

Private Function HtmlColor(ByVal oleColor As Long) As String
    Dim r As Long, g As Long, b As Long
    r = oleColor And &HFF&
    g = (oleColor \ &H100&) And &HFF&
    b = (oleColor \ &H10000) And &HFF&
    HtmlColor = Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function
