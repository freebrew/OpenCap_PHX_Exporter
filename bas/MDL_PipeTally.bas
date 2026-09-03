Attribute VB_Name = "MDL_PipeTally"
Option Explicit

Private Const PT_SHEET As String = "Pipe Tally Calculator"
Private Const SS_SHEET As String = "Slidesheet"
Private Const BTN_NAME As String = "btnAppendTallyBitDepths"
Private Const BTN_CAPTION As String = "Add Bit Depths to Slidesheet"
Private Const BTN_ANCHOR As String = "B16"

Private Const PT_FIRST As Long = 7
Private Const PT_LAST As Long = 315
Private Const MIN_JOINT_M As Double = 0.1
Private Const MAX_JOINT_M As Double = 40#

' ---------------------------------------------------------------------------------
' Sheet resolution
'
' Every routine here reads from the "Pipe Tally Calculator" worksheet.  Calling
' Worksheets(PT_SHEET) directly raises a bare error 9 "Subscript out of range"
' that names nothing, which is what the form used to report.  Resolve through
' TallySheet() instead so the caller can say WHICH object is missing.
' ---------------------------------------------------------------------------------

Private Function TallySheet(Optional ByRef whyNot As String) As Worksheet
    whyNot = ""
    On Error Resume Next
    Set TallySheet = ThisWorkbook.Worksheets(PT_SHEET)
    On Error GoTo 0
    If TallySheet Is Nothing Then
        whyNot = "Worksheet """ & PT_SHEET & """ does not exist in " & ThisWorkbook.name & "."
    End If
End Function

Public Function PipeTallySheetExists() As Boolean
    PipeTallySheetExists = Not (TallySheet() Is Nothing)
End Function

' Text explaining what the missing sheet is expected to contain.
Private Function TallyLayoutHelp() As String
    TallyLayoutHelp = _
        "The calculator reads and writes these cells on that sheet:" & vbCrLf & _
        "    B5        start depth (last bit depth on the Slidesheet)" & vbCrLf & _
        "    C11       stands per joint" & vbCrLf & _
        "    D11       split length" & vbCrLf & _
        "    F" & PT_FIRST & ":F" & PT_LAST & "   pipe joint lengths (input)" & vbCrLf & _
        "    H" & PT_FIRST & ":H" & PT_LAST & "   calculated bit depths (output)" & vbCrLf & vbCrLf & _
        "Restore that sheet from a copy in the backups folder, then try again."
End Function

' ---------------------------------------------------------------------------------
' Form entry (ribbon callback lives in MDL_Ribbon)
' ---------------------------------------------------------------------------------

Public Sub ShowPipeTallyForm()
    Dim why As String
    If TallySheet(why) Is Nothing Then
        MsgBox "Pipe Tally Calculator cannot open." & vbCrLf & vbCrLf & _
               why & vbCrLf & vbCrLf & TallyLayoutHelp(), _
               vbExclamation, "Pipe Tally Calculator"
        Exit Sub
    End If

    On Error GoTo Fail
    frmPipeTally.Show vbModeless
    Exit Sub
Fail:
    Dim msg As String
    msg = "Could not open Pipe Tally form." & vbCrLf & vbCrLf & _
          "Error " & Err.Number & ": " & Err.Description
    If Len(Err.Source) > 0 Then msg = msg & vbCrLf & "Source: " & Err.Source
    If Err.Number = 9 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Error 9 means a named sheet or array item was not found." & vbCrLf & _
              TallyLayoutHelp()
    End If
    MsgBox msg, vbCritical, "Pipe Tally Calculator"
End Sub

' Dismiss the modeless form (also useful if it is left stranded on screen).
Public Sub ClosePipeTallyFormIfOpen()
    On Error Resume Next
    Unload frmPipeTally
    On Error GoTo 0
End Sub

' Form-control button on Slidesheet, in merged Z4:Z6 (under Clear Ranges in Z1:Z3).
Public Sub EnsurePipeTallyFormButton()
    Dim ss As Worksheet
    Dim rng As Range
    Dim btn As Button
    Dim shp As Shape

    Const PT_BTN As String = "btnPipeTallyForm"
    Const PT_CAPTION As String = "Pipe Tally Calculator"
    Const PT_ANCHOR As String = "Z4:Z6"

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    Set rng = ss.Range(PT_ANCHOR)
    If rng.Cells.Count = 1 Then Set rng = rng.MergeArea

    On Error Resume Next
    ss.Buttons(PT_BTN).Delete
    ss.Shapes(PT_BTN).Delete
    On Error GoTo 0

    Set btn = ss.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    btn.name = PT_BTN
    btn.caption = PT_CAPTION
    btn.OnAction = "'" & ThisWorkbook.name & "'!ShowPipeTallyForm"

    Set shp = ss.Shapes(PT_BTN)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.Size = 10
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------
' Sheet helpers used by form + existing append button
' ---------------------------------------------------------------------------------

Public Function PipeTallyStartDepth() As Variant
    Dim pt As Worksheet
    On Error GoTo Fail
    Set pt = TallySheet()
    If pt Is Nothing Then PipeTallyStartDepth = "": Exit Function
    PipeTallyStartDepth = pt.Range("B5").Value2
    Exit Function
Fail:
    PipeTallyStartDepth = ""
End Function

Public Sub PipeTallyGetSettings(ByRef stands As Variant, ByRef splitLen As Variant)
    Dim pt As Worksheet
    stands = Empty: splitLen = Empty
    Set pt = TallySheet()
    If pt Is Nothing Then Exit Sub
    stands = pt.Range("C11").Value2
    splitLen = pt.Range("D11").Value2
End Sub

Public Sub PipeTallySetSettings(ByVal stands As Variant, ByVal splitLen As Variant)
    Dim pt As Worksheet
    Set pt = TallySheet()
    If pt Is Nothing Then Exit Sub
    pt.Range("C11").Value2 = stands
    If Len(Trim$(CStr(splitLen & ""))) = 0 Then
        pt.Range("D11").ClearContents
    Else
        pt.Range("D11").Value2 = splitLen
    End If
End Sub

Public Function PipeTallyReadLengths() As Variant
    Dim pt As Worksheet
    Dim r As Long, n As Long, v As Variant
    Dim tmp() As Double
    Set pt = TallySheet()
    If pt Is Nothing Then PipeTallyReadLengths = Empty: Exit Function
    ReDim tmp(1 To (PT_LAST - PT_FIRST + 1))
    n = 0
    For r = PT_FIRST To PT_LAST
        v = pt.Cells(r, "F").Value2
        If IsNumeric(v) Then
            If CDbl(v) > 0# Then
                n = n + 1
                tmp(n) = CDbl(v)
            End If
        End If
    Next r
    If n = 0 Then
        PipeTallyReadLengths = Empty
        Exit Function
    End If
    ReDim Preserve tmp(1 To n)
    PipeTallyReadLengths = tmp
End Function

Public Sub PipeTallyWriteLengths(ByRef lengths As Variant)
    Dim pt As Worksheet
    Dim r As Long, i As Long, n As Long
    Set pt = TallySheet()
    If pt Is Nothing Then Exit Sub
    pt.Range("F" & PT_FIRST & ":F" & PT_LAST).ClearContents
    If IsEmpty(lengths) Then Exit Sub
    n = UBound(lengths)
    i = 1
    For r = PT_FIRST To PT_LAST
        If i > n Then Exit For
        pt.Cells(r, "F").Value2 = CDbl(lengths(i))
        i = i + 1
    Next r
End Sub

Public Function PipeTallyReadBitDepths() As Variant
    Dim pt As Worksheet
    Dim r As Long, n As Long, v As Variant
    Dim tmp() As Double
    Set pt = TallySheet()
    If pt Is Nothing Then PipeTallyReadBitDepths = Empty: Exit Function
    Application.Calculate
    ReDim tmp(1 To (PT_LAST - PT_FIRST + 1))
    n = 0
    For r = PT_FIRST To PT_LAST
        v = pt.Cells(r, "H").Value2
        If VarType(v) = vbDouble Then
            n = n + 1
            tmp(n) = CDbl(v)
        End If
    Next r
    If n = 0 Then
        PipeTallyReadBitDepths = Empty
        Exit Function
    End If
    ReDim Preserve tmp(1 To n)
    PipeTallyReadBitDepths = tmp
End Function

Public Sub ClearPipeTallyInputs()
    Dim pt As Worksheet
    Set pt = TallySheet()
    If pt Is Nothing Then Exit Sub
    pt.Range("F" & PT_FIRST & ":F" & PT_LAST).ClearContents
End Sub

' Parse clipboard / pasted text into a single column of drill-pipe joint lengths.
' Rejects multi-column data and values that look like depths (not joint lengths).
Public Function ParsePipeLengthColumn(ByVal text As String, ByRef lengths As Variant, _
                                      ByRef errMsg As String) As Boolean
    Dim lines() As String
    Dim i As Long, n As Long, nNonEmpty As Long
    Dim line As String, parts() As String
    Dim v As Double
    Dim tmp() As Double
    Dim tok As String, j As Long, nTok As Long

    ParsePipeLengthColumn = False
    errMsg = ""
    lengths = Empty

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    text = Trim$(text)
    If Len(text) = 0 Then
        errMsg = "Clipboard is empty."
        Exit Function
    End If

    lines = Split(text, vbLf)
    ReDim tmp(1 To UBound(lines) - LBound(lines) + 1)
    n = 0
    nNonEmpty = 0

    For i = LBound(lines) To UBound(lines)
        line = Trim$(lines(i))
        If Len(line) = 0 Then GoTo NextLine

        nNonEmpty = nNonEmpty + 1

        ' Multi-column via tabs
        If InStr(line, vbTab) > 0 Then
            parts = Split(line, vbTab)
            nTok = 0
            For j = LBound(parts) To UBound(parts)
                If Len(Trim$(parts(j))) > 0 Then nTok = nTok + 1
            Next j
            If nTok > 1 Then
                errMsg = "Paste must be a single column of pipe lengths (found multiple columns)."
                Exit Function
            End If
            line = Trim$(parts(LBound(parts)))
            For j = LBound(parts) To UBound(parts)
                If Len(Trim$(parts(j))) > 0 Then line = Trim$(parts(j)): Exit For
            Next j
        End If

        ' Multi-column via commas / semicolons
        If InStr(line, ",") > 0 Or InStr(line, ";") > 0 Then
            errMsg = "Paste must be a single column of pipe lengths (found comma/semicolon separators)."
            Exit Function
        End If

        ' Multiple space-separated tokens on one line
        tok = Application.WorksheetFunction.Trim(line)
        parts = Split(tok, " ")
        nTok = 0
        For j = LBound(parts) To UBound(parts)
            If Len(Trim$(parts(j))) > 0 Then nTok = nTok + 1
        Next j
        If nTok > 1 Then
            errMsg = "Paste must be a single column of pipe lengths (found multiple values on one line)."
            Exit Function
        End If
        line = tok

        If Not IsNumeric(line) Then
            errMsg = "Non-numeric value: """ & line & """."
            Exit Function
        End If
        v = CDbl(line)
        If v < MIN_JOINT_M Or v > MAX_JOINT_M Then
            errMsg = "Value " & CStr(v) & " is outside the expected drill-pipe length range (" & _
                     CStr(MIN_JOINT_M) & "-" & CStr(MAX_JOINT_M) & " m). " & _
                     "Paste joint lengths, not bit depths."
            Exit Function
        End If

        n = n + 1
        tmp(n) = v
NextLine:
    Next i

    If n = 0 Then
        errMsg = "No pipe lengths found in clipboard."
        Exit Function
    End If

    ReDim Preserve tmp(1 To n)
    lengths = tmp
    ParsePipeLengthColumn = True
End Function

Public Function GetClipboardText() As String
    Dim d As Object
    On Error GoTo Fail
    Set d = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") ' MSForms.DataObject
    d.GetFromClipboard
    GetClipboardText = d.GetText(1)
    Exit Function
Fail:
    GetClipboardText = ""
End Function

' ---------------------------------------------------------------------------------
' Append non-blank H7:H315 tally depths into Slidesheet!D
' starting on the first empty row below the last bit-depth entry (matches B5).
' ---------------------------------------------------------------------------------
Public Sub AppendTallyBitDepths()
    Dim pt As Worksheet
    Dim ss As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim r As Long
    Dim i As Long
    Dim nAdded As Long
    Dim nVals As Long
    Dim v As Variant
    Dim startRow As Long
    Dim depths() As Double
    Dim truncated As Boolean

    On Error GoTo Fail

    Dim why As String
    Set pt = TallySheet(why)
    If pt Is Nothing Then
        MsgBox "Cannot add bit depths." & vbCrLf & vbCrLf & why & vbCrLf & vbCrLf & _
               TallyLayoutHelp(), vbExclamation, "Pipe Tally Calculator"
        Exit Sub
    End If
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)

    lastRow = LastBitDepthRow(ss)
    If lastRow < 12 Then
        MsgBox "No bit depth entries found on Slidesheet column D (from row 12).", vbExclamation
        Exit Sub
    End If

    ' H7:H315 is built on B5, which LOOKUPs the last bit depth on the Slidesheet.
    ' Writing one row at a time would move that start depth and re-base every
    ' remaining tally depth, so the appended depths compound. Snapshot first.
    Application.Calculate
    ReDim depths(1 To 309)
    nVals = 0
    For r = 7 To 315
        v = pt.Cells(r, "H").Value2
        If VarType(v) = vbDouble Then
            nVals = nVals + 1
            depths(nVals) = CDbl(v)
        End If
    Next r

    If nVals = 0 Then
        MsgBox "No tally bit depths found in H7:H315 to append.", vbInformation
        Exit Sub
    End If

    destRow = lastRow + 1
    startRow = destRow
    nAdded = 0
    truncated = False

    ScreenBeginBusy "Pipe tally: adding bit depths..."
    For i = 1 To nVals
        If destRow > 305 Then
            truncated = True
            Exit For
        End If
        ss.Cells(destRow, "D").Value2 = depths(i)
        destRow = destRow + 1
        nAdded = nAdded + 1
    Next i

    pt.Range("F7:F315").ClearContents
    RefreshSlideComments
    ScreenEndBusy

    If truncated Then
        MsgBox "Stopped after " & CStr(nAdded) & _
               " values: Slidesheet bit-depth rows only go through D305.", _
               vbExclamation
    End If
    MsgBox "Added " & CStr(nAdded) & " bit depth(s) to Slidesheet starting at D" & _
           CStr(startRow) & ". Pipe tally F7:F315 cleared.", vbInformation
    Exit Sub

Fail:
    ScreenForceReset
    Dim fMsg As String
    fMsg = "AppendTallyBitDepths failed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description
    If Err.Number = 9 Then fMsg = fMsg & vbCrLf & vbCrLf & TallyLayoutHelp()
    MsgBox fMsg, vbCritical, "Pipe Tally Calculator"
End Sub

' Create / re-anchor the append button to merged B16 area.
Public Sub EnsureTallyAppendButton()
    Dim pt As Worksheet
    Dim rng As Range
    Dim btn As Button
    Dim shp As Shape

    Dim why As String
    Set pt = TallySheet(why)
    If pt Is Nothing Then
        MsgBox "Cannot place the tally button." & vbCrLf & vbCrLf & why, _
               vbExclamation, "Pipe Tally Calculator"
        Exit Sub
    End If
    Set rng = pt.Range(BTN_ANCHOR).MergeArea

    On Error Resume Next
    pt.Buttons(BTN_NAME).Delete
    pt.Shapes(BTN_NAME).Delete
    On Error GoTo 0

    Set btn = pt.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    btn.name = BTN_NAME
    btn.caption = BTN_CAPTION
    btn.OnAction = "'" & ThisWorkbook.name & "'!AppendTallyBitDepths"

    Set shp = pt.Shapes(BTN_NAME)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.Size = 10
    On Error GoTo 0
End Sub

Public Sub ResizeTallyAppendButton()
    Dim pt As Worksheet
    Dim rng As Range
    Dim shp As Shape

    On Error Resume Next
    Set pt = TallySheet()
    If pt Is Nothing Then Exit Sub
    Set shp = pt.Shapes(BTN_NAME)
    If shp Is Nothing Then
        EnsureTallyAppendButton
        Exit Sub
    End If
    Set rng = pt.Range(BTN_ANCHOR).MergeArea
    shp.Left = rng.Left
    shp.Top = rng.Top
    shp.Width = rng.Width
    shp.Height = rng.Height
    shp.Placement = xlMoveAndSize
    On Error GoTo 0
End Sub

Private Function LastBitDepthRow(ss As Worksheet) As Long
    Dim r As Long
    Dim last As Long
    last = 11
    ' Survey bit depths live in D12:D305; D306+ are summary LOOKUPs.
    For r = 12 To 305
        If Len(Trim$(CStr(ss.Cells(r, "D").text))) > 0 Then
            last = r
        End If
    Next r
    LastBitDepthRow = last
End Function




