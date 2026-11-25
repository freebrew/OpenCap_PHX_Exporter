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

' Excel cancels CutCopyMode (and empties the clipboard) as soon as a UserForm
' takes focus. Snapshot the text while the sheet still owns the copy, then
' reuse it when they click Paste. Selection stays even after the ants vanish.
Private m_fastDepth As Long
Private m_prevCalc As XlCalculation
Private m_prevEv As Boolean
Private m_prevScr As Boolean
Private mClipCache As String
Private mWatchOn As Boolean
Private mWatchWhen As Date

' Shown on the form after a CSV import. The form clears it on open.
Public gPipeTallyNote As String

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
    ImportPipeTallyCsv
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
    StopPipeTallyClipWatch
    Unload frmPipeTally
    On Error GoTo 0
End Sub

' File chooser replaces Paste from clipboard. Column G (total length) is
' matched to the last Slidesheet bit depth (column D). Lengths start on the
' next row, column F, and run to the bottom of the file.
Public Sub ImportPipeTallyCsv()
    Dim picked As Variant
    Dim lengths As Variant
    Dim errMsg As String

    ' Start in the folder of this workbook, where the rig export is saved.
    ' Restore the caller's folder afterward.
    Dim prevDir As String
    Dim prevDrive As String
    prevDir = CurDir$
    prevDrive = Left$(prevDir, 1)
    On Error Resume Next
    ChDrive Left$(ThisWorkbook.path, 1)
    ChDir ThisWorkbook.path
    On Error GoTo 0
    picked = Application.GetOpenFilename("Pipe tally CSV (*.csv),*.csv", , "Select pipe tally CSV")
    On Error Resume Next
    ChDrive prevDrive
    ChDir prevDir
    On Error GoTo 0
    If VarType(picked) = vbBoolean Then
        gPipeTallyNote = ""
        Exit Sub
    End If
    gPipeTallyNote = ""
    lengths = LoadPipeLengthsFromCsv(CStr(picked), errMsg)
    If Len(errMsg) > 0 Then
        gPipeTallyNote = errMsg
        MsgBox errMsg, vbExclamation, "Pipe Tally Calculator"
        Exit Sub
    End If
    If IsEmpty(lengths) Then Exit Sub
    PipeTallyWriteLengths lengths
    If Len(gPipeTallyNote) > 0 Then
        gPipeTallyNote = "Imported " & CStr(UBound(lengths)) & " pipe length(s). " & gPipeTallyNote
    Else
        gPipeTallyNote = "Imported " & CStr(UBound(lengths)) & " pipe length(s). Set stands and split, then add."
    End If
End Sub

' Whole file in, split on CRLF, CR, or LF. Line Input only breaks on CR, so an
' export with Unix line endings came back as one giant line and column G never
' parsed. Shared read so the file can stay open in Excel.
Private Function ReadCsvLines(ByVal path As String, ByRef errMsg As String) As Variant
    Dim fn As Long
    Dim buf As String
    Dim size As Long

    ReadCsvLines = Empty
    fn = FreeFile
    On Error GoTo BadFile
    Open path For Binary Access Read Shared As #fn
    size = LOF(fn)
    If size <= 0 Then
        Close #fn
        errMsg = "The file is empty: " & path
        Exit Function
    End If
    buf = Space$(size)
    Get #fn, 1, buf
    Close #fn
    On Error GoTo 0

    buf = Replace(buf, vbCrLf, vbLf)
    buf = Replace(buf, vbCr, vbLf)
    ReadCsvLines = Split(buf, vbLf)
    Exit Function

BadFile:
    On Error Resume Next
    Close #fn
    errMsg = "Could not read """ & path & """." & vbCrLf & Err.Description
End Function

Public Function LoadPipeLengthsFromCsv(ByVal path As String, ByRef errMsg As String) As Variant
    Dim bitMd As Double
    Dim rows As Variant
    Dim i As Long
    Dim gTxt As String
    Dim fTxt As String
    Dim gVal As Double
    Dim fVal As Double
    Dim bestRow As Long
    Dim bestG As Double
    Dim matched As Boolean
    Dim tmp() As Double
    Dim n As Long
    Dim capped As Boolean

    LoadPipeLengthsFromCsv = Empty
    errMsg = ""
    If Not LastSlidesheetBitMd(bitMd) Then
        errMsg = "No bit depth in Slidesheet column D."
        Exit Function
    End If

    rows = ReadCsvLines(path, errMsg)
    If Len(errMsg) > 0 Then Exit Function
    If IsEmpty(rows) Then
        errMsg = "Could not read """ & path & """."
        Exit Function
    End If

    ' Column G (totalLength): the row that matches the bit depth, or the
    ' closest total at or below it. Row 0 is the header.
    bestRow = -1
    bestG = 0#
    matched = False
    For i = LBound(rows) To UBound(rows)
        gTxt = CsvField(CStr(rows(i)), 7)
        If IsNumeric(gTxt) And Len(gTxt) > 0 Then
            gVal = val(gTxt)
            If Abs(gVal - bitMd) <= 0.05 Then
                bestRow = i
                bestG = gVal
                matched = True
                Exit For
            End If
            If gVal <= bitMd + 0.001 Then
                If bestRow < 0 Or gVal > bestG Then
                    bestRow = i
                    bestG = gVal
                End If
            End If
        End If
    Next i

    If bestRow < 0 Then
        errMsg = "No tally total in column G is at or below the bit depth " & Format$(bitMd, "0.00") & " m."
        Exit Function
    End If

    ' G only runs upward. If the closest total is more than two joints under
    ' the bit, the file stops before the bit: it is an older export.
    If Not matched And (bitMd - bestG) > 20# Then
        errMsg = "This file only reaches " & Format$(bestG, "0.00") & " m in column G, and the bit is at " & _
                 Format$(bitMd, "0.00") & " m." & vbCrLf & "Pick a newer pipe tally export."
        Exit Function
    End If

    ' One row down from the match, column F (length), to the end of the file.
    ReDim tmp(1 To (PT_LAST - PT_FIRST + 1))
    n = 0
    For i = bestRow + 1 To UBound(rows)
        fTxt = CsvField(CStr(rows(i)), 6)
        If IsNumeric(fTxt) And Len(fTxt) > 0 Then
            fVal = val(fTxt)
            If fVal >= MIN_JOINT_M And fVal <= MAX_JOINT_M Then
                If n >= (PT_LAST - PT_FIRST + 1) Then
                    capped = True
                    Exit For
                End If
                n = n + 1
                tmp(n) = fVal
            End If
        End If
    Next i

    If n = 0 Then
        errMsg = "Bit depth " & Format$(bitMd, "0.00") & " m matched tally total " & Format$(bestG, "0.00") & _
                 " m. No single pipe lengths in column F below that row."
        Exit Function
    End If
    ReDim Preserve tmp(1 To n)
    LoadPipeLengthsFromCsv = tmp
    If Not matched Then
        gPipeTallyNote = "Closest tally total at or below the bit is " & Format$(bestG, "0.00") & " m (bit " & Format$(bitMd, "0.00") & " m)."
    End If
    If capped Then
        gPipeTallyNote = gPipeTallyNote & " Stopped at " & CStr(n) & " joints (the form holds that many)."
    End If
End Function

Private Function LastSlidesheetBitMd(ByRef bitMd As Double) As Boolean
    Dim ss As Worksheet
    Dim r As Long
    Dim v As Variant
    Dim col As Variant
    LastSlidesheetBitMd = False
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    col = ss.Range("D12:D505").Value2
    For r = 494 To 1 Step -1
        v = col(r, 1)
        If isError(v) Then GoTo NextD
        If IsNumeric(v) Then
            If CDbl(v) > 0# Then
                bitMd = CDbl(v)
                LastSlidesheetBitMd = True
                Exit Function
            End If
        End If
NextD:
    Next r
End Function

Private Function CsvField(ByVal line As String, ByVal index1 As Long) As String
    Dim parts() As String
    Dim s As String
    CsvField = ""
    If Len(line) = 0 Then Exit Function
    parts = Split(line, ",")
    If index1 < 1 Or index1 - 1 > UBound(parts) Then Exit Function
    s = Trim$(parts(index1 - 1))
    If Left$(s, 1) = """" Then s = mid$(s, 2)
    If Right$(s, 1) = """" And Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    CsvField = s
End Function

Public Sub CachePipeTallyClipboard()
    Dim t As String
    t = GetClipboardText()
    If Len(Trim$(t)) > 0 Then mClipCache = t
End Sub

Public Sub StartPipeTallyClipWatch()
    mWatchOn = True
    SchedulePipeTallyClipWatch
End Sub

Public Sub StopPipeTallyClipWatch()
    mWatchOn = False
    On Error Resume Next
    If mWatchWhen <> 0 Then Application.OnTime EarliestTime:=mWatchWhen, _
        Procedure:="MDL_PipeTally.PipeTallyClipWatchTick", Schedule:=False
    On Error GoTo 0
    mWatchWhen = 0
End Sub

Public Sub PipeTallyClipWatchTick()
    On Error Resume Next
    CachePipeTallyClipboard
    If mWatchOn Then SchedulePipeTallyClipWatch
End Sub

Private Sub SchedulePipeTallyClipWatch()
    On Error Resume Next
    mWatchWhen = Now + TimeSerial(0, 0, 1)
    Application.OnTime EarliestTime:=mWatchWhen, _
        Procedure:="MDL_PipeTally.PipeTallyClipWatchTick"
End Sub

' Live clipboard, then the snapshot taken before Excel flushed CutCopyMode,
' then the current sheet selection (copy-ants go away; the range does not).
Public Function GetPipeTallyPasteText() As String
    Dim t As String
    t = GetClipboardText()
    If Len(Trim$(t)) > 0 Then
        mClipCache = t
        GetPipeTallyPasteText = t
        Exit Function
    End If
    If Len(Trim$(mClipCache)) > 0 Then
        GetPipeTallyPasteText = mClipCache
        Exit Function
    End If
    GetPipeTallyPasteText = SelectionColumnText()
End Function

Private Function SelectionColumnText() As String
    Dim rng As Range
    Dim area As Range
    Dim c As Range
    Dim buf As String
    Dim n As Long

    SelectionColumnText = ""
    On Error Resume Next
    Set rng = Application.Selection
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    If TypeName(rng) <> "Range" Then Exit Function
    If rng.Areas.Count <> 1 Then Exit Function
    Set area = rng.Columns(1)
    On Error Resume Next
    Set area = Intersect(area, area.Worksheet.UsedRange)
    On Error GoTo 0
    If area Is Nothing Then Exit Function

    For Each c In area.Cells
        If n >= 400 Then Exit For
        If Len(Trim$(CStr(c.Value2 & ""))) = 0 Then GoTo NextSel
        n = n + 1
        buf = buf & CStr(c.Value2) & vbLf
NextSel:
    Next c
    SelectionColumnText = buf
End Function

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

    Set btn = ss.Buttons.Add(rng.Left, rng.Top, rng.width, rng.Height)
    btn.name = PT_BTN
    btn.caption = PT_CAPTION
    btn.OnAction = "'" & ThisWorkbook.name & "'!ShowPipeTallyForm"

    Set shp = ss.Shapes(PT_BTN)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.size = 10
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

' Each cell write on this tab fires a recalc and the sheet's events. Writes
' go through FastBegin / FastEnd so a batch recalculates once, at the end.
Private Sub FastBegin()
    If m_fastDepth = 0 Then
        m_prevCalc = Application.Calculation
        m_prevEv = Application.EnableEvents
        m_prevScr = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If
    m_fastDepth = m_fastDepth + 1
End Sub

Private Sub FastEnd()
    If m_fastDepth = 0 Then Exit Sub
    m_fastDepth = m_fastDepth - 1
    If m_fastDepth = 0 Then
        Application.Calculation = m_prevCalc
        Application.EnableEvents = m_prevEv
        Application.ScreenUpdating = m_prevScr
    End If
End Sub

Public Sub PipeTallySetSettings(ByVal stands As Variant, ByVal splitLen As Variant)
    Dim pt As Worksheet
    Dim errNo As Long
    Dim errTxt As String
    Set pt = TallySheet()
    If pt Is Nothing Then Exit Sub
    FastBegin
    On Error Resume Next
    pt.Range("C11").Value2 = stands
    If Len(Trim$(CStr(splitLen & ""))) = 0 Then
        pt.Range("D11").ClearContents
    Else
        pt.Range("D11").Value2 = splitLen
    End If
    errNo = Err.Number: errTxt = Err.Description
    On Error GoTo 0
    FastEnd
    If errNo <> 0 Then Err.Raise errNo, , errTxt
End Sub

Public Function PipeTallyReadLengths() As Variant
    Dim pt As Worksheet
    Dim r As Long, n As Long, v As Variant
    Dim vals As Variant
    Dim tmp() As Double
    Set pt = TallySheet()
    If pt Is Nothing Then PipeTallyReadLengths = Empty: Exit Function
    vals = pt.Range("F" & PT_FIRST & ":F" & PT_LAST).Value2
    ReDim tmp(1 To (PT_LAST - PT_FIRST + 1))
    n = 0
    For r = 1 To UBound(vals, 1)
        v = vals(r, 1)
        If Not isError(v) Then
            If IsNumeric(v) And Not IsEmpty(v) Then
                If CDbl(v) > 0# Then
                    n = n + 1
                    tmp(n) = CDbl(v)
                End If
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
    Dim i As Long, n As Long, cap As Long
    Dim arr() As Variant
    Dim errNo As Long
    Dim errTxt As String
    Set pt = TallySheet()
    If pt Is Nothing Then Exit Sub
    cap = PT_LAST - PT_FIRST + 1
    FastBegin
    On Error Resume Next
    pt.Range("F" & PT_FIRST & ":F" & PT_LAST).ClearContents
    If Not IsEmpty(lengths) Then
        n = UBound(lengths)
        If n > cap Then n = cap
        If n > 0 Then
            ReDim arr(1 To n, 1 To 1)
            For i = 1 To n
                arr(i, 1) = CDbl(lengths(i))
            Next i
            ' One write for the whole column, not one per joint.
            pt.Range("F" & PT_FIRST).Resize(n, 1).Value2 = arr
        End If
    End If
    errNo = Err.Number: errTxt = Err.Description
    On Error GoTo 0
    FastEnd
    If errNo <> 0 Then Err.Raise errNo, , errTxt
End Sub

Public Function PipeTallyReadBitDepths() As Variant
    Dim pt As Worksheet
    Dim r As Long, n As Long
    Dim vals As Variant
    Dim tmp() As Double
    Set pt = TallySheet()
    If pt Is Nothing Then PipeTallyReadBitDepths = Empty: Exit Function
    ' Only this tab's depths are needed. A workbook-wide Calculate here made
    ' every form action wait on the Slidesheet / Data / Cont DI chain.
    pt.Calculate
    vals = pt.Range("H" & PT_FIRST & ":H" & PT_LAST).Value2
    ReDim tmp(1 To (PT_LAST - PT_FIRST + 1))
    n = 0
    For r = 1 To UBound(vals, 1)
        If VarType(vals(r, 1)) = vbDouble Then
            n = n + 1
            tmp(n) = CDbl(vals(r, 1))
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
    FastBegin
    On Error Resume Next
    pt.Range("F" & PT_FIRST & ":F" & PT_LAST).ClearContents
    On Error GoTo 0
    FastEnd
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

    ' One bulk write of D, then one recalc / comment / gauge pass. A cell-by-cell
    ' write fired the Slidesheet's full calculation on every depth.
    nAdded = nVals
    If destRow + nAdded - 1 > 505 Then
        nAdded = 505 - destRow + 1
        If nAdded < 0 Then nAdded = 0
        truncated = True
    End If

    ScreenBeginBusy "Pipe tally: adding bit depths..."
    If nAdded > 0 Then
        Dim outArr() As Variant
        Dim bulk As Range
        ReDim outArr(1 To nAdded, 1 To 1)
        For i = 1 To nAdded
            outArr(i, 1) = depths(i)
        Next i
        Set bulk = ss.Range(ss.Cells(destRow, "D"), ss.Cells(destRow + nAdded - 1, "D"))
        FastBegin
        bulk.Value2 = outArr
        pt.Range("F7:F315").ClearContents
        FastEnd
        ' The sheet's own Change handler was off for the write; do its work once.
        On Error Resume Next
        ScreenCalcStaggered
        RefreshRotateMetres bulk
        RefreshSlideComments
        MDL_PlanGauge.RenderPlanGauge
        RefreshNewestSurveyIfMoved
        On Error GoTo Fail
    Else
        pt.Range("F7:F315").ClearContents
    End If
    ScreenEndBusy

    If truncated Then
        MsgBox "Stopped after " & CStr(nAdded) & _
               " values: Slidesheet bit-depth rows only go through D505.", _
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

    Set btn = pt.Buttons.Add(rng.Left, rng.Top, rng.width, rng.Height)
    btn.name = BTN_NAME
    btn.caption = BTN_CAPTION
    btn.OnAction = "'" & ThisWorkbook.name & "'!AppendTallyBitDepths"

    Set shp = pt.Shapes(BTN_NAME)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.size = 10
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
    shp.width = rng.width
    shp.Height = rng.Height
    shp.Placement = xlMoveAndSize
    On Error GoTo 0
End Sub

Private Function LastBitDepthRow(ss As Worksheet) As Long
    Dim r As Long
    Dim last As Long
    last = 11
    ' Survey bit depths live in D12:D505; D506+ are summary LOOKUPs.
    For r = 12 To 505
        If Len(Trim$(CStr(ss.Cells(r, "D").text))) > 0 Then
            last = r
        End If
    Next r
    LastBitDepthRow = last
End Function














