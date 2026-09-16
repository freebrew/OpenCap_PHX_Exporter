Attribute VB_Name = "MDL_AcPoc_Pdf"
Option Explicit

' Shared PDF text + sheet helpers for the standalone AC corridor POC.
' pdftotext -layout is the same path the Slide Sheet uses (ResolvePdftotext).

Public Function AcPoc_ExtractPdfText(ByVal pdfPath As String) As String
    Dim tmpOut As String
    Dim fNum As Integer
    Dim content As String
    Dim Ln As String
    Dim Sh As Object
    Dim p2tPath As String

    AcPoc_ExtractPdfText = ""
    tmpOut = Environ$("TEMP") & "\ac_poc_pdf_text.txt"
    On Error Resume Next
    Kill tmpOut
    On Error GoTo 0

    Set Sh = CreateObject("WScript.Shell")
    p2tPath = AcPoc_ResolvePdftotext(Sh)
    If p2tPath = "" Then
        Err.Raise vbObjectError + 1, "AcPoc_ExtractPdfText", _
            "pdftotext not found. Install Poppler or run a Slide Sheet AC import once to bootstrap it."
    End If

    Sh.Run """" & p2tPath & """ -layout """ & pdfPath & """ """ & tmpOut & """", 0, True

    If Dir(tmpOut) = "" Then
        Err.Raise vbObjectError + 2, "AcPoc_ExtractPdfText", "pdftotext produced no output for " & pdfPath
    End If

    fNum = FreeFile
    content = ""
    Open tmpOut For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, Ln
        content = content & Ln & vbLf
    Loop
    Close #fNum
    On Error Resume Next
    Kill tmpOut
    On Error GoTo 0
    AcPoc_ExtractPdfText = content
End Function

Public Function AcPoc_ResolvePdftotext(ByVal Sh As Object) As String
    Dim ptr As String
    Dim fNum As Integer
    Dim exePath As String

    AcPoc_ResolvePdftotext = ""
    On Error Resume Next
    If Sh.Run("cmd /c where pdftotext >nul 2>nul", 0, True) = 0 Then
        AcPoc_ResolvePdftotext = "pdftotext"
    End If
    On Error GoTo 0
    If AcPoc_ResolvePdftotext <> "" Then Exit Function

    ptr = Environ$("LOCALAPPDATA") & "\Poppler\pdftotext_path.txt"
    If Dir(ptr) = "" Then Exit Function

    fNum = FreeFile
    On Error Resume Next
    Open ptr For Input As #fNum
    Line Input #fNum, exePath
    Close #fNum
    On Error GoTo 0

    exePath = Trim$(exePath)
    If exePath <> "" Then
        If Dir(exePath) <> "" Then AcPoc_ResolvePdftotext = exePath
    End If
End Function

Public Function AcPoc_CleanNum(ByVal s As String) As Double
    Dim i As Long
    Dim t As String
    Dim ch As String
    s = Replace(Replace(Trim$(s), ",", ""), " ", "")
    t = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9.-]" Then
            t = t & ch
        ElseIf Len(t) > 0 Then
            Exit For
        End If
    Next i
    If Len(t) = 0 Or t = "-" Or t = "." Or Not IsNumeric(t) Then
        AcPoc_CleanNum = 0#
    Else
        AcPoc_CleanNum = CDbl(t)
    End If
End Function

Public Function AcPoc_CleanWellName(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    AcPoc_CleanWellName = Trim$(s)
End Function

Public Function AcPoc_SplitPages(ByVal txt As String) As String()
    AcPoc_SplitPages = Split(txt, Chr$(12))
End Function

Public Function AcPoc_FindHeaderLine(lines() As String) As Long
    Dim i As Long
    AcPoc_FindHeaderLine = -1
    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), "+N/-S", vbTextCompare) > 0 And _
           InStr(1, lines(i), "Centres", vbTextCompare) > 0 Then
            AcPoc_FindHeaderLine = i
            Exit Function
        End If
    Next i
End Function

Public Function AcPoc_SliceField(ByVal Ln As String, ByVal startCol As Long, ByVal nextCol As Long) As String
    If startCol < 1 Then
        AcPoc_SliceField = ""
        Exit Function
    End If
    If nextCol > startCol Then
        AcPoc_SliceField = Mid$(Ln, startCol, nextCol - startCol)
    Else
        AcPoc_SliceField = Mid$(Ln, startCol)
    End If
End Function

Public Function AcPoc_SliceNum(ByVal Ln As String, ByVal startCol As Long, ByVal nextCol As Long) As Double
    AcPoc_SliceNum = AcPoc_CleanNum(AcPoc_SliceField(Ln, startCol, nextCol))
End Function

Public Function AcPoc_ExtractNums(ByVal s As String, ByRef nums() As Double) As Long
    Dim re As Object
    Dim ms As Object
    Dim m As Object
    Dim n As Long
    Dim t As String

    ReDim nums(0 To 50)
    n = 0
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "-?\d{1,6}(?:,\d{3})*\.\d{2,3}"
    Set ms = re.Execute(s)
    For Each m In ms
        t = Replace(m.Value, ",", "")
        nums(n) = CDbl(t)
        n = n + 1
        If n > 50 Then Exit For
    Next m
    AcPoc_ExtractNums = n
End Function

Public Function AcPoc_WellKey(ByVal s As String) As String
    Dim i As Long, startP As Long, endP As Long
    Dim ch As String
    s = AcPoc_CleanWellName(s)
    For i = 1 To Len(s) - 2
        If UCase$(Mid$(s, i, 1)) = "W" And Mid$(s, i + 2, 1) = "/" Then
            startP = i
            Do While startP > 1
                ch = Mid$(s, startP - 1, 1)
                If ch Like "[0-9/\-]" Then
                    startP = startP - 1
                Else
                    Exit Do
                End If
            Loop
            endP = i + 2
            Do While endP <= Len(s)
                ch = Mid$(s, endP, 1)
                If ch Like "[0-9A-Za-z]" Then
                    endP = endP + 1
                Else
                    Exit Do
                End If
            Loop
            AcPoc_WellKey = UCase$(Mid$(s, startP, endP - startP))
            Exit Function
        End If
    Next i
    AcPoc_WellKey = UCase$(s)
End Function

Public Function AcPoc_WellsMatch(ByVal a As String, ByVal b As String) As Boolean
    Dim ka As String, kb As String
    If InStr(1, b, a, vbTextCompare) > 0 Then AcPoc_WellsMatch = True: Exit Function
    If InStr(1, a, b, vbTextCompare) > 0 Then AcPoc_WellsMatch = True: Exit Function
    ka = AcPoc_WellKey(a)
    kb = AcPoc_WellKey(b)
    AcPoc_WellsMatch = (Len(ka) > 0 And ka = kb)
End Function

Public Function AcPoc_EnsureSheet(ByVal name As String, Optional ByVal hidden As Boolean = False) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = name
    End If
    If hidden Then
        ws.Visible = xlSheetHidden
    Else
        ws.Visible = xlSheetVisible
    End If
    Set AcPoc_EnsureSheet = ws
End Function

Public Sub AcPoc_ClearSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells.ClearFormats
End Sub

Public Function AcPoc_LastDataRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r < 2 Then
        AcPoc_LastDataRow = 1
    Else
        AcPoc_LastDataRow = r
    End If
End Function
