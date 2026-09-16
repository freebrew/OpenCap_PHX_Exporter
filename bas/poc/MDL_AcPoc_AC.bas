Attribute VB_Name = "MDL_AcPoc_AC"
Option Explicit

' Compass AC report: Summary table + per-well Offset Design detail pages.

Public Sub AcPoc_ParseSummary()
    Dim txt As String
    Dim ws As Worksheet
    Dim startPos As Long
    Dim endPos As Long
    Dim p As Long
    Dim region As String
    Dim marker As Variant
    Dim re As Object
    Dim ms As Object
    Dim m As Object
    Dim n As Long
    Dim num2 As String
    Dim sfNum As String

    txt = AcPoc_ExtractPdfText(AcPoc_AcPath())
    If Len(txt) = 0 Then
        AcPoc_Log "Parse AC summary", "FAIL", "Empty AC PDF text"
        Exit Sub
    End If

    Set ws = AcPoc_EnsureSheet("_POC_AC_Summary", True)
    AcPoc_ClearSheet ws
    ws.Range("A1").Value2 = "Well"
    ws.Range("B1").Value2 = "RefMD"
    ws.Range("C1").Value2 = "OffMD"
    ws.Range("D1").Value2 = "BC"
    ws.Range("E1").Value2 = "BE"
    ws.Range("F1").Value2 = "SF"
    ws.Range("G1").Value2 = "Warning"
    ws.Range("A1:G1").Font.Bold = True

    startPos = InStr(1, txt, "Summary", vbTextCompare)
    If startPos = 0 Then startPos = 1
    endPos = Len(txt) + 1
    For Each marker In Array("Offset Design", "Semi Major Axis", "Highside")
        p = InStr(startPos + 8, txt, CStr(marker), vbTextCompare)
        If p > 0 And p < endPos Then endPos = p
    Next marker
    region = Mid$(txt, startPos, endPos - startPos)

    ' Same shape as MDL_Setup.ParseAcSummary: well + 4 distances + SF.
    num2 = "-?\d{1,3}(?:,\s?\d{3})*\.\s?\d{2}"
    sfNum = "\d{1,3}\.\s?\d{3}"
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "(\d{2,3}/\S{5,40}\s*-\s*.{1,60}?)\s+" & _
                 "(" & num2 & ")\s+(" & num2 & ")\s+(" & num2 & ")\s+(" & num2 & ")\s+" & _
                 "(" & sfNum & ")"

    Set ms = re.Execute(region)
    n = 0
    For Each m In ms
        n = n + 1
        ws.Cells(n + 1, 1).Value2 = AcPoc_CleanWellName(CStr(m.SubMatches(0)))
        ws.Cells(n + 1, 2).Value2 = AcPoc_CleanNum(m.SubMatches(1))
        ws.Cells(n + 1, 3).Value2 = AcPoc_CleanNum(m.SubMatches(2))
        ws.Cells(n + 1, 4).Value2 = AcPoc_CleanNum(m.SubMatches(3))
        ws.Cells(n + 1, 5).Value2 = AcPoc_CleanNum(m.SubMatches(4))
        ws.Cells(n + 1, 6).Value2 = AcPoc_CleanNum(m.SubMatches(5))
        ws.Cells(n + 1, 7).Value2 = AcPoc_WarningTail(CStr(m.Value))
    Next m

    If n = 0 Then
        AcPoc_Log "Parse AC summary", "FAIL", "No summary rows matched"
    Else
        AcPoc_Log "Parse AC summary", "OK", n & " summary warning rows"
    End If
End Sub

Private Function AcPoc_WarningTail(ByVal s As String) As String
    Dim t As String
    t = ""
    If InStr(1, s, "CC", vbTextCompare) > 0 Then t = t & "CC "
    If InStr(1, s, "ES", vbTextCompare) > 0 Then t = t & "ES "
    If InStr(1, s, "SF", vbTextCompare) > 0 Then t = t & "SF"
    AcPoc_WarningTail = Trim$(t)
End Function

Public Sub AcPoc_ParseDetail()
    Dim txt As String
    Dim pages() As String
    Dim ws As Worksheet
    Dim i As Long
    Dim well As String
    Dim n As Long
    Dim lineN As Long
    Dim lines() As String
    Dim Ln As String
    Dim nums() As Double
    Dim nn As Long
    Dim warn As String
    Dim hdrIdx As Long
    Dim hdr As String
    Dim cTF As Long, cN As Long, cE As Long, cBC As Long, cBE As Long, cSF As Long

    txt = AcPoc_ExtractPdfText(AcPoc_AcPath())
    pages = AcPoc_SplitPages(txt)

    Set ws = AcPoc_EnsureSheet("_POC_AC_Detail", True)
    AcPoc_ClearSheet ws
    ws.Range("A1").Value2 = "Well"
    ws.Range("B1").Value2 = "RefMD"
    ws.Range("C1").Value2 = "RefTVD"
    ws.Range("D1").Value2 = "OffMD"
    ws.Range("E1").Value2 = "OffTVD"
    ws.Range("F1").Value2 = "TF"
    ws.Range("G1").Value2 = "N"
    ws.Range("H1").Value2 = "E"
    ws.Range("I1").Value2 = "BC"
    ws.Range("J1").Value2 = "BE"
    ws.Range("K1").Value2 = "MinSep"
    ws.Range("L1").Value2 = "SF"
    ws.Range("M1").Value2 = "Warn"
    ws.Range("A1:M1").Font.Bold = True

    n = 0
    For i = LBound(pages) To UBound(pages)
        If InStr(1, pages(i), "Offset Design", vbTextCompare) = 0 Then GoTo NextPg
        well = AcPoc_OffsetDesignName(pages(i))
        If Len(well) = 0 Then GoTo NextPg
        Application.StatusBar = "[STEP 4/6] Parse AC detail  well=" & well
        lines = Split(pages(i), vbLf)
        hdrIdx = AcPoc_FindHeaderLine(lines)
        cTF = 0: cN = 0: cE = 0: cBC = 0: cBE = 0: cSF = 0
        If hdrIdx >= 0 Then
            hdr = lines(hdrIdx)
            cTF = InStr(1, hdr, "Toolface", vbTextCompare)
            cN = InStr(1, hdr, "+N/-S", vbTextCompare)
            cE = InStr(1, hdr, "+E/-W", vbTextCompare)
            cBC = InStr(1, hdr, "Centres", vbTextCompare)
            cBE = InStr(1, hdr, "Ellipses", vbTextCompare)
            cSF = InStr(1, hdr, "Factor", vbTextCompare)
        End If
        For lineN = LBound(lines) To UBound(lines)
            Ln = lines(lineN)
            If InStr(Ln, "Offset Design") > 0 Then GoTo NextDet
            If InStr(1, Ln, "Page ", vbTextCompare) > 0 Then GoTo NextDet
            nn = AcPoc_ExtractNums(Ln, nums)
            If nn < 10 Then GoTo NextDet
            If nums(0) < 0# Or nums(0) > 12000# Then GoTo NextDet
            n = n + 1
            ws.Cells(n + 1, 1).Value2 = well
            ws.Cells(n + 1, 2).Value2 = nums(0)
            ws.Cells(n + 1, 3).Value2 = nums(1)
            ws.Cells(n + 1, 4).Value2 = nums(2)
            ws.Cells(n + 1, 5).Value2 = nums(3)
            If nn >= 13 Then
                ws.Cells(n + 1, 6).Value2 = nums(6)
                ws.Cells(n + 1, 7).Value2 = nums(7)
                ws.Cells(n + 1, 8).Value2 = nums(8)
                ws.Cells(n + 1, 9).Value2 = nums(9)
                ws.Cells(n + 1, 10).Value2 = nums(10)
                ws.Cells(n + 1, 11).Value2 = nums(11)
                ws.Cells(n + 1, 12).Value2 = nums(12)
            ElseIf nn = 12 Then
                ws.Cells(n + 1, 7).Value2 = nums(6)
                ws.Cells(n + 1, 8).Value2 = nums(7)
                ws.Cells(n + 1, 9).Value2 = nums(8)
                ws.Cells(n + 1, 10).Value2 = nums(9)
                ws.Cells(n + 1, 11).Value2 = nums(10)
                ws.Cells(n + 1, 12).Value2 = nums(11)
            ElseIf cN > 0 And cE > 0 And cBC > 0 Then
                ws.Cells(n + 1, 6).Value2 = AcPoc_SliceNum(Ln, cTF, cN)
                ws.Cells(n + 1, 7).Value2 = AcPoc_SliceNum(Ln, cN, cE)
                ws.Cells(n + 1, 8).Value2 = AcPoc_SliceNum(Ln, cE, cBC)
                ws.Cells(n + 1, 9).Value2 = AcPoc_SliceNum(Ln, cBC, cBE)
                ws.Cells(n + 1, 10).Value2 = AcPoc_SliceNum(Ln, cBE, cSF)
                ws.Cells(n + 1, 12).Value2 = AcPoc_SliceNum(Ln, cSF, 0)
            Else
                ws.Cells(n + 1, 6).Value2 = nums(6)
                ws.Cells(n + 1, 7).Value2 = nums(7)
                ws.Cells(n + 1, 8).Value2 = nums(8)
                ws.Cells(n + 1, 9).Value2 = nums(9)
            End If
            warn = ""
            If InStr(1, Ln, " CC", vbTextCompare) > 0 Or Right$(Trim$(Ln), 2) = "CC" Then warn = "CC"
            If InStr(1, Ln, " ES", vbTextCompare) > 0 Or Right$(Trim$(Ln), 2) = "ES" Then warn = "ES"
            If InStr(1, Ln, " SF", vbTextCompare) > 0 Or Right$(Trim$(Ln), 2) = "SF" Then warn = "SF"
            ws.Cells(n + 1, 13).Value2 = warn
NextDet:
        Next lineN
NextPg:
    Next i

    If n < 10 Then
        AcPoc_Log "Parse AC detail", "FAIL", "Only " & n & " detail stations"
    Else
        AcPoc_Log "Parse AC detail", "OK", n & " detail stations across Offset Design pages"
    End If
End Sub

Private Function AcPoc_OffsetDesignName(ByVal pageTxt As String) As String
    Dim p As Long
    Dim rest As String
    Dim cut As Long
    p = InStr(1, pageTxt, "Offset Design", vbTextCompare)
    If p = 0 Then Exit Function
    rest = Mid$(pageTxt, p + Len("Offset Design"))
    rest = Replace(rest, vbLf, " ")
    rest = Trim$(rest)
    cut = InStr(1, rest, "Offset Site", vbTextCompare)
    If cut > 0 Then rest = Left$(rest, cut - 1)
    AcPoc_OffsetDesignName = AcPoc_CleanWellName(rest)
End Function
