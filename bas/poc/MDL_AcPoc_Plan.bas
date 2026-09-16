Attribute VB_Name = "MDL_AcPoc_Plan"
Option Explicit

' Planned Survey + Plan Sections from a Compass planning PDF (pdftotext -layout).

Public Sub AcPoc_ParsePlan()
    Dim txt As String
    Dim ws As Worksheet
    Dim sec As Worksheet
    Dim lines() As String
    Dim i As Long
    Dim n As Long
    Dim nums() As Double
    Dim nn As Long
    Dim inSurvey As Boolean
    Dim md As Double, lastMd As Double
    Dim nSec As Long

    txt = AcPoc_ExtractPdfText(AcPoc_PlanPath())
    If Len(txt) = 0 Then
        AcPoc_Log "Parse plan", "FAIL", "Empty plan PDF text"
        Exit Sub
    End If

    Set ws = AcPoc_EnsureSheet("_POC_Plan", True)
    AcPoc_ClearSheet ws
    ws.Range("A1").Value2 = "MD"
    ws.Range("B1").Value2 = "Inc"
    ws.Range("C1").Value2 = "Azm"
    ws.Range("D1").Value2 = "TVD"
    ws.Range("E1").Value2 = "NS"
    ws.Range("F1").Value2 = "EW"
    ws.Range("G1").Value2 = "VS"
    ws.Range("H1").Value2 = "DLS"
    ws.Range("A1:H1").Font.Bold = True

    Set sec = AcPoc_EnsureSheet("_POC_PlanSec", True)
    AcPoc_ClearSheet sec
    sec.Range("A1").Value2 = "MD"
    sec.Range("B1").Value2 = "Inc"
    sec.Range("C1").Value2 = "Azm"
    sec.Range("D1").Value2 = "TVD"
    sec.Range("E1").Value2 = "NS"
    sec.Range("F1").Value2 = "EW"
    sec.Range("G1").Value2 = "Name"
    sec.Range("A1:G1").Font.Bold = True

    lines = Split(txt, vbLf)
    lastMd = -1#
    n = 0
    nSec = 0
    inSurvey = False

    For i = LBound(lines) To UBound(lines)
        If InStr(1, lines(i), "Planned Survey", vbTextCompare) > 0 Then
            inSurvey = True
            GoTo NextLn
        End If
        If InStr(1, lines(i), "Plan Sections", vbTextCompare) > 0 Then
            inSurvey = False
            GoTo NextLn
        End If
        If InStr(1, lines(i), "Plan Annotations", vbTextCompare) > 0 Then
            inSurvey = False
            GoTo NextLn
        End If

        nn = AcPoc_ExtractNums(lines(i), nums)
        If inSurvey Then
            If nn >= 11 Then
                md = nums(0)
                If md + 0.001 < lastMd Then GoTo NextLn
                If nums(1) < -0.01 Or nums(1) > 180.01 Then GoTo NextLn
                If nums(3) < -50# Or nums(3) > 6000# Then GoTo NextLn
                n = n + 1
                ws.Cells(n + 1, 1).Value2 = nums(0)
                ws.Cells(n + 1, 2).Value2 = nums(1)
                ws.Cells(n + 1, 3).Value2 = nums(2)
                ws.Cells(n + 1, 4).Value2 = nums(3)
                ws.Cells(n + 1, 5).Value2 = nums(5)
                ws.Cells(n + 1, 6).Value2 = nums(6)
                ws.Cells(n + 1, 7).Value2 = nums(7)
                ws.Cells(n + 1, 8).Value2 = nums(8)
                lastMd = md
            End If
        End If
NextLn:
    Next i

    nSec = AcPoc_ParsePlanSections(txt, sec)

    If n < 50 Then
        AcPoc_Log "Parse plan", "FAIL", "Only " & n & " Planned Survey stations (need > 50)"
    Else
        AcPoc_Log "Parse plan", "OK", n & " Planned Survey stations; last MD=" & _
            Format$(ws.Cells(n + 1, 1).Value2, "0.00")
    End If
    If nSec < 4 Then
        AcPoc_Log "Parse plan", "FAIL", "Only " & nSec & " Plan Sections rows"
    Else
        AcPoc_Log "Parse plan", "OK", nSec & " Plan Sections rows"
    End If
End Sub

Private Function AcPoc_ParsePlanSections(ByVal txt As String, ByVal sec As Worksheet) As Long
    Dim startPos As Long, endPos As Long, p As Long
    Dim region As String
    Dim lines() As String
    Dim i As Long, nn As Long, n As Long
    Dim nums() As Double
    Dim lastMd As Double
    Dim nm As String
    Dim marker As Variant

    startPos = InStr(1, txt, "Plan Sections", vbTextCompare)
    If startPos = 0 Then startPos = InStr(1, txt, "Plan Section", vbTextCompare)
    If startPos = 0 Then
        AcPoc_ParsePlanSections = 0
        Exit Function
    End If
    endPos = Len(txt) + 1
    For Each marker In Array("Planning Report", "Planned Survey", "SECTION DETAILS", _
                             "DESIGN TARGET", "Plan Annotations")
        p = InStr(startPos + 14, txt, CStr(marker), vbTextCompare)
        If p > 0 And p < endPos Then endPos = p
    Next marker
    region = Mid$(txt, startPos, endPos - startPos)
    lines = Split(region, vbLf)
    lastMd = -1#
    n = 0
    For i = LBound(lines) To UBound(lines)
        nn = AcPoc_ExtractNums(lines(i), nums)
        If nn >= 6 Then
            If nums(0) + 0.05 < lastMd Then GoTo NextSec
            If nums(1) < -0.01 Or nums(1) > 180.01 Then GoTo NextSec
            n = n + 1
            sec.Cells(n + 1, 1).Value2 = nums(0)
            sec.Cells(n + 1, 2).Value2 = nums(1)
            sec.Cells(n + 1, 3).Value2 = nums(2)
            sec.Cells(n + 1, 4).Value2 = nums(3)
            sec.Cells(n + 1, 5).Value2 = nums(4)
            sec.Cells(n + 1, 6).Value2 = nums(5)
            nm = ""
            If InStr(1, lines(i), "Heel", vbTextCompare) > 0 Then nm = "HEEL"
            If InStr(1, lines(i), "Toe", vbTextCompare) > 0 Then nm = "TD"
            If InStr(1, lines(i), "KOP", vbTextCompare) > 0 Then nm = "KOP"
            sec.Cells(n + 1, 7).Value2 = nm
            lastMd = nums(0)
        End If
NextSec:
    Next i
    AcPoc_ParsePlanSections = n
End Function
