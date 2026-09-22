Attribute VB_Name = "MDL_PlanImport"
Option Explicit

' ================================================================================
'  MDL_PlanImport
'
'  Compass Planning Report extras from the same PDF as Import Plan:
'    Planned Survey  -> _OC_Survey   (dense stations for proximity / PlanAt)
'    Formations      -> _OC_Formations (MD, TVD, Name, Lithology)
'
'  Lithology is not filled on the Compass Formations page. Names are matched
'  to the Alberta Energy Regulator / Alberta Geological Survey
'  "Alberta Table of Formations" (2019) and the AGS Atlas of the Western
'  Canada Sedimentary Basin (ags.aer.ca). Short field labels only.
' ================================================================================

Public Const PI_SURVEY_SHEET As String = "_OC_Survey"
Public Const PI_FORM_SHEET As String = "_OC_Formations"
Private Const PI_LITHO_SRC As String = "AER Table of Formations 2019 / AGS WCSB Atlas"

Public Function ImportPlannedSurveysFromPdf(ByVal pdfText As String, ByVal fPath As String) As Long
    Dim md() As Double, inc() As Double, azm() As Double
    Dim tvd() As Double, ns() As Double, ew() As Double
    Dim n As Long
    ImportPlannedSurveysFromPdf = 0
    n = ParsePlannedSurveys(pdfText, md, inc, azm, tvd, ns, ew)
    If n < 50 Then Exit Function
    WriteSurveySheet fPath, n, md, inc, azm, tvd, ns, ew
    ImportPlannedSurveysFromPdf = n
End Function

Public Function ImportFormationsFromPdf(ByVal pdfText As String, ByVal fPath As String) As Long
    Dim md() As Double, tvd() As Double, nm() As String, lith() As String
    Dim n As Long
    ImportFormationsFromPdf = 0
    n = ParseFormations(pdfText, md, tvd, nm, lith)
    If n < 1 Then Exit Function
    WriteFormationsSheet fPath, n, md, tvd, nm, lith
    ImportFormationsFromPdf = n
End Function

Public Function PI_LoadFormations(ByRef md() As Double, ByRef tvd() As Double, _
                                  ByRef nm() As String, ByRef lith() As String) As Long
    Dim ws As Worksheet
    Dim lastR As Long, r As Long, n As Long
    Dim vM As Variant, vT As Variant
    PI_LoadFormations = 0
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PI_FORM_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then Exit Function
    ReDim md(0 To lastR): ReDim tvd(0 To lastR)
    ReDim nm(0 To lastR): ReDim lith(0 To lastR)
    n = 0
    For r = 3 To lastR
        vM = ws.Cells(r, 1).Value2
        vT = ws.Cells(r, 2).Value2
        If IsNumeric(vM) And IsNumeric(vT) Then
            md(n) = CDbl(vM)
            tvd(n) = CDbl(vT)
            nm(n) = Trim$(CStr(ws.Cells(r, 3).Value2 & ""))
            lith(n) = Trim$(CStr(ws.Cells(r, 4).Value2 & ""))
            n = n + 1
        End If
    Next r
    PI_LoadFormations = n
End Function

Public Function PI_LithologyForName(ByVal formName As String) As String
    PI_LithologyForName = LithologyForNorm(NormFormName(formName))
End Function

' ---- parse ------------------------------------------------------------------

Private Function ParsePlannedSurveys(ByVal pdfText As String, _
        ByRef md() As Double, ByRef inc() As Double, ByRef azm() As Double, _
        ByRef tvd() As Double, ByRef ns() As Double, ByRef ew() As Double) As Long
    Dim lines() As String, i As Long, nn As Long, n As Long
    Dim nums() As Double, lastMd As Double
    Dim inSurvey As Boolean, Ln As String
    ReDim md(0 To 800): ReDim inc(0 To 800): ReDim azm(0 To 800)
    ReDim tvd(0 To 800): ReDim ns(0 To 800): ReDim ew(0 To 800)
    lines = Split(pdfText, Chr$(10))
    lastMd = -1#
    n = 0
    inSurvey = False
    For i = LBound(lines) To UBound(lines)
        Ln = lines(i)
        If InStr(1, Ln, "Planned Survey", vbTextCompare) > 0 Then
            inSurvey = True
            GoTo NextLn
        End If
        If InStr(1, Ln, "Plan Sections", vbTextCompare) > 0 _
                Or InStr(1, Ln, "Plan Annotations", vbTextCompare) > 0 _
                Or InStr(1, Ln, "Formations", vbTextCompare) > 0 _
                Or InStr(1, Ln, "SECTION DETAILS", vbTextCompare) > 0 Then
            If InStr(1, Ln, "Planned Survey", vbTextCompare) = 0 Then inSurvey = False
            GoTo NextLn
        End If
        If Not inSurvey Then GoTo NextLn
        nn = ExtractNums(Ln, nums)
        If nn >= 8 Then
            If nums(0) + 0.001 < lastMd Then GoTo NextLn
            If nums(1) < -0.01 Or nums(1) > 180.01 Then GoTo NextLn
            If nums(3) < -50# Or nums(3) > 12000# Then GoTo NextLn
            If n > 800 Then Exit For
            md(n) = nums(0): inc(n) = nums(1): azm(n) = nums(2): tvd(n) = nums(3)
            If nn >= 7 Then
                ns(n) = nums(5): ew(n) = nums(6)
            Else
                ns(n) = 0#: ew(n) = 0#
            End If
            lastMd = nums(0)
            n = n + 1
        End If
NextLn:
    Next i
    ParsePlannedSurveys = n
End Function

Private Function ParseFormations(ByVal pdfText As String, _
        ByRef md() As Double, ByRef tvd() As Double, _
        ByRef nm() As String, ByRef lith() As String) As Long
    Dim startPos As Long, endPos As Long, p As Long
    Dim region As String, lines() As String
    Dim i As Long, nn As Long, n As Long
    Dim nums() As Double, lastMd As Double, formNm As String
    Dim marker As Variant

    ParseFormations = 0
    startPos = InStr(1, pdfText, "Formations", vbTextCompare)
    If startPos = 0 Then Exit Function
    endPos = Len(pdfText) + 1
    For Each marker In Array("Plan Annotations", "Planning Report", "Planned Survey", "Plan Sections")
        p = InStr(startPos + 10, pdfText, CStr(marker), vbTextCompare)
        If p > 0 And p < endPos Then endPos = p
    Next marker
    region = mid$(pdfText, startPos, endPos - startPos)
    lines = Split(region, Chr$(10))
    ReDim md(0 To 120): ReDim tvd(0 To 120): ReDim nm(0 To 120): ReDim lith(0 To 120)
    lastMd = -1#
    n = 0
    For i = LBound(lines) To UBound(lines)
        nn = ExtractNums(lines(i), nums)
        If nn < 2 Then GoTo NextFm
        If nums(0) + 0.05 < lastMd Then GoTo NextFm
        If nums(0) < 1# Or nums(1) < -50# Then GoTo NextFm
        formNm = NameAfterTwoNums(lines(i))
        If Len(formNm) < 2 Then GoTo NextFm
        If n > 120 Then Exit For
        md(n) = nums(0)
        tvd(n) = nums(1)
        nm(n) = formNm
        lith(n) = LithologyForNorm(NormFormName(formNm))
        lastMd = nums(0)
        n = n + 1
NextFm:
    Next i
    ParseFormations = n
End Function

Private Function NameAfterTwoNums(ByVal Ln As String) As String
    Dim re As Object, ms As Object
    Dim s As String, p As Long
    s = Trim$(Ln)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "-?\d{1,6}(?:,\d{3})*\.\d{2,3}"
    Set ms = re.Execute(s)
    If ms.Count < 2 Then
        NameAfterTwoNums = ""
        Exit Function
    End If
    p = ms(1).FirstIndex + ms(1).Length + 1
    s = Trim$(mid$(s, p))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If InStr(1, s, "Lithology", vbTextCompare) > 0 Then s = ""
    If InStr(1, s, "Depth", vbTextCompare) > 0 Then s = ""
    NameAfterTwoNums = Trim$(s)
End Function

Private Function ExtractNums(ByVal s As String, ByRef nums() As Double) As Long
    Dim re As Object, ms As Object, m As Object
    Dim n As Long, t As String
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
    ExtractNums = n
End Function

' ---- write ------------------------------------------------------------------

Private Sub WriteSurveySheet(ByVal fPath As String, ByVal n As Long, _
        ByRef md() As Double, ByRef inc() As Double, ByRef azm() As Double, _
        ByRef tvd() As Double, ByRef ns() As Double, ByRef ew() As Double)
    Dim ws As Worksheet, i As Long
    Set ws = EnsureHidden(PI_SURVEY_SHEET)
    ws.Cells(1, 1).Value2 = fPath
    ws.Cells(2, 1).Value2 = "MD"
    ws.Cells(2, 2).Value2 = "INC"
    ws.Cells(2, 3).Value2 = "AZI"
    ws.Cells(2, 4).Value2 = "TVD"
    ws.Cells(2, 5).Value2 = "NS"
    ws.Cells(2, 6).Value2 = "EW"
    For i = 0 To n - 1
        ws.Cells(i + 3, 1).Value2 = md(i)
        ws.Cells(i + 3, 2).Value2 = inc(i)
        ws.Cells(i + 3, 3).Value2 = azm(i)
        ws.Cells(i + 3, 4).Value2 = tvd(i)
        ws.Cells(i + 3, 5).Value2 = ns(i)
        ws.Cells(i + 3, 6).Value2 = ew(i)
    Next i
End Sub

Private Sub WriteFormationsSheet(ByVal fPath As String, ByVal n As Long, _
        ByRef md() As Double, ByRef tvd() As Double, _
        ByRef nm() As String, ByRef lith() As String)
    Dim ws As Worksheet, i As Long
    Set ws = EnsureHidden(PI_FORM_SHEET)
    ws.Cells(1, 1).Value2 = fPath
    ws.Cells(1, 2).Value2 = PI_LITHO_SRC
    ws.Cells(2, 1).Value2 = "MD"
    ws.Cells(2, 2).Value2 = "TVD"
    ws.Cells(2, 3).Value2 = "Name"
    ws.Cells(2, 4).Value2 = "Lithology"
    For i = 0 To n - 1
        ws.Cells(i + 3, 1).Value2 = md(i)
        ws.Cells(i + 3, 2).Value2 = tvd(i)
        ws.Cells(i + 3, 3).Value2 = nm(i)
        ws.Cells(i + 3, 4).Value2 = lith(i)
    Next i
End Sub

Private Function EnsureHidden(ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count))
        ws.name = shName
        ws.Visible = xlSheetVeryHidden
    Else
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        ws.Cells.Clear
    End If
    Set EnsureHidden = ws
End Function

' ---- lithology (AER 2019 Table of Formations + AGS WCSB Atlas) --------------

Private Function NormFormName(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    t = Replace(t, ".", "")
    t = Replace(t, ",", "")
    t = Replace(t, "-", " ")
    t = Replace(t, "_", " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Replace(t, " formation", "")
    t = Replace(t, " group", "")
    t = Replace(t, " member", "")
    If right$(t, 3) = " fm" Then t = Left$(t, Len(t) - 3)
    t = Trim$(t)
    If t = "garminia" Or t = "gaminie" Or t = "gramina" Then t = "graminia"
    If t = "wabamum" Then t = "wabamun"
    If t = "sws" Or t = "2ws" Or t = "2nd white specks" Or t = "second white speck" Then _
        t = "second white specks"
    If t = "bfs" Or t = "fish scales" Or t = "base fish scales" Or t = "bfss" Then _
        t = "base of fish scales"
    If t = "cardium sandstone" Then t = "cardium ss"
    If t = "viking sandstone" Then t = "viking ss"
    NormFormName = t
End Function

Private Function LithologyForNorm(ByVal key As String) As String
    ' Short wall labels. Source: AER Table of Formations 2019
    ' https://static.ags.aer.ca/files/document/INF/Table_of_Formations_2019.pdf
    ' and AGS Atlas of the Western Canada Sedimentary Basin.
    Select Case key
        Case "belly river": LithologyForNorm = "ss / silt / mdst, coal"
        Case "lea park": LithologyForNorm = "marine shale"
        Case "colorado": LithologyForNorm = "marine shale"
        Case "cardium": LithologyForNorm = "shale to ss / cgl"
        Case "cardium ss": LithologyForNorm = "sandstone"
        Case "second white specks": LithologyForNorm = "calc. marine shale"
        Case "base of fish scales": LithologyForNorm = "shale / silt marker"
        Case "viking", "viking ss": LithologyForNorm = "ss / silt / shale"
        Case "joli fou": LithologyForNorm = "marine shale"
        Case "mannville", "upper mannville": LithologyForNorm = "ss / shale / coal"
        Case "lower mannville": LithologyForNorm = "ss / shale"
        Case "ellerslie": LithologyForNorm = "sandstone"
        Case "glauconitic", "glauconitic ss": LithologyForNorm = "glauc. sandstone"
        Case "ostracod": LithologyForNorm = "lst / shale"
        Case "detrital": LithologyForNorm = "ss / cgl / shale"
        Case "rock creek": LithologyForNorm = "ss / siltstone"
        Case "poker chip": LithologyForNorm = "black shale"
        Case "nordegg": LithologyForNorm = "phos. lst / shale"
        Case "fernie": LithologyForNorm = "marine shale"
        Case "shunda": LithologyForNorm = "lst / dolomite"
        Case "pekisko": LithologyForNorm = "limestone"
        Case "banff": LithologyForNorm = "lst / shale"
        Case "exshaw": LithologyForNorm = "black shale"
        Case "wabamun": LithologyForNorm = "lst / dolomite"
        Case "graminia": LithologyForNorm = "silt / dolomite"
        Case "blueridge", "blue ridge": LithologyForNorm = "dolomite"
        Case "calmar": LithologyForNorm = "siltstone"
        Case "nisku": LithologyForNorm = "dol / limestone"
        Case "ireton": LithologyForNorm = "calc. shale"
        Case "duvernay": LithologyForNorm = "org. shale / source"
        Case "cooking lake": LithologyForNorm = "limestone"
        Case "leduc": LithologyForNorm = "reef dol / lst"
        Case "waterways": LithologyForNorm = "lst / shale"
        Case "slave point": LithologyForNorm = "limestone"
        Case "fort simpson": LithologyForNorm = "shale"
        Case "muskeg": LithologyForNorm = "anhydrite / dol"
        Case "keg river": LithologyForNorm = "dol / limestone"
        Case "granite wash": LithologyForNorm = "ss / cgl"
        Case "wapiabi": LithologyForNorm = "marine shale"
        Case "kaskapau": LithologyForNorm = "marine shale"
        Case "dunvegan": LithologyForNorm = "ss / shale"
        Case "spirit river": LithologyForNorm = "ss / shale / coal"
        Case "bluesky": LithologyForNorm = "sandstone"
        Case "gething": LithologyForNorm = "ss / coal"
        Case "cadomin": LithologyForNorm = "conglomerate"
        Case "nikanassin": LithologyForNorm = "ss / shale"
        Case "charlie lake": LithologyForNorm = "dol / anhydrite"
        Case "halfway": LithologyForNorm = "sandstone"
        Case "doig": LithologyForNorm = "silt / shale"
        Case "montney": LithologyForNorm = "silt / shale"
        Case "belloy": LithologyForNorm = "ss / dol"
        Case "debolt": LithologyForNorm = "lst / dolomite"
        Case "elkton": LithologyForNorm = "limestone"
        Case "shale": LithologyForNorm = "shale"
        Case Else
            If InStr(key, "ss") > 0 Or InStr(key, "sand") > 0 Then
                LithologyForNorm = "sandstone"
            ElseIf InStr(key, "shale") > 0 Then
                LithologyForNorm = "shale"
            Else
                LithologyForNorm = ""
            End If
    End Select
End Function




