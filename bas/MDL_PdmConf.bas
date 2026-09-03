Attribute VB_Name = "MDL_PdmConf"
Option Explicit

' Auto-sort Data!U32:V55 (PDM Conf | Rev/L) by lobe/stator label, then stage #.
' Typing a config in column U fills Rev/L from the built-in catalog (does not
' overwrite a Rev/L the user just typed in V). Same lobe/stages can have
' different Rev/L by motor OD; OD is guessed from _OC_BHA when present.

Private Const SH_DATA As String = "Data"
Private Const SH_BHA As String = "_OC_BHA"
Private Const PDM_FIRST As Long = 32
Private Const PDM_LAST As Long = 55
Private Const COL_CONF As Long = 21   ' U
Private Const COL_REVL As Long = 22   ' V
Private Const STAGE_TOL As Double = 0.051

Private gCat() As String
Private gCatN As Long

Public Sub SortPdmConfTable()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Dim n As Long
    Dim nInc As Long
    Dim i As Long, j As Long
    Dim conf() As String
    Dim revL() As Variant
    Dim lobe() As Long
    Dim stator() As Long
    Dim stage() As Double
    Dim incConf() As String
    Dim incRevl() As Variant
    Dim r As Long
    Dim cTxt As String
    Dim vTxt As String
    Dim tConf As String
    Dim tRevl As Variant
    Dim tLobe As Long, tStator As Long
    Dim tStage As Double
    Dim swap As Boolean
    Dim cap As Long

    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    cap = PDM_LAST - PDM_FIRST + 1

    ReDim conf(1 To cap)
    ReDim revL(1 To cap)
    ReDim lobe(1 To cap)
    ReDim stator(1 To cap)
    ReDim stage(1 To cap)
    ReDim incConf(1 To cap)
    ReDim incRevl(1 To cap)

    n = 0
    nInc = 0
    For r = PDM_FIRST To PDM_LAST
        cTxt = Trim$(CStr(ws.Cells(r, COL_CONF).Value & ""))
        vTxt = Trim$(CStr(ws.Cells(r, COL_REVL).Value & ""))
        If Len(cTxt) = 0 And Len(vTxt) = 0 Then GoTo NextRead

        If Len(cTxt) > 0 And Len(vTxt) > 0 Then
            n = n + 1
            conf(n) = cTxt
            If IsNumeric(vTxt) Then
                revL(n) = CDbl(vTxt)
            Else
                revL(n) = ws.Cells(r, COL_REVL).Value
            End If
            ParsePdmConf cTxt, lobe(n), stator(n), stage(n)
        Else
            nInc = nInc + 1
            incConf(nInc) = cTxt
            If Len(vTxt) > 0 And IsNumeric(vTxt) Then
                incRevl(nInc) = CDbl(vTxt)
            ElseIf Len(vTxt) > 0 Then
                incRevl(nInc) = ws.Cells(r, COL_REVL).Value
            Else
                incRevl(nInc) = Empty
            End If
        End If
NextRead:
    Next r

    If n = 0 And nInc = 0 Then Exit Sub

    For i = 2 To n
        tConf = conf(i): tRevl = revL(i)
        tLobe = lobe(i): tStator = stator(i): tStage = stage(i)
        j = i - 1
        Do While j >= 1
            swap = False
            If lobe(j) > tLobe Then
                swap = True
            ElseIf lobe(j) = tLobe And stator(j) > tStator Then
                swap = True
            ElseIf lobe(j) = tLobe And stator(j) = tStator And stage(j) > tStage Then
                swap = True
            End If
            If Not swap Then Exit Do
            conf(j + 1) = conf(j)
            revL(j + 1) = revL(j)
            lobe(j + 1) = lobe(j)
            stator(j + 1) = stator(j)
            stage(j + 1) = stage(j)
            j = j - 1
        Loop
        conf(j + 1) = tConf
        revL(j + 1) = tRevl
        lobe(j + 1) = tLobe
        stator(j + 1) = tStator
        stage(j + 1) = tStage
    Next i

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo FailWrite

    ws.Range(ws.Cells(PDM_FIRST, COL_CONF), ws.Cells(PDM_LAST, COL_REVL)).ClearContents
    For i = 1 To n
        r = PDM_FIRST + i - 1
        ws.Cells(r, COL_CONF).Value = conf(i)
        ws.Cells(r, COL_REVL).Value = revL(i)
    Next i
    For i = 1 To nInc
        r = PDM_FIRST + n + i - 1
        If r > PDM_LAST Then Exit For
        If Len(incConf(i)) > 0 Then ws.Cells(r, COL_CONF).Value = incConf(i)
        If Not IsEmpty(incRevl(i)) Then ws.Cells(r, COL_REVL).Value = incRevl(i)
    Next i

    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub

FailWrite:
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
    Err.Raise Err.Number, "SortPdmConfTable", Err.Description
End Sub

' Called from Data Worksheet_Change on any edit in U32:V55.
Public Sub PdmConf_OnDataChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim rng As Range
    Dim area As Range
    Dim confArea As Range
    Dim c As Range
    Dim prevEvents As Boolean
    Dim wasProt As Boolean
    Dim odHint As Double

    On Error GoTo ErrHandler
    If Target Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    Set rng = ws.Range(ws.Cells(PDM_FIRST, COL_CONF), ws.Cells(PDM_LAST, COL_REVL))
    Set area = Intersect(Target, rng)
    If area Is Nothing Then Exit Sub

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ws)

    Set confArea = Intersect(Target, ws.Range(ws.Cells(PDM_FIRST, COL_CONF), ws.Cells(PDM_LAST, COL_CONF)))
    If Not confArea Is Nothing Then
        odHint = GuessMotorOdIn()
        For Each c In confArea.Cells
            FillRevLForRow ws, c.Row, True, odHint
        Next c
    End If

    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
    SortPdmConfTable
    Exit Sub

ErrHandler:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
End Sub

' Excel UDF: =IFERROR(PdmLookupRevL(U32),"")
' Optional motorOdIn (inches) picks the matching size when a label is shared.
Public Function PdmLookupRevL(ByVal confText As Variant, Optional ByVal motorOdIn As Variant) As Variant
    Dim lobe As Long, stator As Long
    Dim stage As Double
    Dim odHint As Double
    Dim v As Variant

    PdmLookupRevL = ""
    If isError(confText) Then Exit Function
    ParsePdmConf CStr(confText & ""), lobe, stator, stage
    If lobe >= 9999 Then Exit Function

    If Not IsMissing(motorOdIn) Then
        If IsNumeric(motorOdIn) Then odHint = CDbl(motorOdIn)
    End If
    If odHint <= 0# Then odHint = GuessMotorOdIn()

    v = ResolveRevL(lobe, stator, stage, odHint)
    If Not IsEmpty(v) Then PdmLookupRevL = CDbl(v)
End Function

Private Sub FillRevLForRow(ByVal ws As Worksheet, ByVal r As Long, _
        ByVal overwrite As Boolean, ByVal odHint As Double)
    Dim cTxt As String
    Dim vTxt As String
    Dim lobe As Long, stator As Long
    Dim stage As Double
    Dim v As Variant

    If r < PDM_FIRST Or r > PDM_LAST Then Exit Sub
    cTxt = Trim$(CStr(ws.Cells(r, COL_CONF).Value & ""))
    If Len(cTxt) = 0 Then Exit Sub
    vTxt = Trim$(CStr(ws.Cells(r, COL_REVL).Value & ""))
    If (Not overwrite) And Len(vTxt) > 0 Then Exit Sub

    ParsePdmConf cTxt, lobe, stator, stage
    If lobe >= 9999 Then Exit Sub

    v = ResolveRevL(lobe, stator, stage, odHint)
    If IsEmpty(v) Then Exit Sub
    ws.Cells(r, COL_REVL).numberFormat = "0.000"
    ws.Cells(r, COL_REVL).Value = CDbl(v)
End Sub

Private Function ResolveRevL(ByVal lobe As Long, ByVal stator As Long, _
        ByVal stage As Double, ByVal odHint As Double) As Variant
    Dim i As Long
    Dim parts() As String
    Dim od As Double
    Dim revL As Double
    Dim prefer As Boolean
    Dim score As Double
    Dim bestScore As Double
    Dim bestRev As Variant
    Dim confLobe As Long, confStator As Long
    Dim confStage As Double

    ResolveRevL = Empty
    EnsureCatalog
    bestScore = -1E+30
    bestRev = Empty

    For i = 1 To gCatN
        parts = Split(gCat(i), "|")
        If UBound(parts) < 3 Then GoTo NextCat
        If Not IsNumeric(parts(0)) Then GoTo NextCat
        If Not IsNumeric(parts(2)) Then GoTo NextCat
        od = CDbl(parts(0))
        revL = CDbl(parts(2))
        prefer = (Trim$(parts(3)) = "1")
        ParsePdmConf Trim$(parts(1)), confLobe, confStator, confStage
        If confLobe <> lobe Or confStator <> stator Then GoTo NextCat
        If Abs(confStage - stage) > STAGE_TOL Then GoTo NextCat

        If odHint > 0# Then
            score = 1000# - Abs(od - odHint) * 25#
            If prefer Then score = score + 1#
        ElseIf prefer Then
            score = 200#
        ElseIf od >= 4.6 And od <= 5.6 Then
            score = 80# - Abs(od - 5#)
        Else
            score = 20# - od * 0.1
        End If

        If score > bestScore Then
            bestScore = score
            bestRev = revL
        End If
NextCat:
    Next i

    ResolveRevL = bestRev
End Function

Private Function GuessMotorOdIn() As Double
    Dim ws As Worksheet
    Dim lastR As Long, r As Long
    Dim desc As String
    Dim od As Double

    GuessMotorOdIn = 0#
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then lastR = 40
    If lastR > 200 Then lastR = 200

    For r = lastR To 2 Step -1
        desc = CStr(ws.Cells(r, 2).Value & "") & " " & CStr(ws.Cells(r, 3).Value & "")
        If InStr(1, desc, "motor", vbTextCompare) = 0 Then GoTo NextBha
        od = OdFromText(desc)
        If od > 0# Then
            GuessMotorOdIn = od
            Exit Function
        End If
NextBha:
    Next r
End Function

Private Function OdFromText(ByVal s As String) As Double
    Dim t As String
    t = LCase$(s)
    OdFromText = 0#
    If InStr(t, "9.62") > 0 Or InStr(t, "9 5/8") > 0 Or InStr(t, "9-5/8") > 0 Then
        OdFromText = 9.62
    ElseIf InStr(t, "8.25") > 0 Or InStr(t, "8 1/4") > 0 Then
        OdFromText = 8.25
    ElseIf InStr(t, "8.00") > 0 Or InStr(t, "8 in") > 0 Or InStr(t, "8""") > 0 _
            Or InStr(t, " 8 ") > 0 And InStr(t, "motor") > 0 Then
        OdFromText = 8#
    ElseIf InStr(t, "7.25") > 0 Or InStr(t, "7 1/4") > 0 Or InStr(t, "7-1/4") > 0 Then
        OdFromText = 7.25
    ElseIf InStr(t, "6.75") > 0 Or InStr(t, "6 3/4") > 0 Or InStr(t, "6-3/4") > 0 Then
        OdFromText = 6.75
    ElseIf InStr(t, "6.62") > 0 Or InStr(t, "6 5/8") > 0 Or InStr(t, "6-5/8") > 0 Then
        OdFromText = 6.625
    ElseIf InStr(t, "6.5") > 0 Or InStr(t, "6 1/2") > 0 Or InStr(t, "6-1/2") > 0 Then
        OdFromText = 6.5
    ElseIf InStr(t, "5.5") > 0 Or InStr(t, "5 1/2") > 0 Then
        OdFromText = 5.5
    ElseIf InStr(t, "5.25") > 0 Or InStr(t, "5 1/4") > 0 Then
        OdFromText = 5.25
    ElseIf InStr(t, "4.75") > 0 Or InStr(t, "4 3/4") > 0 Or InStr(t, "4-3/4") > 0 Then
        OdFromText = 4.75
    ElseIf InStr(t, "5.00") > 0 Or InStr(t, "5 in") > 0 Or InStr(t, "5""") > 0 Then
        OdFromText = 5#
    End If
End Function

Private Sub EnsureCatalog()
    If gCatN > 0 Then Exit Sub
    ' od|conf|revL|prefer  prefer=1 is the Data-sheet number (wins when OD unknown)
    ' Published extras: SLB Dyna-Drill / Discovery DHS / Bico / MDS (gal/3.7854)

    AddCat "5|5/6 6.7|0.166|1"
    AddCat "5|5/6 8.3|0.250|1"
    AddCat "5|6/7 6.4|0.210|1"
    AddCat "5|6/7 8.0|0.210|1"
    AddCat "5|6/7 8.8|0.180|1"
    AddCat "5|6/7 10.5|0.193|1"
    AddCat "5|6/7 11.6|0.159|1"
    AddCat "5|6/7 11.7|0.238|1"
    AddCat "5|7/8 3.7|0.093|1"
    AddCat "5|7/8 3.8|0.138|1"
    AddCat "6.75|7/8 5.0|0.076|1"
    AddCat "5|7/8 5.7|0.146|1"
    AddCat "8|7/8 7.0|0.047|1"
    AddCat "6.75|7/8 7.5|0.063|1"
    AddCat "5.5|7/8 7.6|0.114|1"
    AddCat "8|8/9 4.0|0.024|1"

    AddCat "4.75|5/6 8.3|0.264|0"
    AddCat "5|5/6 8.3|0.264|0"
    AddCat "5|5/6 5.2|0.166|0"
    AddCat "5.25|5/6 8.4|0.185|0"
    AddCat "6.75|5/6 8.1|0.108|0"
    AddCat "7|5/6 9.4|0.106|0"
    AddCat "7|5/6 9.5|0.090|0"

    AddCat "5|6/7 7.0|0.214|0"
    AddCat "5|6/7 7.8|0.159|0"
    AddCat "5|6/7 8.0|0.214|0"
    AddCat "5|6/7 8.8|0.174|0"
    AddCat "5|6/7 10.0|0.230|0"
    AddCat "5|6/7 10.9|0.209|0"
    AddCat "5|6/7 11.3|0.198|0"
    AddCat "6.75|6/7 5.0|0.077|0"
    AddCat "6.75|6/7 7.6|0.087|0"
    AddCat "6.75|6/7 7.8|0.077|0"
    AddCat "6.75|6/7 8.1|0.103|0"

    AddCat "4.75|7/8 2.6|0.069|0"
    AddCat "5|7/8 2.6|0.069|0"
    AddCat "4.75|7/8 3.7|0.097|0"
    AddCat "5|7/8 4.1|0.069|0"
    AddCat "4.75|7/8 5.0|0.170|0"
    AddCat "5|7/8 5.0|0.170|0"
    AddCat "6.75|7/8 3.3|0.037|0"
    AddCat "6.75|7/8 5.7|0.064|0"
    AddCat "6.75|7/8 6.0|0.076|0"
    AddCat "6.75|7/8 6.4|0.076|0"
    AddCat "5|7/8 8.2|0.185|0"
    AddCat "5|7/8 8.3|0.127|0"
    AddCat "5|7/8 8.4|0.185|0"
    AddCat "5|7/8 10.6|0.226|0"
    AddCat "6.625|7/8 6.9|0.066|0"
    AddCat "8|7/8 3.4|0.023|0"
    AddCat "8|7/8 4.0|0.042|0"

    AddCat "5|4/5 6.3|0.269|0"
    AddCat "5|4/5 10.3|0.264|0"
    AddCat "6.75|4/5 7.0|0.131|0"
    AddCat "6.5|8/9 4.0|0.076|0"
End Sub

Private Sub AddCat(ByVal line As String)
    gCatN = gCatN + 1
    ReDim Preserve gCat(1 To gCatN)
    gCat(gCatN) = line
End Sub

Private Sub ParsePdmConf(ByVal confText As String, _
        ByRef lobe As Long, ByRef stator As Long, ByRef stage As Double)
    Dim s As String
    Dim slashPos As Long
    Dim spPos As Long
    Dim leftPart As String
    Dim rightPart As String
    Dim stageTxt As String

    lobe = 9999
    stator = 9999
    stage = 1E+30

    s = Trim$(confText)
    If Len(s) = 0 Then Exit Sub

    slashPos = InStr(1, s, "/")
    If slashPos <= 1 Then Exit Sub

    leftPart = Trim$(Left$(s, slashPos - 1))
    rightPart = Trim$(mid$(s, slashPos + 1))
    If Len(leftPart) = 0 Or Len(rightPart) = 0 Then Exit Sub

    spPos = InStr(1, rightPart, " ")
    If spPos > 0 Then
        stageTxt = Trim$(mid$(rightPart, spPos + 1))
        rightPart = Trim$(Left$(rightPart, spPos - 1))
    Else
        stageTxt = ""
    End If

    If Not IsNumeric(leftPart) Or Not IsNumeric(rightPart) Then Exit Sub
    lobe = CLng(val(leftPart))
    stator = CLng(val(rightPart))
    If Len(stageTxt) > 0 And IsNumeric(stageTxt) Then
        stage = CDbl(stageTxt)
    Else
        stage = 0#
    End If
End Sub



