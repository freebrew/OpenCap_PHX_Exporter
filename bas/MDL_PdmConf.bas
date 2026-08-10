Attribute VB_Name = "MDL_PdmConf"
Option Explicit

' Auto-sort Data!U32:V55 (PDM Conf | Rev/L) by lobe/stator label, then stage #.

Private Const SH_DATA As String = "Data"
Private Const PDM_FIRST As Long = 32
Private Const PDM_LAST As Long = 55
Private Const COL_CONF As Long = 21   ' U
Private Const COL_REVL As Long = 22   ' V

Public Sub SortPdmConfTable()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Dim n As Long
    Dim nInc As Long
    Dim i As Long, j As Long
    Dim conf() As String
    Dim revl() As Variant
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
    ReDim revl(1 To cap)
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
                revl(n) = CDbl(vTxt)
            Else
                revl(n) = ws.Cells(r, COL_REVL).Value
            End If
            ParsePdmConf cTxt, lobe(n), stator(n), stage(n)
        Else
            ' Incomplete pair — keep relative order, pack after completes
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

    ' Insertion sort complete pairs: lobe, stator, stage ascending
    For i = 2 To n
        tConf = conf(i): tRevl = revl(i)
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
            revl(j + 1) = revl(j)
            lobe(j + 1) = lobe(j)
            stator(j + 1) = stator(j)
            stage(j + 1) = stage(j)
            j = j - 1
        Loop
        conf(j + 1) = tConf
        revl(j + 1) = tRevl
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
        ws.Cells(r, COL_REVL).Value = revl(i)
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

' Called from Data Worksheet_Change on any edit in U32:V55 (add, edit, or clear).
' Complete pairs re-sort to the top; clears pack remaining rows upward.
Public Sub PdmConf_OnDataChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim rng As Range
    Dim area As Range

    If Target Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    Set rng = ws.Range(ws.Cells(PDM_FIRST, COL_CONF), ws.Cells(PDM_LAST, COL_REVL))
    Set area = Intersect(Target, rng)
    If area Is Nothing Then Exit Sub

    SortPdmConfTable
End Sub

' Parse "7/8 5.7" -> lobe=7, stator=8, stage=5.7
' Unparseable conf sorts last within its group (lobe/stator/stage = large sentinels).
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
