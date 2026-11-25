Attribute VB_Name = "MDL_ContDI"
Option Explicit

' Continuous DI / BURR tab: one live grid (rows 8:44) tied to the newest
' Slidesheet survey. "New Set" clears the typed readings and re-pulls the
' tie-on; nothing is archived and there is no course paging. ContDI_Data
' (very hidden) only carries diagnostic lookups and projection seeds.
' Layout and patched rates live here so Rebuild keeps them.

Public Const SH_CONTDI As String = "Cont DI"
Public Const SH_CONTDI_DATA As String = "ContDI_Data"

Private Const FIRST_DATA As Long = 8
Private Const LAST_DATA As Long = 44

Private Const SPN_NAME As String = "ContDI_SpinCtl"     ' legacy course spinner, removed on sight
Private Const BTN_NEW As String = "ContDI_NewCourse"     ' shape name kept for old workbooks; caption "New Set"
Private Const BTN_PULL As String = "ContDI_PullSurvey"
Private Const CAP_NEW As String = "New Set"
Private Const CAP_PULL As String = "Last Survey"
Private Const SS_SHEET As String = "Slidesheet"
Private Const SS_FIRST As Long = 13
Private Const SS_LAST As Long = 505
Private Const PLANSEC_SHEET As String = "_OC_PlanSec"
Private Const TARGET_PICK_CELL As String = "P2"

' Column A (the "Last MWD Survey" label) is not on the tab. Rate block H:Q,
' left to right: MWD DLS | MWD Average | Between sensors | PD DLS | PD Average.
' Targets R1:S4 and rates T1:U4 sit over the bit-projection columns. Row-3
' averages sit on the MWD (I:L) and PD (P:Q) average columns. The Bit to
' Sensor Offsets box stays on L1:O2.
Private Const COL_MWD_DLS As Long = 8     ' H
Private Const COL_MWD_BURC As Long = 9    ' I
Private Const COL_MWD_TRC As Long = 10    ' J
Private Const COL_MWD_BURS As Long = 11   ' K
Private Const COL_MWD_TRS As Long = 12    ' L
Private Const COL_SENS_DLS As Long = 13   ' M  3D DLS MWD(C,D) <-> PD(F,G)
Private Const COL_SENS_BUR As Long = 14   ' N  signed BUR, PD minus MWD
Private Const COL_PD_DLS As Long = 15     ' O
Private Const COL_PD_BUR As Long = 16     ' P
Private Const COL_PD_TR As Long = 17      ' Q
Private Const COL_HOLE As Long = 18       ' R
Private Const COL_INC_BIT As Long = 19    ' S
Private Const COL_AZ_BIT As Long = 20     ' T
Private Const COL_TVD_BIT As Long = 21    ' U
Private Const COL_ROP As Long = 22        ' V
Private Const HDR_SENS_DLS As String = "Sens DLS"
Private Const HDR_PD_DLS As String = "PD DLS"

' Everything the user types on the tab, held in memory across a repaint.
Private Type ContDIInputs
    tieOn As Variant      ' A2:D2
    pdOff As Variant      ' M2
    diOff As Variant      ' O2
    targetName As Variant ' P2
    targets As Variant    ' S1:S4
    bit As Variant        ' A8:A44
    mwd As Variant        ' C8:D44
    pD As Variant         ' F8:G44
    rop As Variant        ' V8:V44
End Type

' ================================================================================
'  PUBLIC
' ================================================================================

' Full repaint. Typed readings, tie-on, offsets and targets survive the repaint.
Public Sub BuildContDITab()
    Dim ws As Worksheet
    Dim snap As ContDIInputs
    Dim ev As Boolean
    Dim su As Boolean
    Dim wp As Boolean

    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    su = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    EnsureContDIData
    Set ws = EnsureContDISheet()
    wp = SheetUnprotectForVba(ws)
    ' Bring an old-layout tab up to date first so the snapshot reads the
    ' right cells (ROP at V, targets at S1:S4).
    DropLabelColumn ws
    EnsureInterSensorColumns ws
    EnsureHeaderShift ws
    snap = SnapshotInputs(ws)
    PaintContDI ws
    AddContDIControls ws
    ApplyRowFormulas ws
    SeedProjectionHelpers ws
    ApplyLookupsAndRates ws
    RestoreInputs ws, snap
    ContDI_FollowSlideTarget
    ApplyContDILock ws
    SheetReprotectAfterVba ws, True

    Application.EnableEvents = ev
    Application.ScreenUpdating = su
    Exit Sub

ErrHandler:
    Application.EnableEvents = ev
    Application.ScreenUpdating = su
    MsgBox "BuildContDITab: " & Err.Description, vbCritical
End Sub

' "New Set" button: wipe the typed readings (A/C/D/F/G/V 8:44) and re-pull the
' newest green Slidesheet survey into A2:D2. Offsets (M2/O2) stay. The yellow
' target (S1:S4) is the slidesheet's active station, not a typed setup value.
Public Sub ContDI_NewSet()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim ev As Boolean

    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    Application.EnableEvents = False

    Set ws = ContDISheet()
    wasProt = SheetUnprotectForVba(ws)
    ClearInputs ws
    WriteLastSurveyTieOn ws, True
    EnsureContDIControls ws
    ApplyContDILock ws
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    Exit Sub

ErrHandler:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    MsgBox "ContDI_NewSet: " & Err.Description, vbCritical
End Sub

' Legacy entry point: buttons painted as "New Course" still point here.
Public Sub ContDI_NewCourse()
    ContDI_NewSet
End Sub

' Button / macro: refill the tie-on survey (A2:D2) from the Slidesheet,
' clear typed A/C/D/F/G/V from row 8 down, and keep the last entered grid
' row as row 8.
Public Sub ContDI_PullLastSurvey()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim ev As Boolean

    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    Application.EnableEvents = False

    Set ws = ContDISheet()
    wasProt = SheetUnprotectForVba(ws)
    PromoteLastGridRow ws
    WriteLastSurveyTieOn ws, True
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    Exit Sub

ErrHandler:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    MsgBox "ContDI_PullLastSurvey: " & Err.Description, vbCritical
End Sub

' Called from Workbook_SheetActivate: bring the buttons up to date on tabs
' painted by an older build (shapes only; no cells, no repaint).
Public Sub EnsureContDIPullButton()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONTDI)
    If ws Is Nothing Then Exit Sub
    wasProt = SheetUnprotectForVba(ws)
    If Not ControlsCurrent(ws) Then EnsureContDIControls ws
    SheetReprotectAfterVba ws, wasProt
End Sub

' The yellow target follows the slidesheet. A typed edit in that block is put
' back onto the slidesheet's active station.
Public Function ContDI_ApplyTargetSelection(ByVal ws As Worksheet, _
                                             ByVal Target As Range) As Boolean
    Dim hit As Range
    On Error GoTo Fail
    Set hit = Application.Intersect(Target, ws.Range("P2:Q2,S1:S4"))
    If hit Is Nothing Then Exit Function
    ContDI_FollowSlideTarget
    ContDI_ApplyTargetSelection = True
    Exit Function
Fail:
    ContDI_ApplyTargetSelection = False
End Function

' Target S1:S4 / P2 is the slidesheet's active target (AO) on the as-drilled
' survey row this tab ties on to (A2), so it advances through the named
' targets exactly as the slidesheet does. Rates U1:U4 are the same plan-point
' arc as Slidesheet AR:AT, run from this tab's bit (ContDI_Data T1/T4/U4/V1)
' toward S1. Writes only cells that differ.
Public Sub ContDI_FollowSlideTarget()
    Dim ws As Worksheet
    Dim h As Worksheet
    Dim wasProt As Boolean
    Dim wasProtH As Boolean
    Dim prevEvents As Boolean
    Dim tgts As String
    Dim aim As String
    Dim aimH As String
    Dim fX2 As String
    Dim aimCells As Variant
    Dim fmls As Variant
    Dim i As Long
    Dim dirty As Boolean

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_CONTDI)
    If ws Is Nothing Then Exit Sub
    Set h = ContDIDataSheet()
    If h Is Nothing Then Exit Sub

    tgts = "ProjTargets_MD,ProjTargets_INC,ProjTargets_AZM,ProjTargets_TVD"
    aim = "A2,B2,C2,ContDI_Data!T1,ContDI_Data!T4,ContDI_Data!U4,ContDI_Data!V1," & tgts & ",S1"
    aimH = "'" & SH_CONTDI & "'!A2,'" & SH_CONTDI & "'!B2,'" & SH_CONTDI & "'!C2,T1,T4,U4,V1," & _
           tgts & ",'" & SH_CONTDI & "'!S1"

    aimCells = Array("P2", "S1", "S2", "S3", "S4", "U1", "U2", "U3", "U4")
    fmls = Array( _
        "=IFERROR(INDEX(ProjTargets_NAME,MATCH(S1,ProjTargets_MD,0)),"""")", _
        "=IFERROR(INDEX(Slidesheet!$AO$13:$AO$505,MATCH(A2,Slidesheet!$E$13:$E$505,0)),"""")", _
        "=IFERROR(INDEX(ProjTargets_INC,MATCH(S1,ProjTargets_MD,0)),"""")", _
        "=IFERROR(INDEX(ProjTargets_AZM,MATCH(S1,ProjTargets_MD,0)),"""")", _
        "=IFERROR(INDEX(ProjTargets_TVD,MATCH(S1,ProjTargets_MD,0)),"""")", _
        "=IF(S1="""","""",IFERROR(ROUND(ProjContDiAim(""BRR""," & aim & "),2),""""))", _
        "=IF(S1="""","""",IFERROR(ROUND(ProjContDiAim(""TRR""," & aim & "),2),""""))", _
        "=IF(S1="""","""",IFERROR(ROUND(ProjContDiAim(""DLR""," & aim & "),2),""""))", _
        "=IF(S1="""","""",IFERROR(ProjContDiAim(""TF""," & aim & "),""""))")
    fX2 = "=IF('" & SH_CONTDI & "'!S1="""","""",IFERROR(ROUND(ProjContDiAim(""TRRS""," & aimH & "),2),""""))"

    dirty = (StrComp(CStr(h.Range("X2").formula), fX2, vbTextCompare) <> 0)
    For i = 0 To UBound(aimCells)
        If StrComp(CStr(ws.Range(aimCells(i)).formula), CStr(fmls(i)), vbTextCompare) <> 0 Then dirty = True
    Next i
    If Not dirty Then Exit Sub

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ws)
    wasProtH = SheetUnprotectForVba(h)
    ws.Range("P1").value = "Active Target"
    On Error Resume Next
    ws.Range("P2").Validation.Delete
    On Error GoTo Fail
    For i = 0 To UBound(aimCells)
        If StrComp(CStr(ws.Range(aimCells(i)).formula), CStr(fmls(i)), vbTextCompare) <> 0 Then
            ws.Range(aimCells(i)).formula = fmls(i)
        End If
    Next i
    ws.Range("P2").numberFormat = "General"
    ws.Range("S1:S4").numberFormat = "0.00"
    h.Range("X2").formula = fX2
    SheetReprotectAfterVba h, wasProtH
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    If Not h Is Nothing Then SheetReprotectAfterVba h, wasProtH
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
End Sub

' ================================================================================
'  LAST MWD SURVEY (Slidesheet -> A2:D2)
' ================================================================================

' The Slidesheet survey-entry fill: light green RGB(204,255,204) (Excel
' ColorIndex 35). A row is a GOOD survey only when Inc (F) and Azm (G) are
' both numeric AND both carry exactly this green. Any other fill - the bright
' yellow the DD uses for plan-ahead / projected stations, notes, highlights -
' means "not a survey" and is never counted, whatever its depth.
Private Function GoodSurveyFill() As Long
    GoodSurveyFill = RGB(204, 255, 204)
End Function

Private Function FillLooksGreen(ByVal c As Range) As Boolean
    Dim fill As Long
    Dim idx As Long
    FillLooksGreen = False
    On Error Resume Next
    idx = CLng(c.Interior.ColorIndex)
    fill = CLng(c.Interior.Color)
    If idx = 35 Or fill = GoodSurveyFill() Then
        FillLooksGreen = True
        Exit Function
    End If
    idx = CLng(c.DisplayFormat.Interior.ColorIndex)
    fill = CLng(c.DisplayFormat.Interior.Color)
    On Error GoTo 0
    FillLooksGreen = (idx = 35 Or fill = GoodSurveyFill())
End Function

Private Function IsGoodSurveyCell(ByVal c As Range) As Boolean
    IsGoodSurveyCell = False
    If Not IsNumberValue(c.Value2) Then Exit Function
    IsGoodSurveyCell = FillLooksGreen(c)
End Function

' True when survey MD (E) is a number and Inc (F) / Azm (G) are numeric and
' painted survey-green. Any other fill is a user projection / note, and an
' Inc/Azm with no MD is not a station (it made bit depth 0 for the target window).
Private Function IsGoodSurveyRowCells(ByVal ss As Worksheet, ByVal r As Long) As Boolean
    IsGoodSurveyRowCells = False
    If Not IsNumberValue(ss.Cells(r, "E").Value2) Then Exit Function
    IsGoodSurveyRowCells = IsGoodSurveyCell(ss.Cells(r, "F")) _
                        And IsGoodSurveyCell(ss.Cells(r, "G"))
End Function

Public Function IsSlidesheetGoodSurveyRow(ByVal r As Long) As Boolean
    Dim ss As Worksheet
    IsSlidesheetGoodSurveyRow = False
    On Error GoTo Done
    If r < SS_FIRST Or r > SS_LAST Then Exit Function
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    IsSlidesheetGoodSurveyRow = IsGoodSurveyRowCells(ss, r)
Done:
End Function

Public Function IsSlidesheetGoodSurveyRowOn(ByVal ss As Worksheet, ByVal r As Long) As Boolean
    IsSlidesheetGoodSurveyRowOn = False
    If ss Is Nothing Then Exit Function
    If r < SS_FIRST Or r > SS_LAST Then Exit Function
    IsSlidesheetGoodSurveyRowOn = IsGoodSurveyRowCells(ss, r)
End Function

' Deepest good-survey row on the Slidesheet (13:505); 0 = none.
Public Function LastSlidesheetSurveyRow() As Long
    Dim ss As Worksheet
    Dim r As Long

    LastSlidesheetSurveyRow = 0
    On Error GoTo Done
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    For r = SS_LAST To SS_FIRST Step -1
        If IsGoodSurveyRowCells(ss, r) Then
            LastSlidesheetSurveyRow = r
            Exit Function
        End If
    Next r
Done:
End Function

' A2 = survey MD (E), B2 = Inc (F), C2 = Azm (G), D2 = TVD (H).
' Stand length (Slidesheet C, next survey course) is stored on ContDI_Data!X1
' for ProjBurr's short-MD floor - not shown on the tab.
' Caller owns protection / events. verbose:=True reports when nothing qualifies.
Private Sub WriteLastSurveyTieOn(ByVal ws As Worksheet, ByVal verbose As Boolean)
    Dim ss As Worksheet
    Dim r As Long
    Dim nextC As Variant

    r = LastSlidesheetSurveyRow()
    If r = 0 Then
        If verbose Then
            MsgBox "No Slidesheet row has Inc and Azm in the green survey fill.", vbExclamation
        End If
        Exit Sub
    End If
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    ws.Range("A2").value = ss.Cells(r, "E").Value2
    ws.Range("B2").value = ss.Cells(r, "F").Value2
    ws.Range("C2").value = ss.Cells(r, "G").Value2
    ws.Range("D2").value = ss.Cells(r, "H").Value2

    nextC = Empty
    If r + 1 <= SS_LAST Then
        If IsNumberValue(ss.Cells(r + 1, "C").Value2) Then nextC = ss.Cells(r + 1, "C").Value2
    End If
    If IsEmpty(nextC) Then
        If IsNumberValue(ss.Cells(r, "C").Value2) Then nextC = ss.Cells(r, "C").Value2
    End If
    If Not IsEmpty(nextC) Then ContDIDataSheet().Range("X1").value = Round(CDbl(nextC), 2)
    SeedProjectionHelpers ws, r
End Sub

' ContDI_Data W1 = Slidesheet bit MD (D), W2/W3 = Inc@Bit / Az@Bit (W/X) for
' the tie-on row. Does not touch A2:D2. Lets BRR start from the same projected
' bit as Slidesheet Y while the Cont DI grid is still empty.
Private Sub SeedProjectionHelpers(ByVal ws As Worksheet, Optional ByVal ssRow As Long = 0)
    Dim ss As Worksheet
    Dim r As Long
    Dim md As Double
    Dim i As Long

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    r = ssRow
    If r < SS_FIRST Then
        r = 0
        If IsNumberValue(ws.Range("A2").Value2) Then
            md = CDbl(ws.Range("A2").Value2)
            For i = SS_LAST To SS_FIRST Step -1
                If IsNumberValue(ss.Cells(i, "E").Value2) Then
                    If Abs(CDbl(ss.Cells(i, "E").Value2) - md) < 0.05 Then
                        r = i
                        Exit For
                    End If
                End If
            Next i
        End If
        If r < SS_FIRST Then r = LastSlidesheetSurveyRow()
    End If
    If r < SS_FIRST Then Exit Sub
    With ContDIDataSheet()
        If IsNumberValue(ss.Cells(r, "D").Value2) Then .Range("W1").value = ss.Cells(r, "D").Value2
        If IsNumberValue(ss.Cells(r, "W").Value2) Then .Range("W2").value = ss.Cells(r, "W").Value2
        If IsNumberValue(ss.Cells(r, "X").Value2) Then .Range("W3").value = ss.Cells(r, "X").Value2
        If r + 1 <= SS_LAST And IsNumberValue(ss.Cells(r + 1, "C").Value2) Then
            .Range("X1").value = Round(CDbl(ss.Cells(r + 1, "C").Value2), 2)
        ElseIf IsNumberValue(ss.Cells(r, "C").Value2) Then
            .Range("X1").value = Round(CDbl(ss.Cells(r, "C").Value2), 2)
        End If
    End With
    Err.Clear
End Sub

Private Function IsNumberValue(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            IsNumberValue = True
        Case Else
            IsNumberValue = False
    End Select
End Function

' ================================================================================
'  SHEETS
' ================================================================================

Public Function ContDISheet() As Worksheet
    Set ContDISheet = ThisWorkbook.Worksheets(SH_CONTDI)
End Function

Public Function ContDIDataSheet() As Worksheet
    Set ContDIDataSheet = ThisWorkbook.Worksheets(SH_CONTDI_DATA)
End Function

Private Function EnsureContDISheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONTDI)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SH_CONTDI
    End If
    Set EnsureContDISheet = ws
End Function

' Very-hidden helper sheet: T1:V7 last-value lookups, W1:W3 projection seeds,
' X1 stand length. Nothing else is read from it.
Private Function EnsureContDIData() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONTDI_DATA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SH_CONTDI_DATA
        ws.Range("A1").value = "ContDI_Data"
        ws.Range("D1").value = "helpers only: T:V lookups, W1:W3 projection seeds, X1 stand length"
    End If
    ws.Visible = xlSheetVeryHidden
    Set EnsureContDIData = ws
End Function

' ================================================================================
'  INPUTS (snapshot / restore / clear)
' ================================================================================

Private Function SnapshotInputs(ByVal ws As Worksheet) As ContDIInputs
    Dim s As ContDIInputs
    s.tieOn = ws.Range("A2:D2").value
    s.pdOff = ws.Range("M2").value
    s.diOff = ws.Range("O2").value
    s.targetName = ws.Range(TARGET_PICK_CELL).value
    s.targets = ws.Range("S1:S4").value
    s.bit = ws.Range("A8:A44").value
    s.mwd = ws.Range("C8:D44").value
    s.pD = ws.Range("F8:G44").value
    s.rop = ws.Range("V8:V44").value
    SnapshotInputs = s
End Function

Private Sub RestoreInputs(ByVal ws As Worksheet, ByRef s As ContDIInputs)
    ws.Range("A2:D2").value = s.tieOn
    If Not IsEmpty(s.pdOff) Then ws.Range("M2").value = s.pdOff
    If Not IsEmpty(s.diOff) Then ws.Range("O2").value = s.diOff
    If Not IsEmpty(s.targetName) Then ws.Range(TARGET_PICK_CELL).value = s.targetName
    ws.Range("S1:S4").value = s.targets
    ws.Range("A8:A44").value = s.bit
    ws.Range("C8:D44").value = s.mwd
    ws.Range("F8:G44").value = s.pD
    ws.Range("V8:V44").value = s.rop
End Sub

Private Sub ClearInputs(ByVal ws As Worksheet)
    ws.Range("A8:A44").ClearContents
    ws.Range("C8:D44").ClearContents
    ws.Range("F8:G44").ClearContents
    ws.Range("V8:V44").ClearContents
End Sub

' Last Survey: the bottom-most row with anything in A/C/D/F/G becomes row 8,
' including that row's ROP. A/C/D/F/G/V are cleared from row 8 down first.
Private Sub PromoteLastGridRow(ByVal ws As Worksheet)
    Dim r As Long
    Dim lastR As Long
    Dim b As Variant
    Dim d As Variant
    Dim e As Variant
    Dim g As Variant
    Dim h As Variant
    Dim w As Variant

    lastR = 0
    For r = LAST_DATA To FIRST_DATA Step -1
        If GridRowHasEntry(ws, r) Then
            lastR = r
            Exit For
        End If
    Next r
    If lastR > 0 Then
        b = ws.Cells(lastR, "A").value
        d = ws.Cells(lastR, "C").value
        e = ws.Cells(lastR, "D").value
        g = ws.Cells(lastR, "F").value
        h = ws.Cells(lastR, "G").value
        w = ws.Cells(lastR, "V").value
    End If
    ws.Range("A8:A44").ClearContents
    ws.Range("C8:D44").ClearContents
    ws.Range("F8:G44").ClearContents
    ws.Range("V8:V44").ClearContents
    If lastR > 0 Then
        ws.Cells(FIRST_DATA, "A").value = b
        ws.Cells(FIRST_DATA, "C").value = d
        ws.Cells(FIRST_DATA, "D").value = e
        ws.Cells(FIRST_DATA, "F").value = g
        ws.Cells(FIRST_DATA, "G").value = h
        ws.Cells(FIRST_DATA, "V").value = w
    End If
End Sub

Private Function GridRowHasEntry(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    GridRowHasEntry = CellHasEntry(ws.Cells(r, "A")) Or CellHasEntry(ws.Cells(r, "C")) _
                   Or CellHasEntry(ws.Cells(r, "D")) Or CellHasEntry(ws.Cells(r, "F")) _
                   Or CellHasEntry(ws.Cells(r, "G"))
End Function

Private Function CellHasEntry(ByVal c As Range) As Boolean
    Dim v As Variant
    v = c.Value2
    CellHasEntry = False
    If isError(v) Then
        CellHasEntry = True
    ElseIf IsEmpty(v) Then
        CellHasEntry = False
    ElseIf VarType(v) = vbString Then
        CellHasEntry = (Len(Trim$(CStr(v))) > 0)
    Else
        CellHasEntry = True
    End If
End Function

' ================================================================================
'  LOCK / CONTROLS
' ================================================================================

' Inputs (tie-on, offsets, A/C/D/F/G/V 8:44) unlocked. The target block
' (P2 name, S1:S4) stays locked: those cells are the slidesheet aim.
Private Sub ApplyContDILock(ByVal ws As Worksheet)
    ws.Cells.Locked = True
    ws.Range("A2:D2").Locked = False
    ws.Range("M2").Locked = False
    ws.Range("O2").Locked = False
    ws.Range("A8:A44").Locked = False
    ws.Range("C8:D44").Locked = False
    ws.Range("F8:G44").Locked = False
    ws.Range("V8:V44").Locked = False
End Sub

Private Function HasShape(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(nm)
    HasShape = Not shp Is Nothing
End Function

Private Function ControlsCurrent(ByVal ws As Worksheet) As Boolean
    ControlsCurrent = False
    If HasShape(ws, SPN_NAME) Then Exit Function
    If Not HasShape(ws, BTN_NEW) Then Exit Function
    If Not HasShape(ws, BTN_PULL) Then Exit Function
    On Error Resume Next
    If ws.Buttons(BTN_NEW).caption <> CAP_NEW Then Exit Function
    If ws.Buttons(BTN_PULL).caption <> CAP_PULL Then Exit Function
    ControlsCurrent = True
End Function

' Idempotent: "New Set" at E1:G2, "Last Survey" at E3:G4, no spinner.
Private Sub EnsureContDIControls(ByVal ws As Worksheet)
    Dim btn As Button

    On Error Resume Next
    ws.Spinners(SPN_NAME).Delete
    On Error GoTo 0

    If HasShape(ws, BTN_NEW) Then
        Set btn = ws.Buttons(BTN_NEW)
    Else
        Set btn = ws.Buttons.Add(ws.Range("E1").Left + 3, ws.Range("E1").Top + 3, _
                                 ws.Range("E1:G1").width - 6, ws.Range("E1:E2").Height - 6)
        btn.name = BTN_NEW
        btn.Placement = xlMoveAndSize
    End If
    btn.caption = CAP_NEW
    btn.Font.bold = True
    btn.OnAction = "'" & ThisWorkbook.name & "'!ContDI_NewSet"

    If HasShape(ws, BTN_PULL) Then
        Set btn = ws.Buttons(BTN_PULL)
    'Else
     '   Set btn = ws.Buttons.Add(ws.Range("E3").Left + 3, ws.Range("E3").top + 2, _
     '                            ws.Range("E3:G3").Width - 6, ws.Range("E3:E4").Height - 4)
     '   btn.name = BTN_PULL
     '   btn.Placement = xlMoveAndSize
    End If
    btn.caption = CAP_PULL
    btn.Font.bold = True
    btn.OnAction = "'" & ThisWorkbook.name & "'!ContDI_PullLastSurvey"
    Err.Clear
End Sub

Private Sub AddContDIControls(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Buttons(BTN_NEW).Delete
    ws.Buttons(BTN_PULL).Delete
    ws.Spinners(SPN_NAME).Delete
    On Error GoTo 0
    EnsureContDIControls ws
End Sub

' ================================================================================
'  LAYOUT + FORMULAS (patched rates)
' ================================================================================

' -- palette (matches the original Continuous DI workbook) -----------------------
Private Function cBand() As Long: cBand = RGB(191, 191, 191): End Function
Private Function cBandLt() As Long: cBandLt = RGB(217, 217, 217): End Function
Private Function cInput() As Long: cInput = RGB(255, 255, 204): End Function
Private Function cTarget() As Long: cTarget = RGB(255, 255, 0): End Function
Private Function cGood() As Long: cGood = RGB(146, 208, 80): End Function
Private Function cBad() As Long: cBad = RGB(255, 0, 0): End Function
Private Function cGridLine() As Long: cGridLine = RGB(150, 150, 150): End Function
' Rate-block tints: Between sensors stays light blue; MWD Average (incl. DLS)
' and PD Average (incl. PD DLS) use the green the DD painted on Field.
Private Function cPdBlock() As Long: cPdBlock = RGB(248, 203, 173): End Function
Private Function cMwdBlock() As Long: cMwdBlock = RGB(217, 225, 242): End Function
Private Function cRateHead() As Long: cRateHead = RGB(198, 224, 180): End Function

' Green when the seen average keeps pace with the required rate, red when short.
Private Sub AddPaceCF(ByVal c As Range, ByVal reqAddr As String)
    Dim a As String
    a = c.Address
    With c.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & a & "),ISNUMBER(" & reqAddr & ")," & a & ">=" & reqAddr & ")")
        .Interior.Color = cGood()
    End With
    With c.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & a & "),ISNUMBER(" & reqAddr & ")," & a & "<" & reqAddr & ")")
        .Interior.Color = cBad()
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

' Signed turn-rate pace (+ right / - left), compared at the 2-decimal display
' precision. On target = turning the needed way at least as fast, or within
' TURN_DEADBAND of the required turn (so holding azimuth when the target needs
' only a hair of turn reads green).
Private Sub AddTurnPaceCF(ByVal c As Range, ByVal reqAddr As String)
    Const TURN_DEADBAND As String = "0.05"
    Dim a As String, got As String, want As String, onTarget As String
    a = c.Address
    got = "ROUND(" & a & ",2)"
    want = "ROUND(" & reqAddr & ",2)"
    onTarget = "OR(AND(SIGN(" & got & ")=SIGN(" & want & "),ABS(" & got & ")>=ABS(" & want & "))," & _
               "ROUND(ABS(" & got & "-" & want & "),2)<=" & TURN_DEADBAND & ")"
    With c.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & a & "),ISNUMBER(" & reqAddr & ")," & onTarget & ")")
        .Interior.Color = cGood()
    End With
    With c.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(ISNUMBER(" & a & "),ISNUMBER(" & reqAddr & "),NOT(" & onTarget & "))")
        .Interior.Color = cBad()
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Private Sub PaintContDI(ByVal ws As Worksheet)
    Dim wasProt As Boolean

    wasProt = SheetUnprotectForVba(ws)
    ws.Cells.Clear
    On Error Resume Next
    ws.Cells.UnMerge
    ws.Cells.FormatConditions.Delete
    ws.Buttons.Delete
    ws.Spinners.Delete
    On Error GoTo 0

    ' -- last survey tie-on block ---------------------------------------------------
    ws.Range("A1").value = "SD"
    ws.Range("B1").value = "Inc."
    ws.Range("C1").value = "Azimuth"
    ws.Range("D1").value = "TVD"
    With ws.Range("A1:D1")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBandLt()
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A2:D2")
        .Interior.Color = cInput()
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .numberFormat = "0.00"
    End With

    ' -- offsets block ---------------------------------------------------------------
    With ws.Range("L1:O1")
        .Merge
        .value = "Bit to Sensor Offsets"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("L2").value = "PD Inc."
    ws.Range("N2").value = "D&I"
    ws.Range("M2").value = 2.25
    ws.Range("O2").value = 35
    PaintContDITopChrome ws

    ' -- target block + required rates (over bit-projection columns R:U) ------------
    ws.Range("R1").value = "Target MD"
    ws.Range("R2").value = "Target INC"
    ws.Range("R3").value = "Target AZM"
    ws.Range("R4").value = "Target TVD"
    ws.Range("T1").value = "BRR"
    ws.Range("T2").value = "TRR"
    ws.Range("T3").value = "DLR"
    ws.Range("T4").value = "TFR"

    PaintContDIGridHeaders ws
    HideContDIHelpers ws

    ' -- widths / freeze -------------------------------------------------------------
    ws.Columns("A:G").ColumnWidth = 9
    ws.Columns("H").ColumnWidth = 9.5
    ws.Columns("I:L").ColumnWidth = 9
    ws.Columns("M:O").ColumnWidth = 9.5
    ws.Columns("P:Q").ColumnWidth = 9
    ws.Columns("R").ColumnWidth = 10.5
    ws.Columns("S:V").ColumnWidth = 9.5
    ws.Columns("W").ColumnWidth = 9
    ws.Columns("X:Y").ColumnWidth = 8
    ws.Columns("Z:AB").ColumnWidth = 8
    On Error Resume Next
    ws.Activate
    ws.Range("A8").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    SheetReprotectAfterVba ws, wasProt
End Sub

' Rows 1:3 chrome that carries no values: button backdrop behind New Set
' (E1:G2) and Last Survey (E3:G4), no fill where the course spinner used to
' sit (H1:I2), offsets row L2:O2 centred and boxed. Safe on a live tab.
Private Sub PaintContDITopChrome(ByVal ws As Worksheet)
    ws.Range("E1:G4").Interior.Color = cBandLt()
    ws.Range("H1:I2").Interior.ColorIndex = xlColorIndexNone
    With ws.Range("L2:O2")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("M2,O2").Interior.Color = cInput()
    PaintContDITargetSelector ws
    Err.Clear
End Sub

' P1:Q2 names the station the yellow target box is following. The name and
' S1:S4 are formulas (ContDI_FollowSlideTarget); this only paints the chrome.
Private Sub PaintContDITargetSelector(ByVal ws As Worksheet)
    Dim nameFormula As String

    On Error Resume Next
    nameFormula = CStr(ws.Range(TARGET_PICK_CELL).formula & "")
    ws.Range("P1:Q2").UnMerge
    ws.Range("P1:Q2").ClearFormats
    ws.Range(TARGET_PICK_CELL).Validation.Delete

    With ws.Range("P1:Q1")
        .Merge
        .value = "Active Target"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBandLt()
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("P2:Q2")
        .Merge
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cTarget()
        .Borders.LineStyle = xlContinuous
        .numberFormat = "General"
    End With
    If InStr(1, nameFormula, "ProjTargets_NAME", vbTextCompare) > 0 Then
        ws.Range(TARGET_PICK_CELL).formula = nameFormula
    End If
    On Error GoTo 0
    Err.Clear
End Sub

' Column A was the "Last MWD Survey" label. Delete it once so the tie-on
' survey starts in A and every formula on this sheet and ContDI_Data shifts
' with it. A later refresh sees SD already in A1 and leaves the grid alone.
Private Sub DropLabelColumn(ByVal ws As Worksheet)
    Dim a1 As String, a2 As String, b1 As String
    a1 = Trim$(CStr(ws.Range("A1").value & ""))
    a2 = Trim$(CStr(ws.Range("A2").value & ""))
    b1 = Trim$(CStr(ws.Range("B1").value & ""))
    If StrComp(a2, "Last MWD Survey", vbTextCompare) <> 0 _
            And Not (StrComp(b1, "SD", vbTextCompare) = 0 _
                     And StrComp(a1, "SD", vbTextCompare) <> 0) Then
        Exit Sub
    End If
    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False
    End If
    ws.Columns(1).Delete
    ws.Activate
    ws.Range("A8").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    Err.Clear
End Sub

' Insert I:J in the data band only (rows 5:44) when the tab still has the
' pre-sensor-DLS layout. Rows 1:4 stay put so Target / BRR / offsets do not move.
Private Sub EnsureInterSensorColumns(ByVal ws As Worksheet)
    Dim hdr As String
    hdr = Trim$(CStr(ws.Range("I6").value & ""))
    If StrComp(hdr, HDR_PD_DLS, vbTextCompare) <> 0 Then Exit Sub
    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False
    End If
    ws.Range("H4:O4").UnMerge
    ws.Range("A5:T5").UnMerge
    ws.Range("H6:T7").UnMerge
    On Error GoTo 0
    ws.Range("I5:J44").Insert Shift:=xlToRight
    Err.Clear
End Sub

' Labels / bands / input fill for H:V plus the R1:U4 target block. Safe to
' re-run; does not touch A2:D2, M2, O2, S1:S4, or typed A/C/D/F/G/V values.
Private Sub PaintContDIGridHeaders(ByVal ws As Worksheet)
    Dim deg As String, degRate As String
    Dim c As Range
    deg = "(" & ChrW$(176) & ")"
    degRate = "(" & ChrW$(176) & "/30m)"

    On Error Resume Next
    ws.Range("A5:V5").UnMerge
    ws.Range("H4:Q4").UnMerge
    ws.Range("R5:V5").UnMerge
    On Error GoTo 0
    ' Group labels are re-placed below; clear first so a merge never has to
    ' pick between two leftover labels.
    ws.Range("H4:Q4").ClearContents

    ' Averages sit on I3:L3 (MWD Average) and P3:Q3 (PD Average), directly
    ' over their rate columns. M3:O3 and H3 carry no average.
    ws.Range("H3:Q3").FormatConditions.Delete
    ws.Range("H3:Q3").Borders.LineStyle = xlNone
    With ws.Range("I3:L3,P3:Q3")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With
    ' BUR averages pace against BRR (U1). TR averages are signed, so they pace
    ' against the signed required turn (ContDI_Data!X2; U2 shows its size):
    ' green only when turning the way the target needs and at least as fast.
    For Each c In ws.Range("I3,K3,P3").Cells
        AddPaceCF c, "$U$1"
    Next c
    For Each c In ws.Range("J3,L3,Q3").Cells
        AddTurnPaceCF c, SH_CONTDI_DATA & "!$X$2"
    Next c

    ' Rate-block bands (row 4): H:L MWD Average green, M:N Between blue,
    ' O:Q PD Average green. Group labels include the DLS columns.
    With ws.Range("H4:Q4")
        .Font.bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Interior.Color = cMwdBlock()
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("H4:L4").Interior.Color = cRateHead()
    ws.Range("O4:Q4").Interior.Color = cRateHead()
    With ws.Range("H4:L4")
        .Merge
        .value = "MWD Average"
    End With
    With ws.Range("M4:N4")
        .Merge
        .value = "Between sensors"
    End With
    With ws.Range("O4:Q4")
        .Merge
        .value = "PD Average"
    End With

    With ws.Range("A5:G5")
        .Merge
        .value = "Continuous D & I"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("H5").value = "f/ survey"
    ws.Range("I5").value = "f/ cont."
    ws.Range("J5").value = "f/ cont."
    ws.Range("K5").value = "f/ survey"
    ws.Range("L5").value = "f/ survey"
    ws.Range("M5").value = "btw sens."
    ws.Range("N5").value = "btw sens."
    ws.Range("O5").value = "f/ survey"
    ws.Range("P5").value = "f/ cont."
    ws.Range("Q5").value = "f/ cont."
    With ws.Range("H5:Q5")
        .Font.size = 9
        .Font.bold = True
        .Font.Italic = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Interior.Color = cMwdBlock()
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("H5:L5").Interior.Color = cRateHead()
    ws.Range("O5:Q5").Interior.Color = cRateHead()
    With ws.Range("R5:V5")
        .Merge
        .value = "Bit Projection f/ last survey"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With

    ws.Range("A6").value = "Bit Depth"
    ws.Range("B6").value = "MWD SD"
    ws.Range("C6").value = "Inc."
    ws.Range("D6").value = "Azi."
    ws.Range("E6").value = "PD SD"
    ws.Range("F6").value = "PD Inc."
    ws.Range("G6").value = "PD Azm."
    ws.Range("H6").value = "MWD DLS"
    ws.Range("I6").value = "BUR"
    ws.Range("J6").value = "TR"
    ws.Range("K6").value = "BUR"
    ws.Range("L6").value = "TR"
    ws.Range("M6").value = HDR_SENS_DLS
    ws.Range("N6").value = "Sens BUR"
    ws.Range("O6").value = HDR_PD_DLS
    ws.Range("P6").value = "BUR"
    ws.Range("Q6").value = "TR"
    ws.Range("R6").value = "Hole Depth"
    ws.Range("S6").value = "Inc. @ Bit"
    ws.Range("T6").value = "Az. @ Bit"
    ws.Range("U6").value = "TVD @ Bit"
    ws.Range("V6").value = "ROP"
    ws.Range("A7").value = "(m)"
    ws.Range("B7").value = "(m)"
    ws.Range("C7").value = deg
    ws.Range("D7").value = deg
    ws.Range("E7").value = "(m)"
    ws.Range("F7").value = deg
    ws.Range("G7").value = deg
    ws.Range("H7").value = "f/ survey"
    ws.Range("I7").value = degRate
    ws.Range("J7").value = degRate
    ws.Range("K7").value = degRate
    ws.Range("L7").value = degRate
    ws.Range("M7").value = "PD-D&I"
    ws.Range("N7").value = degRate
    ws.Range("O7").value = "f/ survey"
    ws.Range("P7").value = degRate
    ws.Range("Q7").value = degRate
    ws.Range("R7").value = "(m)"
    ws.Range("S7").value = deg
    ws.Range("T7").value = deg
    ws.Range("U7").value = "(m)"
    ws.Range("V7").value = "(m/hr)"
    With ws.Range("A6:V7")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("A7:V7").Font.size = 9
    ws.Range("A6:G7").Interior.Color = cBandLt()
    ' Rate-block column heads (rows 6:7) carry the row-4 tints.
    With ws.Range("H6:Q7")
        .Interior.Color = cMwdBlock()
        .Font.Color = RGB(0, 0, 0)
    End With
    ws.Range("H6:L7").Interior.Color = cRateHead()
    ws.Range("O6:Q7").Interior.Color = cRateHead()
    ws.Range("R6:V7").Interior.Color = cBandLt()

    With ws.Range("A8:V44")
        .numberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = cGridLine()
    End With
    ws.Range("A8:A44,C8:D44,F8:G44,V8:V44").Interior.Color = cInput()

    ' Target / rate block. Labels re-seeded if a prior layout pass cleared them.
    If Len(Trim$(CStr(ws.Range("R1").value & ""))) = 0 Then ws.Range("R1").value = "Target MD"
    If Len(Trim$(CStr(ws.Range("R2").value & ""))) = 0 Then ws.Range("R2").value = "Target INC"
    If Len(Trim$(CStr(ws.Range("R3").value & ""))) = 0 Then ws.Range("R3").value = "Target AZM"
    If Len(Trim$(CStr(ws.Range("R4").value & ""))) = 0 Then ws.Range("R4").value = "Target TVD"
    If Len(Trim$(CStr(ws.Range("T1").value & ""))) = 0 Then ws.Range("T1").value = "BRR"
    If Len(Trim$(CStr(ws.Range("T2").value & ""))) = 0 Then ws.Range("T2").value = "TRR"
    If Len(Trim$(CStr(ws.Range("T3").value & ""))) = 0 Then ws.Range("T3").value = "DLR"
    If Len(Trim$(CStr(ws.Range("T4").value & ""))) = 0 Then ws.Range("T4").value = "TFR"
    ws.Range("R1:U4").FormatConditions.Delete
    With ws.Range("R1:R4,T1:T4")
        .Font.bold = True
        .Interior.Color = cBandLt()
        .Borders.LineStyle = xlContinuous
        .numberFormat = "General"
    End With
    ws.Range("R1:R4").HorizontalAlignment = xlLeft
    ws.Range("T1:T4").HorizontalAlignment = xlCenter
    With ws.Range("S1:S4")
        .Interior.Color = cTarget()
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With
    With ws.Range("U1:U4")
        .Interior.Color = cGood()
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With
    Err.Clear
End Sub

' Park every diagnostic lookup on the very-hidden ContDI_Data sheet and wipe
' leftover on-grid helper cells (old Average Stand Length, V1:W4 lookups,
' X1:Y2 dBURR, W5:AB7 projections) so only the working tab shows.
Private Sub HideContDIHelpers(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range("V1:W4").Clear
    ws.Range("X1:Y2").Clear
    ws.Range("W5:AB7").Clear
    ws.Range("Z1:AB7").Clear
    ws.Columns("W").Hidden = True
    ws.Columns("X:Y").Hidden = True
    ws.Columns("Z:AB").Hidden = True
    On Error GoTo 0
    Err.Clear
End Sub

' One-time +2 column shift: P1:S4 (targets + rates) -> R1:U4 so averages
' can sit on L3:Q3 over columns L:Q. Idempotent once R1 is already Target MD.
Private Sub EnsureHeaderShift(ByVal ws As Worksheet)
    Dim q1 As String
    q1 = Trim$(CStr(ws.Range("P1").value & ""))
    If StrComp(q1, "Target MD", vbTextCompare) = 0 Then
        ws.Range("P1:S4").Copy Destination:=ws.Range("R1")
        Application.CutCopyMode = False
        ws.Range("P1:Q4").Clear
    End If
    ' Old averages sat on J3:O3 (formulas already pointed at L9:Q44).
    ' Do not Copy - relative AVERAGE refs would slide +2. Just drop leftovers.
    If InStr(1, UCase$(CStr(ws.Range("J3").formula)), "AVERAGE(L9", vbBinaryCompare) > 0 Then
        ws.Range("J3:K3").Clear
    End If
    Err.Clear
End Sub

Private Sub ApplyLookupsAndRates(ByVal ws As Worksheet)
    ' Averages ride with their rate columns: MWD Average I:L, PD Average P:Q.
    ' H3 and M3:O3 (MWD DLS, Between sensors, PD DLS) carry none.
    Dim avgCol As Variant
    ws.Range("H3:Q3").ClearContents
    For Each avgCol In Array("I", "J", "K", "L", "P", "Q")
        ws.Range(avgCol & "3").formula = "=IFERROR(AVERAGE(" & avgCol & "9:" & avgCol & "44),"""")"
    Next avgCol
    ws.Range("I3:L3,P3:Q3").numberFormat = "0.00"

    ' Diagnostics live on very-hidden ContDI_Data (T:W), not on the grid.
    Dim h As Worksheet
    Dim cd As String
    Set h = ContDIDataSheet()
    cd = "'" & SH_CONTDI & "'!"
    h.Range("T1").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "A8:A44<>"""")," & cd & "A8:A44),IF(ISNUMBER(W1),W1,IF(AND(ISNUMBER(" & cd & "A2),ISNUMBER(" & cd & "O2))," & cd & "A2+" & cd & "O2,IF(" & cd & "A2="""",""""," & cd & "A2))))"
    h.Range("T2").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "C8:C44<>"""")," & cd & "C8:C44),IF(" & cd & "B2="""",""""," & cd & "B2))"
    h.Range("T3").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "D8:D44<>"""")," & cd & "D8:D44),IF(" & cd & "C2="""",""""," & cd & "C2))"
    h.Range("T4").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "S8:S44<>"""")," & cd & "S8:S44),IF(ISNUMBER(W2),W2,IF(" & cd & "B2="""",""""," & cd & "B2)))"
    h.Range("U1").formula = "=IFERROR(" & cd & "S1-T1,"""")"
    h.Range("U2").formula = "=IFERROR((" & cd & "S2-T4)/U1,"""")"
    h.Range("U3").formula = "=IFERROR((MOD(" & cd & "S3-U4+180,360)-180)/U1,"""")"
    h.Range("U4").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "T8:T44<>"""")," & cd & "T8:T44),IF(ISNUMBER(W3),W3,IF(" & cd & "C2="""",""""," & cd & "C2)))"
    h.Range("V1").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "U8:U44<>"""")," & cd & "U8:U44),IF(" & cd & "D2="""",""""," & cd & "D2))"
    h.Range("T5").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "A8:A44<>"""")," & cd & "A8:A44),"""")"
    h.Range("U5").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "C8:C44<>"""")," & cd & "C8:C44),"""")"
    h.Range("V5").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "D8:D44<>"""")," & cd & "D8:D44),"""")"
    h.Range("T6").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "F8:F44<>"""")," & cd & "F8:F44),"""")"
    h.Range("U6").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "G8:G44<>"""")," & cd & "G8:G44),"""")"
    h.Range("V6").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "R8:R44<>"""")," & cd & "R8:R44),"""")"
    h.Range("T7").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "S8:S44<>"""")," & cd & "S8:S44),"""")"
    h.Range("U7").formula = "=IFERROR(LOOKUP(2,1/(" & cd & "T8:T44<>"""")," & cd & "T8:T44),"""")"
    h.Range("V7").formula = "=IFERROR(ROUND(LOOKUP(2,1/(" & cd & "U8:U44<>"""")," & cd & "U8:U44),2),"""")"

    ' Target S1:S4 / P2, rates U1:U4 (BRR, TRR, DLR, TFR) and the signed TRR
    ' twin (X2, turn-average colouring) are one plan-point arc.
    ContDI_FollowSlideTarget
    HideContDIHelpers ws
End Sub

' Re-apply every formula cell (grid rows 8:44 + header lookups / rates) on a
' painted tab without touching inputs (A2:D2, M2, O2, S1:S4, A/C/D/F/G/V 8:44).
' Migrates an old tab in place: inserts I:J once (PD-DLS-at-I layout), shifts
' P1:S4 -> R1:U4 once, drops the course spinner and the stand-length row.
Public Sub ContDI_RefreshFormulas()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim ev As Boolean
    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    Application.EnableEvents = False
    EnsureContDIData
    Set ws = ContDISheet()
    wasProt = SheetUnprotectForVba(ws)
    DropLabelColumn ws
    EnsureInterSensorColumns ws
    Err.Clear
    EnsureHeaderShift ws
    Err.Clear
    PaintContDIGridHeaders ws
    Err.Clear
    PaintContDITopChrome ws
    Err.Clear
    EnsureContDIControls ws
    Err.Clear
    ApplyRowFormulas ws
    Err.Clear
    SeedProjectionHelpers ws
    Err.Clear
    ApplyLookupsAndRates ws
    Err.Clear
    ContDI_FollowSlideTarget
    ApplyContDILock ws
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    Exit Sub
ErrHandler:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    MsgBox "ContDI_RefreshFormulas: [" & CStr(en) & "] " & ed, vbCritical
End Sub

Private Sub ApplyRowFormulas(ByVal ws As Worksheet)
    Dim r As Long
    Dim p As Long
    ' A blank inclination or azimuth is not zero. The equation uses the last
    ' number above it in that column. Azimuth falls back to the tie-on azimuth
    ' (C2). MWD inclination falls back to the tie-on inclination (B2). PD
    ' inclination stays blank until that column has a number, so a missing PD
    ' reading still projects from the MWD. An empty row (no bit depth) stays
    ' blank. ROP (V) is an entry; no equation on this sheet reads it.
    ' Keep the formula text in step with contDiCalcFormula in the preview.
    Dim tie As String, dMWD As String, dPD As String
    Dim nD As String, nG As String
    Dim pD As String, pG As String
    Dim stepMWD As String, stepPD As String
    Dim sens As String, rowOn As String
    Dim sensBur As String, mwdBurs As String
    Dim azE As String, azH As String, azEp As String, azHp As String
    Dim incC As String, incF As String, incCp As String, incFp As String
    Dim seenG As String
    tie = "ISNUMBER($A$2)"
    sens = "ISNUMBER($O$2),ISNUMBER($M$2),$O$2<>$M$2"   ' both offsets present and distinct
    For r = FIRST_DATA To LAST_DATA
        p = r - 1
        rowOn = "ISNUMBER(A" & r & ")"
        incC = Held("C", r, "$B$2")
        incF = Held("F", r, """""")
        azE = Held("D", r, "$C$2")
        azH = Held("G", r, "$C$2")
        seenG = Held("G", r, """""")
        nD = "ISNUMBER(" & incC & ")"
        nG = "ISNUMBER(" & incF & ")"
        dMWD = "ISNUMBER(B" & r & "),B" & r & ">$A$2"      ' MWD sensor below the survey
        dPD = "ISNUMBER(E" & r & "),E" & r & ">$A$2"       ' PD sensor below the survey
        stepMWD = "ISNUMBER(B" & r & "),ISNUMBER(B" & p & "),B" & r & "<>B" & p
        stepPD = "ISNUMBER(E" & r & "),ISNUMBER(E" & p & "),E" & r & "<>E" & p
        ws.Cells(r, 2).formula = "=IF(ISNUMBER(A" & r & "),A" & r & "-$O$2,"""")"
        ws.Cells(r, 5).formula = "=IF(ISNUMBER(A" & r & "),A" & r & "-$M$2,"""")"
        ' MWD DLS f/ survey (H) and PD DLS f/ survey (O).
        ws.Cells(r, COL_MWD_DLS).formula = "=IF(AND(" & tie & "," & nD & "," & dMWD & ",ISNUMBER(" & azE & ")),(30/(B" & r & "-$A$2))*DEGREES(ACOS(MIN(1,MAX(-1,COS(RADIANS($B$2))*COS(RADIANS(" & incC & "))+SIN(RADIANS($B$2))*SIN(RADIANS(" & incC & "))*COS(RADIANS(" & azE & "-$C$2)))))),"""")"
        ws.Cells(r, COL_PD_DLS).formula = "=IF(AND(" & tie & "," & nG & "," & dPD & ",ISNUMBER(" & azH & ")),(30/(E" & r & "-$A$2))*DEGREES(ACOS(MIN(1,MAX(-1,COS(RADIANS($B$2))*COS(RADIANS(" & incF & "))+SIN(RADIANS($B$2))*SIN(RADIANS(" & incF & "))*COS(RADIANS(" & azH & "-$C$2)))))),"""")"
        ' Sens DLS / Sens BUR: dogleg and signed BUR between the PD sensor (F,G)
        ' and the D&I (C,D) over O2-M2. A blank reading holds the previous one.
        ws.Cells(r, COL_SENS_DLS).formula = "=IF(AND(" & rowOn & "," & nD & "," & nG & ",ISNUMBER(" & azE & "),ISNUMBER(" & azH & "),ISNUMBER($O$2),ISNUMBER($M$2),$O$2<>$M$2),ProjDoglegDeg(" & incC & "," & azE & "," & incF & "," & azH & ")/ABS($O$2-$M$2)*30,"""")"
        ws.Cells(r, COL_SENS_BUR).formula = "=IF(AND(" & rowOn & "," & nD & "," & nG & ",ISNUMBER($O$2),ISNUMBER($M$2),$O$2<>$M$2),(" & incF & "-" & incC & ")/($O$2-$M$2)*30,"""")"
        If r = FIRST_DATA Then
            ws.Cells(r, COL_PD_BUR).ClearContents
            ws.Cells(r, COL_PD_TR).ClearContents
            ws.Cells(r, COL_MWD_BURC).ClearContents
            ws.Cells(r, COL_MWD_TRC).ClearContents
        End If
        If r >= 9 Then
            incCp = Held("C", p, "$B$2")
            incFp = Held("F", p, """""")
            azEp = Held("D", p, "$C$2")
            azHp = Held("G", p, "$C$2")
            pD = "ISNUMBER(" & incCp & ")"
            pG = "ISNUMBER(" & incFp & ")"
            ws.Cells(r, COL_PD_BUR).formula = "=IF(AND(" & nG & "," & pG & "," & stepPD & "),(" & incF & "-" & incFp & ")/(E" & r & "-E" & p & ")*30,"""")"
            ws.Cells(r, COL_PD_TR).formula = "=IF(AND(" & stepPD & ",ISNUMBER(" & azH & "),ISNUMBER(" & azHp & ")),(MOD(" & azH & "-" & azHp & "+180,360)-180)/(E" & r & "-E" & p & ")*30,"""")"
            ws.Cells(r, COL_MWD_BURC).formula = "=IF(AND(" & nD & "," & pD & "," & stepMWD & "),(" & incC & "-" & incCp & ")/(B" & r & "-B" & p & ")*30,"""")"
            ws.Cells(r, COL_MWD_TRC).formula = "=IF(AND(" & stepMWD & ",ISNUMBER(" & azE & "),ISNUMBER(" & azEp & ")),(MOD(" & azE & "-" & azEp & "+180,360)-180)/(B" & r & "-B" & p & ")*30,"""")"
        End If
        ws.Cells(r, COL_MWD_BURS).formula = "=IF(AND(" & tie & "," & nD & "," & dMWD & "),(" & incC & "-$B$2)/(B" & r & "-$A$2)*30,"""")"
        ' Signed like the f/ cont. TR columns: + right, - left.
        ws.Cells(r, COL_MWD_TRS).formula = "=IF(AND(" & tie & "," & dMWD & ",ISNUMBER(" & azE & ")),(MOD(" & azE & "-$C$2+180,360)-180)/(B" & r & "-$A$2)*30,"""")"
        ws.Cells(r, COL_HOLE).formula = "=IF(ISNUMBER(A" & r & "),A" & r & ","""")"
        ' Inc / Az @ Bit: the PD sensor is M2 (~2 m) off the bit, so walk the
        ' PD reading over M2 at the rate seen BETWEEN the sensors (Sens BUR).
        ' A blank PD inclination holds the last PD inclination. Only when the
        ' PD column has never had a number, carry the MWD reading over the
        ' whole D&I offset (O2) at the f/ survey rate.
        sensBur = ColLetter(COL_SENS_BUR) & r
        mwdBurs = ColLetter(COL_MWD_BURS) & r
        ws.Cells(r, COL_INC_BIT).formula = "=IF(NOT(" & rowOn & "),"""",IF(" & nG & ",IF(ISNUMBER(" & sensBur & ")," & incF & "+(" & sensBur & "/30)*$M$2," & incF & ")," & _
            "IF(NOT(" & nD & "),"""",IF(ISNUMBER(" & mwdBurs & "),(" & mwdBurs & "/30)*O$2+" & incC & "," & incC & "))))"
        ' Az @ Bit on any row that has a bit depth. A blank azimuth holds.
        ' The PD path is used once either PD column has a reading above or here.
        ws.Cells(r, COL_AZ_BIT).formula = "=IF(NOT(" & rowOn & "),"""",IF(OR(" & nG & ",ISNUMBER(" & seenG & ")),IF(AND(" & sens & ",ISNUMBER(" & azE & ")),MOD(" & azH & "+(MOD(" & azH & "-" & azE & "+180,360)-180)/($O$2-$M$2)*$M$2,360),MOD(" & azH & ",360))," & _
            "IF(NOT(" & nD & "),"""",IF(AND(" & tie & "," & dMWD & ",ISNUMBER(" & azE & ")),MOD(((MOD(" & azE & "-$C$2+180,360)-180)/(B" & r & "-$A$2))*O$2+" & azE & ",360),IF(ISNUMBER(" & azE & "),MOD(" & azE & ",360),"""")))))"
        ' TVD @ Bit: minimum curvature tie-on -> every D&I reading so far -> this
        ' row's PD reading -> bit. Blank D&I inclination keeps the previous one.
        ws.Cells(r, COL_TVD_BIT).formula = "=IF(AND(ISNUMBER($R" & r & "),ISNUMBER($S" & r & "),ISNUMBER($T" & r & ")," & _
            "ISNUMBER($A$2),ISNUMBER($B$2),ISNUMBER($C$2),ISNUMBER($D$2))," & _
            "ContDITvdAtBit($R" & r & ",$S" & r & ",$T" & r & ",$A$2,$B$2,$C$2,$D$2,$M$2," & _
            "$B$" & FIRST_DATA & ":$B" & r & ",$C$" & FIRST_DATA & ":$C" & r & ",$D$" & FIRST_DATA & ":$D" & r & "," & _
            incF & "," & azH & "),"""")"
    Next r
End Sub

' TVD @ Bit for one Cont DI row by minimum curvature through every measured
' station: tie-on, each D&I reading up to this row (MD = MWD SD), this row's
' PD reading (MD = bit - PD offset), then the bit. A station with MD but a
' blank inclination or azimuth keeps the previous one, starting from the
' tie-on. A blank PD inclination skips the PD station. Stations that do not
' advance in MD are skipped.
Public Function ContDITvdAtBit(ByVal bitMd As Double, ByVal bitInc As Double, ByVal bitAzm As Double, _
                               ByVal tieMd As Double, ByVal tieInc As Double, ByVal tieAzm As Double, _
                               ByVal tieTvd As Double, ByVal pdOffset As Variant, _
                               ByVal diMd As Range, ByVal diInc As Range, ByVal diAzm As Range, _
                               ByVal pdInc As Variant, ByVal pdAzm As Variant) As Variant
    Dim k As Long
    Dim lastMd As Double, lastInc As Double, lastAzm As Double
    Dim tvd As Double
    Dim m As Variant, i As Variant, a As Variant
    Dim useInc As Double, useAzm As Double

    On Error GoTo Fail
    If bitMd <= tieMd Then ContDITvdAtBit = "": Exit Function
    lastMd = tieMd: lastInc = tieInc: lastAzm = tieAzm
    tvd = tieTvd

    For k = 1 To diMd.rows.Count
        m = diMd.Cells(k, 1).Value2
        i = diInc.Cells(k, 1).Value2
        a = diAzm.Cells(k, 1).Value2
        If IsNumericCell(m) Then
            If CDbl(m) > lastMd And CDbl(m) < bitMd Then
                useInc = lastInc
                useAzm = lastAzm
                If IsNumericCell(i) Then useInc = CDbl(i)
                If IsNumericCell(a) Then useAzm = CDbl(a)
                AddTvdStep tvd, lastMd, lastInc, lastAzm, CDbl(m), useInc, useAzm
            End If
        End If
    Next k

    If IsNumericCell(pdOffset) And IsNumericCell(pdInc) Then
        m = bitMd - CDbl(pdOffset)
        If m > lastMd And m < bitMd Then
            If IsNumericCell(pdAzm) Then
                AddTvdStep tvd, lastMd, lastInc, lastAzm, CDbl(m), CDbl(pdInc), CDbl(pdAzm)
            Else
                AddTvdStep tvd, lastMd, lastInc, lastAzm, CDbl(m), CDbl(pdInc), lastAzm
            End If
        End If
    End If

    AddTvdStep tvd, lastMd, lastInc, lastAzm, bitMd, bitInc, bitAzm
    ContDITvdAtBit = tvd
    Exit Function
Fail:
    ContDITvdAtBit = ""
End Function

Private Sub AddTvdStep(ByRef tvd As Double, ByRef lastMd As Double, ByRef lastInc As Double, _
                       ByRef lastAzm As Double, ByVal md As Double, ByVal inc As Double, ByVal azm As Double)
    Dim dV As Double, dN As Double, dE As Double
    MDL_PlanGauge.PG_McStep lastMd, lastInc, lastAzm, md, inc, azm, dV, dN, dE
    tvd = tvd + dV
    lastMd = md: lastInc = inc: lastAzm = azm
End Sub

Private Function IsNumericCell(ByVal v As Variant) As Boolean
    If isError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    IsNumericCell = IsNumeric(v)
End Function

' Reading used by the row equations. This cell when it is a number; otherwise
' the last number above it in the same column; otherwise tieFallback.
' tieFallback is a formula fragment ("$B$2", "$C$2") or """""" when a column
' that has never been typed stays blank. A blank is never passed through as 0.
Private Function Held(ByVal col As String, ByVal r As Long, ByVal tieFallback As String) As String
    Dim prev As String
    If r > FIRST_DATA Then
        prev = "IFERROR(LOOKUP(2,1/(" & col & "$" & FIRST_DATA & ":" & col & CStr(r - 1) & "<>"""")," & col & "$" & FIRST_DATA & ":" & col & CStr(r - 1) & ")," & tieFallback & ")"
    Else
        prev = tieFallback
    End If
    Held = "IF(ISNUMBER(" & col & CStr(r) & ")," & col & CStr(r) & "," & prev & ")"
End Function

' Column number -> letter (1..26 only; the Cont DI grid ends at V).
Private Function ColLetter(ByVal col As Long) As String
    ColLetter = Chr$(64 + col)
End Function



























