Attribute VB_Name = "MDL_ContDI"
Option Explicit

' Continuous DI / BURR tab: one visible grid, hidden course archive, spin to
' page past course lengths (view-only). NewCourse archives instead of copying a sheet.
' Layout and patched rates live here so Rebuild keeps them.

Public Const SH_CONTDI As String = "Cont DI"
Public Const SH_CONTDI_DATA As String = "ContDI_Data"

Private Const FIRST_DATA As Long = 8
Private Const LAST_DATA As Long = 44
Private Const DATA_ROWS As Long = 37
Private Const IDX_FIRST As Long = 4
Private Const DUMP_START As Long = 50
Private Const DUMP_STRIDE As Long = 40

Private Const SPN_NAME As String = "ContDI_SpinCtl"
Private Const BTN_NEW As String = "ContDI_NewCourse"
Private Const BTN_BUILD As String = "ContDI_Build"
Private Const BTN_PULL As String = "ContDI_PullSurvey"
Private Const SS_SHEET As String = "Slidesheet"
Private Const SS_FIRST As Long = 13
Private Const SS_LAST As Long = 305

' Grid columns after the inter-sensor pair was inserted at J:K (rows 5:44).
' Targets S1:T4 and rates U1:V4 sit over the bit-projection columns.
' Averages M3:R3 sit on PD / MWD rate columns M:R.
Private Const COL_SENS_DLS As Long = 10   ' J  3D DLS MWD(D,E) <-> PD(G,H)
Private Const COL_SENS_BUR As Long = 11   ' K  signed BUR, PD minus MWD
Private Const COL_PD_DLS As Long = 12     ' L
Private Const COL_PD_BUR As Long = 13     ' M
Private Const COL_PD_TR As Long = 14      ' N
Private Const COL_MWD_BURC As Long = 15   ' O
Private Const COL_MWD_TRC As Long = 16    ' P
Private Const COL_MWD_BURS As Long = 17   ' Q
Private Const COL_MWD_TRS As Long = 18    ' R
Private Const COL_HOLE As Long = 19       ' S
Private Const COL_INC_BIT As Long = 20    ' T
Private Const COL_AZ_BIT As Long = 21     ' U
Private Const COL_TVD_BIT As Long = 22    ' V
Private Const COL_ROP As Long = 23        ' W
Private Const HDR_SENS_DLS As String = "Sens DLS"
Private Const HDR_PD_DLS As String = "PD DLS"

Private Type ContDISeed
    bitMd As Variant
    mwdInc As Variant
    mwdAzm As Variant
    pdInc As Variant
    pdAzm As Variant
    holeMd As Variant
    bitInc As Variant
    bitAzm As Variant
    bitTvd As Variant
    standLen As Variant
    pdOff As Variant
    diOff As Variant
    tgtMd As Variant
    tgtInc As Variant
    tgtAzm As Variant
    tgtTvd As Variant
End Type

' ================================================================================
'  PUBLIC
' ================================================================================

Public Sub BuildContDITab()
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim view As Long
    Dim ev As Boolean
    Dim su As Boolean

    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    su = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set dataWs = EnsureContDIData()
    view = CourseCount()
    If view < 1 Then
        dataWs.Range("B1").Value = 1
        dataWs.Range("C1").Value = 1
        view = 1
    End If

    Set ws = EnsureContDISheet()
    ' Archive the live grid before repaint wipes it (rebuild safety).
    If CourseCount() >= 1 And ViewIdx() = CourseCount() _
       And Len(Trim$(CStr(ws.Range("B8").Value & ""))) > 0 Then
        Dim wp As Boolean
        wp = SheetUnprotectForVba(ws)
        ArchiveCurrent ws, dataWs, CourseCount()
        SheetReprotectAfterVba ws, wp
    End If
    PaintContDI ws
    AddContDIControls ws
    ApplyRowFormulas ws
    ApplyLookupsAndRates ws
    ApplyContDILock ws, True
    If DumpHasRows(1) Then
        LoadCourseIntoGrid CLng(Nz(dataWs.Range("C1").Value, 1))
        ApplyContDILock ws, (CLng(Nz(dataWs.Range("C1").Value, 1)) >= CourseCount())
    End If
    UpdateNavigator ws
    SheetReprotectAfterVba ws, True

    Application.EnableEvents = ev
    Application.ScreenUpdating = su
    Exit Sub

ErrHandler:
    Application.EnableEvents = ev
    Application.ScreenUpdating = su
    MsgBox "BuildContDITab: " & Err.Description, vbCritical
End Sub

Public Sub ContDI_NewCourse()
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim wasProt As Boolean
    Dim seed As ContDISeed
    Dim n As Long
    Dim ev As Boolean

    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    Application.EnableEvents = False

    Set ws = ContDISheet()
    Set dataWs = ContDIDataSheet()
    If ViewIdx() <> CourseCount() Then
        Application.EnableEvents = ev
        MsgBox "Spin to the latest course before starting a new one.", vbExclamation
        Exit Sub
    End If

    wasProt = SheetUnprotectForVba(ws)
    seed = CaptureSeed(ws)
    ArchiveCurrent ws, dataWs, CourseCount()

    n = CourseCount() + 1
    dataWs.Range("B1").Value = n
    dataWs.Range("C1").Value = n

    ClearInputs ws
    ApplySeed ws, seed
    ' Tie-on (B2:E2) is always a real MWD survey: the last green Slidesheet
    ' row, or - when none qualifies - the previous course's tie-on. The
    ' projected bit (V5:X7) never becomes a survey; it only seeds row 8.
    WriteLastSurveyTieOn ws, False
    ' Persist the header AFTER the tie-on is set, so spinning between courses
    ' reloads the real survey and not the projection.
    WriteIndexRow dataWs, n, HeaderSeed(ws)
    EnsurePullSurveyButton ws
    ApplyContDILock ws, True
    SyncSpinner ws, n, n
    UpdateNavigator ws
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    Exit Sub

ErrHandler:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    MsgBox "ContDI_NewCourse: " & Err.Description, vbCritical
End Sub

' Button / macro: refill "Last MWD Survey" (B2:E2) from the Slidesheet.
Public Sub ContDI_PullLastSurvey()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim ev As Boolean

    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    Application.EnableEvents = False

    Set ws = ContDISheet()
    If CourseCount() >= 1 And ViewIdx() <> CourseCount() Then
        Application.EnableEvents = ev
        MsgBox "Spin to the latest course before pulling a survey.", vbExclamation
        Exit Sub
    End If

    wasProt = SheetUnprotectForVba(ws)
    WriteLastSurveyTieOn ws, True
    ' Persist so spinning away and back reloads this survey, not a stale header.
    If CourseCount() >= 1 Then WriteIndexRow ContDIDataSheet(), CourseCount(), HeaderSeed(ws)
    UpdateNavigator ws
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    Exit Sub

ErrHandler:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    MsgBox "ContDI_PullLastSurvey: " & Err.Description, vbCritical
End Sub

' Lazy control creation for workbooks whose Cont DI tab was painted before the
' pull button existed (called from Workbook_SheetActivate; no repaint needed).
Public Sub EnsureContDIPullButton()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONTDI)
    If ws Is Nothing Then Exit Sub
    If HasShape(ws, BTN_PULL) Then
        If ws.Buttons(BTN_PULL).caption = "Last Survey" Then Exit Sub
    End If
    wasProt = SheetUnprotectForVba(ws)
    EnsurePullSurveyButton ws
    SheetReprotectAfterVba ws, wasProt
End Sub

' ================================================================================
'  LAST MWD SURVEY (Slidesheet -> B2:E2)
' ================================================================================

' The Slidesheet survey-entry fill: light green RGB(204,255,204) (Excel
' ColorIndex 35). A row is a GOOD survey only when Inc (F) and Azm (G) are
' both numeric AND both carry exactly this green. Any other fill — the bright
' yellow the DD uses for plan-ahead / projected stations, notes, highlights —
' means "not a survey" and is never counted, whatever its depth.
Private Function GoodSurveyFill() As Long
    GoodSurveyFill = RGB(204, 255, 204)
End Function

Private Function IsGoodSurveyCell(ByVal c As Range) As Boolean
    IsGoodSurveyCell = False
    If Not IsNumberValue(c.Value2) Then Exit Function
    If c.Interior.Pattern <> xlSolid Then Exit Function
    IsGoodSurveyCell = (CLng(c.Interior.Color) = GoodSurveyFill())
End Function

' Deepest good-survey row on the Slidesheet (13:305); 0 = none.
Public Function LastSlidesheetSurveyRow() As Long
    Dim ss As Worksheet
    Dim r As Long

    LastSlidesheetSurveyRow = 0
    On Error GoTo Done
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    For r = SS_LAST To SS_FIRST Step -1
        If IsGoodSurveyCell(ss.Cells(r, "F")) And IsGoodSurveyCell(ss.Cells(r, "G")) Then
            LastSlidesheetSurveyRow = r
            Exit Function
        End If
    Next r
Done:
End Function

' B2 = survey MD (E), C2 = Inc (F), D2 = Azm (G), E2 = TVD (H).
' Stand length (Slidesheet C, next survey course) is stored on ContDI_Data!X1
' for ProjBurr's short-MD floor — not shown on the tab.
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
    ws.Range("B2").Value = ss.Cells(r, "E").Value2
    ws.Range("C2").Value = ss.Cells(r, "F").Value2
    ws.Range("D2").Value = ss.Cells(r, "G").Value2
    ws.Range("E2").Value = ss.Cells(r, "H").Value2

    nextC = Empty
    If r + 1 <= SS_LAST Then
        If IsNumberValue(ss.Cells(r + 1, "C").Value2) Then nextC = ss.Cells(r + 1, "C").Value2
    End If
    If IsEmpty(nextC) Then
        If IsNumberValue(ss.Cells(r, "C").Value2) Then nextC = ss.Cells(r, "C").Value2
    End If
    If Not IsEmpty(nextC) Then ContDIDataSheet().Range("X1").Value = Round(CDbl(nextC), 2)
    SeedProjectionHelpers ws, r
End Sub

' AA1 = Slidesheet bit MD (D), AA2/AB2 = Inc@Bit / Az@Bit (W/X) for the
' tie-on row. Does not touch B2:E2. Used so T1 can start from the same
' projected bit as Slidesheet Y while the Cont DI grid is still empty.
Private Sub SeedProjectionHelpers(ByVal ws As Worksheet, Optional ByVal ssRow As Long = 0)
    Dim ss As Worksheet
    Dim r As Long
    Dim md As Double

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    r = ssRow
    If r < SS_FIRST Then
        r = 0
        If IsNumberValue(ws.Range("B2").Value2) Then
            md = CDbl(ws.Range("B2").Value2)
            Dim i As Long
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
        If IsNumberValue(ss.Cells(r, "D").Value2) Then .Range("W1").Value = ss.Cells(r, "D").Value2
        If IsNumberValue(ss.Cells(r, "W").Value2) Then .Range("W2").Value = ss.Cells(r, "W").Value2
        If IsNumberValue(ss.Cells(r, "X").Value2) Then .Range("W3").Value = ss.Cells(r, "X").Value2
        If r + 1 <= SS_LAST And IsNumberValue(ss.Cells(r + 1, "C").Value2) Then
            .Range("X1").Value = Round(CDbl(ss.Cells(r + 1, "C").Value2), 2)
        ElseIf IsNumberValue(ss.Cells(r, "C").Value2) Then
            .Range("X1").Value = Round(CDbl(ss.Cells(r, "C").Value2), 2)
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

Private Function HasShape(ByVal ws As Worksheet, ByVal nm As String) As Boolean
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(nm)
    HasShape = Not shp Is Nothing
End Function

' "Last Survey" button under the New Course button (F3:H4). Idempotent.
Private Sub EnsurePullSurveyButton(ByVal ws As Worksheet)
    Dim btn As Button
    If HasShape(ws, BTN_PULL) Then
        On Error Resume Next
        ws.Buttons(BTN_PULL).caption = "Last Survey"
        On Error GoTo 0
        Exit Sub
    End If
    Set btn = ws.Buttons.Add(ws.Range("F3").Left + 3, ws.Range("F3").Top + 2, _
                             ws.Range("F3:H3").Width - 6, ws.Range("F3:F4").Height - 4)
    btn.name = BTN_PULL
    btn.caption = "Last Survey"
    btn.Font.bold = True
    btn.OnAction = "'" & ThisWorkbook.name & "'!ContDI_PullLastSurvey"
    btn.Placement = xlMoveAndSize
End Sub

Public Sub ContDI_Spin()
    Dim ws As Worksheet
    Dim sp As Spinner
    Dim idx As Long

    On Error GoTo ErrHandler
    Set ws = ContDISheet()
    Set sp = ws.Spinners(SPN_NAME)
    idx = CLng(sp.Value)
    If idx < 1 Then idx = 1
    If idx > CourseCount() Then idx = CourseCount()
    ContDI_LoadCourse idx
    Exit Sub

ErrHandler:
    MsgBox "ContDI_Spin: " & Err.Description, vbCritical
End Sub

Public Sub ContDI_LoadCourse(ByVal idx As Long)
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim wasProt As Boolean
    Dim ev As Boolean
    Dim live As Boolean

    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    Application.EnableEvents = False

    If idx < 1 Or idx > CourseCount() Then
        Application.EnableEvents = ev
        Exit Sub
    End If

    Set ws = ContDISheet()
    Set dataWs = ContDIDataSheet()
    wasProt = SheetUnprotectForVba(ws)

    ' Archive whatever course is on screen before switching (inputs are
    ' editable on archived courses too, so their edits must round-trip).
    If idx <> ViewIdx() And ViewIdx() >= 1 And ViewIdx() <= CourseCount() Then
        ArchiveCurrent ws, dataWs, ViewIdx()
    End If

    dataWs.Range("C1").Value = idx
    LoadCourseIntoGrid idx
    live = (idx = CourseCount())
    ApplyContDILock ws, live
    SyncSpinner ws, idx, CourseCount()
    UpdateNavigator ws
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    Exit Sub

ErrHandler:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = ev
    MsgBox "ContDI_LoadCourse: " & Err.Description, vbCritical
End Sub

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

Private Function EnsureContDIData() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONTDI_DATA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SH_CONTDI_DATA
        ws.Range("A1").Value = "ContDI_Data"
        ws.Range("B1").Value = 0
        ws.Range("C1").Value = 0
        ws.Range("D1").Value = "courseCount / viewIdx"
        ws.Range("A3").Value = "idx"
        ws.Range("B3").Value = "label"
        ws.Range("C3").Value = "tieMD"
        ws.Range("D3").Value = "tieInc"
        ws.Range("E3").Value = "tieAzm"
        ws.Range("F3").Value = "tieTVD"
        ws.Range("G3").Value = "standLen"
        ws.Range("H3").Value = "pdOff"
        ws.Range("I3").Value = "diOff"
        ws.Range("J3").Value = "tgtMD"
        ws.Range("K3").Value = "tgtInc"
        ws.Range("L3").Value = "tgtAzm"
        ws.Range("M3").Value = "tgtTVD"
        ws.Range("N3").Value = "dumpRow"
    End If
    ws.Visible = xlSheetVeryHidden
    Set EnsureContDIData = ws
End Function

Private Function CourseCount() As Long
    On Error Resume Next
    CourseCount = CLng(Nz(ContDIDataSheet().Range("B1").Value, 0))
End Function

Private Function ViewIdx() As Long
    On Error Resume Next
    ViewIdx = CLng(Nz(ContDIDataSheet().Range("C1").Value, 1))
End Function

Private Function Nz(ByVal v As Variant, ByVal fallback As Variant) As Variant
    If isError(v) Then
        Nz = fallback
    ElseIf IsEmpty(v) Or IsNull(v) Then
        Nz = fallback
    ElseIf VarType(v) = vbString And Len(Trim$(CStr(v))) = 0 Then
        Nz = fallback
    Else
        Nz = v
    End If
End Function

Private Function DumpRowFor(ByVal idx As Long) As Long
    DumpRowFor = DUMP_START + (idx - 1) * DUMP_STRIDE
End Function

' ================================================================================
'  ARCHIVE / LOAD
' ================================================================================

Private Function CaptureSeed(ByVal ws As Worksheet) As ContDISeed
    Dim s As ContDISeed
    Dim h As Worksheet
    Set h = ContDIDataSheet()
    s.bitMd = h.Range("T5").Value
    s.mwdInc = h.Range("U5").Value
    s.mwdAzm = h.Range("V5").Value
    s.pdInc = h.Range("T6").Value
    s.pdAzm = h.Range("U6").Value
    s.holeMd = h.Range("V6").Value
    s.bitInc = h.Range("T7").Value
    s.bitAzm = h.Range("U7").Value
    s.bitTvd = h.Range("V7").Value
    s.standLen = ContDIDataSheet().Range("X1").Value
    s.pdOff = ws.Range("N2").Value
    s.diOff = ws.Range("P2").Value
    s.tgtMd = ws.Range("T1").Value
    s.tgtInc = ws.Range("T2").Value
    s.tgtAzm = ws.Range("T3").Value
    s.tgtTvd = ws.Range("T4").Value
    CaptureSeed = s
End Function

' Header snapshot for the index row: CaptureSeed's offsets / targets plus the
' tie-on actually shown in B2:E2 (never the next-course projection).
Private Function HeaderSeed(ByVal ws As Worksheet) As ContDISeed
    Dim s As ContDISeed
    s = CaptureSeed(ws)
    s.bitMd = ws.Range("B2").Value
    s.bitInc = ws.Range("C2").Value
    s.bitAzm = ws.Range("D2").Value
    s.bitTvd = ws.Range("E2").Value
    HeaderSeed = s
End Function

' New course: B2:E2 is left as the previous course's tie-on (a real survey)
' until WriteLastSurveyTieOn replaces it with the latest green Slidesheet
' survey. The projected bit only seeds row 8 (last readings carried forward).
Private Sub ApplySeed(ByVal ws As Worksheet, ByRef s As ContDISeed)
    If Not IsEmpty(s.standLen) Then ContDIDataSheet().Range("X1").Value = s.standLen
    If Not IsEmpty(s.pdOff) Then ws.Range("N2").Value = s.pdOff
    If Not IsEmpty(s.diOff) Then ws.Range("P2").Value = s.diOff
    If Not IsEmpty(s.tgtMd) Then ws.Range("T1").Value = s.tgtMd
    If Not IsEmpty(s.tgtInc) Then ws.Range("T2").Value = s.tgtInc
    If Not IsEmpty(s.tgtAzm) Then ws.Range("T3").Value = s.tgtAzm
    If Not IsEmpty(s.tgtTvd) Then ws.Range("T4").Value = s.tgtTvd
    ws.Range("B8").Value = s.bitMd
    ws.Range("D8").Value = s.mwdInc
    ws.Range("E8").Value = s.mwdAzm
    ws.Range("G8").Value = s.pdInc
    ws.Range("H8").Value = s.pdAzm
End Sub

Private Sub WriteIndexRow(ByVal dataWs As Worksheet, ByVal idx As Long, ByRef s As ContDISeed)
    Dim r As Long
    r = IDX_FIRST + idx - 1
    dataWs.Cells(r, 1).Value = idx
    dataWs.Cells(r, 2).Value = s.bitMd
    dataWs.Cells(r, 3).Value = s.bitMd
    dataWs.Cells(r, 4).Value = s.bitInc
    dataWs.Cells(r, 5).Value = s.bitAzm
    dataWs.Cells(r, 6).Value = s.bitTvd
    dataWs.Cells(r, 7).Value = s.standLen
    dataWs.Cells(r, 8).Value = s.pdOff
    dataWs.Cells(r, 9).Value = s.diOff
    dataWs.Cells(r, 10).Value = s.tgtMd
    dataWs.Cells(r, 11).Value = s.tgtInc
    dataWs.Cells(r, 12).Value = s.tgtAzm
    dataWs.Cells(r, 13).Value = s.tgtTvd
    dataWs.Cells(r, 14).Value = DumpRowFor(idx)
End Sub

Private Function DumpHasRows(ByVal idx As Long) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = ContDIDataSheet().Cells(DumpRowFor(idx), 1).Value
    DumpHasRows = (Len(Trim$(CStr(v & ""))) > 0)
End Function

Private Sub ArchiveCurrent(ByVal ws As Worksheet, ByVal dataWs As Worksheet, ByVal idx As Long)
    Dim r As Long
    Dim dest As Long
    Dim i As Long
    Dim s As ContDISeed

    If idx < 1 Then Exit Sub
    s = HeaderSeed(ws)
    WriteIndexRow dataWs, idx, s

    dest = DumpRowFor(idx)
    For i = 0 To DATA_ROWS - 1
        r = FIRST_DATA + i
        dataWs.Cells(dest + i, 1).Value = ws.Cells(r, 2).Value   ' B bit
        dataWs.Cells(dest + i, 2).Value = ws.Cells(r, 4).Value   ' D MWD inc
        dataWs.Cells(dest + i, 3).Value = ws.Cells(r, 5).Value   ' E MWD azm
        dataWs.Cells(dest + i, 4).Value = ws.Cells(r, 7).Value   ' G PD inc
        dataWs.Cells(dest + i, 5).Value = ws.Cells(r, 8).Value   ' H PD azm
        dataWs.Cells(dest + i, 6).Value = ws.Cells(r, COL_ROP).Value  ' W ROP
    Next i
End Sub

Private Sub LoadCourseIntoGrid(ByVal idx As Long)
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim r As Long
    Dim dest As Long
    Dim i As Long
    Dim ir As Long

    Set ws = ContDISheet()
    Set dataWs = ContDIDataSheet()
    ir = IDX_FIRST + idx - 1

    ws.Range("B2").Value = dataWs.Cells(ir, 3).Value
    ws.Range("C2").Value = dataWs.Cells(ir, 4).Value
    ws.Range("D2").Value = dataWs.Cells(ir, 5).Value
    ws.Range("E2").Value = dataWs.Cells(ir, 6).Value
    If Len(Trim$(CStr(dataWs.Cells(ir, 7).Value & ""))) > 0 Then dataWs.Range("X1").Value = dataWs.Cells(ir, 7).Value
    If Len(Trim$(CStr(dataWs.Cells(ir, 8).Value & ""))) > 0 Then ws.Range("N2").Value = dataWs.Cells(ir, 8).Value
    If Len(Trim$(CStr(dataWs.Cells(ir, 9).Value & ""))) > 0 Then ws.Range("P2").Value = dataWs.Cells(ir, 9).Value
    If Len(Trim$(CStr(dataWs.Cells(ir, 10).Value & ""))) > 0 Then ws.Range("T1").Value = dataWs.Cells(ir, 10).Value
    If Len(Trim$(CStr(dataWs.Cells(ir, 11).Value & ""))) > 0 Then ws.Range("T2").Value = dataWs.Cells(ir, 11).Value
    If Len(Trim$(CStr(dataWs.Cells(ir, 12).Value & ""))) > 0 Then ws.Range("T3").Value = dataWs.Cells(ir, 12).Value
    If Len(Trim$(CStr(dataWs.Cells(ir, 13).Value & ""))) > 0 Then ws.Range("T4").Value = dataWs.Cells(ir, 13).Value

    ClearInputs ws
    dest = DumpRowFor(idx)
    For i = 0 To DATA_ROWS - 1
        r = FIRST_DATA + i
        ws.Cells(r, 2).Value = dataWs.Cells(dest + i, 1).Value
        ws.Cells(r, 4).Value = dataWs.Cells(dest + i, 2).Value
        ws.Cells(r, 5).Value = dataWs.Cells(dest + i, 3).Value
        ws.Cells(r, 7).Value = dataWs.Cells(dest + i, 4).Value
        ws.Cells(r, 8).Value = dataWs.Cells(dest + i, 5).Value
        ws.Cells(r, COL_ROP).Value = dataWs.Cells(dest + i, 6).Value
    Next i
End Sub

Private Sub ClearInputs(ByVal ws As Worksheet)
    ws.Range("B8:B44").ClearContents
    ws.Range("D8:E44").ClearContents
    ws.Range("G8:H44").ClearContents
    ws.Range("W8:W44").ClearContents
End Sub

' ================================================================================
'  LOCK / NAV
' ================================================================================

' Input columns (B bit depth, D:E MWD, G:H PD, U ROP) stay unlocked on every
' course so an archived course can be corrected under protection; edits are
' archived when the user spins away (ContDI_LoadCourse). Header inputs
' (tie-on, stand length, offsets, targets) are editable on the live course only.
Private Sub ApplyContDILock(ByVal ws As Worksheet, ByVal live As Boolean)
    ws.Cells.Locked = True
    ws.Range("B2:E2").Locked = Not live
    ws.Range("N2").Locked = Not live
    ws.Range("P2").Locked = Not live
    ws.Range("T1:T4").Locked = Not live
    ws.Range("B8:B44").Locked = False
    ws.Range("D8:E44").Locked = False
    ws.Range("G8:H44").Locked = False
    ws.Range("W8:W44").Locked = False
End Sub

Private Sub UpdateNavigator(ByVal ws As Worksheet)
    Dim n As Long
    Dim v As Long
    Dim lbl As String
    n = CourseCount()
    v = ViewIdx()
    If n < 1 Then
        lbl = "Course 0 of 0"
    Else
        lbl = "Course " & v & " of " & n
        If Len(Trim$(CStr(ws.Range("B2").Value & ""))) > 0 Then
            lbl = lbl & " - " & Format$(ws.Range("B2").Value, "0") & "m"
        End If
    End If
    ws.Range("A1").Value = lbl
    If n >= 1 And v < n Then
        ' Archived course: amber banner so view-only state is unmissable.
        ws.Range("A1").Interior.Color = RGB(255, 192, 0)
        With ws.Range("A4")
            .Value = "VIEW ONLY - spin to latest"
            .Font.bold = True
            .Font.Color = RGB(192, 0, 0)
        End With
    Else
        ws.Range("A1").Interior.ColorIndex = xlColorIndexNone
        ws.Range("A4").ClearContents
    End If
End Sub

Private Sub SyncSpinner(ByVal ws As Worksheet, ByVal cur As Long, ByVal maxN As Long)
    Dim sp As Spinner
    On Error Resume Next
    Set sp = ws.Spinners(SPN_NAME)
    On Error GoTo 0
    If sp Is Nothing Then Exit Sub
    If maxN < 1 Then maxN = 1
    If cur < 1 Then cur = 1
    If cur > maxN Then cur = maxN
    sp.Min = 1
    sp.Max = maxN
    sp.Value = cur
End Sub

Private Sub AddContDIControls(ByVal ws As Worksheet)
    Dim btn As Button
    Dim sp As Spinner

    On Error Resume Next
    ws.Buttons(BTN_NEW).Delete
    ws.Buttons(BTN_PULL).Delete
    ws.Spinners(SPN_NAME).Delete
    On Error GoTo 0

    EnsurePullSurveyButton ws

    Set btn = ws.Buttons.Add(ws.Range("F1").Left + 3, ws.Range("F1").Top + 3, _
                             ws.Range("F1:H1").Width - 6, ws.Range("F1:F2").Height - 6)
    btn.name = BTN_NEW
    btn.caption = "New Course"
    btn.Font.bold = True
    btn.OnAction = "'" & ThisWorkbook.name & "'!ContDI_NewCourse"
    btn.Placement = xlMoveAndSize

    Set sp = ws.Spinners.Add(ws.Range("I1").Left + 3, ws.Range("I1").Top + 3, 16, _
                             ws.Range("I1:I2").Height - 6)
    sp.name = SPN_NAME
    sp.OnAction = "'" & ThisWorkbook.name & "'!ContDI_Spin"
    sp.Min = 1
    sp.Max = Application.Max(1, CourseCount())
    sp.Value = Application.Max(1, ViewIdx())
    sp.SmallChange = 1
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
Private Function cHdrRed() As Long: cHdrRed = RGB(192, 0, 0): End Function
Private Function cGridLine() As Long: cGridLine = RGB(150, 150, 150): End Function
Private Function cHelper() As Long: cHelper = RGB(128, 128, 128): End Function

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

Private Sub PaintContDI(ByVal ws As Worksheet)
    Dim wasProt As Boolean
    Dim deg As String
    Dim degRate As String
    Dim c As Range

    wasProt = SheetUnprotectForVba(ws)
    ws.Cells.Clear
    On Error Resume Next
    ws.Cells.UnMerge
    ws.Cells.FormatConditions.Delete
    ws.Buttons.Delete
    ws.Spinners.Delete
    On Error GoTo 0

    deg = "(" & ChrW$(176) & ")"
    degRate = "(" & ChrW$(176) & "/30m)"

    ' -- course navigator label ---------------------------------------------------
    With ws.Range("A1")
        .Value = "Course 0 of 0"
        .Font.bold = True
        .Font.Size = 12
    End With

    ' -- last survey tie-on block ---------------------------------------------------
    ws.Range("B1").Value = "SD"
    ws.Range("C1").Value = "Inc."
    ws.Range("D1").Value = "Azimuth"
    ws.Range("E1").Value = "TVD"
    With ws.Range("B1:E1")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBandLt()
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("A2").Value = "Last MWD Survey"
    ws.Range("A2").Font.bold = True
    With ws.Range("B2:E2")
        .Interior.Color = cInput()
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .numberFormat = "0.00"
    End With

    ' -- button zone backdrop -------------------------------------------------------
    ws.Range("F1:J2").Interior.Color = cBandLt()

    ' -- offsets block ---------------------------------------------------------------
    With ws.Range("M1:P1")
        .Merge
        .Value = "Bit to Sensor Offsets"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("M2").Value = "PD Inc."
    ws.Range("O2").Value = "D&I"
    ws.Range("M2,O2").Font.bold = True
    ws.Range("M2,O2").HorizontalAlignment = xlRight
    ws.Range("N2").Value = 2.25
    ws.Range("P2").Value = 35
    With ws.Range("N2,P2")
        .Interior.Color = cInput()
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .Font.bold = True
    End With

    ' -- target block + required rates (over bit-projection columns S:V) ------------
    ws.Range("S1").Value = "Target MD"
    ws.Range("S2").Value = "Target INC"
    ws.Range("S3").Value = "Target AZM"
    ws.Range("S4").Value = "Target TVD"
    With ws.Range("S1:S4")
        .Font.bold = True
        .Interior.Color = cBandLt()
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("T1:T4")
        .Interior.Color = cTarget()
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With
    ws.Range("U1").Value = "BRR"
    ws.Range("U2").Value = "TRR"
    ws.Range("U3").Value = "DLR"
    ws.Range("U4").Value = "TFR"
    With ws.Range("U1:U4")
        .Font.bold = True
        .Interior.Color = cBandLt()
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("V1:V4")
        .Interior.Color = cGood()
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With

    PaintContDIGridHeaders ws
    HideContDIHelpers ws

    ' -- widths / freeze -------------------------------------------------------------
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B:H").ColumnWidth = 9
    ws.Columns("I:L").ColumnWidth = 9.5
    ws.Columns("M:R").ColumnWidth = 9
    ws.Columns("S").ColumnWidth = 10.5
    ws.Columns("T:W").ColumnWidth = 9.5
    ws.Columns("X").ColumnWidth = 9
    ws.Columns("Y:Z").ColumnWidth = 8
    ws.Columns("AA:AC").ColumnWidth = 8
    On Error Resume Next
    ws.Activate
    ws.Range("B8").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    SheetReprotectAfterVba ws, wasProt
End Sub

' Insert J:K in the data band only (rows 5:44) when the tab still has the
' pre-sensor-DLS layout. Rows 1:4 stay put so Target / BRR / offsets do not move.
Private Sub EnsureInterSensorColumns(ByVal ws As Worksheet)
    Dim hdr As String
    hdr = Trim$(CStr(ws.Range("J6").Value & ""))
    If StrComp(hdr, HDR_PD_DLS, vbTextCompare) <> 0 Then Exit Sub
    On Error Resume Next
    If Not ActiveWindow Is Nothing Then
        If ActiveWindow.FreezePanes Then ActiveWindow.FreezePanes = False
    End If
    ws.Range("I4:P4").UnMerge
    ws.Range("B5:U5").UnMerge
    ws.Range("I6:U7").UnMerge
    On Error GoTo 0
    ws.Range("J5:K44").Insert Shift:=xlToRight
    Err.Clear
End Sub

' Labels / bands / input fill for I:W. Safe to re-run; does not touch B2:E2,
' B3, N2, P2, T1:T4, or typed B/D/E/G/H/W values.
Private Sub PaintContDIGridHeaders(ByVal ws As Worksheet)
    Dim deg As String, degRate As String
    Dim c As Range
    deg = "(" & ChrW$(176) & ")"
    degRate = "(" & ChrW$(176) & "/30m)"

    On Error Resume Next
    ws.Range("B5:W5").UnMerge
    ws.Range("J4:K4").UnMerge
    ws.Range("M4:N4").UnMerge
    ws.Range("O4:R4").UnMerge
    ws.Range("S5:W5").UnMerge
    On Error GoTo 0

    ' Averages sit on M3:R3 — directly over PD / MWD rate columns M:R.
    With ws.Range("M3:R3")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With
    ws.Range("K3:R3").FormatConditions.Delete
    For Each c In ws.Range("M3,O3,Q3").Cells
        AddPaceCF c, "$V$1"
    Next c
    For Each c In ws.Range("N3,P3,R3").Cells
        AddPaceCF c, "$V$2"
    Next c

    With ws.Range("J4:K4")
        .Merge
        .Value = "Between sensors"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("M4:N4")
        .Merge
        .Value = "PD Average"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("O4:R4")
        .Merge
        .Value = "MWD Average"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range("B5:H5")
        .Merge
        .Value = "Continuous D & I"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("I5").Value = "f/ survey"
    ws.Range("J5").Value = "btw sens."
    ws.Range("K5").Value = "btw sens."
    ws.Range("L5").Value = "f/ survey"
    ws.Range("M5").Value = "f/ cont."
    ws.Range("N5").Value = "f/ cont."
    ws.Range("O5").Value = "f/ cont."
    ws.Range("P5").Value = "f/ cont."
    ws.Range("Q5").Value = "f/ survey"
    ws.Range("R5").Value = "f/ survey"
    With ws.Range("I5:R5")
        .Font.Size = 9
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("S5:W5")
        .Merge
        .Value = "Bit Projection f/ last survey"
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = cBand()
        .Borders.LineStyle = xlContinuous
    End With

    ws.Range("B6").Value = "Bit Depth"
    ws.Range("C6").Value = "MWD SD"
    ws.Range("D6").Value = "Inc."
    ws.Range("E6").Value = "Azi."
    ws.Range("F6").Value = "PD SD"
    ws.Range("G6").Value = "PD Inc."
    ws.Range("H6").Value = "PD Azm."
    ws.Range("I6").Value = "MWD DLS"
    ws.Range("J6").Value = HDR_SENS_DLS
    ws.Range("K6").Value = "Sens BUR"
    ws.Range("L6").Value = HDR_PD_DLS
    ws.Range("M6").Value = "BUR"
    ws.Range("N6").Value = "TR"
    ws.Range("O6").Value = "BUR"
    ws.Range("P6").Value = "TR"
    ws.Range("Q6").Value = "BUR"
    ws.Range("R6").Value = "TR"
    ws.Range("S6").Value = "Hole Depth"
    ws.Range("T6").Value = "Inc. @ Bit"
    ws.Range("U6").Value = "Az. @ Bit"
    ws.Range("V6").Value = "TVD @ Bit"
    ws.Range("W6").Value = "ROP"
    ws.Range("B7").Value = "(m)"
    ws.Range("C7").Value = "(m)"
    ws.Range("D7").Value = deg
    ws.Range("E7").Value = deg
    ws.Range("F7").Value = "(m)"
    ws.Range("G7").Value = deg
    ws.Range("H7").Value = deg
    ws.Range("I7").Value = "f/ survey"
    ws.Range("J7").Value = "PD-D&I"
    ws.Range("K7").Value = degRate
    ws.Range("L7").Value = "f/ survey"
    ws.Range("M7").Value = degRate
    ws.Range("N7").Value = degRate
    ws.Range("O7").Value = degRate
    ws.Range("P7").Value = degRate
    ws.Range("Q7").Value = degRate
    ws.Range("R7").Value = degRate
    ws.Range("S7").Value = "(m)"
    ws.Range("T7").Value = deg
    ws.Range("U7").Value = deg
    ws.Range("V7").Value = "(m)"
    ws.Range("W7").Value = "(m/hr)"
    With ws.Range("B6:W7")
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws.Range("B7:W7").Font.Size = 9
    With ws.Range("I6:L7")
        .Interior.Color = cHdrRed()
        .Font.Color = RGB(255, 255, 0)
    End With
    ws.Range("B6:H7").Interior.Color = cBandLt()
    ws.Range("M6:R7").Interior.Color = cBandLt()
    ws.Range("S6:W7").Interior.Color = cBandLt()

    With ws.Range("B8:W44")
        .numberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = cGridLine()
    End With
    ws.Range("B8:B44,D8:E44,G8:H44,W8:W44").Interior.Color = cInput()
    ' Keep the yellow target labels if a prior layout pass cleared them.
    If Len(Trim$(CStr(ws.Range("S1").Value & ""))) = 0 Then ws.Range("S1").Value = "Target MD"
    If Len(Trim$(CStr(ws.Range("S2").Value & ""))) = 0 Then ws.Range("S2").Value = "Target INC"
    If Len(Trim$(CStr(ws.Range("S3").Value & ""))) = 0 Then ws.Range("S3").Value = "Target AZM"
    If Len(Trim$(CStr(ws.Range("S4").Value & ""))) = 0 Then ws.Range("S4").Value = "Target TVD"
    If Len(Trim$(CStr(ws.Range("U1").Value & ""))) = 0 Then ws.Range("U1").Value = "BRR"
    If Len(Trim$(CStr(ws.Range("U2").Value & ""))) = 0 Then ws.Range("U2").Value = "TRR"
    If Len(Trim$(CStr(ws.Range("U3").Value & ""))) = 0 Then ws.Range("U3").Value = "DLR"
    If Len(Trim$(CStr(ws.Range("U4").Value & ""))) = 0 Then ws.Range("U4").Value = "TFR"
    ws.Range("S1:V4").FormatConditions.Delete
    With ws.Range("S1:S4,U1:U4")
        .Font.bold = True
        .Interior.Color = cBandLt()
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("T1:T4")
        .Interior.Color = cTarget()
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With
    With ws.Range("V1:V4")
        .Interior.Color = cGood()
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .numberFormat = "0.00"
    End With
    Err.Clear
End Sub

' Park every diagnostic lookup on the very-hidden ContDI_Data sheet and wipe
' leftover on-grid helper cells (W1:X4 / X5:AC7) so only the working tab shows.
Private Sub HideContDIHelpers(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range("A3:C3").Clear
    ws.Range("W1:X4").Clear
    ws.Range("Y1:Z2").Clear
    ws.Range("X5:AC7").Clear
    ws.Range("AA1:AC7").Clear
    ws.Columns("X").Hidden = True
    ws.Columns("Y:Z").Hidden = True
    ws.Columns("AA:AC").Hidden = True
    On Error GoTo 0
    Err.Clear
End Sub

' One-time +2 column shift: Q1:T4 (targets + rates) -> S1:V4 so averages
' can sit on M3:R3 over columns M:R. Idempotent once S1 is already Target MD.
Private Sub EnsureHeaderShift(ByVal ws As Worksheet)
    Dim q1 As String
    q1 = Trim$(CStr(ws.Range("Q1").Value & ""))
    If StrComp(q1, "Target MD", vbTextCompare) = 0 Then
        ws.Range("Q1:T4").Copy Destination:=ws.Range("S1")
        Application.CutCopyMode = False
        ws.Range("Q1:R4").Clear
    End If
    ' Old averages sat on K3:P3 (formulas already pointed at M9:R44).
    ' Do not Copy — relative AVERAGE refs would slide +2. Just drop leftovers.
    If InStr(1, UCase$(CStr(ws.Range("K3").Formula)), "AVERAGE(M9", vbBinaryCompare) > 0 Then
        ws.Range("K3:L3").Clear
    End If
    Err.Clear
End Sub

Private Sub ApplyLookupsAndRates(ByVal ws As Worksheet)
    ws.Range("K3:L3").Clear
    ws.Range("M3").Formula = "=IFERROR(AVERAGE(M9:M44),"""")"
    ws.Range("N3").Formula = "=IFERROR(AVERAGE(N9:N44),"""")"
    ws.Range("O3").Formula = "=IFERROR(AVERAGE(O9:O44),"""")"
    ws.Range("P3").Formula = "=IFERROR(AVERAGE(P9:P44),"""")"
    ws.Range("Q3").Formula = "=IFERROR(AVERAGE(Q9:Q44),"""")"
    ws.Range("R3").Formula = "=IFERROR(AVERAGE(R9:R44),"""")"
    ws.Range("M3:R3").numberFormat = "0.00"

    ' Diagnostics live on very-hidden ContDI_Data (T:W), not on the grid.
    Dim h As Worksheet
    Dim cd As String
    Set h = ContDIDataSheet()
    cd = "'" & SH_CONTDI & "'!"
    h.Range("T1").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "B8:B44<>"""")," & cd & "B8:B44),IF(ISNUMBER(W1),W1,IF(AND(ISNUMBER(" & cd & "B2),ISNUMBER(" & cd & "P2))," & cd & "B2+" & cd & "P2,IF(" & cd & "B2="""",""""," & cd & "B2))))"
    h.Range("T2").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "D8:D44<>"""")," & cd & "D8:D44),IF(" & cd & "C2="""",""""," & cd & "C2))"
    h.Range("T3").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "E8:E44<>"""")," & cd & "E8:E44),IF(" & cd & "D2="""",""""," & cd & "D2))"
    h.Range("T4").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "T8:T44<>"""")," & cd & "T8:T44),IF(ISNUMBER(W2),W2,IF(" & cd & "C2="""",""""," & cd & "C2)))"
    h.Range("U1").Formula = "=IFERROR(" & cd & "T1-T1,"""")"
    h.Range("U2").Formula = "=IFERROR((" & cd & "T2-T4)/U1,"""")"
    h.Range("U3").Formula = "=IFERROR((MOD(" & cd & "T3-U4+180,360)-180)/U1,"""")"
    h.Range("U4").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "U8:U44<>"""")," & cd & "U8:U44),IF(ISNUMBER(W3),W3,IF(" & cd & "D2="""",""""," & cd & "D2)))"
    h.Range("V1").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "V8:V44<>"""")," & cd & "V8:V44),IF(" & cd & "E2="""",""""," & cd & "E2))"
    h.Range("T5").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "B8:B44<>"""")," & cd & "B8:B44),"""")"
    h.Range("U5").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "D8:D44<>"""")," & cd & "D8:D44),"""")"
    h.Range("V5").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "E8:E44<>"""")," & cd & "E8:E44),"""")"
    h.Range("T6").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "G8:G44<>"""")," & cd & "G8:G44),"""")"
    h.Range("U6").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "H8:H44<>"""")," & cd & "H8:H44),"""")"
    h.Range("V6").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "S8:S44<>"""")," & cd & "S8:S44),"""")"
    h.Range("T7").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "T8:T44<>"""")," & cd & "T8:T44),"""")"
    h.Range("U7").Formula = "=IFERROR(LOOKUP(2,1/(" & cd & "U8:U44<>"""")," & cd & "U8:U44),"""")"
    h.Range("V7").Formula = "=IFERROR(ROUND(LOOKUP(2,1/(" & cd & "V8:V44<>"""")," & cd & "V8:V44),2),"""")"

    ' V1 = Slidesheet plan BURR (next named station). V2/V4 = yellow box.
    ' V3 = 3D dogleg to the same yellow box as V2 (land), not a hypot of V1.
    ws.Range("V1").Formula = "=IFERROR(ROUND(ProjBurr(ContDI_Data!T1,ContDI_Data!T4,ContDI_Data!U4,0,IF(ISNUMBER(ContDI_Data!X1),ContDI_Data!X1,19.2),ProjTargets_MD,ProjTargets_INC,ProjTargets_AZM,ProjTargets_TVD,ContDI_Data!V1),2),IFERROR(ROUND(ABS(T2-ContDI_Data!T4)/ContDI_Data!U1*30,2),""""))"
    ws.Range("V2").Formula = "=IFERROR(ROUND(ABS(MOD(T3-ContDI_Data!U4+180,360)-180)/ContDI_Data!U1*30,2),"""")"
    ws.Range("V3").Formula = "=IFERROR(ROUND(ProjDoglegDeg(ContDI_Data!T4,ContDI_Data!U4,T2,T3)/ContDI_Data!U1*30,2),"""")"
    ws.Range("V4").Formula = "=IFERROR(ProjTfToTarget(ContDI_Data!T4,ContDI_Data!U4,T2,T3,5),"""")"
    HideContDIHelpers ws
End Sub

' Re-apply every formula cell (grid rows 8:44 + header lookups / rates) on a
' painted tab without touching inputs (B2:E2, B3, N2, P2, T1:T4,
' B/D/E/G/H/W 8:44). Inserts J:K once if the old PD-DLS-at-J layout is still on.
' Shifts Q1:T4 -> S1:V4 and K3:P3 leftovers once if the old header is present.
Public Sub ContDI_RefreshFormulas()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    Dim ev As Boolean
    On Error GoTo ErrHandler
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Set ws = ContDISheet()
    wasProt = SheetUnprotectForVba(ws)
    EnsureInterSensorColumns ws
    Err.Clear
    EnsureHeaderShift ws
    Err.Clear
    PaintContDIGridHeaders ws
    Err.Clear
    ApplyRowFormulas ws
    Err.Clear
    SeedProjectionHelpers ws
    Err.Clear
    ApplyLookupsAndRates ws
    Err.Clear
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
    ' Every rate / dogleg cell tests exactly the inputs it uses with ISNUMBER
    ' (this row, the previous row for f/ cont., and the tie-on) so a blank or
    ' half-typed D/E/G/H yields "" instead of #VALUE!. A missing azimuth blanks
    ' the TR and the dogleg but not the BUR, and vice versa.
    Dim tie As String, dMWD As String, dPD As String
    Dim nD As String, nE As String, nG As String, nH As String
    Dim pD As String, pE As String, pG As String, pH As String
    Dim stepMWD As String, stepPD As String
    tie = "ISNUMBER($B$2)"
    For r = FIRST_DATA To LAST_DATA
        p = r - 1
        nD = "ISNUMBER(D" & r & ")": nE = "ISNUMBER(E" & r & ")"
        nG = "ISNUMBER(G" & r & ")": nH = "ISNUMBER(H" & r & ")"
        pD = "ISNUMBER(D" & p & ")": pE = "ISNUMBER(E" & p & ")"
        pG = "ISNUMBER(G" & p & ")": pH = "ISNUMBER(H" & p & ")"
        dMWD = "ISNUMBER(C" & r & "),C" & r & ">$B$2"      ' MWD sensor below the survey
        dPD = "ISNUMBER(F" & r & "),F" & r & ">$B$2"       ' PD sensor below the survey
        stepMWD = "ISNUMBER(C" & r & "),ISNUMBER(C" & p & "),C" & r & "<>C" & p
        stepPD = "ISNUMBER(F" & r & "),ISNUMBER(F" & p & "),F" & r & "<>F" & p
        ws.Cells(r, 3).Formula = "=IF(ISNUMBER(B" & r & "),B" & r & "-$P$2,"""")"
        ws.Cells(r, 6).Formula = "=IF(ISNUMBER(B" & r & "),B" & r & "-$N$2,"""")"
        ' I: MWD DLS f/ survey. L: PD DLS f/ survey.
        ws.Cells(r, 9).Formula = "=IF(AND(" & tie & "," & nD & "," & nE & "," & dMWD & "),(30/(C" & r & "-$B$2))*DEGREES(ACOS(MIN(1,MAX(-1,COS(RADIANS($C$2))*COS(RADIANS(D" & r & "))+SIN(RADIANS($C$2))*SIN(RADIANS(D" & r & "))*COS(RADIANS(E" & r & "-$D$2)))))),"""")"
        ws.Cells(r, COL_PD_DLS).Formula = "=IF(AND(" & tie & "," & nG & "," & nH & "," & dPD & "),(30/(F" & r & "-$B$2))*DEGREES(ACOS(MIN(1,MAX(-1,COS(RADIANS($C$2))*COS(RADIANS(G" & r & "))+SIN(RADIANS($C$2))*SIN(RADIANS(G" & r & "))*COS(RADIANS(H" & r & "-$D$2)))))),"""")"
        ' J / K: dogleg and signed BUR between the PD sensor (G,H) and the D&I
        ' (D,E) over P2-N2. Blank either sensor -> "".
        ws.Cells(r, COL_SENS_DLS).Formula = "=IF(AND(" & nD & "," & nE & "," & nG & "," & nH & ",ISNUMBER($P$2),ISNUMBER($N$2),$P$2<>$N$2),ProjDoglegDeg(D" & r & ",E" & r & ",G" & r & ",H" & r & ")/ABS($P$2-$N$2)*30,"""")"
        ws.Cells(r, COL_SENS_BUR).Formula = "=IF(AND(" & nD & "," & nG & ",ISNUMBER($P$2),ISNUMBER($N$2),$P$2<>$N$2),(G" & r & "-D" & r & ")/($P$2-$N$2)*30,"""")"
        If r = FIRST_DATA Then
            ws.Cells(r, COL_PD_BUR).ClearContents
            ws.Cells(r, COL_PD_TR).ClearContents
            ws.Cells(r, COL_MWD_BURC).ClearContents
            ws.Cells(r, COL_MWD_TRC).ClearContents
        End If
        If r >= 9 Then
            ws.Cells(r, COL_PD_BUR).Formula = "=IF(AND(" & nG & "," & pG & "," & stepPD & "),(G" & r & "-G" & p & ")/(F" & r & "-F" & p & ")*30,"""")"
            ws.Cells(r, COL_PD_TR).Formula = "=IF(AND(" & nH & "," & pH & "," & stepPD & "),(MOD(H" & r & "-H" & p & "+180,360)-180)/(F" & r & "-F" & p & ")*30,"""")"
            ws.Cells(r, COL_MWD_BURC).Formula = "=IF(AND(" & nD & "," & pD & "," & stepMWD & "),(D" & r & "-D" & p & ")/(C" & r & "-C" & p & ")*30,"""")"
            ws.Cells(r, COL_MWD_TRC).Formula = "=IF(AND(" & nE & "," & pE & "," & stepMWD & "),(MOD(E" & r & "-E" & p & "+180,360)-180)/(C" & r & "-C" & p & ")*30,"""")"
        End If
        ws.Cells(r, COL_MWD_BURS).Formula = "=IF(AND(" & tie & "," & nD & "," & dMWD & "),(D" & r & "-$C$2)/(C" & r & "-$B$2)*30,"""")"
        ws.Cells(r, COL_MWD_TRS).Formula = "=IF(AND(" & tie & "," & nE & "," & dMWD & "),ABS(MOD(E" & r & "-$D$2+180,360)-180)/(C" & r & "-$B$2)*30,"""")"
        ws.Cells(r, COL_HOLE).Formula = "=IF(ISNUMBER(B" & r & "),B" & r & ","""")"
        ' Inc / Az @ Bit: carry the reading over the D&I offset (P2) at the
        ' rate seen since the tie-on (Q = BUR f/ survey).
        ws.Cells(r, COL_INC_BIT).Formula = "=IF(NOT(" & nD & "),"""",IF(ISNUMBER(Q" & r & "),(Q" & r & "/30)*P$2+$D" & r & ",$D" & r & "))"
        ws.Cells(r, COL_AZ_BIT).Formula = "=IF(NOT(" & nE & "),"""",IF(AND(" & tie & "," & dMWD & "),MOD(((MOD(E" & r & "-$D$2+180,360)-180)/(C" & r & "-$B$2))*P$2+E" & r & ",360),MOD(E" & r & ",360)))"
        ws.Cells(r, COL_TVD_BIT).Formula = "=IF(AND(" & nD & ",ISNUMBER($S" & r & "),ISNUMBER($T" & r & "),ISNUMBER(B$2),ISNUMBER(E$2)),(($S" & r & "-B$2)*COS((RADIANS($T" & r & ")+RADIANS($D" & r & "))/2))+E$2,"""")"
    Next r
End Sub




