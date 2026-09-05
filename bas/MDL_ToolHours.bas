Attribute VB_Name = "MDL_ToolHours"
Option Explicit

' ================================================================================
'  MDL_ToolHours - Data tab 3rd-party + motor hours
' ================================================================================
' Front tables (equal share of O28:S55):
'   O28:S29 title   "Enter any 3rd Party Tools/hours below"
'   O30             headers
'   O31:S41         11 tool rows  (active BHA only)
'   O42:S43 title   "Enter Motors & hours below"
'   O44             headers
'   O45:S55         11 motor rows (active BHA only)
'
' Hidden tracker _TH_Hours (VeryHidden): every tool/motor seen on the job.
'   A Serial | B Kind | C PreJob | D JobHours | E Total | F BhaList
'   M2:M50   motor serials (dropdown source for Motors On Location)
'
' Hours:
'   Q Previous = PreJob (typed once; auto-filled when a serial is reused)
'   R Current  = this-job hours: sum of BHA Total Hrs for every BHA that
'                carried the serial, with the selected BHA using live Q24
'   S Total    = Q + R
'
' Only FieldCap Other Tools on the selected BHA (H3): Source/Category
' "Other Inventory"/"Other Tools", or SubCategory "DD other".
' DD inventory (crossovers, rentals, MWD), drill bits, and tubulars stay out.
' Orbit RSS and iCruise seed into Motors.
' ================================================================================

Private Const SH_DATA As String = "Data"
Private Const SH_BHA As String = "_OC_BHA"
Private Const SH_INVENTORY As String = "_OC_Inventory"
Private Const SH_TRACK As String = "_TH_Hours"
Private Const OTHER_INVENTORY_TAG As String = "Other Inventory"
Private Const OTHER_TOOLS_TAG As String = "Other Tools"
Private Const KIND_TOOL As String = "Tool"
Private Const KIND_MOTOR As String = "Motor"

Private Const CELL_BHA As String = "H3"
Private Const CELL_HOURS As String = "Q24"
Private Const CELL_MOTOR As String = "C22"
Private Const CELL_PREV_MOTOR As String = "C31"
Private Const RNG_PICKS As String = "D24:D32"
Private Const RNG_HOURS_SRC As String = "P4:Q22"
Private Const RNG_MOL_SN As String = "B45:B55"

Private Const COL_SERIAL As Long = 15   ' O
Private Const COL_PREV As Long = 17     ' Q
Private Const COL_CURRENT As Long = 18  ' R
Private Const COL_TOTAL As Long = 19    ' S

Private Const TOOLS_TITLE_ROW As Long = 28
Private Const TOOLS_HEADER_ROW As Long = 30
Private Const TOOLS_FIRST As Long = 31
Private Const TOOLS_LAST As Long = 41
Private Const MOTOR_TITLE_ROW As Long = 42
Private Const MOTOR_HEADER_ROW As Long = 44
Private Const MOTOR_FIRST As Long = 45
Private Const MOTOR_LAST As Long = 55

Private Const TOOLS_TITLE_TOKEN As String = "3rd Party"
Private Const MOTOR_TITLE_TOKEN As String = "Enter Motors"
Private Const TITLE_TYPO As String = "Toosl"
Private Const TITLE_FIX As String = "Tools"

Private Const TR_SERIAL As Long = 1
Private Const TR_KIND As Long = 2
Private Const TR_PREJOB As Long = 3
Private Const TR_JOB As Long = 4
Private Const TR_TOTAL As Long = 5
Private Const TR_BHAS As Long = 6
Private Const TR_MOTOR_LIST As Long = 13  ' M

Private mBusy As Boolean

Public Sub ToolHours_OnDataChange(ByVal Target As Range)
    Dim ws As Worksheet
    Dim watch As Range

    On Error GoTo Quiet
    If mBusy Then Exit Sub
    If Target Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_DATA)
    If Not Target.Worksheet Is ws Then Exit Sub

    Set watch = Union(ws.Range(CELL_BHA), ws.Range(CELL_MOTOR), _
                      ws.Range(RNG_PICKS), ws.Range(RNG_HOURS_SRC), _
                      ws.Range(ws.Cells(TOOLS_FIRST, COL_SERIAL), ws.Cells(MOTOR_LAST, COL_SERIAL)), _
                      ws.Range(ws.Cells(TOOLS_FIRST, COL_PREV), ws.Cells(MOTOR_LAST, COL_PREV)))
    If Intersect(Target, watch) Is Nothing Then Exit Sub

    ToolHours_Sync
    Exit Sub
Quiet:
End Sub

Public Sub ToolHours_SeedFromInventory()
    ToolHours_Sync
End Sub

Public Sub ToolHours_Sync()
    Dim ws As Worksheet, tr As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean
    Dim hrs As Double
    Dim bha As Long

    If mBusy Then Exit Sub
    If Not SheetExistsTH(SH_DATA) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_DATA)

    mBusy = True
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Fail

    On Error Resume Next
    Application.Calculate
    On Error GoTo Fail

    hrs = NumOrZero(ws.Range(CELL_HOURS).Value)
    bha = BhaNumber(ws.Range(CELL_BHA).Value)

    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo FailProt

    FixTitleSpelling ws
    HarvestFrontPrevious ws
    EnsureLayout ws
    Set tr = EnsureTracker()
    HarvestFrontPrevious ws
    UpsertKnownSerials tr
    RecomputeJobHours tr, bha, hrs
    PaintFrontTools ws, tr, bha
    PaintFrontMotors ws, tr, bha
    WriteMotorList tr
    FixDependentFormulas ws
    UnlockFrontTables ws

    SheetReprotectAfterVba ws, wasProt
    Application.EnableEvents = prevEvents
    mBusy = False
    Exit Sub

FailProt:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
Fail:
    On Error Resume Next
    Application.EnableEvents = prevEvents
    mBusy = False
End Sub

' ------------------------------------------------------------------------------
'  Hidden tracker
' ------------------------------------------------------------------------------

Private Function EnsureTracker() As Worksheet
    Dim ws As Worksheet
    If SheetExistsTH(SH_TRACK) Then
        Set ws = ThisWorkbook.Worksheets(SH_TRACK)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SH_TRACK
    End If
    ws.Visible = xlSheetVeryHidden
    If StrComp(CStr(ws.Cells(1, TR_SERIAL).Value & ""), "Serial", vbTextCompare) <> 0 Then
        ws.Cells(1, TR_SERIAL).Value = "Serial"
        ws.Cells(1, TR_KIND).Value = "Kind"
        ws.Cells(1, TR_PREJOB).Value = "PreJobHours"
        ws.Cells(1, TR_JOB).Value = "JobHours"
        ws.Cells(1, TR_TOTAL).Value = "Total"
        ws.Cells(1, TR_BHAS).Value = "BhaList"
        ws.Cells(1, TR_MOTOR_LIST).Value = "MotorList"
    End If
    Set EnsureTracker = ws
End Function

Private Function TrackerLastRow(ByVal tr As Worksheet) As Long
    Dim r As Long
    r = tr.Cells(tr.Rows.Count, TR_SERIAL).End(xlUp).Row
    If r < 1 Then r = 1
    TrackerLastRow = r
End Function

Private Function TrackerFindRow(ByVal tr As Worksheet, ByVal sn As String) As Long
    Dim r As Long, lastR As Long, key As String
    TrackerFindRow = 0
    key = NormSerial(sn)
    If Len(key) = 0 Then Exit Function
    lastR = TrackerLastRow(tr)
    For r = 2 To lastR
        If NormSerial(CStr(tr.Cells(r, TR_SERIAL).Value & "")) = key Then
            TrackerFindRow = r
            Exit Function
        End If
    Next r
End Function

Private Function TrackerUpsert(ByVal tr As Worksheet, ByVal sn As String, _
                               ByVal kind As String) As Long
    Dim r As Long
    TrackerUpsert = 0
    If Not LooksLikeSerial(sn) Then Exit Function
    r = TrackerFindRow(tr, sn)
    If r = 0 Then
        r = TrackerLastRow(tr) + 1
        If r < 2 Then r = 2
        tr.Cells(r, TR_SERIAL).numberFormat = "@"
        tr.Cells(r, TR_SERIAL).Value = sn
        tr.Cells(r, TR_PREJOB).Value = 0
        tr.Cells(r, TR_JOB).Value = 0
        tr.Cells(r, TR_TOTAL).Value = 0
    End If
    If Len(kind) > 0 Then tr.Cells(r, TR_KIND).Value = kind
    TrackerUpsert = r
End Function

Private Sub HarvestFrontPrevious(ByVal ws As Worksheet)
    Dim tr As Worksheet
    Set tr = EnsureTracker()
    If LayoutIsNew(ws) Then
        HarvestRange ws, tr, TOOLS_FIRST, TOOLS_LAST, KIND_TOOL
        HarvestRange ws, tr, MOTOR_FIRST, MOTOR_LAST, KIND_MOTOR
    Else
        HarvestRange ws, tr, 33, 38, KIND_TOOL
        HarvestRange ws, tr, 42, 55, KIND_MOTOR
    End If
End Sub

Private Sub HarvestRange(ByVal ws As Worksheet, ByVal tr As Worksheet, _
                         ByVal firstR As Long, ByVal lastR As Long, ByVal kind As String)
    Dim r As Long, trR As Long
    Dim sn As String
    Dim prev As Double
    For r = firstR To lastR
        sn = SerialAt(ws, r)
        If LooksLikeSerial(sn) Then
            prev = NumOrZero(ws.Cells(r, COL_PREV).Value)
            trR = TrackerUpsert(tr, sn, kind)
            If prev > 0 Then tr.Cells(trR, TR_PREJOB).Value = prev
        End If
    Next r
End Sub

Private Function LooksLikeSerial(ByVal sn As String) As Boolean
    LooksLikeSerial = False
    sn = Trim$(sn)
    If Len(NormSerial(sn)) < 3 Then Exit Function
    If InStr(1, sn, "hours below", vbTextCompare) > 0 Then Exit Function
    If StrComp(sn, "Serial Number", vbTextCompare) = 0 Then Exit Function
    LooksLikeSerial = True
End Function

Private Sub DedupTracker(ByVal tr As Worksheet)
    Dim r As Long, other As Long
    Dim sn As String, key As String
    For r = TrackerLastRow(tr) To 2 Step -1
        sn = CanonicalSerial(CStr(tr.Cells(r, TR_SERIAL).Value & ""))
        If LooksLikeSerial(sn) Then
            tr.Cells(r, TR_SERIAL).numberFormat = "@"
            tr.Cells(r, TR_SERIAL).Value = sn
        End If
        key = NormSerial(sn)
        If Len(key) = 0 Then GoTo NextDedup
        For other = 2 To r - 1
            If NormSerial(CStr(tr.Cells(other, TR_SERIAL).Value & "")) = key Then
                If NumOrZero(tr.Cells(r, TR_PREJOB).Value) > NumOrZero(tr.Cells(other, TR_PREJOB).Value) Then
                    tr.Cells(other, TR_PREJOB).Value = tr.Cells(r, TR_PREJOB).Value
                End If
                tr.Rows(r).Delete
                Exit For
            End If
        Next other
NextDedup:
    Next r
End Sub

Private Sub PurgeBadTrackerRows(ByVal tr As Worksheet)
    Dim r As Long
    For r = TrackerLastRow(tr) To 2 Step -1
        If Not LooksLikeSerial(CStr(tr.Cells(r, TR_SERIAL).Value & "")) Then
            tr.Rows(r).Delete
        End If
    Next r
End Sub

Private Sub UpsertKnownSerials(ByVal tr As Worksheet)
    Dim bhaWs As Worksheet
    Dim cNum As Long, cSn As Long, cSrc As Long, cDesc As Long, cSub As Long
    Dim lastR As Long, r As Long
    Dim sn As String, src As String, desc As String, subDesc As String
    Dim cat As String, subCat As String, itemName As String, ship As String

    PurgeBadTrackerRows tr
    DedupTracker tr

    If SheetExistsTH(SH_BHA) Then
        Set bhaWs = ThisWorkbook.Worksheets(SH_BHA)
        cNum = HeaderCol(bhaWs, "BHA #")
        cSn = HeaderCol(bhaWs, "Serial #")
        cSrc = HeaderCol(bhaWs, "Source")
        cDesc = HeaderCol(bhaWs, "Description")
        cSub = HeaderCol(bhaWs, "Sub Description")
        If cSn > 0 Then
            lastR = bhaWs.Cells(bhaWs.Rows.Count, cSn).End(xlUp).Row
            For r = 2 To lastR
                sn = CellSerial(bhaWs.Cells(r, cSn))
                If LooksLikeSerial(sn) Then
                    src = ""
                    desc = ""
                    subDesc = ""
                    If cSrc > 0 Then src = CStr(bhaWs.Cells(r, cSrc).Value & "")
                    If cDesc > 0 Then desc = CStr(bhaWs.Cells(r, cDesc).Value & "")
                    If cSub > 0 Then subDesc = CStr(bhaWs.Cells(r, cSub).Value & "")
                    InventoryMeta sn, cat, subCat, itemName, ship
                    If Len(Trim$(subCat)) = 0 Then subCat = subDesc
                    If IsMotorLike(sn, desc, cat, subCat, itemName) Then
                        TrackerUpsert tr, sn, KIND_MOTOR
                    ElseIf IsThirdPartySeed(src, sn, desc, cat, subCat, itemName) Then
                        TrackerUpsert tr, sn, KIND_TOOL
                    End If
                End If
            Next r
        End If
    End If

    UpsertMotorsFromInventory tr
End Sub

Private Sub UpsertMotorsFromInventory(ByVal tr As Worksheet)
    Dim inv As Worksheet
    Dim cSn As Long, cName As Long, cSub As Long
    Dim lastR As Long, r As Long
    Dim sn As String, itemName As String, subCat As String

    If Not SheetExistsTH(SH_INVENTORY) Then Exit Sub
    Set inv = ThisWorkbook.Worksheets(SH_INVENTORY)
    cSn = HeaderCol(inv, "SerialNumber")
    If cSn = 0 Then cSn = HeaderCol(inv, "Serial #")
    cName = HeaderCol(inv, "ItemName")
    If cName = 0 Then cName = HeaderCol(inv, "Description")
    cSub = HeaderCol(inv, "SubCategory")
    If cSn = 0 Then Exit Sub
    lastR = inv.Cells(inv.Rows.Count, cSn).End(xlUp).Row
    For r = 2 To lastR
        sn = CellSerial(inv.Cells(r, cSn))
        itemName = ""
        subCat = ""
        If cName > 0 Then itemName = CStr(inv.Cells(r, cName).Value & "")
        If cSub > 0 Then subCat = CStr(inv.Cells(r, cSub).Value & "")
        If Len(sn) > 0 Then
            If IsMotorLike(sn, itemName, "", subCat, itemName) Then
                TrackerUpsert tr, sn, KIND_MOTOR
            End If
        End If
    Next r
End Sub

Private Sub RecomputeJobHours(ByVal tr As Worksheet, ByVal curBha As Long, ByVal liveHrs As Double)
    Dim bhaHrs As Collection
    Dim bhaMembers As Collection
    Dim lastR As Long, r As Long
    Dim sn As String, key As String
    Dim job As Double
    Dim bhas As String
    Dim v As Variant

    Set bhaHrs = BhaHoursMap()
    Set bhaMembers = BhaMembersMap()
    lastR = TrackerLastRow(tr)
    For r = 2 To lastR
        sn = Trim$(CStr(tr.Cells(r, TR_SERIAL).Value & ""))
        key = NormSerial(sn)
        job = 0
        bhas = ""
        If Len(key) > 0 Then
            If InSet(bhaMembers, key) Then
                bhas = CStr(bhaMembers.Item(key))
                job = SumBhaHoursFor(bhas, bhaHrs, curBha, liveHrs)
            End If
        End If
        tr.Cells(r, TR_JOB).Value = job
        tr.Cells(r, TR_BHAS).Value = bhas
        tr.Cells(r, TR_TOTAL).Value = NumOrZero(tr.Cells(r, TR_PREJOB).Value) + job
    Next r
End Sub

' "3,1,2" -> hours of those BHAs; current BHA uses live Q24.
Private Function SumBhaHoursFor(ByVal bhaList As String, ByVal bhaHrs As Collection, _
                                ByVal curBha As Long, ByVal liveHrs As Double) As Double
    Dim parts() As String
    Dim i As Long, n As Long
    Dim seen As New Collection
    Dim tot As Double
    Dim key As String

    tot = 0
    parts = Split(bhaList, ",")
    For i = LBound(parts) To UBound(parts)
        n = BhaNumber(Trim$(parts(i)))
        If n > 0 Then
            key = CStr(n)
            If Not InSet(seen, key) Then
                AddUnique seen, key
                If n = curBha Then
                    tot = tot + liveHrs
                ElseIf InSet(bhaHrs, key) Then
                    tot = tot + CDbl(bhaHrs.Item(key))
                End If
            End If
        End If
    Next i
    SumBhaHoursFor = tot
End Function

' BHA# -> BHA Total Hrs (first row of that BHA).
Private Function BhaHoursMap() As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    Dim cNum As Long, cHrs As Long
    Dim lastR As Long, r As Long
    Dim n As Long
    Dim key As String

    Set BhaHoursMap = col
    If Not SheetExistsTH(SH_BHA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    cNum = HeaderCol(ws, "BHA #")
    cHrs = HeaderCol(ws, "BHA Total Hrs")
    If cNum = 0 Or cHrs = 0 Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, cNum).End(xlUp).Row
    For r = 2 To lastR
        n = BhaNumber(ws.Cells(r, cNum).Value)
        If n > 0 Then
            key = CStr(n)
            If Not InSet(col, key) Then
                On Error Resume Next
                col.Add NumOrZero(ws.Cells(r, cHrs).Value), key
                On Error GoTo 0
            End If
        End If
    Next r
End Function

' normSerial -> comma-separated BHA# list
Private Function BhaMembersMap() As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    Dim cNum As Long, cSn As Long
    Dim lastR As Long, r As Long
    Dim n As Long
    Dim sn As String, key As String, cur As String

    Set BhaMembersMap = col
    If Not SheetExistsTH(SH_BHA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    cNum = HeaderCol(ws, "BHA #")
    cSn = HeaderCol(ws, "Serial #")
    If cNum = 0 Or cSn = 0 Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, cSn).End(xlUp).Row
    For r = 2 To lastR
        n = BhaNumber(ws.Cells(r, cNum).Value)
        sn = CellSerial(ws.Cells(r, cSn))
        key = NormSerial(sn)
        If n > 0 And Len(key) > 0 Then
            cur = ""
            On Error Resume Next
            cur = CStr(col.Item(key))
            On Error GoTo 0
            If Len(cur) = 0 Then
                col.Add CStr(n), key
            ElseIf InStr("," & cur & ",", "," & CStr(n) & ",") = 0 Then
                col.Remove key
                col.Add cur & "," & CStr(n), key
            End If
        End If
    Next r
End Function

Private Sub WriteMotorList(ByVal tr As Worksheet)
    Dim lastR As Long, r As Long, dest As Long
    Dim sn As String

    tr.Range("M2:M80").ClearContents
    dest = 2
    lastR = TrackerLastRow(tr)
    For r = 2 To lastR
        If StrComp(CStr(tr.Cells(r, TR_KIND).Value & ""), KIND_MOTOR, vbTextCompare) = 0 Then
            sn = Trim$(CStr(tr.Cells(r, TR_SERIAL).Value & ""))
            If Len(sn) > 0 Then
                tr.Cells(dest, TR_MOTOR_LIST).numberFormat = "@"
                tr.Cells(dest, TR_MOTOR_LIST).Value = sn
                dest = dest + 1
            End If
        End If
    Next r
End Sub

' ------------------------------------------------------------------------------
'  Front tables
' ------------------------------------------------------------------------------

Private Function LayoutIsNew(ByVal ws As Worksheet) As Boolean
    Dim h As String
    h = Trim$(CStr(ws.Cells(TOOLS_HEADER_ROW, COL_SERIAL).Value & ""))
    LayoutIsNew = (StrComp(h, "Serial Number", vbTextCompare) = 0) _
                  And (InStr(1, CStr(ws.Cells(TOOLS_TITLE_ROW, COL_SERIAL).Value & ""), TOOLS_TITLE_TOKEN, vbTextCompare) > 0)
End Function

Private Sub EnsureLayout(ByVal ws As Worksheet)
    If LayoutIsNew(ws) Then Exit Sub
    RebuildEqualTables ws
End Sub

Private Sub RebuildEqualTables(ByVal ws As Worksheet)
    Dim rng As Range

    On Error Resume Next
    ws.Range("O28:S55").UnMerge
    On Error GoTo 0

    ws.Range("O28:S55").ClearContents
    ws.Range("O28:S55").ClearFormats

    Set rng = ws.Range("O28:S29")
    rng.Merge
    rng.Value = "Enter any 3rd Party Tools/hours below"
    rng.Font.bold = True
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.Interior.Color = RGB(221, 235, 247)

    PaintHeaderRow ws, TOOLS_HEADER_ROW
    PaintDataRows ws, TOOLS_FIRST, TOOLS_LAST

    Set rng = ws.Range("O42:S43")
    rng.Merge
    rng.Value = "Enter Motors & hours below"
    rng.Font.bold = True
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.Interior.Color = RGB(252, 228, 214)

    PaintHeaderRow ws, MOTOR_HEADER_ROW
    PaintDataRows ws, MOTOR_FIRST, MOTOR_LAST
End Sub

Private Sub PaintHeaderRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, COL_SERIAL), ws.Cells(r, COL_SERIAL + 1))
    rng.Merge
    rng.Value = "Serial Number"
    rng.Font.bold = True
    rng.HorizontalAlignment = xlCenter
    ws.Cells(r, COL_PREV).Value = "Previous hours"
    ws.Cells(r, COL_PREV).Font.bold = True
    ws.Cells(r, COL_CURRENT).Value = "Current Hours"
    ws.Cells(r, COL_CURRENT).Font.bold = True
    ws.Cells(r, COL_TOTAL).Value = "Total"
    ws.Cells(r, COL_TOTAL).Font.bold = True
    ws.Range(ws.Cells(r, COL_SERIAL), ws.Cells(r, COL_TOTAL)).Interior.Color = RGB(217, 217, 217)
    ws.Range(ws.Cells(r, COL_SERIAL), ws.Cells(r, COL_TOTAL)).Borders.LineStyle = xlContinuous
End Sub

Private Sub PaintDataRows(ByVal ws As Worksheet, ByVal firstR As Long, ByVal lastR As Long)
    Dim r As Long
    Dim rng As Range
    For r = firstR To lastR
        On Error Resume Next
        ws.Range(ws.Cells(r, COL_SERIAL), ws.Cells(r, COL_SERIAL + 1)).UnMerge
        On Error GoTo 0
        Set rng = ws.Range(ws.Cells(r, COL_SERIAL), ws.Cells(r, COL_SERIAL + 1))
        rng.Merge
        rng.numberFormat = "@"
        rng.HorizontalAlignment = xlLeft
        ws.Cells(r, COL_PREV).numberFormat = "0.00"
        ws.Cells(r, COL_CURRENT).numberFormat = "0.00"
        ws.Cells(r, COL_TOTAL).Formula = "=IFERROR(IF(O" & r & "<>"""",SUM(Q" & r & ":R" & r & "),""""),"""")"
        ws.Range(ws.Cells(r, COL_SERIAL), ws.Cells(r, COL_TOTAL)).Borders.LineStyle = xlContinuous
    Next r
End Sub

Private Sub PaintFrontTools(ByVal ws As Worksheet, ByVal tr As Worksheet, ByVal bha As Long)
    Dim sns As Collection
    Dim i As Long, r As Long

    Set sns = ActiveBhaToolSerials(bha)

    r = TOOLS_FIRST
    For i = 1 To sns.Count
        If r > TOOLS_LAST Then Exit For
        WriteFrontRow ws, tr, r, CStr(sns.Item(i)), KIND_TOOL
        r = r + 1
    Next i
    For r = r To TOOLS_LAST
        ClearFrontRow ws, r
    Next r
End Sub

Private Sub PaintFrontMotors(ByVal ws As Worksheet, ByVal tr As Worksheet, ByVal bha As Long)
    Dim sns As Collection
    Dim motorSn As String
    Dim i As Long, r As Long

    Set sns = New Collection
    motorSn = Trim$(CStr(ws.Range(CELL_MOTOR).Value & ""))
    If Len(motorSn) = 0 Then motorSn = BhaMotorSerial(bha)
    If Len(motorSn) > 0 Then AddUnique sns, motorSn

    AppendBhaMotors sns, bha

    r = MOTOR_FIRST
    For i = 1 To sns.Count
        If r > MOTOR_LAST Then Exit For
        WriteFrontRow ws, tr, r, CStr(sns.Item(i)), KIND_MOTOR
        r = r + 1
    Next i
    For r = r To MOTOR_LAST
        ClearFrontRow ws, r
    Next r
End Sub

Private Sub WriteFrontRow(ByVal ws As Worksheet, ByVal tr As Worksheet, _
                          ByVal r As Long, ByVal sn As String, ByVal kind As String)
    Dim trR As Long
    trR = TrackerUpsert(tr, sn, kind)
    SetSerialAt ws, r, sn
    WriteNum ws, r, COL_PREV, NumOrZero(tr.Cells(trR, TR_PREJOB).Value)
    WriteNum ws, r, COL_CURRENT, NumOrZero(tr.Cells(trR, TR_JOB).Value)
End Sub

Private Sub ClearFrontRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Range
    Set c = ws.Cells(r, COL_SERIAL)
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    If Len(Trim$(CStr(c.Value & ""))) = 0 _
       And IsEmpty(ws.Cells(r, COL_PREV).Value) _
       And IsEmpty(ws.Cells(r, COL_CURRENT).Value) Then Exit Sub
    c.numberFormat = "@"
    c.Value = ""
    ws.Cells(r, COL_PREV).ClearContents
    ws.Cells(r, COL_CURRENT).ClearContents
End Sub

Private Function ActiveBhaToolSerials(ByVal bha As Long) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    Dim cNum As Long, cSn As Long, cSrc As Long, cDesc As Long, cSub As Long
    Dim lastR As Long, r As Long
    Dim sn As String, src As String, desc As String, subDesc As String
    Dim cat As String, subCat As String, itemName As String, ship As String

    Set ActiveBhaToolSerials = col
    If bha <= 0 Then Exit Function
    If Not SheetExistsTH(SH_BHA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    cNum = HeaderCol(ws, "BHA #")
    cSn = HeaderCol(ws, "Serial #")
    cSrc = HeaderCol(ws, "Source")
    cDesc = HeaderCol(ws, "Description")
    cSub = HeaderCol(ws, "Sub Description")
    If cNum = 0 Or cSn = 0 Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, cNum).End(xlUp).Row
    For r = 2 To lastR
        If BhaNumber(ws.Cells(r, cNum).Value) = bha Then
            sn = CellSerial(ws.Cells(r, cSn))
            If LooksLikeSerial(sn) Then
                src = ""
                desc = ""
                subDesc = ""
                If cSrc > 0 Then src = CStr(ws.Cells(r, cSrc).Value & "")
                If cDesc > 0 Then desc = CStr(ws.Cells(r, cDesc).Value & "")
                If cSub > 0 Then subDesc = CStr(ws.Cells(r, cSub).Value & "")
                InventoryMeta sn, cat, subCat, itemName, ship
                If Len(Trim$(subCat)) = 0 Then subCat = subDesc
                If IsThirdPartySeed(src, sn, desc, cat, subCat, itemName) Then
                    AddUnique col, sn
                End If
            End If
        End If
    Next r
End Function

Private Sub AppendBhaMotors(ByVal col As Collection, ByVal bha As Long)
    Dim ws As Worksheet
    Dim cNum As Long, cSn As Long, cDesc As Long
    Dim lastR As Long, r As Long
    Dim sn As String

    If bha <= 0 Then Exit Sub
    If Not SheetExistsTH(SH_BHA) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    cNum = HeaderCol(ws, "BHA #")
    cSn = HeaderCol(ws, "Serial #")
    cDesc = HeaderCol(ws, "Description")
    If cNum = 0 Or cSn = 0 Or cDesc = 0 Then Exit Sub
    lastR = ws.Cells(ws.Rows.Count, cNum).End(xlUp).Row
    For r = 2 To lastR
        If BhaNumber(ws.Cells(r, cNum).Value) = bha Then
            sn = CellSerial(ws.Cells(r, cSn))
            If LooksLikeSerial(sn) Then
                If IsMotorLike(sn, CStr(ws.Cells(r, cDesc).Value & ""), "", "", "") Then
                    AddUnique col, sn
                End If
            End If
        End If
    Next r
End Sub

Private Sub FixDependentFormulas(ByVal ws As Worksheet)
    Dim r As Long
    Dim toolTbl As String, motorTbl As String
    Dim f As String

    toolTbl = "$O$" & TOOLS_FIRST & ":$S$" & TOOLS_LAST
    motorTbl = "$O$" & MOTOR_FIRST & ":$S$" & MOTOR_LAST

    For r = 24 To 28
        SetLookup ws.Cells(r, 5), "D" & r, toolTbl, 3   ' E Start = Previous
        SetLookupBlank ws.Cells(r, 6), "D" & r, toolTbl, 5  ' F Total
        SetListValidation ws.Cells(r, 4), "=$O$" & TOOLS_FIRST & ":$O$" & TOOLS_LAST
    Next r

    ' D29 is "Activated Agitator Hours" header — leave it.
    SetLookup ws.Cells(30, 5), "D30", toolTbl, 3
    SetLookupBlank ws.Cells(30, 6), "D30", toolTbl, 5
    SetListValidation ws.Cells(30, 4), "=$O$" & TOOLS_FIRST & ":$O$" & TOOLS_LAST

    For r = 31 To 32
        If r = 31 Then
            SetLookup ws.Cells(r, 5), "$D31", toolTbl, 3
            SetLookup ws.Cells(r, 6), "$D31", toolTbl, 5
        Else
            SetLookup ws.Cells(r, 5), "D32", toolTbl, 3
        End If
        SetListValidation ws.Cells(r, 4), "=$O$" & TOOLS_FIRST & ":$O$" & TOOLS_LAST
    Next r

    ws.Range(CELL_PREV_MOTOR).Formula = "=IFERROR(VLOOKUP($C$22," & motorTbl & ",3,0),"""")"

    For r = 45 To 55
        ws.Cells(r, 6).Formula = "=IFERROR(VLOOKUP(B" & r & "," & motorTbl & ",5,FALSE),IFERROR(VLOOKUP(B" & r & "," & SH_TRACK & "!$A$2:$E$80,5,FALSE),""""))"
    Next r

    SetListValidation ws.Range(CELL_MOTOR), "=$B$45:$B$55"

    f = CStr(ws.Range("C23").Formula & "")
    If InStr(1, f, "B45:F52", vbTextCompare) > 0 Then
        ws.Range("C23").Formula = Replace(f, "B45:F52", "B45:F55", 1, -1, vbTextCompare)
    End If
    f = CStr(ws.Range("C24").Formula & "")
    If InStr(1, f, "B45:F52", vbTextCompare) > 0 Then
        ws.Range("C24").Formula = Replace(f, "B45:F52", "B45:F55", 1, -1, vbTextCompare)
    End If
End Sub

Private Sub SetLookup(ByVal c As Range, ByVal keyAddr As String, _
                      ByVal tbl As String, ByVal col As Long)
    Dim f As String
    f = CStr(c.Formula & "")
    If Len(f) = 0 Then Exit Sub
    If InStr(1, f, "VLOOKUP", vbTextCompare) = 0 Then Exit Sub
    c.Formula = "=IFERROR(VLOOKUP(" & keyAddr & "," & tbl & "," & col & ",FALSE),"""")"
End Sub

Private Sub SetLookupBlank(ByVal c As Range, ByVal keyAddr As String, _
                           ByVal tbl As String, ByVal col As Long)
    Dim f As String
    f = CStr(c.Formula & "")
    If Len(f) = 0 Then Exit Sub
    If InStr(1, f, "VLOOKUP", vbTextCompare) = 0 Then Exit Sub
    c.Formula = "=IFERROR(IF(ISBLANK(" & keyAddr & "),"""",VLOOKUP(" & keyAddr & "," & tbl & "," & col & ",0)),"""")"
End Sub

Private Sub SetListValidation(ByVal c As Range, ByVal listRef As String)
    On Error Resume Next
    If c.Validation.Type = 3 Or Len(CStr(c.Validation.Formula1 & "")) > 0 Then
        c.Validation.Delete
        c.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:=listRef
        c.Validation.IgnoreBlank = True
        c.Validation.InCellDropdown = True
    End If
    On Error GoTo 0
End Sub

Private Sub UnlockFrontTables(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Range(ws.Cells(TOOLS_FIRST, COL_SERIAL), ws.Cells(MOTOR_LAST, COL_TOTAL)).Locked = False
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
'  Internals
' ------------------------------------------------------------------------------

Private Sub FixTitleSpelling(ByVal ws As Worksheet)
    Dim c As Range
    Dim txt As String
    Set c = FindInColumn(ws, COL_SERIAL, TITLE_TYPO)
    If c Is Nothing Then Exit Sub
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    txt = CStr(c.Value & "")
    If InStr(1, txt, TITLE_TYPO, vbTextCompare) = 0 Then Exit Sub
    c.Value = Replace(txt, TITLE_TYPO, TITLE_FIX, 1, -1, vbTextCompare)
End Sub

Private Function FindInColumn(ByVal ws As Worksheet, ByVal col As Long, _
                              ByVal token As String) As Range
    On Error Resume Next
    Set FindInColumn = ws.Columns(col).Find(What:=token, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)
    On Error GoTo 0
End Function

Private Function SerialAt(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim c As Range
    Set c = ws.Cells(r, COL_SERIAL)
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    SerialAt = CellSerial(c)
End Function

' Prefer .Text so a text-formatted "0261" is not read as 261.
' If Excel already coerced a leading-zero serial to a number, recover the
' inventory spelling (0261) by matching digit-keys.
Private Function CellSerial(ByVal c As Range) As String
    Dim t As String
    If c Is Nothing Then Exit Function
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    t = Trim$(CStr(c.text & ""))
    If Len(t) = 0 Or t = "0" Then t = Trim$(CStr(c.Value & ""))
    CellSerial = CanonicalSerial(t)
End Function

Private Function CanonicalSerial(ByVal sn As String) As String
    Dim inv As Worksheet
    Dim cSn As Long, lastR As Long, r As Long
    Dim other As String, want As String
    CanonicalSerial = Trim$(sn)
    want = DigitKey(sn)
    If Len(want) = 0 Then Exit Function
    If Not SheetExistsTH(SH_INVENTORY) Then Exit Function
    Set inv = ThisWorkbook.Worksheets(SH_INVENTORY)
    cSn = HeaderCol(inv, "SerialNumber")
    If cSn = 0 Then cSn = HeaderCol(inv, "Serial #")
    If cSn = 0 Then Exit Function
    lastR = inv.Cells(inv.Rows.Count, cSn).End(xlUp).Row
    For r = 2 To lastR
        other = Trim$(CStr(inv.Cells(r, cSn).text & ""))
        If Len(other) = 0 Then other = Trim$(CStr(inv.Cells(r, cSn).Value & ""))
        If DigitKey(other) = want And Len(other) >= Len(sn) Then
            CanonicalSerial = other
            Exit Function
        End If
    Next r
End Function

' Digit-only key with leading zeros stripped, so 0261 and 261 match.
Private Function DigitKey(ByVal sn As String) As String
    Dim s As String
    s = NormSerial(sn)
    DigitKey = ""
    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9]*" Then Exit Function
    Do While Len(s) > 1 And Left$(s, 1) = "0"
        s = mid$(s, 2)
    Loop
    DigitKey = s
End Function

Private Sub SetSerialAt(ByVal ws As Worksheet, ByVal r As Long, ByVal sn As String)
    Dim c As Range
    Set c = ws.Cells(r, COL_SERIAL)
    If c.MergeCells Then
        c.MergeArea.numberFormat = "@"
        Set c = c.MergeArea.Cells(1, 1)
    End If
    c.numberFormat = "@"
    If StrComp(CStr(c.text & ""), sn, vbTextCompare) = 0 Then Exit Sub
    c.Value = CStr(sn)
End Sub

Private Sub WriteNum(ByVal ws As Worksheet, ByVal r As Long, ByVal col As Long, ByVal hrs As Double)
    Dim c As Range
    Set c = ws.Cells(r, col)
    If c.MergeCells Then Set c = c.MergeArea.Cells(1, 1)
    If Not IsEmpty(c.Value) Then
        If IsNumeric(c.Value) Then
            If Abs(CDbl(c.Value) - hrs) < 0.000001 Then Exit Sub
        End If
    ElseIf hrs = 0 Then
        c.Value = 0
        Exit Sub
    End If
    c.Value = hrs
End Sub

Private Function CollectPicks(ByVal ws As Worksheet) As Collection
    Dim col As New Collection
    Dim c As Range
    Dim sn As String
    For Each c In ws.Range(RNG_PICKS).Cells
        sn = Trim$(CStr(c.Value & ""))
        If Len(sn) > 0 And InStr(1, sn, "Agitator Hours", vbTextCompare) = 0 Then
            AddUnique col, sn
        End If
    Next c
    Set CollectPicks = col
End Function

Private Function BhaMotorSerial(ByVal bha As Long) As String
    Dim ws As Worksheet
    Dim cNum As Long, cSn As Long, cDesc As Long
    Dim lastR As Long, r As Long
    Dim sn As String

    BhaMotorSerial = ""
    If bha <= 0 Then Exit Function
    If Not SheetExistsTH(SH_BHA) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    cNum = HeaderCol(ws, "BHA #")
    cSn = HeaderCol(ws, "Serial #")
    cDesc = HeaderCol(ws, "Description")
    If cNum = 0 Or cSn = 0 Or cDesc = 0 Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, cNum).End(xlUp).Row
    For r = 2 To lastR
        If BhaNumber(ws.Cells(r, cNum).Value) = bha Then
            sn = CellSerial(ws.Cells(r, cSn))
            If LooksLikeSerial(sn) Then
                If IsMotorLike(sn, CStr(ws.Cells(r, cDesc).Value & ""), "", "", "") Then
                    BhaMotorSerial = sn
                    Exit Function
                End If
            End If
        End If
    Next r
End Function

Private Function InventoryMeta(ByVal sn As String, ByRef cat As String, _
                               ByRef subCat As String, ByRef itemName As String, _
                               ByRef ship As String) As Boolean
    Dim inv As Worksheet
    Dim cSn As Long, cCat As Long, cSub As Long, cName As Long, cShip As Long
    Dim lastR As Long, r As Long
    Dim want As String

    cat = ""
    subCat = ""
    itemName = ""
    ship = ""
    InventoryMeta = False
    want = NormSerial(sn)
    If Len(want) = 0 Then Exit Function
    If Not SheetExistsTH(SH_INVENTORY) Then Exit Function
    Set inv = ThisWorkbook.Worksheets(SH_INVENTORY)
    cSn = HeaderCol(inv, "SerialNumber")
    If cSn = 0 Then cSn = HeaderCol(inv, "Serial #")
    cCat = HeaderCol(inv, "Category")
    cSub = HeaderCol(inv, "SubCategory")
    cName = HeaderCol(inv, "ItemName")
    If cName = 0 Then cName = HeaderCol(inv, "Description")
    cShip = HeaderCol(inv, "ShippingStatus")
    If cSn = 0 Then Exit Function
    lastR = inv.Cells(inv.Rows.Count, cSn).End(xlUp).Row
    For r = 2 To lastR
        If NormSerial(CellSerial(inv.Cells(r, cSn))) = want Then
            If cCat > 0 Then cat = CStr(inv.Cells(r, cCat).Value & "")
            If cSub > 0 Then subCat = CStr(inv.Cells(r, cSub).Value & "")
            If cName > 0 Then itemName = CStr(inv.Cells(r, cName).Value & "")
            If cShip > 0 Then ship = CStr(inv.Cells(r, cShip).Value & "")
            InventoryMeta = True
            Exit Function
        End If
    Next r
End Function

Private Function IsSteerMotorText(ByVal text As String) As Boolean
    Dim t As String
    t = CStr(text & "")
    IsSteerMotorText = (InStr(1, t, "Orbit RSS", vbTextCompare) > 0) _
                    Or (InStr(1, t, "iCruise", vbTextCompare) > 0) _
                    Or (InStr(1, t, "i-Cruise", vbTextCompare) > 0)
End Function

Private Function IsMotorLike(ByVal sn As String, ByVal desc As String, _
                             ByVal cat As String, ByVal subCat As String, _
                             ByVal itemName As String) As Boolean
    IsMotorLike = False
    If Len(Trim$(sn)) = 0 Then Exit Function
    If InStr(1, desc, "Mud Motor", vbTextCompare) > 0 Then
        IsMotorLike = True
        Exit Function
    End If
    If InStr(1, itemName, "Mud Motor", vbTextCompare) > 0 Then
        IsMotorLike = True
        Exit Function
    End If
    ' Bare SubCategory "Motor" is a FieldCap Other Tools type — not a mud motor.
    IsMotorLike = IsSteerMotorText(desc) Or IsSteerMotorText(itemName)
End Function

Private Function IsSkippedToolKind(ByVal cat As String, ByVal subCat As String, _
                                   ByVal desc As String, ByVal itemName As String) As Boolean
    Dim blob As String
    blob = LCase$(Trim$(cat) & " " & Trim$(subCat) & " " & Trim$(desc) & " " & Trim$(itemName))
    If StrComp(Trim$(cat), "MWD", vbTextCompare) = 0 Then
        IsSkippedToolKind = True
        Exit Function
    End If
    If InStr(1, blob, "mwd kit", vbTextCompare) > 0 _
       Or InStr(1, blob, "mwd other", vbTextCompare) > 0 Then
        IsSkippedToolKind = True
        Exit Function
    End If
    If InStr(1, blob, "dd tubular", vbTextCompare) > 0 Then
        IsSkippedToolKind = True
        Exit Function
    End If
    If InStr(1, blob, "drill bit", vbTextCompare) > 0 Then
        IsSkippedToolKind = True
        Exit Function
    End If
    If InStr(1, " " & LCase$(Trim$(subCat)) & " ", " tubular ", vbTextCompare) > 0 Then
        IsSkippedToolKind = True
        Exit Function
    End If
    If InStr(1, blob, "slick collar", vbTextCompare) > 0 _
       Or InStr(1, blob, "pony collar", vbTextCompare) > 0 _
       Or InStr(1, blob, "drill pipe", vbTextCompare) > 0 _
       Or InStr(1, blob, "hwdp", vbTextCompare) > 0 Then
        IsSkippedToolKind = True
        Exit Function
    End If
    IsSkippedToolKind = False
End Function

Private Function IsOtherToolsCategory(ByVal cat As String) As Boolean
    Dim c As String
    c = Trim$(cat)
    IsOtherToolsCategory = (StrComp(c, OTHER_INVENTORY_TAG, vbTextCompare) = 0) _
                        Or (StrComp(c, OTHER_TOOLS_TAG, vbTextCompare) = 0) _
                        Or (StrComp(c, "Other", vbTextCompare) = 0)
End Function

' FieldCap Other Tools types on the selected BHA. "Motor" / "Other" are exact Type/SubCategory only.
Private Function IsOtherToolType(ByVal cat As String, ByVal subCat As String, _
                                 ByVal desc As String, ByVal itemName As String, _
                                 Optional ByVal toolType As String = "") As Boolean
    Dim blob As String
    Dim subT As String
    Dim typeT As String
    blob = Trim$(cat) & " " & Trim$(subCat) & " " & Trim$(desc) & " " & Trim$(itemName) & " " & Trim$(toolType)
    subT = Trim$(subCat)
    typeT = Trim$(toolType)
    If ContainsTypeToken(blob, "Agitator") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Bit Sub") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Crossover") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Digger") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Exciter") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Hole Opener") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Jar") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "OBL Sub") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Reamer") Then IsOtherToolType = True: Exit Function
    If ContainsTypeToken(blob, "Shock Sub") Then IsOtherToolType = True: Exit Function
    If StrComp(subT, "Motor", vbTextCompare) = 0 Or StrComp(typeT, "Motor", vbTextCompare) = 0 Then
        IsOtherToolType = True
        Exit Function
    End If
    If StrComp(subT, "Other", vbTextCompare) = 0 Or StrComp(typeT, "Other", vbTextCompare) = 0 Then
        IsOtherToolType = True
        Exit Function
    End If
    ' Chrome tags FieldCap Other Tools as "DD other" (not the word Other alone).
    If InStr(1, subT, "other", vbTextCompare) > 0 _
       And InStr(1, subT, "mwd", vbTextCompare) = 0 Then
        IsOtherToolType = True
        Exit Function
    End If
    If InStr(1, typeT, "other", vbTextCompare) > 0 _
       And InStr(1, typeT, "mwd", vbTextCompare) = 0 Then
        IsOtherToolType = True
        Exit Function
    End If
    If IsOtherToolsCategory(cat) And Len(subT) = 0 And Len(typeT) = 0 Then
        IsOtherToolType = True
        Exit Function
    End If
    IsOtherToolType = False
End Function

Private Function ContainsTypeToken(ByVal blob As String, ByVal token As String) As Boolean
    ContainsTypeToken = (InStr(1, blob, token, vbTextCompare) > 0)
End Function

Private Function IsThirdPartySeed(ByVal src As String, ByVal sn As String, _
                                  ByVal desc As String, ByVal cat As String, _
                                  ByVal subCat As String, ByVal itemName As String) As Boolean
    IsThirdPartySeed = False
    If Not LooksLikeSerial(sn) Then Exit Function
    If IsMotorLike(sn, desc, cat, subCat, itemName) Then Exit Function
    If IsSkippedToolKind(cat, subCat, desc, itemName) Then Exit Function

    IsThirdPartySeed = IsOtherToolsFamily(src, cat, subCat)
End Function

' FieldCap Other Tools table only — not PHX DD/MWD inventory.
Private Function IsOtherToolsFamily(ByVal src As String, ByVal cat As String, _
                                    ByVal subCat As String) As Boolean
    Dim subT As String
    subT = Trim$(subCat)
    If StrComp(Trim$(src), OTHER_INVENTORY_TAG, vbTextCompare) = 0 Then
        IsOtherToolsFamily = True
        Exit Function
    End If
    If IsOtherToolsCategory(cat) Then
        IsOtherToolsFamily = True
        Exit Function
    End If
    If InStr(1, subT, "other", vbTextCompare) > 0 _
       And InStr(1, subT, "mwd", vbTextCompare) = 0 Then
        IsOtherToolsFamily = True
        Exit Function
    End If
    IsOtherToolsFamily = False
End Function

Private Function InventorySerialsOfKind(ParamArray kinds() As Variant) As Collection
    Dim col As New Collection
    Dim inv As Worksheet
    Dim cSn As Long, cSub As Long
    Dim lastR As Long, r As Long, k As Long
    Dim sub_ As String, sn As String

    Set InventorySerialsOfKind = col
    If Not SheetExistsTH(SH_INVENTORY) Then Exit Function
    Set inv = ThisWorkbook.Worksheets(SH_INVENTORY)
    cSn = HeaderCol(inv, "SerialNumber")
    cSub = HeaderCol(inv, "SubCategory")
    If cSn = 0 Or cSub = 0 Then Exit Function

    lastR = inv.Cells(inv.Rows.Count, cSn).End(xlUp).Row
    For r = 2 To lastR
        sub_ = " " & Trim$(CStr(inv.Cells(r, cSub).Value & "")) & " "
        For k = LBound(kinds) To UBound(kinds)
            If InStr(1, sub_, " " & CStr(kinds(k)) & " ", vbTextCompare) > 0 Then
                sn = CellSerial(inv.Cells(r, cSn))
                If Len(sn) > 0 Then AddUnique col, sn
                Exit For
            End If
        Next k
    Next r
End Function

Private Function HeaderCol(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim lastC As Long, c As Long
    HeaderCol = 0
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value & "")), header, vbTextCompare) = 0 Then
            HeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Function NormSerial(ByVal sn As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(sn)
        ch = mid$(sn, i, 1)
        If ch Like "[A-Za-z0-9]" Then out = out & UCase$(ch)
    Next i
    NormSerial = out
End Function

Private Sub AddUnique(ByVal col As Collection, ByVal sn As String)
    Dim key As String
    key = NormSerial(sn)
    If Len(key) = 0 Then Exit Sub
    On Error Resume Next
    col.Add sn, key
    On Error GoTo 0
End Sub

Private Function InSet(ByVal col As Collection, ByVal sn As String) As Boolean
    Dim v As Variant
    InSet = False
    If col Is Nothing Then Exit Function
    On Error Resume Next
    If Len(NormSerial(sn)) = 0 Then Exit Function
    v = col.Item(NormSerial(sn))
    InSet = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function BhaNumber(ByVal v As Variant) As Long
    BhaNumber = 0
    On Error Resume Next
    If IsNumeric(v) Then BhaNumber = CLng(v)
    On Error GoTo 0
End Function

Private Function NumOrZero(ByVal v As Variant) As Double
    NumOrZero = 0#
    On Error Resume Next
    If IsNumeric(v) Then NumOrZero = CDbl(v)
    On Error GoTo 0
End Function

Private Function SheetExistsTH(ByVal name As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    SheetExistsTH = Not ws Is Nothing
End Function


