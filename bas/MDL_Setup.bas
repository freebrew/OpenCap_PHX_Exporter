Attribute VB_Name = "MDL_Setup"
Option Explicit

' ================================================================================
'  MODULE: MDL_Setup
'  OpenCap Field Workbook - Well Database & Setup Tab
'  v1.0.0
'
'  QUICK START:
'   1. Alt+F11 > Insert > Module > Import this file
'   2. Immediate window:  InitSetup
'   3. Place CSV exports in the same folder as this workbook
'   4. Click  REFRESH CSVs  on the Setup sheet
'
'  SHEET ARCHITECTURE:
'   "Setup"         = visible Well Database + File Status dashboard
'   "_OC_Job"       = hidden  Job Details CSV
'   "_OC_Crew"      = hidden  Crew / Personnel CSV
'   "_OC_BHA"       = hidden  BHA Equipment CSV
'   "_OC_Slide"     = hidden  Slide-Rotate by Day CSV
'   "_OC_Inventory" = hidden  Equipment Inventory CSV
'   "_OC_Costs"     = hidden  Ticket Costs by Day CSV
'   "_OC_PlanSec"   = hidden  Plan Sections (PDF) + auto/user target names
'
'  LAYOUT WARNING:
'   BuildSetupUI / RebuildSetup / InitSetup CLEAR the entire Setup sheet and
'   redraw it from this module. Manual Setup tweaks are lost unless they are
'   encoded here (section gap rowHeight=20, notes box = 5 rows, etc.).
'   Agents must NOT run Rebuild/Init unless the user explicitly asks.
' ================================================================================

' -- Sheet names ---------------------------------------------------------------
Private Const SH_SETUP     As String = "Setup"
Private Const SH_JOB       As String = "_OC_Job"
Private Const SH_CREW      As String = "_OC_Crew"
Private Const SH_BHA       As String = "_OC_BHA"
Private Const SH_SLIDE     As String = "_OC_Slide"
Private Const SH_INVENTORY As String = "_OC_Inventory"
Private Const SH_COSTS     As String = "_OC_Costs"
Private Const SH_SURVEY    As String = "_OC_Survey"
Private Const SH_AC        As String = "_OC_AC"
Private Const SH_PLANSEC   As String = "_OC_PlanSec"
Private Const SH_COSTS_TAB As String = "Costs"

' -- Anti-collision status bands -----------------------------------------------
' SF <  AC_SF_RED     -> faded red     (hard concern)
' SF <  AC_SF_YELLOW  -> faded yellow  (watch)
' SF >= AC_SF_YELLOW  -> faded green   (clear)
Private Const AC_SF_RED    As Double = 1.5
Private Const AC_SF_YELLOW As Double = 2#

' -- CSV filename tokens (case-insensitive partial match) ----------------------
Private Const TOK_JOB       As String = "job-details"
Private Const TOK_CREW      As String = "crew"
Private Const TOK_BHA       As String = "bha-equipment"

' OpenCap extension / source (Setup → OPENCAP EXPORT FILES footer)
Private Const URL_OC_CHROME As String = _
    "https://chromewebstore.google.com/detail/opencap-data-exporter/egafclljokidcoiipaiojlaamgkooffb"
Private Const URL_OC_GITHUB As String = _
    "https://github.com/freebrew/OpenCap_PHX_Exporter"
Private Const TOK_SLIDE     As String = "slide-rotate"
Private Const TOK_INVENTORY As String = "inventory"
Private Const TOK_COSTS     As String = "ticket-costs"

' -- Costs tab layout ----------------------------------------------------------
Private Const COSTS_FIRST_ROW As Long = 3
Private Const COSTS_LAST_ROW  As Long = 72
Private Const COSTS_COL_DATE  As Long = 3   ' C
Private Const COSTS_COL_DAILY As Long = 5   ' E
Private Const COSTS_COL_TOTAL As Long = 7   ' G

' -- Data sheet Mud Motor table (inventory.csv on location; BHA as fallback) ---
Private Const SH_DATA          As String = "Data"
Private Const MM_TITLE_TEXT    As String = "Enter Motors"
Private Const MM_COL_SERIAL    As Long = 2   ' B  (Motors On Location B45:B55)
Private Const MM_COL_HOURS     As Long = 18  ' R  (front motor table; hours owned by MDL_ToolHours)
Private Const MM_ROWS          As Long = 11  ' B45:B55
Private Const MM_TITLE_TO_DATA As Long = 0
Private Const MM_FIRST_ROW     As Long = 45  ' Motors On Location first serial row
Private Const NM_LAST_JOB      As String = "OC_LastJobId"
Private Const NM_CSV_ROOT      As String = "OC_CsvRoot"

' -- Column layout (1-based) ---------------------------------------------------
'  A(1)       = section accent strip (1.5 wide)
'  B-C(2-3)   = label col 1 (merged, right-aligned dim text)
'  D-E(4-5)   = value col 1 (merged, bold)
'  F(6)       = label col 2 (single, right-aligned dim text)
'  G-H(7-8)   = value col 2 (merged, bold)
'  I(9)       = gap column
'  J-K(10-11) = crew role (merged)
'  L-M(12-13) = crew name (merged)
'  N(14)      = crew phone / email
'  O-R(15-18) = overflow / file path
Private Const C_ACCENT  As Long = 1
Private Const C_L1S     As Long = 2    ' label 1 start
Private Const C_L1E     As Long = 3    ' label 1 end
Private Const C_V1S     As Long = 4    ' value 1 start
Private Const C_V1E     As Long = 5    ' value 1 end
Private Const C_L2      As Long = 6    ' label 2 (single)
Private Const C_V2S     As Long = 7    ' value 2 start
Private Const C_V2E     As Long = 8    ' value 2 end
Private Const C_GAP     As Long = 9
Private Const C_CRL     As Long = 10   ' crew left
Private Const C_CRR     As Long = 17   ' crew right
Private Const C_LAST    As Long = 18

' -- Row anchors ---------------------------------------------------------------
Private Const R_HDR  As Long = 1
Private Const R_DIV  As Long = 2
Private Const R_BODY As Long = 3

' Re-entrance guard
Private mBusy As Boolean
Private mNewJob As Boolean
Private mCsvExtraRoots As Collection

' Poppler self-bootstrap: offer the download at most once per Excel session
Private mPopplerPrompted As Boolean

' ================================================================================
'  COLOR PALETTE  (Slidesheet: grey table headers, green tab / status only)
' ================================================================================
Private Function cWh() As Long:     cWh = RGB(255, 255, 255): End Function
Private Function cBg() As Long:     cBg = RGB(242, 242, 242): End Function
Private Function cMed() As Long:    cMed = RGB(217, 217, 217): End Function
Private Function cDk() As Long:     cDk = RGB(89, 89, 89): End Function
Private Function cBlk() As Long:    cBlk = RGB(0, 0, 0): End Function
' Table / section headers — same grey as Slidesheet SAIL / Way Points headers
Private Function cTeal() As Long:   cTeal = RGB(217, 217, 217): End Function
Private Function cTealLt() As Long: cTealLt = RGB(242, 242, 242): End Function
Private Function cLine() As Long:   cLine = RGB(166, 166, 166): End Function
Private Function cAccent() As Long: cAccent = RGB(217, 217, 217): End Function
' Green — sheet tab + file-found badges only (not headers)
Private Function cGrnBadge() As Long: cGrnBadge = RGB(0, 176, 80): End Function
Private Function cGrnRow() As Long:   cGrnRow = RGB(206, 239, 198): End Function
Private Function cGrnTxt() As Long:   cGrnTxt = RGB(0, 110, 60): End Function
' Red - missing files
Private Function cRedBadge() As Long: cRedBadge = RGB(176, 40, 40): End Function
Private Function cRedRow() As Long:   cRedRow = RGB(255, 240, 240): End Function
Private Function cRedTxt() As Long:   cRedTxt = RGB(160, 30, 30): End Function
Private Function cLink() As Long:   cLink = RGB(0, 102, 204): End Function

' ================================================================================
'  PUBLIC ENTRY POINTS
' ================================================================================

Public Sub InitSetup()
    If mBusy Then Exit Sub
    mBusy = True

    Dim en As Long
    Dim ed As String
    Dim setupWasProt As Boolean

    ScreenBeginBusy "OpenCap: starting up..."

    On Error GoTo ErrHandler

    SetupDataSheets
    EnsureSheet SH_SETUP, True

    Dim wbPath As String: wbPath = ThisWorkbook.Path
    If wbPath <> "" Then
        Dim fJob As String:   fJob = FindCsvByToken(wbPath, TOK_JOB)
        Dim fCrew As String:  fCrew = FindCsvByToken(wbPath, TOK_CREW)
        Dim fBha As String:   fBha = FindCsvByToken(wbPath, TOK_BHA)
        Dim fSlide As String: fSlide = FindCsvByToken(wbPath, TOK_SLIDE)
        Dim fInv As String:   fInv = FindCsvByToken(wbPath, TOK_INVENTORY)
        Dim fCosts As String: fCosts = FindCsvByToken(wbPath, TOK_COSTS)
        If fJob <> "" Then LoadCsv fJob, SH_JOB
        If fCrew <> "" Then LoadCsv fCrew, SH_CREW
        If fBha <> "" Then LoadCsv fBha, SH_BHA
        If fSlide <> "" Then LoadCsv fSlide, SH_SLIDE
        If fInv <> "" Then LoadCsv fInv, SH_INVENTORY
        If fCosts <> "" Then LoadCsv fCosts, SH_COSTS
    End If

    RememberCsvRootsFromSetup
    DetectNewJob

    setupWasProt = SheetUnprotectForVba(ThisWorkbook.Worksheets(SH_SETUP))
    BuildSetupUI
    SheetReprotectAfterVba ThisWorkbook.Worksheets(SH_SETUP), setupWasProt

    ScreenEndBusy
    mBusy = False

    SyncCostsFromOpenCap
    SyncMudMotorsFromBha
    Exit Sub

ErrHandler:
    en = Err.Number
    ed = Err.Description
    ScreenForceReset
    mBusy = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = "INIT ERR " & en & ": " & ed
    On Error GoTo 0
End Sub

Public Sub RefreshAllCsvData()
    ' Single refresh entry: import ALL OpenCap CSVs, rebuild Setup, sync Costs.
    If mBusy Then mBusy = False
    mBusy = True

    Dim errNum As Long
    Dim setupWasProt As Boolean
    Dim errMsg As String

    Dim wbPath As String: wbPath = ThisWorkbook.Path
    If wbPath = "" Then
        mBusy = False
        MsgBox "Save the workbook to a folder first." & vbCrLf & _
               "Refresh scans <workbook>\OpenCap\ for CSV exports.", _
               vbInformation, "OpenCap Refresh"
        Exit Sub
    End If

    ScreenBeginBusy "OpenCap: scanning for CSV files..."

    On Error GoTo ErrHandler

    SetupDataSheets

    Dim fJob As String:   fJob = FindCsvByToken(wbPath, TOK_JOB)
    Dim fCrew As String:  fCrew = FindCsvByToken(wbPath, TOK_CREW)
    Dim fBha As String:   fBha = FindCsvByToken(wbPath, TOK_BHA)
    Dim fSlide As String: fSlide = FindCsvByToken(wbPath, TOK_SLIDE)
    Dim fInv As String:   fInv = FindCsvByToken(wbPath, TOK_INVENTORY)
    Dim fCosts As String: fCosts = FindCsvByToken(wbPath, TOK_COSTS)

    If fJob <> "" Then LoadCsv fJob, SH_JOB
    If fCrew <> "" Then LoadCsv fCrew, SH_CREW
    If fBha <> "" Then LoadCsv fBha, SH_BHA
    If fSlide <> "" Then LoadCsv fSlide, SH_SLIDE
    If fInv <> "" Then LoadCsv fInv, SH_INVENTORY
    If fCosts <> "" Then LoadCsv fCosts, SH_COSTS

    RememberCsvRootsFromSetup
    DetectNewJob

    Application.StatusBar = "OpenCap: building Setup sheet..."
    EnsureSheet SH_SETUP, True
    setupWasProt = SheetUnprotectForVba(ThisWorkbook.Worksheets(SH_SETUP))
    BuildSetupUI
    SheetReprotectAfterVba ThisWorkbook.Worksheets(SH_SETUP), setupWasProt

    Application.StatusBar = "OpenCap: syncing Costs sheet..."
    SyncCostsFromOpenCap

    Application.StatusBar = "OpenCap: syncing Mud Motors (inventory)..."
    SyncMudMotorsFromBha

    ScreenEndBusy
    mBusy = False
    Exit Sub

ErrHandler:
    errNum = Err.Number
    errMsg = Err.Description
    On Error Resume Next
    SheetReprotectAfterVba ThisWorkbook.Worksheets(SH_SETUP), setupWasProt
    On Error GoTo 0
    ScreenForceReset
    mBusy = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = "REFRESH ERR " & errNum & ": " & errMsg
    MsgBox "Refresh failed:" & vbCrLf & vbCrLf & errNum & " - " & errMsg, vbExclamation, "OpenCap Refresh"
    On Error GoTo 0
End Sub

' Load latest OpenCap CSVs and sync Costs / motors / Setup view.
' Does NOT rebuild the Setup sheet (preserves notes and layout).
Public Sub AdoptCurrentJobFromOpenCap()
    Dim wbPath As String
    Dim fJob As String, fCrew As String, fBha As String
    Dim fSlide As String, fInv As String, fCosts As String

    On Error GoTo Fail
    RememberCsvRootsFromSetup
    wbPath = ThisWorkbook.Path
    If wbPath = "" Then Exit Sub

    SetupDataSheets
    fJob = FindCsvByToken(wbPath, TOK_JOB)
    fCrew = FindCsvByToken(wbPath, TOK_CREW)
    fBha = FindCsvByToken(wbPath, TOK_BHA)
    fSlide = FindCsvByToken(wbPath, TOK_SLIDE)
    fInv = FindCsvByToken(wbPath, TOK_INVENTORY)
    fCosts = FindCsvByToken(wbPath, TOK_COSTS)
    If fJob <> "" Then LoadCsv fJob, SH_JOB
    If fCrew <> "" Then LoadCsv fCrew, SH_CREW
    If fBha <> "" Then LoadCsv fBha, SH_BHA
    If fSlide <> "" Then LoadCsv fSlide, SH_SLIDE
    If fInv <> "" Then LoadCsv fInv, SH_INVENTORY
    If fCosts <> "" Then LoadCsv fCosts, SH_COSTS

    ' Caller is starting / adopting a job: drop leftover prior-job costs & motor hours.
    mNewJob = True
    SyncCostsFromOpenCap
    SyncMudMotorsFromBha
    If CurrentOcJobId() <> "" Then SetNamedText NM_LAST_JOB, CurrentOcJobId()
    FixSetupView
    Exit Sub
Fail:
    MsgBox "AdoptCurrentJobFromOpenCap: " & Err.Description, vbExclamation, "OpenCap"
End Sub

Public Sub RefreshSetup()
    ' Alias � Setup button and DD Tools RefreshData both use RefreshAllCsvData.
    RefreshAllCsvData
End Sub

Public Sub RebuildSetup()
    If mBusy Then Exit Sub
    mBusy = True

    Dim eR As Long
    Dim em As String
    Dim wasProt As Boolean

    ScreenBeginBusy "OpenCap: rebuilding Setup sheet..."

    On Error GoTo ErrHandler
    EnsureSheet SH_SETUP, True
    wasProt = SheetUnprotectForVba(ThisWorkbook.Worksheets(SH_SETUP))
    BuildSetupUI
    SheetReprotectAfterVba ThisWorkbook.Worksheets(SH_SETUP), wasProt

    ScreenEndBusy
    mBusy = False
    Exit Sub

ErrHandler:
    eR = Err.Number
    em = Err.Description
    ScreenForceReset
    mBusy = False
    On Error Resume Next
    SheetReprotectAfterVba ThisWorkbook.Worksheets(SH_SETUP), wasProt
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = "REBUILD ERR " & eR & ": " & em
    On Error GoTo 0
End Sub

' ================================================================================
'  SHEET HOUSEKEEPING
' ================================================================================

Private Sub SetupDataSheets()
    Dim n(8) As String
    n(0) = SH_JOB: n(1) = SH_CREW: n(2) = SH_BHA: n(3) = SH_SLIDE: n(4) = SH_INVENTORY
    n(5) = SH_COSTS: n(6) = SH_SURVEY: n(7) = SH_AC: n(8) = SH_PLANSEC
    Dim i As Integer
    For i = 0 To 8
        If Not SheetExists(n(i)) Then
            ThisWorkbook.sheets.Add(After:=ThisWorkbook.sheets( _
                ThisWorkbook.sheets.Count)).name = n(i)
        End If
        Worksheets(n(i)).Visible = xlSheetVeryHidden
    Next i
End Sub

Private Sub EnsureSheet(nm As String, addBefore As Boolean)
    If Not SheetExists(nm) Then
        Dim ws As Worksheet
        If addBefore Then
            Set ws = ThisWorkbook.sheets.Add(Before:=ThisWorkbook.sheets(1))
        Else
            Set ws = ThisWorkbook.sheets.Add( _
                After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count))
        End If
        ws.name = nm
    End If
End Sub

' ================================================================================
'  MAIN BUILDER
' ================================================================================

Private Sub BuildSetupUI()
    Dim ws As Worksheet: Set ws = Worksheets(SH_SETUP)

    ws.Cells.UnMerge
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes: shp.Delete: Next shp

    ConfigSheet ws
    DrawHeader ws

    DrawJobAndCrewSection ws, R_BODY

    DrawFilesPanel ws, 18
    DrawImportFilesPanel ws, 27
    DrawNotesPanel ws, 32

    AttachButtons ws

    ' Ensure every email on Setup is a mailto: hyperlink (crew + contacts)
    LinkMailtoInUsedRange ws

    ' Draw the AC summary skeleton so the column headers are always present
    ' on the Setup sheet even before any AC PDF has been imported.
    Dim dEmpty(0) As Double
    BuildAcTable 0, dEmpty, dEmpty, dEmpty

    ResetSetupWindow ws
End Sub

Private Sub ConfigSheet(ws As Worksheet)
    With ws
        .Cells.Interior.Color = cWh()
        .Cells.Font.Color = cBlk()
        .Cells.Font.name = "Consolas"
        .Cells.Font.Size = 9
        .Tab.Color = cGrnBadge()
    End With

    ApplySetupSheetLayout ws
End Sub

' Column widths, row heights, and kill print page-break ghosts (dashed lines).
' Safe to run without rebuilding Setup content.
Public Sub FixSetupView()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    wasProt = SheetUnprotectForVba(ws)
    ApplySetupSheetLayout ws
    ResetSetupWindow ws
    SheetReprotectAfterVba ws, wasProt
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
End Sub

' Unfreeze, jump to A1, then freeze only the title/divider (rows 1-2).
' A stale freeze at col H / row 21 plus ScrollColumn=M clips the header to "E & SETUP".
Private Sub ResetSetupWindow(ws As Worksheet)
    On Error Resume Next
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    ws.Activate
    With ActiveWindow
        .view = xlNormalView
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayPageBreaks = False
        .FreezePanes = False
        .Split = False
        .ScrollColumn = 1
        .ScrollRow = 1
        .Zoom = 100
    End With
    ws.Cells(R_BODY, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
    On Error GoTo 0
End Sub

Private Sub ApplySetupSheetLayout(ws As Worksheet)
    On Error Resume Next
    Dim owh As Object: Set owh = ActiveWindow
    If Not owh Is Nothing Then
        If StrComp(ActiveSheet.name, ws.name, vbTextCompare) = 0 Then
            owh.view = xlNormalView
            owh.DisplayGridlines = False
            owh.DisplayHeadings = False
            owh.DisplayPageBreaks = False
        End If
    End If
    On Error GoTo 0

    ' Column widths (N-P were previously unset and drifted)
    ws.Columns(C_ACCENT).ColumnWidth = 1.4
    ws.Columns(C_L1S).ColumnWidth = 11
    ws.Columns(C_L1E).ColumnWidth = 8
    ws.Columns(C_V1S).ColumnWidth = 15
    ws.Columns(C_V1E).ColumnWidth = 9
    ws.Columns(C_L2).ColumnWidth = 14
    ws.Columns(C_V2S).ColumnWidth = 13
    ws.Columns(C_V2E).ColumnWidth = 8
    ws.Columns(C_GAP).ColumnWidth = 1.8
    ws.Columns(C_CRL).ColumnWidth = 10       ' J crew role
    ws.Columns(C_CRL + 1).ColumnWidth = 4    ' K
    ws.Columns(C_CRL + 2).ColumnWidth = 17   ' L crew name
    ws.Columns(C_CRL + 3).ColumnWidth = 4    ' M
    ws.Columns(14).ColumnWidth = 10          ' N filename
    ws.Columns(15).ColumnWidth = 12          ' O filename
    ws.Columns(16).ColumnWidth = 7           ' P rows
    ws.Columns(C_CRR).ColumnWidth = 18       ' Q path / email
    ws.Columns(C_CRR + 1).ColumnWidth = 8
    ws.Columns(C_CRR + 2).ColumnWidth = 8
    ws.Columns(C_CRR + 3).ColumnWidth = 8

    ws.Rows("1:130").rowHeight = 16
    ws.Rows(R_HDR).rowHeight = 30
    ws.Rows(R_DIV).rowHeight = 3

    ' Dashed lines through the sheet are leftover print page breaks.
    ' Do not touch PageSetup here — talking to the printer can hang Excel.
    On Error Resume Next
    ws.ResetAllPageBreaks
    ws.DisplayPageBreaks = False
    On Error GoTo 0
End Sub

' ================================================================================
'  HEADER BAND
' ================================================================================

Private Sub DrawHeader(ws As Worksheet)
    ' Full-width title (A-R). AC summary title band removed (table lives at J9).
    With ws.Range(ws.Cells(R_HDR, 1), ws.Cells(R_HDR, C_LAST))
        .Merge
        .Value = "  OPENCAP  |  WELL DATABASE & SETUP"
        .Interior.Color = cTeal()
        .Font.Color = cBlk()
        .Font.Size = 14
        .Font.bold = True
        .Font.name = "Consolas"
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
    End With
    With ws.Range(ws.Cells(R_DIV, 1), ws.Cells(R_DIV, C_LAST))
        .Interior.Color = cLine()
    End With
End Sub

' ================================================================================
'  JOB + CREW SECTION  (Left: A-H, Right: J-N, same rows)
' ================================================================================

Private Sub DrawJobAndCrewSection(ws As Worksheet, startRow As Long)
    Dim j As Object: Set j = ReadJobFields()

    Dim r As Long: r = startRow

    ' -- JOB IDENTITY -------------------------------------------------------------
    SectionBar ws, r, C_ACCENT, C_V2E, "  JOB IDENTITY", cTeal(), cBlk()
    r = r + 1
    Pair ws, r, "JOB ID", GF(j, "Job Code", "Job ID"), _
                 "OPS STATUS", GF(j, "opsStatus", "Ops Status")
    r = r + 1
    Pair ws, r, "JOB NAME", GF(j, "Job Name"), _
                 "PROFILE", GF(j, "WellProfile", "Job Type")
    r = r + 1
    Pair ws, r, "CLIENT", JobClientName(j), _
                 "AFE", GF(j, "AFE", "AFE (core)")
    r = r + 1
    Pair ws, r, "WELL / UWI", GF(j, "UWI", "Well (core)"), _
                 "WELL LICENCE", GF(j, "WellLicenceNumber")
    r = r + 1
    Pair ws, r, "SURFACE COORDS", GF(j, "SurfaceCoordinates"), _
                 "PROVINCE", GF(j, "Province")
    r = r + 1
    Pair ws, r, "SPUD DATE", FmtDate(GF(j, "SpudDate")), _
                 "START DATE", GF(j, "Start Date", "Planned Start Date")
    r = r + 1
    Pair ws, r, "MAP SYSTEM", GF(j, "MapSystem"), _
                 "NORTH REF", GF(j, "NorthReference")
    r = r + 1

    ' -- WELL GEOMETRY ------------------------------------------------------------
    ' Section gap height is intentional UI (do not shrink back to 7 - it gets wiped
    ' every Rebuild and has been re-applied by hand too many times).
    ws.Rows(r).rowHeight = 20: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  WELL GEOMETRY", cTealLt(), cBlk()
    r = r + 1
    Pair ws, r, "GROUND MSL", ValUnit(GF(j, "GroundMSL"), " m"), _
                 "TVD", ValUnit(GF(j, "TVD"), " m")
    r = r + 1
    Pair ws, r, "RKB HEIGHT", ValUnit(GF(j, "RKB"), " m"), _
                 "TOTAL KB", ValUnit(GF(j, "TotalKB"), " m")
    r = r + 1
    Pair ws, r, "LATITUDE", GF(j, "LatitudeDMS", "Latitude"), _
                 "LONGITUDE", GF(j, "LongitudeDMS", "Longitude")
    r = r + 1
    Pair ws, r, "VSD", GF(j, "VSD"), _
                 "WELL TYPE", GF(j, "WellProfile", "Job Type")
    r = r + 1

    ' -- DIRECTIONAL PARAMETERS ---------------------------------------------------
    ws.Rows(r).rowHeight = 20: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  DIRECTIONAL PARAMETERS", cTealLt(), cBlk()
    r = r + 1
    Pair ws, r, "MAG DECL", ValUnit(GF(j, "MagneticDeclination"), Chr(176)), _
                 "REF MODEL", GF(j, "MagneticReferenceModel")
    r = r + 1
    Pair ws, r, "DIP", ValUnit(GF(j, "DIP"), Chr(176)), _
                 "REF DATE", FmtDate(GF(j, "MagneticReferenceDate"))
    r = r + 1
    Pair ws, r, "TGF", GF(j, "TGF"), _
                 "TMF", GF(j, "TMF")
    r = r + 1
    Pair ws, r, "CONVERGENCE", ValUnit(GF(j, "Convergence"), Chr(176)), _
                 "TOL DIP", GF(j, "ToleranceDIP")
    r = r + 1
    Pair ws, r, "SURVEY CORR.", GF(j, "SurveyCorrection"), _
                 "DRILL MEAS.", GF(j, "DrillMeasuredFrom")
    r = r + 1

    ' -- EQUIPMENT & RIG ----------------------------------------------------------
    ws.Rows(r).rowHeight = 20: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  EQUIPMENT & RIG", cTeal(), cBlk()
    r = r + 1
    Pair ws, r, "RIG TYPE", GF(j, "RigType"), _
                 "TOP DRIVE", GF(j, "TopDrive")
    r = r + 1
    Pair ws, r, "DATA RECORDER", GF(j, "DataRecorder"), _
                 "GAMMA", GF(j, "Gamma")
    r = r + 1
    Pair ws, r, "MWD GUIDANCE", GF(j, "MWDGuidanceType"), _
                 "PIPE ARM", GF(j, "PipeArm")
    r = r + 1
    Pair ws, r, "RIG NAME", GF(j, "Rig Name", "RigName"), _
                 "LOADER", GF(j, "Loader")
    r = r + 1

    ' -- CONTACTS -----------------------------------------------------------------
    ws.Rows(r).rowHeight = 20: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  CONTACTS", cTeal(), cBlk()
    r = r + 1
    Pair ws, r, "COMPANY MAN", GF(j, "CompanyMan"), _
                 "CM PHONE", GF(j, "CompanyManPhone")
    r = r + 1
    Pair ws, r, "2ND COMP. MAN", GF(j, "SecondCompanyMan"), _
                 "2ND CM PHONE", GF(j, "SecondCompanyManPhone")
    r = r + 1
    Pair ws, r, "GEOLOGIST", GF(j, "Geologist"), _
                 "GEOL. PHONE", GF(j, "GeologistPhone")
    r = r + 1
    Pair ws, r, "DD COORDINATOR", GF(j, "DDCoordinator"), _
                 "DD COORD EMAIL", GF(j, "DDCoordinatorEmail")
    r = r + 1
    Pair ws, r, "MWD COORDINATOR", GF(j, "MWDCoordinator"), _
                 "MWD EMAIL", GF(j, "MWDCoordinatorEmail")
    r = r + 1
    Pair ws, r, "SALES REP", GF(j, "SalesRep"), _
                 "SALES EMAIL", GF(j, "SalesRepEmail")
    r = r + 1

    ' Left panel outer border
    With ws.Range(ws.Cells(startRow, C_ACCENT), ws.Cells(r - 1, C_V2E))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = cMed()
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = cMed()
    End With

    ' Crew owns J3:R10 (title + header + 6 people). AC table starts at J11.
    DrawCrew ws, startRow, 10
End Sub

' -- Crew panel (columns J:N) --------------------------------------------------
Private Sub DrawCrew(ws As Worksheet, topRow As Long, bottomRow As Long)
    Dim r As Long: r = topRow

    Dim crew() As String: crew = ReadCrewRows()
    Dim nCrew As Long
    On Error Resume Next
    nCrew = UBound(crew, 1) - LBound(crew, 1) + 1
    If Err.Number <> 0 Then nCrew = 0
    On Error GoTo 0
    If nCrew = 1 Then
        If Trim$(crew(0, 1)) = "" And Trim$(crew(0, 0)) = "" Then nCrew = 0
    End If

    Dim crewCap As Long: crewCap = bottomRow - topRow - 1
    If crewCap < 1 Then crewCap = 1
    Dim crewTitle As String
    crewTitle = "  CREW MANIFEST"
    If nCrew > crewCap Then
        crewTitle = crewTitle & "   " & CStr(crewCap) & "+" & CStr(nCrew - crewCap)
    ElseIf nCrew > 0 Then
        crewTitle = crewTitle & "   " & CStr(nCrew)
    End If
    SectionBar ws, r, C_CRL, C_LAST, crewTitle, cTeal(), cBlk()
    ws.Rows(r).rowHeight = 16
    r = r + 1

    ' Sub-header: ROLE | NAME | EMAIL | PHONE
    ws.Rows(r).rowHeight = 11
    Dim hdrBg As Long: hdrBg = cBg()
    Sgl ws, r, C_CRL, "ROLE", hdrBg, cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, C_CRL + 1, C_CRL + 2, "NAME", hdrBg, cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, C_CRL + 3, C_CRL + 5, "EMAIL", hdrBg, cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, C_CRL + 6, C_LAST, "PHONE", hdrBg, cDk(), True, 7, xlHAlignLeft
    With ws.Range(ws.Cells(r, C_CRL), ws.Cells(r, C_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(235, 235, 235)
        .Weight = xlHairline
    End With
    r = r + 1

    Dim dataTop As Long: dataTop = r

    Dim i As Long
    Dim slot As Long: slot = 0
    For i = 0 To UBound(crew, 1)
        If r > bottomRow Then Exit For
        If nCrew = 0 Then Exit For
        PaintCrewRow ws, r, slot, crew(i, 0), crew(i, 1), crew(i, 2), crew(i, 3)
        r = r + 1
        slot = slot + 1
    Next i

    Do While r <= bottomRow
        PaintCrewRow ws, r, slot, "", "", "", ""
        r = r + 1
        slot = slot + 1
    Loop

    Dim dataBot As Long: dataBot = r - 1

    ' Add DD/MWD dropdown validation on each Role cell
    If dataBot >= dataTop Then
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(dataTop, C_CRL), ws.Cells(dataBot, C_CRL))
        With rng.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="DD,MWD"
            .InCellDropdown = True
            .ShowError = False
        End With
        rng.Locked = False

        On Error Resume Next
        ws.names("OC_CrewTop").Delete
        ws.names("OC_CrewBot").Delete
        On Error GoTo 0
        ws.names.Add name:="OC_CrewTop", RefersToR1C1:="=R" & dataTop & "C" & C_CRL
        ws.names.Add name:="OC_CrewBot", RefersToR1C1:="=R" & dataBot & "C" & C_CRL

        InstallCrewSortEvent ws
    End If

    With ws.Range(ws.Cells(topRow, C_CRL), ws.Cells(dataBot, C_LAST))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Color = cMed()
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Color = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = cMed()
    End With
End Sub

Private Sub PaintCrewRow(ws As Worksheet, ByVal r As Long, ByVal slot As Long, _
        ByVal role As String, ByVal nm As String, ByVal em As String, ByVal pH As String)
    Dim bg As Long
    If slot Mod 2 = 0 Then bg = cWh() Else bg = cBg()
    ws.Rows(r).rowHeight = 12
    Sgl ws, r, C_CRL, role, bg, cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, C_CRL + 1, C_CRL + 2, nm, bg, cBlk(), (nm <> ""), 8, xlHAlignLeft
    MrgCell ws, r, C_CRL + 3, C_CRL + 5, em, bg, cDk(), False, 7, xlHAlignLeft
    ApplyMailtoLink ws.Range(ws.Cells(r, C_CRL + 3), ws.Cells(r, C_CRL + 5))
    MrgCell ws, r, C_CRL + 6, C_LAST, pH, bg, cDk(), False, 7, xlHAlignLeft
    With ws.Range(ws.Cells(r, C_CRL), ws.Cells(r, C_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(235, 235, 235): .Weight = xlHairline
    End With
End Sub

' Inject Worksheet_Change into the Setup sheet's code module so dropdown
' changes auto-trigger SortCrewByRole. Requires "Trust access to VBA project".
Private Sub InstallCrewSortEvent(ws As Worksheet)
    On Error GoTo NoAccess
    Dim cm As Object
    Dim shName As String: shName = ws.CodeName
    Set cm = ThisWorkbook.VBProject.VBComponents(shName).CodeModule

    Dim tag As String: tag = "SortCrewByRole"
    Dim i As Long
    For i = 1 To cm.CountOfLines
        If InStr(1, cm.lines(i, 1), tag, vbTextCompare) > 0 Then Exit Sub
    Next i

    Dim code As String
    code = ""
    code = code & "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
    code = code & "    Dim rT As Long, rB As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    rT = Me.Range(""OC_CrewTop"").Row" & vbCrLf
    code = code & "    rB = Me.Range(""OC_CrewBot"").Row" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "    If rT = 0 Or rB = 0 Then Exit Sub" & vbCrLf
    code = code & "    If Target.Column <> 10 Then Exit Sub" & vbCrLf
    code = code & "    If Target.Row < rT Or Target.Row > rB Then Exit Sub" & vbCrLf
    code = code & "    Application.EnableEvents = False" & vbCrLf
    code = code & "    MDL_Setup.SortCrewByRole" & vbCrLf
    code = code & "    Application.EnableEvents = True" & vbCrLf
    code = code & "End Sub"
    cm.AddFromString code
    Exit Sub
NoAccess:
    Application.StatusBar = "Crew auto-sort: enable Trust Access to VBA Project, or run SortCrewByRole manually."
End Sub

' Sort crew rows: DD first, MWD second, blank last; alpha by name within each group
Public Sub SortCrewByRole()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim rTop As Long, rBot As Long
    On Error Resume Next
    rTop = ws.Range("OC_CrewTop").Row
    rBot = ws.Range("OC_CrewBot").Row
    On Error GoTo 0
    If rTop = 0 Or rBot = 0 Or rBot < rTop Then Exit Sub

    Dim n As Long: n = rBot - rTop + 1
    If n < 2 Then Exit Sub

    ' Read current grid into arrays
    Dim aRole() As String, aName() As String
    Dim aEmail() As String, aPhone() As String
    ReDim aRole(1 To n): ReDim aName(1 To n)
    ReDim aEmail(1 To n): ReDim aPhone(1 To n)
    Dim idx As Long
    For idx = 1 To n
        Dim ro As Long: ro = rTop + idx - 1
        aRole(idx) = Trim(SafeStr(ws.Cells(ro, C_CRL)))
        aName(idx) = Trim(SafeStr(ws.Cells(ro, C_CRL + 1)))
        aEmail(idx) = Trim(SafeStr(ws.Cells(ro, C_CRL + 3)))
        aPhone(idx) = Trim(SafeStr(ws.Cells(ro, C_CRL + 6)))
    Next idx

    ' Insertion sort by (priority, name)
    Dim j As Long, k As Long
    For j = 2 To n
        Dim tr As String: tr = aRole(j)
        Dim tN As String: tN = aName(j)
        Dim tE As String: tE = aEmail(j)
        Dim tP As String: tP = aPhone(j)
        Dim pj As Long: pj = RolePriority(tr)
        k = j - 1
        Do While k >= 1
            Dim pk As Long: pk = RolePriority(aRole(k))
            If pk > pj Or (pk = pj And LCase(aName(k)) > LCase(tN)) Then
                aRole(k + 1) = aRole(k)
                aName(k + 1) = aName(k)
                aEmail(k + 1) = aEmail(k)
                aPhone(k + 1) = aPhone(k)
                k = k - 1
            Else
                Exit Do
            End If
        Loop
        aRole(k + 1) = tr
        aName(k + 1) = tN
        aEmail(k + 1) = tE
        aPhone(k + 1) = tP
    Next j

    ' Write back, restyle, and re-apply dropdown
    ScreenBeginBusy
    For idx = 1 To n
        ro = rTop + idx - 1
        Dim bg As Long
        If (idx - 1) Mod 2 = 0 Then bg = cWh() Else bg = cBg()

        With ws.Cells(ro, C_CRL)
            .Value = aRole(idx)
            .Interior.Color = bg
            .Font.Color = cDk(): .Font.bold = True: .Font.Size = 7
            .Font.name = "Consolas"
        End With

        ' Reapply dropdown validation
        With ws.Cells(ro, C_CRL).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="DD,MWD"
            .InCellDropdown = True
            .ShowError = False
        End With

        Dim mrg As Range
        Set mrg = ws.Range(ws.Cells(ro, C_CRL + 1), ws.Cells(ro, C_CRL + 2))
        mrg.Cells(1, 1).Value = aName(idx)
        mrg.Interior.Color = bg: mrg.Font.Color = cBlk()
        mrg.Font.bold = (aName(idx) <> ""): mrg.Font.Size = 8
        mrg.Font.name = "Consolas"

        Set mrg = ws.Range(ws.Cells(ro, C_CRL + 3), ws.Cells(ro, C_CRL + 5))
        mrg.Cells(1, 1).Value = aEmail(idx)
        mrg.Interior.Color = bg: mrg.Font.Color = cDk()
        mrg.Font.Size = 7: mrg.Font.name = "Consolas"
        ApplyMailtoLink mrg

        Set mrg = ws.Range(ws.Cells(ro, C_CRL + 6), ws.Cells(ro, C_LAST))
        mrg.Cells(1, 1).Value = aPhone(idx)
        mrg.Interior.Color = bg: mrg.Font.Color = cDk()
        mrg.Font.Size = 7: mrg.Font.name = "Consolas"
    Next idx
    ScreenEndBusy
End Sub

Private Function RolePriority(rl As String) As Long
    Select Case UCase(Trim(rl))
        Case "DD":  RolePriority = 1
        Case "MWD": RolePriority = 2
        Case Else:  RolePriority = 3
    End Select
End Function

' ================================================================================
'  FILES STATUS PANEL
' ================================================================================

Private Sub DrawFilesPanel(ws As Worksheet, startRow As Long)
    Dim cL As Long: cL = C_CRL   ' left col  = J (10)
    Dim cR As Long: cR = C_LAST  ' right col = R (18)
    Dim r As Long: r = startRow
    Dim wbPath As String: wbPath = ThisWorkbook.Path

    ' Header
    SectionBar ws, r, cL, cR, "  OPENCAP EXPORT FILES", cTeal(), cBlk()
    r = r + 1

    ' Column sub-headers
    ws.Rows(r).rowHeight = 14
    FHdr ws, r, cL, cL, "STATUS"
    FHdr ws, r, cL + 1, cL + 2, "FILE TYPE"
    FHdr ws, r, cL + 3, cL + 5, "FILENAME"
    FHdr ws, r, cL + 6, cL + 6, "ROWS"
    FHdr ws, r, cL + 7, cR, "PATH"
    With ws.Range(ws.Cells(r, cL), ws.Cells(r, cR)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = cMed(): .Weight = xlHairline
    End With
    r = r + 1

    Dim panelTop As Long: panelTop = r

    Dim tokens(5) As String, labels(5) As String, sheets(5) As String
    tokens(0) = TOK_JOB:       labels(0) = "Job Details":           sheets(0) = SH_JOB
    tokens(1) = TOK_CREW:      labels(1) = "Crew / Personnel":      sheets(1) = SH_CREW
    tokens(2) = TOK_BHA:       labels(2) = "BHA Equipment":         sheets(2) = SH_BHA
    tokens(3) = TOK_SLIDE:     labels(3) = "Slide / Rotate by Day": sheets(3) = SH_SLIDE
    tokens(4) = TOK_INVENTORY: labels(4) = "Equipment Inventory":   sheets(4) = SH_INVENTORY
    tokens(5) = TOK_COSTS:     labels(5) = "Ticket Costs by Day":   sheets(5) = SH_COSTS

    Dim fi As Long
    For fi = 0 To 5
        ws.Rows(r).rowHeight = 18

        Dim fPath As String: fPath = ""
        If wbPath <> "" Then fPath = FindCsvByToken(wbPath, tokens(fi))
        Dim found As Boolean: found = (fPath <> "")

        If found Then
            Sgl ws, r, cL, ChrW(10003), cGrnBadge(), RGB(255, 255, 255), True, 8, xlHAlignCenter
            MrgCell ws, r, cL + 1, cL + 2, labels(fi), cGrnRow(), cGrnTxt(), True, 7, xlHAlignLeft

            Dim fn As String: fn = mid(fPath, InStrRev(fPath, Application.PathSeparator) + 1)
            MrgCell ws, r, cL + 3, cL + 5, fn, cGrnRow(), cGrnTxt(), False, 7, xlHAlignLeft

            Dim rc As Long: rc = 0
            If SheetExists(sheets(fi)) Then
                On Error Resume Next
                Dim dw As Worksheet: Set dw = Worksheets(sheets(fi))
                rc = dw.Cells(dw.Rows.Count, 1).End(xlUp).Row - 1
                If rc < 0 Then rc = 0
                On Error GoTo 0
            End If
            Sgl ws, r, cL + 6, CStr(rc), cGrnRow(), cGrnTxt(), False, 7, xlHAlignRight

            MrgCell ws, r, cL + 7, cR, fPath, cGrnRow(), cDk(), False, 6, xlHAlignLeft
        Else
            Sgl ws, r, cL, ChrW(10007), cRedBadge(), RGB(255, 255, 255), True, 8, xlHAlignCenter
            MrgCell ws, r, cL + 1, cL + 2, labels(fi), cRedRow(), cRedTxt(), True, 7, xlHAlignLeft
            Dim hint As String
            hint = "*" & tokens(fi) & "*.csv"
            If wbPath = "" Then hint = "(save workbook first)"
            MrgCell ws, r, cL + 3, cL + 6, hint, cRedRow(), cRedTxt(), False, 7, xlHAlignLeft
            MrgCell ws, r, cL + 7, cR, IIf(wbPath = "", "", wbPath), cRedRow(), cDk(), False, 6, xlHAlignLeft
        End If

        With ws.Range(ws.Cells(r, cL), ws.Cells(r, cR)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Color = RGB(210, 210, 210): .Weight = xlHairline
        End With
        r = r + 1
    Next fi

    ' Extension + repo links (sit under the file checklist; natural "where do I get CSVs?" spot)
    ws.Rows(r).rowHeight = 16
    MrgCell ws, r, cL, cL + 1, "GET DATA", cBg(), cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, cL + 2, cL + 5, "Chrome Web Store", cBg(), cLink(), False, 7, xlHAlignLeft
    ApplyHttpLink ws.Cells(r, cL + 2), URL_OC_CHROME, "Chrome Web Store"
    MrgCell ws, r, cL + 6, cR, "GitHub", cBg(), cLink(), False, 7, xlHAlignLeft
    ApplyHttpLink ws.Cells(r, cL + 6), URL_OC_GITHUB, "GitHub"
    With ws.Range(ws.Cells(r, cL), ws.Cells(r, cR)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Color = cMed(): .Weight = xlHairline
    End With

    With ws.Range(ws.Cells(panelTop, cL), ws.Cells(r, cR))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Color = cMed()
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Color = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = cMed()
    End With
End Sub

Private Sub DrawNotesPanel(ws As Worksheet, startRow As Long)
    Dim cL As Long: cL = C_CRL
    Dim cR As Long: cR = C_LAST
    Dim r As Long: r = startRow

    SectionBar ws, r, cL, cR, "  EOW / JOB NOTES", cTeal(), cBlk()
    ws.Rows(r).rowHeight = 18
    r = r + 1

    ' Five note rows (r .. r+4) - matches the hand-tuned Field layout (J33:R37).
    Dim nr As Long
    For nr = r To r + 4: ws.Rows(nr).rowHeight = 16: Next nr
    With ws.Range(ws.Cells(r, cL), ws.Cells(r + 4, cR))
        .Merge
        .Interior.Color = cBg()
        .Font.Color = cDk()
        .Font.Size = 9
        .Font.name = "Consolas"
        .VerticalAlignment = xlVAlignTop
        .HorizontalAlignment = xlHAlignLeft
        .WrapText = True
        With .Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Color = cMed(): End With
        With .Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Color = cMed(): End With
        With .Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Color = cMed(): End With
        With .Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Color = cMed(): End With
    End With
End Sub

' ================================================================================
'  IMPORT FILES PANEL  (rows 26-28, cols J-R)
' ================================================================================

Private Sub DrawImportFilesPanel(ws As Worksheet, startRow As Long)
    Dim cL As Long: cL = C_CRL
    Dim cR As Long: cR = C_LAST
    ' Button host = J:L (exact fit for Import buttons). Path = M:R (no overlap under buttons).
    Dim cBtnEnd As Long: cBtnEnd = cL + 2
    Dim cPath As Long: cPath = cL + 3
    Dim r As Long: r = startRow

    SectionBar ws, r, cL, cR, "  IMPORT FILES", cTealLt(), cBlk()
    ws.Rows(r).rowHeight = 20
    r = r + 1

    ' Sub-header
    ws.Rows(r).rowHeight = 12
    MrgCell ws, r, cL, cBtnEnd, "FILE", cBg(), cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, cPath, cR, "SELECTED PATH", cBg(), cDk(), True, 7, xlHAlignLeft
    With ws.Range(ws.Cells(r, cL), ws.Cells(r, cR)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = cMed(): .Weight = xlHairline
    End With
    r = r + 1

    Dim planPath As String: planPath = GetImportPath(SH_SURVEY)
    ws.Rows(r).rowHeight = 22
    ' Empty merged host — Import Plan button is sized to this MergeArea exactly
    MrgCell ws, r, cL, cBtnEnd, "", cWh(), cBlk(), False, 8, xlHAlignCenter
    MrgCell ws, r, cPath, cR, planPath, cWh(), cDk(), False, 7, xlHAlignLeft
    With ws.Range(ws.Cells(r, cL), ws.Cells(r, cR)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(230, 230, 230): .Weight = xlHairline
    End With
    On Error Resume Next: ws.names("OC_ImpPlanRow").Delete: On Error GoTo 0
    ws.names.Add name:="OC_ImpPlanRow", RefersToR1C1:="=R" & r & "C" & cL
    r = r + 1

    Dim acPath As String: acPath = GetImportPath(SH_AC)
    ws.Rows(r).rowHeight = 22
    MrgCell ws, r, cL, cBtnEnd, "", cBg(), cBlk(), False, 8, xlHAlignCenter
    MrgCell ws, r, cPath, cR, acPath, cBg(), cDk(), False, 7, xlHAlignLeft
    On Error Resume Next: ws.names("OC_ImpAcRow").Delete: On Error GoTo 0
    ws.names.Add name:="OC_ImpAcRow", RefersToR1C1:="=R" & r & "C" & cL

    With ws.Range(ws.Cells(startRow, cL), ws.Cells(r, cR))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Color = cMed()
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Color = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = cMed()
    End With
End Sub

Private Function GetImportPath(shName As String) As String
    GetImportPath = ""
    If Not SheetExists(shName) Then Exit Function
    On Error Resume Next
    Dim v As String: v = Trim(CStr(Worksheets(shName).Cells(1, 1).Value))
    If Err.Number = 0 And Len(v) > 0 Then GetImportPath = v
    On Error GoTo 0
End Function

' ================================================================================
'  CONTROL BUTTONS
' ================================================================================

Private Sub AttachButtons(ws As Worksheet)
    Dim b As Button
    For Each b In ws.Buttons: b.Delete: Next b

    ' REFRESH CSVs (right side of title bar)
    Dim rc As Range: Set rc = ws.Cells(R_HDR, C_LAST - 3)
    Dim bR As Button
    Set bR = ws.Buttons.Add(rc.Left, rc.Top + 4, rc.Width * 3.5, rc.Height - 8)
    bR.name = "BtnOCRefresh"
    bR.OnAction = "'" & ThisWorkbook.name & "'!RefreshSetup"
    bR.Placement = xlMoveAndSize
    StyleBtn bR, "REFRESH CSVs", cBg(), cMed(), cBlk(), True

    ' REBUILD (a bit further left)
    Dim rB As Range: Set rB = ws.Cells(R_HDR, C_LAST - 7)
    Dim bB As Button
    Set bB = ws.Buttons.Add(rB.Left, rB.Top + 4, rB.Width * 3, rB.Height - 8)
    bB.name = "BtnOCRebuild"
    bB.OnAction = "'" & ThisWorkbook.name & "'!RebuildSetup"
    bB.Placement = xlMoveAndSize
    StyleBtn bB, "REBUILD", cBg(), cMed(), cBlk(), True

    ' Import Plan / AC — embed exactly in the FILE host MergeArea (J:L)
    Dim planCell As Range
    On Error Resume Next: Set planCell = ws.Range("OC_ImpPlanRow").MergeArea: On Error GoTo 0
    If Not planCell Is Nothing Then
        Dim bIP As Button
        Set bIP = ws.Buttons.Add(planCell.Left, planCell.Top, planCell.Width, planCell.Height)
        bIP.name = "BtnImportPlan"
        bIP.OnAction = "'" & ThisWorkbook.name & "'!ImportSurveyPlan"
        bIP.Placement = xlMoveAndSize
        StyleBtn bIP, "Import Plan", cWh(), cMed(), cBlk(), False
    End If

    Dim acCell As Range
    On Error Resume Next: Set acCell = ws.Range("OC_ImpAcRow").MergeArea: On Error GoTo 0
    If Not acCell Is Nothing Then
        Dim bAC As Button
        Set bAC = ws.Buttons.Add(acCell.Left, acCell.Top, acCell.Width, acCell.Height)
        bAC.name = "BtnImportAC"
        bAC.OnAction = "'" & ThisWorkbook.name & "'!ImportAntiCollision"
        bAC.Placement = xlMoveAndSize
        StyleBtn bAC, "Import AC", cBg(), cMed(), cBlk(), False
    End If
End Sub

Private Sub StyleBtn(btn As Button, cap As String, bg As Long, brd As Long, fg As Long, bold As Boolean)
    On Error Resume Next
    btn.caption = cap
    btn.Font.name = "Consolas": btn.Font.Size = 9: btn.Font.bold = bold: btn.Font.Color = fg
    btn.ShapeRange.fill.ForeColor.RGB = bg
    btn.ShapeRange.line.ForeColor.RGB = brd
    btn.ShapeRange.line.Weight = 0.75
    On Error GoTo 0
End Sub

' ================================================================================
'  CELL / RANGE DRAWING HELPERS
' ================================================================================

Private Sub SectionBar(ws As Worksheet, r As Long, c1 As Long, c2 As Long, _
                        title As String, bg As Long, fg As Long)
    ws.Rows(r).rowHeight = 18
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Merge
        .Value = title
        .Interior.Color = bg
        .Font.Color = fg
        .Font.bold = True
        .Font.Size = 8
        .Font.name = "Consolas"
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
    End With
End Sub

' Pair row: left-panel label+value pair + right-panel label+value pair
Private Sub Pair(ws As Worksheet, r As Long, lbl1 As String, val1 As String, _
                                              lbl2 As String, val2 As String)
    ws.Rows(r).rowHeight = 17

    ' Accent strip
    ws.Cells(r, C_ACCENT).Interior.Color = cAccent()

    ' Label 1 (B:C)
    MrgCell ws, r, C_L1S, C_L1E, lbl1, cBg(), cDk(), True, 7, xlHAlignRight

    ' Value 1 (D:E)
    With ws.Range(ws.Cells(r, C_V1S), ws.Cells(r, C_V1E))
        .Merge
        .numberFormat = "@"
        .Value = val1
        .Interior.Color = cWh()
        .Font.Color = cBlk()
        .Font.bold = (val1 <> "")
        .Font.Size = 9
        .Font.name = "Consolas"
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
        .IndentLevel = 1
    End With

    ' Label 2 (F)
    Sgl ws, r, C_L2, lbl2, cBg(), cDk(), True, 7, xlHAlignRight

    ' Value 2 (G:H) — emails use slightly smaller type so mailto links fit the row
    If InStr(1, lbl2, "EMAIL", vbTextCompare) > 0 Or LooksLikeEmail(val2) Then
        MrgCell ws, r, C_V2S, C_V2E, val2, cWh(), cBlk(), False, 8, xlHAlignLeft
        ApplyMailtoLink ws.Range(ws.Cells(r, C_V2S), ws.Cells(r, C_V2E))
    Else
        MrgCell ws, r, C_V2S, C_V2E, val2, cWh(), cBlk(), (val2 <> ""), 9, xlHAlignLeft
    End If

    If InStr(1, lbl1, "EMAIL", vbTextCompare) > 0 Or LooksLikeEmail(val1) Then
        With ws.Range(ws.Cells(r, C_V1S), ws.Cells(r, C_V1E))
            .Font.Size = 8
            .Font.bold = False
        End With
        ApplyMailtoLink ws.Range(ws.Cells(r, C_V1S), ws.Cells(r, C_V1E))
    End If

    With ws.Range(ws.Cells(r, C_ACCENT), ws.Cells(r, C_V2E)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(232, 232, 232): .Weight = xlHairline
    End With
End Sub

' mailto: hyperlink so click opens the default mail client (Outlook / webmail / etc.)
Private Function LooksLikeEmail(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    LooksLikeEmail = (Len(t) >= 5 And InStr(t, "@") > 1 And InStr(t, ".") > 0 And InStr(t, " ") = 0)
End Function

' http(s) hyperlink — keep Consolas sizing (Hyperlinks.Add resets font otherwise)
Private Sub ApplyHttpLink(rng As Range, ByVal url As String, ByVal displayText As String)
    Dim anchor As Range
    Dim h As Hyperlink
    Dim sz As Double
    Dim wasBold As Boolean
    Dim fName As String
    On Error GoTo Fail
    If rng Is Nothing Then Exit Sub
    If Len(Trim$(url)) = 0 Then Exit Sub

    Set anchor = rng.MergeArea.Cells(1, 1)
    If Len(Trim$(displayText)) > 0 Then anchor.Value = displayText

    sz = anchor.Font.Size
    wasBold = anchor.Font.bold
    fName = anchor.Font.name
    If sz < 6# Or sz > 10# Then sz = 7#

    On Error Resume Next
    For Each h In anchor.Hyperlinks
        h.Delete
    Next h
    On Error GoTo Fail

    anchor.Worksheet.Hyperlinks.Add anchor:=anchor, Address:=url, TextToDisplay:=displayText
    With anchor.Font
        .Size = sz
        .bold = wasBold
        .name = fName
        If Len(.name) = 0 Then .name = "Consolas"
        .Color = cLink()
        .Underline = xlUnderlineStyleSingle
    End With
    Exit Sub
Fail:
End Sub

Private Sub ApplyMailtoLink(rng As Range)
    Dim em As String
    Dim anchor As Range
    Dim h As Hyperlink
    Dim sz As Double
    Dim wasBold As Boolean
    Dim fName As String
    On Error GoTo Fail
    If rng Is Nothing Then Exit Sub
    Set anchor = rng.MergeArea.Cells(1, 1)
    em = SafeStr(anchor)
    If Not LooksLikeEmail(em) Then Exit Sub

    ' Preserve size/name before Hyperlinks.Add (Excel resets font to ~11pt)
    sz = anchor.Font.Size
    wasBold = anchor.Font.bold
    fName = anchor.Font.name
    If sz < 6# Or sz > 10# Then sz = 8#

    On Error Resume Next
    For Each h In anchor.Hyperlinks
        h.Delete
    Next h
    On Error GoTo Fail

    ' Address only — keep the cell's existing display text
    anchor.Worksheet.Hyperlinks.Add anchor:=anchor, Address:="mailto:" & em
    With anchor.Font
        .Size = sz
        .bold = wasBold
        .name = fName
        If Len(.name) = 0 Then .name = "Consolas"
        .Color = cLink()
        .Underline = xlUnderlineStyleSingle
    End With
    Exit Sub
Fail:
End Sub

' Sweep a sheet for email-looking values and ensure mailto: links exist.
Private Sub LinkMailtoInUsedRange(ws As Worksheet)
    Dim c As Range
    On Error Resume Next
    For Each c In ws.UsedRange.Cells
        If c.MergeCells Then
            If c.Address <> c.MergeArea.Cells(1, 1).Address Then GoTo NextCell
        End If
        If LooksLikeEmail(SafeStr(c)) Then ApplyMailtoLink c
NextCell:
    Next c
    On Error GoTo 0
End Sub

' Public entry: link all emails on Setup (protect-aware).
Public Sub LinkSetupEmails()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    wasProt = SheetUnprotectForVba(ws)
    LinkMailtoInUsedRange ws
    SheetReprotectAfterVba ws, wasProt
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
End Sub

Private Sub MrgCell(ws As Worksheet, r As Long, c1 As Long, c2 As Long, _
                    val As String, bg As Long, fg As Long, bold As Boolean, _
                    sz As Long, align As Long)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Merge
        .Value = val
        .Interior.Color = bg
        .Font.Color = fg
        .Font.bold = bold
        .Font.Size = sz
        .Font.name = "Consolas"
        .HorizontalAlignment = align
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub Sgl(ws As Worksheet, r As Long, c As Long, val As String, _
                bg As Long, fg As Long, bold As Boolean, sz As Long, align As Long)
    With ws.Cells(r, c)
        .Value = val
        .Interior.Color = bg
        .Font.Color = fg
        .Font.bold = bold
        .Font.Size = sz
        .Font.name = "Consolas"
        .HorizontalAlignment = align
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub FHdr(ws As Worksheet, r As Long, c1 As Long, c2 As Long, lbl As String)
    If c1 = c2 Then
        Sgl ws, r, c1, lbl, cBg(), cDk(), True, 7, xlHAlignLeft
    Else
        MrgCell ws, r, c1, c2, lbl, cBg(), cDk(), True, 7, xlHAlignLeft
    End If
End Sub

' ================================================================================
'  DATA READERS
' ================================================================================

Private Function ReadJobFields() As Object
    Set ReadJobFields = CreateObject("Scripting.Dictionary")
    ReadJobFields.CompareMode = vbTextCompare
    If Not SheetExists(SH_JOB) Then Exit Function
    Dim ws As Worksheet: Set ws = Worksheets(SH_JOB)
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long
    For c = 1 To lastCol
        Dim hd As String: hd = Trim(SafeStr(ws.Cells(1, c)))
        If hd <> "" And Not ReadJobFields.Exists(hd) Then
            ReadJobFields.Add hd, Trim(SafeStr(ws.Cells(2, c)))
        End If
    Next c
End Function

Private Function GF(dict As Object, k1 As String, Optional k2 As String = "", _
                    Optional k3 As String = "") As String
    GF = ""
    If dict Is Nothing Then Exit Function
    If dict.Exists(k1) Then GF = CStr(dict(k1)): If GF <> "" Then Exit Function
    If k2 <> "" And dict.Exists(k2) Then GF = CStr(dict(k2)): If GF <> "" Then Exit Function
    If k3 <> "" And dict.Exists(k3) Then GF = CStr(dict(k3))
End Function

' OpenCap Job Details often ships a blank "Client" column. Fall back to
' Company/Operator, then the first non-PHX address on ClientEmail.
Private Function JobClientName(j As Object) As String
    JobClientName = GF(j, "Client", "Company", "Operator")
    If JobClientName = "" Then JobClientName = GF(j, "Client Name", "ClientName")
    If JobClientName = "" Then JobClientName = ClientNameFromEmails(GF(j, "ClientEmail"))
End Function

Private Function ClientNameFromEmails(ByVal raw As String) As String
    Dim parts() As String
    Dim i As Long, em As String, at As Long, domain As String, Dot As Long
    ClientNameFromEmails = ""
    If Len(Trim$(raw)) = 0 Then Exit Function
    raw = Replace(Replace(raw, ",", ";"), " ", "")
    parts = Split(raw, ";")
    For i = LBound(parts) To UBound(parts)
        em = LCase$(Trim$(parts(i)))
        If em = "" Then GoTo NextClientEm
        at = InStr(em, "@")
        If at < 2 Then GoTo NextClientEm
        If InStr(em, "phxtech.com") > 0 Then GoTo NextClientEm
        domain = mid$(em, at + 1)
        Dot = InStrRev(domain, ".")
        If Dot > 1 Then domain = Left$(domain, Dot - 1)
        domain = Replace(Replace(domain, "-", " "), ".", " ")
        ClientNameFromEmails = StrConv(domain, vbProperCase)
        Exit Function
NextClientEm:
    Next i
End Function

Public Function OcJobField(ByVal k1 As String, Optional ByVal k2 As String = "") As String
    OcJobField = GF(ReadJobFields(), k1, k2)
End Function

' Fill Setup JOB ID (Job Code) + CLIENT only — does not RebuildSetup.
Public Sub RefreshJobClientFromImport()
    Dim ws As Worksheet
    Dim r As Long
    Dim j As Object
    Dim wasProt As Boolean
    Dim lbl As String
    Dim v As String
    On Error GoTo Fail
    If Not SheetExists(SH_SETUP) Then Exit Sub
    Set j = ReadJobFields()
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    wasProt = SheetUnprotectForVba(ws)
    For r = 3 To 24
        lbl = UCase$(Trim$(CStr(ws.Cells(r, C_L1S).Value2 & "")))
        v = ""
        If lbl = "JOB ID" Then v = GF(j, "Job Code", "Job ID")
        If lbl = "CLIENT" Then v = JobClientName(j)
        If v <> "" Then
            ws.Range(ws.Cells(r, C_V1S), ws.Cells(r, C_V1E)).Value = v
            ws.Range(ws.Cells(r, C_V1S), ws.Cells(r, C_V1E)).Font.bold = True
        End If
    Next r
    SheetReprotectAfterVba ws, wasProt
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
End Sub

Private Function ReadCrewRows() As String()
    Dim blankArr() As String
    ReDim blankArr(0, 3)
    ReadCrewRows = blankArr
    If Not SheetExists(SH_CREW) Then Exit Function

    Dim ws As Worksheet: Set ws = Worksheets(SH_CREW)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim cName  As Long: cName = ColByName(ws, "Name", "FullName")
    Dim cRole  As Long: cRole = ColByName(ws, "Role")
    Dim cWork  As Long: cWork = ColByName(ws, "Work Type", "WorkType", "Position")
    Dim cEmail As Long: cEmail = ColByName(ws, "Email")
    Dim cPhone As Long: cPhone = ColByName(ws, "Phone", "Mobile", "MobilePhone")
    If cName = 0 Then cName = 2

    Dim buf(200, 3) As String
    Dim n As Long: n = 0
    Dim r As Long
    For r = 2 To lastRow
        Dim nm As String:  nm = Trim(SafeStr(ws.Cells(r, cName)))
        Dim rl As String:  rl = ""
        Dim em As String:  em = ""
        Dim pH As String:  pH = ""

        If cRole > 0 Then rl = Trim(SafeStr(ws.Cells(r, cRole)))
        If rl = "" And cWork > 0 Then rl = Trim(SafeStr(ws.Cells(r, cWork)))
        If cEmail > 0 Then em = Trim(SafeStr(ws.Cells(r, cEmail)))
        If cPhone > 0 Then pH = Trim(SafeStr(ws.Cells(r, cPhone)))

        If nm = "" Then GoTo SkipRow
        buf(n, 0) = rl
        buf(n, 1) = nm
        buf(n, 2) = em
        buf(n, 3) = pH
        n = n + 1
SkipRow:
    Next r

    If n = 0 Then Exit Function
    Dim out() As String: ReDim out(n - 1, 3)
    Dim i As Long
    For i = 0 To n - 1
        out(i, 0) = buf(i, 0)
        out(i, 1) = buf(i, 1)
        out(i, 2) = buf(i, 2)
        out(i, 3) = buf(i, 3)
    Next i
    ReadCrewRows = out
End Function

' ================================================================================
'  CSV IMPORT
' ================================================================================

Private Sub LoadCsv(filePath As String, shName As String)
    If Not SheetExists(shName) Then Exit Sub
    Dim ws As Worksheet: Set ws = Worksheets(shName)
    ws.Cells.Clear

    ' -- Phase 1: read every raw line -------------------------------------------
    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Input As #fNum
    Dim rawLines() As String
    ReDim rawLines(0 To 4999)
    Dim nLines As Long: nLines = 0
    Do While Not EOF(fNum) And nLines < 5000
        Line Input #fNum, rawLines(nLines)
        nLines = nLines + 1
    Loop
    Close #fNum
    If nLines = 0 Then Exit Sub
    ReDim Preserve rawLines(0 To nLines - 1)

    ' -- Phase 2: parse all rows, track actual max column count ----------------
    Dim parsedRows() As String
    ReDim parsedRows(0 To nLines - 1)   ' store each parsed row as a temp placeholder
    Dim allFields() As Variant
    ReDim allFields(0 To nLines - 1)    ' holds each row's String() result

    Dim nCols As Long: nCols = 0
    Dim ri As Long
    For ri = 0 To nLines - 1
        Dim rowArr() As String: rowArr = CsvParseLine(rawLines(ri))
        Dim rCols As Long: rCols = UBound(rowArr) + 1
        If rCols > nCols Then nCols = rCols
        allFields(ri) = rowArr
    Next ri
    If nCols = 0 Then Exit Sub

    ' -- Phase 3: build Variant 2D array then bulk-write -----------------------
    Dim data() As Variant
    ReDim data(1 To nLines, 1 To nCols)
    Dim c As Long
    Dim fa() As String
    For ri = 0 To nLines - 1
        fa = allFields(ri)
        Dim nc As Long: nc = UBound(fa)
        For c = 0 To nc
            data(ri + 1, c + 1) = CsvConv(fa(c))
        Next c
    Next ri
    ws.Range(ws.Cells(1, 1), ws.Cells(nLines, nCols)).Value = data
End Sub

' Public hunt used by Setup REFRESH CSVs and Data REFRESH / DD Tools.
' Same folders: workbook\OpenCap, workbook root, OC_CsvRoot, Setup path list.
Public Function FindOpenCapCsv(ByVal token As String) As String
    Dim wbPath As String
    FindOpenCapCsv = ""
    wbPath = ThisWorkbook.Path
    If Len(wbPath) = 0 Then Exit Function
    RememberCsvRootsFromSetup
    FindOpenCapCsv = FindCsvByToken(wbPath, token)
End Function

Private Function FindCsvByToken(folder As String, token As String) As String
    ' Newest matching CSV across workbook\OpenCap, workbook root, and remembered job folders.
    Dim bestPath As String
    Dim bestDate As Date
    bestPath = ""
    bestDate = 0

    ConsiderCsvFolder folder & Application.PathSeparator & "OpenCap", token, bestPath, bestDate
    ConsiderCsvFolder folder, token, bestPath, bestDate
    ConsiderCsvFolder GetNamedText(NM_CSV_ROOT), token, bestPath, bestDate

    Dim extra As Variant
    If Not mCsvExtraRoots Is Nothing Then
        For Each extra In mCsvExtraRoots
            ConsiderCsvFolder CStr(extra), token, bestPath, bestDate
        Next extra
    End If

    FindCsvByToken = bestPath
    If bestPath <> "" Then
        SetNamedText NM_CSV_ROOT, ParentFolderOf(bestPath)
    End If
End Function

Private Sub ConsiderCsvFolder(ByVal folder As String, ByVal token As String, _
                              ByRef bestPath As String, ByRef bestDate As Date)
    Dim fp As String
    Dim st As Date
    fp = FindCsvByTokenInFolder(folder, token)
    If fp = "" Then Exit Sub
    On Error Resume Next
    st = FileDateTime(fp)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Sub
    On Error GoTo 0
    If bestPath = "" Or st >= bestDate Then
        bestDate = st
        bestPath = fp
    End If
End Sub

Private Function ParentFolderOf(ByVal filePath As String) As String
    Dim p As Long
    ParentFolderOf = ""
    p = InStrRev(filePath, Application.PathSeparator)
    If p > 1 Then ParentFolderOf = Left$(filePath, p - 1)
End Function

Private Sub RememberCsvRootsFromSetup()
    Dim ws As Worksheet
    Dim v As String
    Dim folder As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Set mCsvExtraRoots = New Collection
    folder = GetNamedText(NM_CSV_ROOT)
    AddCsvExtraRoot folder

    Dim r As Long, c As Long
    For r = 18 To 30
        For c = 10 To 18
            v = CStr(ws.Cells(r, c).Value2 & "")
            If InStr(1, v, "OpenCap", vbTextCompare) > 0 Then
                If InStr(1, v, ".csv", vbTextCompare) > 0 Then
                    AddCsvExtraRoot ParentFolderOf(v)
                Else
                    AddCsvExtraRoot v
                End If
            End If
        Next c
    Next r
End Sub

Private Sub AddCsvExtraRoot(ByVal folder As String)
    Dim k As String
    If Len(Trim$(folder)) = 0 Then Exit Sub
    If mCsvExtraRoots Is Nothing Then Set mCsvExtraRoots = New Collection
    k = LCase$(Trim$(folder))
    On Error Resume Next
    mCsvExtraRoots.Add folder, k
    On Error GoTo 0
End Sub

Private Sub DetectNewJob()
    Dim cur As String
    Dim prev As String
    cur = CurrentOcJobId()
    prev = GetNamedText(NM_LAST_JOB)
    mNewJob = (Len(cur) > 0 And Len(prev) > 0 And StrComp(cur, prev, vbTextCompare) <> 0)
    If Len(cur) > 0 Then SetNamedText NM_LAST_JOB, cur
End Sub

Private Function CurrentOcJobId() As String
    CurrentOcJobId = Trim$(GF(ReadJobFields(), "Job ID"))
End Function

Private Function GetNamedText(ByVal nm As String) As String
    Dim rf As String
    GetNamedText = ""
    On Error Resume Next
    rf = CStr(ThisWorkbook.names(nm).refersTo)
    On Error GoTo 0
    If Len(rf) >= 3 Then
        If Left$(rf, 2) = "=""" And right$(rf, 1) = """" Then
            GetNamedText = mid$(rf, 3, Len(rf) - 3)
        ElseIf Left$(rf, 1) = "=" Then
            GetNamedText = Replace(mid$(rf, 2), """", "")
        End If
    End If
End Function

Private Sub SetNamedText(ByVal nm As String, ByVal v As String)
    Dim safe As String
    safe = Replace(Trim$(v), """", "")
    On Error Resume Next
    ThisWorkbook.names(nm).Delete
    On Error GoTo 0
    If Len(safe) = 0 Then Exit Sub
    ThisWorkbook.names.Add name:=nm, refersTo:="=""" & safe & """"
End Sub

Private Function FindCsvByTokenInFolder(folder As String, token As String) As String
    FindCsvByTokenInFolder = ""
    If folder = "" Then Exit Function

    Dim newest As Date: newest = 0
    Dim fn As String
    On Error Resume Next
    fn = Dir(folder & Application.PathSeparator & "*.csv")
    On Error GoTo 0
    Do While fn <> ""
        If InStr(LCase(fn), LCase(token)) > 0 Then
            Dim fp As String: fp = folder & Application.PathSeparator & fn
            On Error Resume Next
            Dim st As Date: st = FileDateTime(fp)
            If Err.Number = 0 And st >= newest Then newest = st: FindCsvByTokenInFolder = fp
            Err.Clear: On Error GoTo 0
        End If
        fn = Dir()
    Loop
End Function

Private Function CsvParseLine(ByVal line As String) As String()
    Dim buf(500) As String
    Dim idx As Long: idx = 0
    Dim pos As Long: pos = 1
    Dim inQ As Boolean: inQ = False
    Dim tok As String: tok = ""
    Do While pos <= Len(line)
        Dim ch As String: ch = mid(line, pos, 1)
        If ch = Chr(34) Then
            If inQ And mid(line, pos + 1, 1) = Chr(34) Then
                tok = tok & Chr(34): pos = pos + 1
            Else: inQ = Not inQ
            End If
        ElseIf ch = "," And Not inQ Then
            buf(idx) = tok: idx = idx + 1: tok = ""
        Else: tok = tok & ch
        End If
        pos = pos + 1
    Loop
    buf(idx) = tok
    Dim out() As String: ReDim out(idx)
    Dim k As Long
    For k = 0 To idx: out(k) = buf(k): Next k
    CsvParseLine = out
End Function

Private Function CsvConv(s As String) As Variant
    s = Trim(s)
    If s = "" Then Exit Function   ' returns Empty variant -> cell stays blank
    If IsNumeric(s) Then CsvConv = CDbl(s) Else CsvConv = s
End Function

' ================================================================================
'  MUD MOTORS ON LOCATION
' ================================================================================
' Primary source: _OC_Inventory (inventory.csv) — ItemName contains "Mud Motor"
' or SubCategory = motor. SerialNumber is written as-is (PHX-*-PTS and 24X-*-PTS).
' Fallback: _OC_BHA Description contains "Mud Motor".
' Writes serials into Data B45:B55 (Motors On Location). Front O-table and
' hours are owned by MDL_ToolHours (hidden _TH_Hours tracker + active BHA).
' ================================================================================

Public Sub SyncMudMotorsFromBha()
    On Error GoTo Fail
    WriteMudMotorsFromInventory
    Exit Sub
Fail:
    Dim errN As Long, errD As String
    errN = Err.Number
    errD = Err.Description
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = _
        "MOTOR SYNC ERR " & errN & ": " & errD
    Application.StatusBar = "Mud motor sync failed: " & errD
End Sub

Private Sub WriteMudMotorsFromInventory()
    Dim wsD As Worksheet
    Dim sns() As String
    Dim seqs() As Double
    Dim n As Long
    Dim i As Long, j As Long
    Dim tmpS As String, tmpQ As Double
    Dim firstRow As Long, lastRow As Long, capacity As Long
    Dim wasProt As Boolean
    Dim r As Long
    Dim wbPath As String, fInv As String, fBha As String

    If Not SheetExists(SH_DATA) Then Exit Sub
    Set wsD = ThisWorkbook.Worksheets(SH_DATA)

    wbPath = ThisWorkbook.Path
    If wbPath <> "" Then
        fInv = FindCsvByToken(wbPath, TOK_INVENTORY)
        If fInv <> "" Then
            EnsureSheet SH_INVENTORY, False
            LoadCsv fInv, SH_INVENTORY
        End If
        fBha = FindCsvByToken(wbPath, TOK_BHA)
        If fBha <> "" Then
            EnsureSheet SH_BHA, False
            LoadCsv fBha, SH_BHA
        End If
    End If

    ReDim sns(1 To 64)
    ReDim seqs(1 To 64)
    n = 0
    CollectMudMotorsFromInventory sns, seqs, n
    If n = 0 Then CollectMudMotorsFromBha sns, seqs, n

    For i = 1 To n - 1
        For j = i + 1 To n
            If seqs(j) < seqs(i) Then
                tmpQ = seqs(i): seqs(i) = seqs(j): seqs(j) = tmpQ
                tmpS = sns(i): sns(i) = sns(j): sns(j) = tmpS
            End If
        Next j
    Next i

    firstRow = MM_FIRST_ROW
    lastRow = firstRow + MM_ROWS - 1
    capacity = MM_ROWS

    wasProt = SheetUnprotectForVba(wsD)
    On Error GoTo ReprotectFail

    For r = firstRow To lastRow
        ClearMergedCell wsD.Cells(r, MM_COL_SERIAL)
    Next r

    If n > capacity Then n = capacity
    For i = 1 To n
        SetMergedCellValue wsD.Cells(firstRow + i - 1, MM_COL_SERIAL), sns(i)
    Next i

    On Error Resume Next
    ToolHours_Sync
    On Error GoTo ReprotectFail

    SheetReprotectAfterVba wsD, wasProt
    Exit Sub

ReprotectFail:
    On Error Resume Next
    SheetReprotectAfterVba wsD, wasProt
End Sub

Private Sub CollectMudMotorsFromInventory(ByRef sns() As String, _
                                          ByRef seqs() As Double, _
                                          ByRef n As Long)
    Dim ws As Worksheet
    Dim colSn As Long, colName As Long, colSub As Long
    Dim lastR As Long, r As Long
    Dim sn As String, itemName As String, subCat As String
    If Not SheetExists(SH_INVENTORY) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_INVENTORY)
    colSn = FindHeaderCol(ws, "SerialNumber")
    If colSn = 0 Then colSn = FindHeaderCol(ws, "Serial #")
    colName = FindHeaderCol(ws, "ItemName")
    If colName = 0 Then colName = FindHeaderCol(ws, "Description")
    colSub = FindHeaderCol(ws, "SubCategory")
    If colSn = 0 Then Exit Sub
    lastR = ws.Cells(ws.Rows.Count, colSn).End(xlUp).Row
    For r = 2 To lastR
        sn = Trim$(CStr(ws.Cells(r, colSn).Value2 & ""))
        itemName = ""
        subCat = ""
        If colName > 0 Then itemName = CStr(ws.Cells(r, colName).Value2 & "")
        If colSub > 0 Then subCat = CStr(ws.Cells(r, colSub).Value2 & "")
        If IsMudMotorInventory(sn, itemName, subCat) Then
            AddUniqueMotor sns, seqs, n, sn
        End If
    Next r
End Sub

Private Sub CollectMudMotorsFromBha(ByRef sns() As String, _
                                    ByRef seqs() As Double, _
                                    ByRef n As Long)
    Dim ws As Worksheet
    Dim colSn As Long, colDesc As Long
    Dim lastR As Long, r As Long
    Dim sn As String, desc As String
    If Not SheetExists(SH_BHA) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_BHA)
    colSn = FindHeaderCol(ws, "Serial #")
    If colSn = 0 Then colSn = FindHeaderCol(ws, "SerialNumber")
    colDesc = FindHeaderCol(ws, "Description")
    If colSn = 0 Then Exit Sub
    lastR = ws.Cells(ws.Rows.Count, colSn).End(xlUp).Row
    For r = 2 To lastR
        sn = Trim$(CStr(ws.Cells(r, colSn).Value2 & ""))
        desc = ""
        If colDesc > 0 Then desc = CStr(ws.Cells(r, colDesc).Value2 & "")
        If IsMudMotorSerial(sn, desc) Then
            AddUniqueMotor sns, seqs, n, sn
        End If
    Next r
End Sub

Private Sub AddUniqueMotor(ByRef sns() As String, ByRef seqs() As Double, _
                           ByRef n As Long, ByVal sn As String)
    Dim i As Long
    If Len(sn) = 0 Then Exit Sub
    For i = 1 To n
        If StrComp(sns(i), sn, vbTextCompare) = 0 Then Exit Sub
    Next i
    n = n + 1
    If n > UBound(sns) Then
        ReDim Preserve sns(1 To n + 32)
        ReDim Preserve seqs(1 To n + 32)
    End If
    sns(n) = sn
    seqs(n) = MudMotorSequenceNumber(sn)
End Sub

' PHX-525010-PTS -> 525010; 24X-19722-PTS -> 19722; else sorts last.
Private Function MudMotorSequenceNumber(ByVal sn As String) As Double
    Dim s As String, p As Long, tok As String
    MudMotorSequenceNumber = 1E+99
    s = UCase$(Trim$(sn))
    If right$(s, 4) <> "-PTS" Then Exit Function
    s = Left$(s, Len(s) - 4)
    p = InStrRev(s, "-")
    If p > 0 Then tok = mid$(s, p + 1) Else tok = s
    If Len(tok) = 0 Then Exit Function
    If Not IsNumeric(tok) Then Exit Function
    MudMotorSequenceNumber = CDbl(tok)
End Function

Private Sub ClearMergedCell(ByVal cell As Range)
    Dim topLeft As Range
    If cell Is Nothing Then Exit Sub
    If cell.MergeCells Then
        Set topLeft = cell.MergeArea.Cells(1, 1)
    Else
        Set topLeft = cell
    End If
    topLeft.Value2 = ""
End Sub

Private Sub SetMergedCellValue(ByVal cell As Range, ByVal v As Variant)
    Dim topLeft As Range
    If cell Is Nothing Then Exit Sub
    If cell.MergeCells Then
        Set topLeft = cell.MergeArea.Cells(1, 1)
    Else
        Set topLeft = cell
    End If
    topLeft.Value2 = v
End Sub

Private Function IsSteerMotorText(ByVal text As String) As Boolean
    Dim t As String
    t = CStr(text & "")
    IsSteerMotorText = (InStr(1, t, "Orbit RSS", vbTextCompare) > 0) _
                    Or (InStr(1, t, "iCruise", vbTextCompare) > 0) _
                    Or (InStr(1, t, "i-Cruise", vbTextCompare) > 0)
End Function

Private Function IsMudMotorInventory(ByVal sn As String, ByVal itemName As String, _
                                    ByVal subCat As String) As Boolean
    IsMudMotorInventory = False
    If Len(Trim$(sn)) = 0 Then Exit Function
    If InStr(1, itemName, "Mud Motor", vbTextCompare) > 0 Then
        IsMudMotorInventory = True
        Exit Function
    End If
    If StrComp(Trim$(subCat), "motor", vbTextCompare) = 0 Then
        IsMudMotorInventory = True
        Exit Function
    End If
    IsMudMotorInventory = IsSteerMotorText(itemName)
End Function

Private Function IsMudMotorSerial(ByVal sn As String, ByVal desc As String) As Boolean
    IsMudMotorSerial = False
    If Len(sn) = 0 Then Exit Function
    If InStr(1, desc, "Mud Motor", vbTextCompare) > 0 Then
        IsMudMotorSerial = True
        Exit Function
    End If
    IsMudMotorSerial = IsSteerMotorText(desc)
End Function

Private Function FindHeaderCol(ws As Worksheet, ByVal headerName As String) As Long
    Dim c As Long, lastC As Long, h As String
    FindHeaderCol = 0
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastC < 1 Then Exit Function
    For c = 1 To lastC
        h = Trim$(CStr(ws.Cells(1, c).Value2 & ""))
        If StrComp(h, headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
End Function

' ================================================================================
'  COSTS TAB SYNC  (ticket-costs-by-day.csv -> Costs)
' ================================================================================
' Merges OpenCap daily ticket costs into the visible Costs sheet by DATE:
'  - keeps prior Costs rows for dates not in the CSV
'  - CSV values win on matching dates
'  - appends new CSV dates in chronological order
'  - column C dates follow the existing Day1 seed + cascade formulas when
'    the merged dates are consecutive; otherwise writes explicit dates
' ================================================================================

Public Sub SyncCostsFromOpenCap()
    On Error GoTo Fail
    SyncCostsSheet
    GoTo Fin
Fail:
    Dim errN As Long, errD As String
    errN = Err.Number
    errD = Err.Description
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = _
        "COSTS SYNC ERR " & errN & ": " & errD
    Application.StatusBar = "Costs sync failed: " & errD
    MsgBox "Costs sync failed:" & vbCrLf & vbCrLf & errN & " - " & errD, _
           vbExclamation, "OpenCap Costs"
Fin:
    ' Hide Costs tab + ensure Data toolbar Costs button (CSV sync still owns data).
    On Error Resume Next
    EnsureCostsUi
    On Error GoTo 0
End Sub

Private Sub WipeCostsDailyColumn(ByVal wsC As Worksheet)
    Dim r As Long
    Dim wasProt As Boolean
    If wsC Is Nothing Then Exit Sub
    wasProt = SheetUnprotectForVba(wsC)
    On Error Resume Next
    For r = COSTS_FIRST_ROW To COSTS_LAST_ROW
        wsC.Cells(r, COSTS_COL_DAILY).ClearContents
    Next r
    SheetReprotectAfterVba wsC, wasProt
    On Error GoTo 0
End Sub

Private Sub SyncCostsSheet()
    If Not SheetExists(SH_COSTS_TAB) Then Exit Sub

    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Worksheets(SH_COSTS_TAB)
    Dim tmpL As Long

    On Error Resume Next
    wsC.Calculate
    On Error GoTo 0

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' 1) Prior Costs entries (date -> daily cost). Skip on a new Job ID.
    Dim r As Long
    Dim seed As Double
    seed = 0
    If Not mNewJob Then
        If IsNumeric(wsC.Cells(COSTS_FIRST_ROW, COSTS_COL_DATE).Value2) Then
            seed = CDbl(wsC.Cells(COSTS_FIRST_ROW, COSTS_COL_DATE).Value2)
        End If

        For r = COSTS_FIRST_ROW To COSTS_LAST_ROW
            If Len(Trim$(CStr(wsC.Cells(r, COSTS_COL_DAILY).Value & ""))) = 0 Then GoTo NextPrior

            Dim dSerial As Long
            dSerial = ParseCostDateSerial(wsC.Cells(r, COSTS_COL_DATE).Value)
            If dSerial = 0 And seed > 0 Then
                dSerial = CLng(seed) + (r - COSTS_FIRST_ROW)
            End If
            If dSerial > 0 Then
                dict(CStr(dSerial)) = CDbl(wsC.Cells(r, COSTS_COL_DAILY).Value)
            End If
NextPrior:
        Next r
    End If

    ' 2) Overlay OpenCap ticket-costs CSV (CSV wins on same date)
    If SheetExists(SH_COSTS) Then
        Dim wsO As Worksheet
        Set wsO = ThisWorkbook.Worksheets(SH_COSTS)
        Dim lastO As Long
        lastO = wsO.Cells(wsO.Rows.Count, 1).End(xlUp).Row
        Dim oc As Long
        For oc = 2 To lastO
            Dim ocDate As Long
            Dim ocCost As Double
            ocDate = ParseCostDateSerial(wsO.Cells(oc, 2).Value)
            If ocDate = 0 Then ocDate = ParseCostDateSerial(wsO.Cells(oc, 2).Value2)
            If ocDate > 0 And IsNumeric(wsO.Cells(oc, 3).Value) Then
                ocCost = CDbl(wsO.Cells(oc, 3).Value)
                dict(CStr(ocDate)) = ocCost
            End If
        Next oc
    End If

    If dict.Count = 0 Then
        If mNewJob Then
            WipeCostsDailyColumn wsC
            Exit Sub
        End If
        Err.Raise vbObjectError + 602, , _
            "No cost rows found in Costs sheet or ticket-costs CSV (_OC_Costs)."
    End If
    If dict.Count > (COSTS_LAST_ROW - COSTS_FIRST_ROW + 1) Then
        Err.Raise vbObjectError + 601, , "Too many cost days for Costs sheet rows " & _
            COSTS_FIRST_ROW & "-" & COSTS_LAST_ROW
    End If

    ' 3) Sort date keys ascending (Variant, NOT Variant() � dict.Keys assignment fails otherwise)
    Dim keys As Variant
    Dim k As Variant
    Dim keyList() As Long
    ReDim keyList(0 To dict.Count - 1)
    Dim ki As Long
    ki = 0
    For Each k In dict.keys
        keyList(ki) = CLng(k)
        ki = ki + 1
    Next k

    Dim i As Long, j As Long
    For i = LBound(keyList) To UBound(keyList) - 1
        For j = i + 1 To UBound(keyList)
            If keyList(j) < keyList(i) Then
                tmpL = keyList(i)
                keyList(i) = keyList(j)
                keyList(j) = tmpL
            End If
        Next j
    Next i

    Dim consecutive As Boolean
    consecutive = True
    For i = LBound(keyList) + 1 To UBound(keyList)
        If keyList(i) <> keyList(i - 1) + 1 Then
            consecutive = False
            Exit For
        End If
    Next i

    ' 4) Only NOW clear + write (never clear before keys are ready)
    Dim prevCalc As XlCalculation
    Dim costsWasProt As Boolean
    Dim synErrNum As Long
    Dim synErrSrc As String
    Dim synErrDesc As String
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    costsWasProt = SheetUnprotectForVba(wsC)

    On Error GoTo SyncWriteFail

    For r = COSTS_FIRST_ROW To COSTS_LAST_ROW
        wsC.Cells(r, COSTS_COL_DAILY).ClearContents
    Next r

    For i = LBound(keyList) To UBound(keyList)
        r = COSTS_FIRST_ROW + (i - LBound(keyList))
        wsC.Cells(r, COSTS_COL_DAILY).numberFormat = "$#,##0.00"
        wsC.Cells(r, COSTS_COL_DAILY).Value = CDbl(dict(CStr(keyList(i))))

        ' Date column is merged C:D � write via MergeArea
        With wsC.Cells(r, COSTS_COL_DATE).MergeArea.Cells(1, 1)
            .numberFormat = "[$-F800]dddd, mmmm dd, yyyy"
            If i = LBound(keyList) Or (Not consecutive) Then
                .Value = keyList(i)
            Else
                .Formula = "=IF(E" & r & "<>"""",C" & (r - 1) & "+1,"""")"
            End If
        End With

        wsC.Cells(r, COSTS_COL_TOTAL).numberFormat = "$#,##0.00"
        If r = COSTS_FIRST_ROW Then
            wsC.Cells(r, COSTS_COL_TOTAL).Formula = _
                "=IF(ISBLANK(E" & r & "),"""",E" & r & ")"
        Else
            wsC.Cells(r, COSTS_COL_TOTAL).Formula = _
                "=IF(ISBLANK(E" & r & "),"""",SUM($E$" & COSTS_FIRST_ROW & ":E" & r & "))"
        End If
    Next i

    ' Restore date/total formulas on unused trailing day rows.
    Dim firstEmpty As Long
    firstEmpty = COSTS_FIRST_ROW + dict.Count
    For r = firstEmpty To COSTS_LAST_ROW
        If Len(Trim$(CStr(wsC.Cells(r, 2).Value & ""))) = 0 Then Exit For
        With wsC.Cells(r, COSTS_COL_DATE).MergeArea.Cells(1, 1)
            .numberFormat = "[$-F800]dddd, mmmm dd, yyyy"
            .Formula = "=IF(E" & r & "<>"""",C" & (r - 1) & "+1,"""")"
        End With
        wsC.Cells(r, COSTS_COL_TOTAL).numberFormat = "$#,##0.00"
        wsC.Cells(r, COSTS_COL_TOTAL).Formula = _
            "=IF(ISBLANK(E" & r & "),"""",SUM($E$" & COSTS_FIRST_ROW & ":E" & r & "))"
    Next r

    Application.Calculation = prevCalc
    On Error Resume Next
    wsC.Calculate
    On Error GoTo 0
    SheetReprotectAfterVba wsC, costsWasProt
    Exit Sub

SyncWriteFail:
    synErrNum = Err.Number
    synErrSrc = Err.Source
    synErrDesc = Err.Description
    Application.Calculation = prevCalc
    SheetReprotectAfterVba wsC, costsWasProt
    Err.Raise synErrNum, synErrSrc, synErrDesc
End Sub
Private Function ParseCostDateSerial(ByVal v As Variant) As Long
    ParseCostDateSerial = 0
    On Error GoTo Fail

    If isError(v) Then Exit Function
    If IsDate(v) Then
        ParseCostDateSerial = CLng(CDate(v))
        Exit Function
    End If
    If IsNumeric(v) Then
        Dim n As Double
        n = CDbl(v)
        ' Excel serial dates for this workbook era are roughly 40000+.
        If n >= 30000 And n < 100000 Then
            ParseCostDateSerial = CLng(n)
            Exit Function
        End If
    End If

    Dim s As String
    s = Trim$(CStr(v & ""))
    If Len(s) = 0 Then Exit Function

    ' YYYY-MM-DD or YYYY/MM/DD
    If Len(s) >= 10 Then
        Dim y As String, mo As String, d As String
        y = Left$(s, 4)
        mo = mid$(s, 6, 2)
        d = mid$(s, 9, 2)
        If IsNumeric(y) And IsNumeric(mo) And IsNumeric(d) Then
            ParseCostDateSerial = CLng(dateSerial(CInt(y), CInt(mo), CInt(d)))
            Exit Function
        End If
    End If

    If IsDate(s) Then ParseCostDateSerial = CLng(CDate(s))
Fail:
End Function

' ================================================================================
'  STRING HELPERS
' ================================================================================

' Appends a unit suffix if value is non-empty
Private Function ValUnit(val As String, unit As String) As String
    If Trim(val) = "" Then ValUnit = "" Else ValUnit = Trim(val) & unit
End Function

' Converts YYYYMMDDHHMI integer to a readable date string
Private Function FmtDate(raw As String) As String
    FmtDate = raw
    If Len(raw) < 8 Then Exit Function
    ' Handle "YYYYMMDD0000" or "YYYYMMDDHHmm" format
    On Error Resume Next
    Dim y As String: y = Left(raw, 4)
    Dim mo As String: mo = mid(raw, 5, 2)
    Dim d As String: d = mid(raw, 7, 2)
    If Not IsNumeric(y) Or Not IsNumeric(mo) Or Not IsNumeric(d) Then Exit Function
    FmtDate = y & "-" & mo & "-" & d
    On Error GoTo 0
End Function

' ================================================================================
'  UTILITY
' ================================================================================

Private Function SheetExists(nm As String) As Boolean
    On Error Resume Next: SheetExists = Not (ThisWorkbook.sheets(nm) Is Nothing): On Error GoTo 0
End Function

Private Function SafeStr(cell As Range) As String
    On Error Resume Next: SafeStr = Trim(CStr(cell.Value)): On Error GoTo 0
End Function

Private Function ColByName(ws As Worksheet, ParamArray names() As Variant) As Long
    ColByName = 0
    Dim last As Long: last = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim n As Variant, c As Long
    For Each n In names
        For c = 1 To last
            If LCase(Trim(SafeStr(ws.Cells(1, c)))) = LCase(Trim(CStr(n))) Then
                ColByName = c: Exit Function
            End If
        Next c
    Next n
End Function

' ================================================================================
'  SURVEY PLAN CSV IMPORT
'  Finds the "MD" header row in any well-plan CSV, imports survey stations
'  (MD, INC, AZI, Sub-Sea, TVD, NS, EW, VS, DLS, UTM(N), UTM(E), Lat, Long, Comment)
'  into hidden sheet _OC_Survey.  Row 1 = source path; Row 2 = headers; Row 3+ = data.
' ================================================================================

Public Sub ImportSurveyPlan()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select Well Plan PDF (Plan Sections) or CSV"
    fd.Filters.Clear
    fd.Filters.Add "Well Plan PDF", "*.pdf", 1
    fd.Filters.Add "CSV Files", "*.csv", 2
    fd.Filters.Add "All Files", "*.*", 3
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub

    ImportSurveyPlanFile fd.SelectedItems(1)
End Sub

' Testable core (Application.Run-able without the file dialog).
Public Sub ImportSurveyPlanFile(ByVal fPath As String)
    Application.StatusBar = "Importing survey plan..."
    On Error GoTo ImportPlanErr

    Dim ext As String
    ext = LCase$(mid$(fPath, InStrRev(fPath, ".") + 1))

    If ext = "pdf" Then
        ImportSurveyPlanPdf fPath
    Else
        ImportSurveyPlanCsv fPath
        BuildPlanSectionsFromSurveyComments
    End If

    UpdateImportPathDisplay SH_SURVEY

    On Error Resume Next
    MDL_SlidesheetClear.SyncPlanTargetWindow
    MDL_PlanGauge.RenderPlanGauge
    MDL_TDCalc.RefreshTdPlannedFromPlan
    On Error GoTo ImportPlanErr

    Application.StatusBar = "Plan imported from " & _
                             mid$(fPath, InStrRev(fPath, Application.PathSeparator) + 1)
    Exit Sub
ImportPlanErr:
    Application.StatusBar = "Survey import failed: " & Err.Description
End Sub

Private Function EnsureHiddenSheet(ByVal shName As String) As Worksheet
    If SheetExists(shName) Then
        Set EnsureHiddenSheet = Worksheets(shName)
        EnsureHiddenSheet.Cells.Clear
    Else
        Set EnsureHiddenSheet = ThisWorkbook.sheets.Add( _
            After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count))
        EnsureHiddenSheet.name = shName
        EnsureHiddenSheet.Visible = xlSheetVeryHidden
    End If
End Function

Private Sub ImportSurveyPlanCsv(ByVal fPath As String)
    Dim survWs As Worksheet
    Set survWs = EnsureHiddenSheet(SH_SURVEY)
    survWs.Cells(1, 1).Value = fPath

    Dim fNum As Integer: fNum = FreeFile
    Open fPath For Input As #fNum

    Dim lineText As String, outRow As Long
    Dim headerFound As Boolean: headerFound = False
    Dim colCount As Long: colCount = 0
    outRow = 3

    Do While Not EOF(fNum)
        Line Input #fNum, lineText
        lineText = Trim(lineText)
        If lineText = "" Then GoTo NextSurveyLine

        Dim fields() As String: fields = CsvParseLine(lineText)

        If Not headerFound Then
            Dim fi As Long
            For fi = 0 To UBound(fields)
                If LCase(Trim(fields(fi))) = "md" Then
                    headerFound = True
                    colCount = UBound(fields) + 1
                    Dim hc As Long
                    For hc = 0 To UBound(fields)
                        survWs.Cells(2, hc + 1).Value = Trim(fields(hc))
                    Next hc
                    Exit For
                End If
            Next fi
        Else
            If colCount = 0 Then GoTo NextSurveyLine
            If Not IsNumeric(Trim(fields(0))) Then GoTo NextSurveyLine
            Dim dc As Long
            For dc = 0 To UBound(fields)
                If dc < colCount Then
                    Dim cv As String: cv = Trim(fields(dc))
                    If IsNumeric(cv) Then
                        survWs.Cells(outRow, dc + 1).Value = CDbl(cv)
                    Else
                        survWs.Cells(outRow, dc + 1).Value = cv
                    End If
                End If
            Next dc
            outRow = outRow + 1
        End If
NextSurveyLine:
    Loop
    Close #fNum
End Sub

Private Sub ImportSurveyPlanPdf(ByVal fPath As String)
    Dim pdfText As String
    pdfText = ExtractPdfText(fPath)
    If pdfText = "" Then
        Application.StatusBar = "PDF extraction unavailable. Install Poppler or Adobe Acrobat."
        MsgBox "Could not read the well-plan PDF automatically." & Chr(10) & Chr(10) & _
               "For best results, re-run the import and choose Yes when " & _
               "offered the automatic Poppler download (one-time, ~16 MB)." & Chr(10) & Chr(10) & _
               "Manual alternatives:" & Chr(10) & _
               "  - Poppler for Windows (adds pdftotext.exe to PATH)" & Chr(10) & _
               "  - Adobe Acrobat (full version, not Reader)", _
               vbExclamation, "PDF Reader Not Available"
        Exit Sub
    End If

    Dim nSec As Long
    Dim aMd() As Double, aInc() As Double, aAzm() As Double
    Dim aTvd() As Double, ans() As Double, aEW() As Double
    Dim aDls() As Double, aBld() As Double, aTrn() As Double
    Dim aAnn() As String
    nSec = ParsePlanSections(pdfText, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn)
    If nSec < 1 Then
        MsgBox "No Plan Sections / SECTION DETAILS / Plan Annotations table found in that PDF.", _
               vbExclamation, "Import Plan"
        Exit Sub
    End If

    Dim aAuto() As String
    ReDim aAuto(0 To nSec - 1)
    NamePlanSectionTargets nSec, aMd, aInc, aAzm, aDls, aBld, aTrn, aAnn, aAuto
    WritePlanSecSheet fPath, nSec, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn, aAuto

    Dim sibCsv As String
    sibCsv = Left$(fPath, InStrRev(fPath, ".") - 1) & ".csv"
    If Dir(sibCsv) <> "" Then
        ImportSurveyPlanCsv sibCsv
        Worksheets(SH_SURVEY).Cells(1, 1).Value = fPath
    Else
        WriteSurveyFromSections fPath, nSec, aMd, aInc, aAzm, aTvd, ans, aEW
    End If
End Sub

' COMPASS Planning Report, in order:
'   1) SECTION DETAILS packed table (older / full reports)
'   2) Plan Sections visual table (newer extract: Y-grouped or layout rows)
'   3) Plan Annotations (MD + comment only)
' Overlay annotation comments onto (2) by matching MD.
Private Function ParsePlanSections(ByVal pdfText As String, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String) As Long

    ReDim aMd(0 To 80): ReDim aInc(0 To 80): ReDim aAzm(0 To 80)
    ReDim aTvd(0 To 80): ReDim ans(0 To 80): ReDim aEW(0 To 80)
    ReDim aDls(0 To 80): ReDim aBld(0 To 80): ReDim aTrn(0 To 80)
    ReDim aAnn(0 To 80)

    Dim n As Long
    n = ParseSectionDetailsBlock(pdfText, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn)
    If n < 2 Then
        n = ParsePlanSectionsTable(pdfText, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn)
    End If
    If n < 2 Then
        n = ParsePlanAnnotationsBlock(pdfText, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn)
    End If
    If n >= 2 Then OverlayPlanAnnotations pdfText, n, aMd, aAnn
    ParsePlanSections = n
End Function

Private Function ParseSectionDetailsBlock(ByVal pdfText As String, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String) As Long

    Dim startPos As Long
    startPos = InStr(1, pdfText, "SECTION DETAILS", vbTextCompare)
    If startPos = 0 Then Exit Function

    Dim endPos As Long: endPos = Len(pdfText) + 1
    Dim marker As Variant
    For Each marker In Array("DESIGN TARGET", "LOCAL COORDINATES", "Planning Report")
        Dim p As Long: p = InStr(startPos + 20, pdfText, CStr(marker), vbTextCompare)
        If p > 0 And p < endPos Then endPos = p
    Next marker

    Dim region As String
    region = mid$(pdfText, startPos, endPos - startPos)
    region = Replace(region, "\", "")
    region = Replace(region, ",", "")
    ' Extractor injects a space after a decimal: "0. 000"
    Do While InStr(region, ". ") > 0
        region = Replace(region, ". ", ".")
    Loop

    ' 9 packed numbers only (VBScript has no lookahead). Annotation is the
    ' text between this pack and the next.
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(-?\d{1,5}\.\d{2})(-?\d{1,3}\.\d{2})(-?\d{1,3}\.\d{2})" & _
                 "(-?\d{1,5}\.\d{2})(-?\d{1,5}\.\d{2})(-?\d{1,5}\.\d{2})" & _
                 "(-?\d{1,2}\.\d{3})(-?\d{1,3}\.\d{2})(-?\d{1,5}\.\d{2})"

    Dim ms As Object: Set ms = re.Execute(region)
    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To ms.Count - 1
        If n > 80 Then Exit For
        Dim m As Object: Set m = ms(i)
        aMd(n) = CDbl(m.SubMatches(0))
        aInc(n) = CDbl(m.SubMatches(1))
        aAzm(n) = CDbl(m.SubMatches(2))
        aTvd(n) = CDbl(m.SubMatches(3))
        ans(n) = CDbl(m.SubMatches(4))
        aEW(n) = CDbl(m.SubMatches(5))
        aDls(n) = CDbl(m.SubMatches(6))
        aBld(n) = 0#
        aTrn(n) = 0#
        Dim annStart As Long, annLen As Long
        annStart = m.FirstIndex + m.Length + 1
        If i < ms.Count - 1 Then
            annLen = ms(i + 1).FirstIndex - m.FirstIndex - m.Length
        Else
            annLen = Len(region) - m.FirstIndex - m.Length
        End If
        If annLen > 0 Then
            aAnn(n) = CleanSectionAnnot(mid$(region, annStart, annLen))
        Else
            aAnn(n) = ""
        End If
        n = n + 1
    Next i
    ParseSectionDetailsBlock = n
    If n < 2 Then
        n = ParseSectionDetailsSpaced(region, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn)
        ParseSectionDetailsBlock = n
    End If
End Function

' Spaced SECTION DETAILS (some extracts put a gap between the 9 packed fields).
Private Function ParseSectionDetailsSpaced(ByVal region As String, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String) As Long

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(-?\d{1,5}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+" & _
                 "(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+" & _
                 "(-?\d{1,2}\.\d{3})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,5}\.\d{2})"

    Dim ms As Object: Set ms = re.Execute(region)
    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To ms.Count - 1
        If n > 80 Then Exit For
        Dim m As Object: Set m = ms(i)
        If Not ValidPlanStation(CDbl(m.SubMatches(0)), CDbl(m.SubMatches(1)), CDbl(m.SubMatches(2))) Then GoTo NextSpaced
        If n > 0 And CDbl(m.SubMatches(0)) + 0.05 < aMd(n - 1) Then GoTo NextSpaced
        aMd(n) = CDbl(m.SubMatches(0))
        aInc(n) = CDbl(m.SubMatches(1))
        aAzm(n) = CDbl(m.SubMatches(2))
        aTvd(n) = CDbl(m.SubMatches(3))
        ans(n) = CDbl(m.SubMatches(4))
        aEW(n) = CDbl(m.SubMatches(5))
        aDls(n) = CDbl(m.SubMatches(6))
        aBld(n) = 0#
        aTrn(n) = 0#
        aAnn(n) = ""
        n = n + 1
NextSpaced:
    Next i
    ParseSectionDetailsSpaced = n
End Function

' Visual "Plan Sections" / "Plan Section" table. Built-in extractor groups glyphs
' by Y, so a row lands as: EW  NS  TVD  Azi  Inc  MD  [Target].
' Layout-preserved extracts (pdftotext -layout) are MD Inc Azi TVD NS EW DLS BR TR.
Private Function ParsePlanSectionsTable(ByVal pdfText As String, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String) As Long

    Dim startPos As Long
    startPos = InStr(1, pdfText, "Plan Sections", vbTextCompare)
    If startPos = 0 Then startPos = InStr(1, pdfText, "Plan Section", vbTextCompare)
    If startPos = 0 Then Exit Function

    Dim endPos As Long: endPos = Len(pdfText) + 1
    Dim marker As Variant
    For Each marker In Array("Planning Report", "Planned Survey", "SECTION DETAILS", _
                             "DESIGN TARGET", "Plan Annotations", "Plan  Annotations")
        Dim p As Long: p = InStr(startPos + 14, pdfText, CStr(marker), vbTextCompare)
        If p > 0 And p < endPos Then endPos = p
    Next marker

    Dim region As String
    region = mid$(pdfText, startPos, endPos - startPos)
    region = Replace(region, "\", "")
    region = Replace(region, ",", "")
    Do While InStr(region, ". ") > 0
        region = Replace(region, ". ", ".")
    Loop

    Dim n As Long
    n = ParsePlanSectionsYGrouped(region, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn)
    If n < 2 Then
        n = ParsePlanSectionsLayout(region, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn)
    End If
    ParsePlanSectionsTable = n
End Function

Private Function ParsePlanSectionsYGrouped(ByVal region As String, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String) As Long

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+" & _
                 "(-?\d{1,3}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,5}\.\d{2})" & _
                 "(?:\s+([A-Za-z][^0-9]{0,40}?))?(?=\s+-?\d{1,5}\.\d{2}|\s*$)"

    Dim ms As Object: Set ms = re.Execute(region)
    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To ms.Count - 1
        If n > 80 Then Exit For
        Dim m As Object: Set m = ms(i)
        Dim md As Double, inc As Double, azi As Double
        md = CDbl(m.SubMatches(5))
        inc = CDbl(m.SubMatches(4))
        azi = CDbl(m.SubMatches(3))
        If Not ValidPlanStation(md, inc, azi) Then GoTo NextY
        If n > 0 And md + 0.05 < aMd(n - 1) Then GoTo NextY
        If n > 0 And Abs(md - aMd(n - 1)) < 0.05 Then GoTo NextY
        aEW(n) = CDbl(m.SubMatches(0))
        ans(n) = CDbl(m.SubMatches(1))
        aTvd(n) = CDbl(m.SubMatches(2))
        aAzm(n) = azi
        aInc(n) = inc
        aMd(n) = md
        aDls(n) = 0#: aBld(n) = 0#: aTrn(n) = 0#
        aAnn(n) = CleanSectionAnnot(CStr(m.SubMatches(6)))
        n = n + 1
NextY:
    Next i
    ParsePlanSectionsYGrouped = n
End Function

Private Function ParsePlanSectionsLayout(ByVal region As String, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String) As Long

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(-?\d{1,5}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+(-?\d{1,3}\.\d{2})\s+" & _
                 "(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+(-?\d{1,5}\.\d{2})\s+" & _
                 "(-?\d{1,2}\.\d{3})\s+(-?\d{1,3}\.\d{3})\s+(-?\d{1,3}\.\d{3})" & _
                 "(?:\s+(-?\d{1,3}\.\d{2}))?(?:\s+([A-Za-z][^0-9]{0,40}?))?"

    Dim ms As Object: Set ms = re.Execute(region)
    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To ms.Count - 1
        If n > 80 Then Exit For
        Dim m As Object: Set m = ms(i)
        Dim md As Double, inc As Double, azi As Double
        md = CDbl(m.SubMatches(0))
        inc = CDbl(m.SubMatches(1))
        azi = CDbl(m.SubMatches(2))
        If Not ValidPlanStation(md, inc, azi) Then GoTo NextLay
        If n > 0 And md + 0.05 < aMd(n - 1) Then GoTo NextLay
        If n > 0 And Abs(md - aMd(n - 1)) < 0.05 Then GoTo NextLay
        aMd(n) = md
        aInc(n) = inc
        aAzm(n) = azi
        aTvd(n) = CDbl(m.SubMatches(3))
        ans(n) = CDbl(m.SubMatches(4))
        aEW(n) = CDbl(m.SubMatches(5))
        aDls(n) = CDbl(m.SubMatches(6))
        aBld(n) = CDbl(m.SubMatches(7))
        aTrn(n) = CDbl(m.SubMatches(8))
        aAnn(n) = CleanSectionAnnot(CStr(m.SubMatches(10)))
        n = n + 1
NextLay:
    Next i
    ParsePlanSectionsLayout = n
End Function

Private Function ValidPlanStation(ByVal md As Double, ByVal inc As Double, ByVal azi As Double) As Boolean
    ValidPlanStation = (md >= 0# And md < 20000# And inc >= 0# And inc <= 180# And azi >= 0# And azi <= 360#)
End Function

Private Sub OverlayPlanAnnotations(ByVal pdfText As String, ByVal n As Long, _
        ByRef aMd() As Double, ByRef aAnn() As String)

    Dim tMd() As Double, tInc() As Double, tAzm() As Double
    Dim tTvd() As Double, tNS() As Double, tEW() As Double
    Dim tDls() As Double, tBld() As Double, tTrn() As Double
    Dim tAnn() As String
    ReDim tMd(0 To 80): ReDim tInc(0 To 80): ReDim tAzm(0 To 80)
    ReDim tTvd(0 To 80): ReDim tNS(0 To 80): ReDim tEW(0 To 80)
    ReDim tDls(0 To 80): ReDim tBld(0 To 80): ReDim tTrn(0 To 80)
    ReDim tAnn(0 To 80)

    Dim tN As Long
    tN = ParsePlanAnnotationsBlock(pdfText, tMd, tInc, tAzm, tTvd, tNS, tEW, tDls, tBld, tTrn, tAnn)
    If tN < 1 Then Exit Sub

    Dim i As Long, j As Long
    For i = 0 To n - 1
        If Len(Trim$(aAnn(i))) > 0 Then GoTo NextOverlay
        For j = 0 To tN - 1
            If Abs(aMd(i) - tMd(j)) < 0.25 Then
                aAnn(i) = tAnn(j)
                Exit For
            End If
        Next j
NextOverlay:
    Next i
End Sub

Private Function CleanSectionAnnot(ByVal s As String) As String
    s = Trim$(Replace(Replace(s, Chr(10), " "), Chr(13), " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Dim eqPos As Long: eqPos = InStrRev(s, "=")
    If eqPos > 0 Then
        Dim mPos As Long: mPos = InStr(eqPos, s, "m")
        If mPos > 0 Then s = Left$(s, mPos)
    ElseIf Len(s) > 60 Then
        s = Left$(s, 60)
    End If
    CleanSectionAnnot = Trim$(s)
End Function

Private Function ParsePlanAnnotationsBlock(ByVal pdfText As String, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String) As Long

    Dim startPos As Long
    startPos = InStr(1, pdfText, "Plan  Annotations", vbTextCompare)
    If startPos = 0 Then startPos = InStr(1, pdfText, "Plan Annotations", vbTextCompare)
    If startPos = 0 Then Exit Function

    Dim endPos As Long: endPos = Len(pdfText) + 1
    Dim p As Long: p = InStr(startPos + 20, pdfText, "Planning Report", vbTextCompare)
    If p > 0 Then endPos = p
    Dim region As String: region = mid$(pdfText, startPos, endPos - startPos)
    region = Replace(region, "\", "")

    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(\d{1,3}(?:,\d{3})*\.\d{2})\s+(\d{1,3}(?:,\d{3})*\.\d{2})\s+" & _
                 "(-?\d{1,3}(?:,\d{3})*\.\d{2})\s+(-?\d{1,3}(?:,\d{3})*\.\d{2})\s+" & _
                 "([A-Za-z][^0-9]{0,80}?)" & _
                 "(?=\s+\d{1,3}(?:,\d{3})*\.\d{2}\s+\d{1,3}(?:,\d{3})*\.\d{2}|$)"

    Dim ms As Object: Set ms = re.Execute(region)
    Dim n As Long: n = 0
    Dim m As Object
    For Each m In ms
        If n > 80 Then Exit For
        aMd(n) = CleanNum(CStr(m.SubMatches(0)))
        aTvd(n) = CleanNum(CStr(m.SubMatches(1)))
        aEW(n) = CleanNum(CStr(m.SubMatches(2)))
        ans(n) = CleanNum(CStr(m.SubMatches(3)))
        aInc(n) = 0#: aAzm(n) = 0#: aDls(n) = 0#: aBld(n) = 0#: aTrn(n) = 0#
        aAnn(n) = Trim$(CStr(m.SubMatches(4)))
        n = n + 1
    Next m
    ParsePlanAnnotationsBlock = n
End Function

Private Sub NamePlanSectionTargets(ByVal n As Long, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String, ByRef aAuto() As String)

    Dim i As Long
    For i = 0 To n - 1
        aAuto(i) = SeedNameFromText(aAnn(i))
    Next i

    Dim kopI As Long: kopI = -1
    Dim heelI As Long: heelI = -1
    Dim sotI As Long: sotI = -1
    Dim eotI As Long: eotI = -1
    Dim tangI As Long: tangI = -1

    For i = 0 To n - 1
        If aAuto(i) = "KOP" And kopI < 0 Then kopI = i
        If aAuto(i) = "HEEL" And heelI < 0 Then heelI = i
        If aAuto(i) = "SOT" And sotI < 0 Then sotI = i
        If aAuto(i) = "EOT" Then eotI = i
        If aAuto(i) = "TANGENT" And tangI < 0 Then tangI = i
    Next i

    If kopI < 0 Then
        For i = 1 To n - 1
            If aInc(i) > 1# And aInc(i - 1) < 1# And aMd(i) > 100# Then
                If InStr(1, aAnn(i), "Nudge", vbTextCompare) = 0 Then
                    kopI = i: aAuto(i) = "KOP": Exit For
                End If
            End If
            If InStr(1, aAnn(i), "KOP", vbTextCompare) > 0 Then
                kopI = i: aAuto(i) = "KOP": Exit For
            End If
        Next i
    End If

    If heelI < 0 Then
        For i = 0 To n - 1
            If aInc(i) >= 88# Then
                heelI = i: aAuto(i) = "HEEL": Exit For
            End If
        Next i
    End If

    If n > 0 Then
        If aAuto(n - 1) = "" Then aAuto(n - 1) = "TD"
        If InStr(1, aAnn(n - 1), "Toe", vbTextCompare) > 0 Then aAuto(n - 1) = "TD"
    End If

    ' SOT: first significant turn after KOP (or "B & T" / build+turn).
    If sotI < 0 Then
        Dim startS As Long: startS = 0
        If kopI >= 0 Then startS = kopI + 1
        For i = startS To n - 1
            If aAuto(i) = "HEEL" Or aAuto(i) = "TD" Then Exit For
            If Abs(aTrn(i)) > 0.5 Or InStr(1, aAnn(i), "B & T", vbTextCompare) > 0 _
                    Or InStr(1, aAnn(i), "B&T", vbTextCompare) > 0 Then
                sotI = i: aAuto(i) = "SOT": Exit For
            End If
        Next i
    End If

    ' EOT: last hold / turn-off before heel after SOT.
    If eotI < 0 And sotI >= 0 Then
        Dim lastTurn As Long: lastTurn = sotI
        Dim lim As Long: lim = n - 1
        If heelI > sotI Then lim = heelI - 1
        For i = sotI + 1 To lim
            If Abs(aTrn(i)) > 0.5 Or aDls(i) > 0.3 Then lastTurn = i
            If aAuto(i) = "TANGENT" And i > sotI Then
                eotI = i
            End If
        Next i
        If eotI < 0 Then
            For i = lastTurn To lim
                If InStr(1, aAnn(i), "Hold", vbTextCompare) > 0 Then
                    eotI = i: Exit For
                End If
            Next i
        End If
        If eotI < 0 And lastTurn > sotI Then eotI = lastTurn
        If eotI >= 0 And aAuto(eotI) = "" Then aAuto(eotI) = "EOT"
    End If

    ' NUDGE arrival: first 3–20° station after a Nudge annotation (5° / 10° hold).
    ' Only walk from the annotated Nudge row so later holds at the same inc stay unnamed.
    For i = 0 To n - 2
        If InStr(1, aAnn(i), "Nudge", vbTextCompare) = 0 Then GoTo NextNudgeSrc
        Dim jN As Long
        For jN = i + 1 To n - 1
            If aAuto(jN) = "KOP" Or aAuto(jN) = "VERTICAL" Or aAuto(jN) = "TD" Then Exit For
            If aInc(jN) >= 3# And aInc(jN) < 20# Then
                If aAuto(jN) = "" Then aAuto(jN) = "NUDGE"
                Exit For
            End If
        Next jN
NextNudgeSrc:
    Next i

    ' TANGENT: first hold after KOP with meaningful inclination, before SOT/HEEL.
    If tangI < 0 And kopI >= 0 Then
        Dim tangLim As Long: tangLim = n - 1
        If sotI > kopI Then tangLim = sotI - 1
        If heelI > kopI And heelI < tangLim Then tangLim = heelI - 1
        For i = kopI + 1 To tangLim
            If aAuto(i) <> "" Then GoTo NextTang
            If aInc(i) > 15# And aInc(i) < 85# Then
                If aDls(i) < 0.3 Or InStr(1, aAnn(i), "Hold", vbTextCompare) > 0 Then
                    tangI = i: aAuto(i) = "TANGENT": Exit For
                End If
            End If
NextTang:
        Next i
    End If
End Sub

Private Function SeedNameFromText(ByVal t As String) As String
    Dim s As String: s = LCase$(Trim$(t))
    SeedNameFromText = ""
    If s = "" Then Exit Function
    If InStr(s, "toe") > 0 Or InStr(s, "td") = 1 Or InStr(s, " td") > 0 _
            Or Left$(s, 3) = "td " Or s = "td" Then
        SeedNameFromText = "TD": Exit Function
    End If
    If InStr(s, "heel") > 0 Then SeedNameFromText = "HEEL": Exit Function
    If InStr(s, "kop") > 0 Or InStr(s, "kick") > 0 Then SeedNameFromText = "KOP": Exit Function
    If InStr(s, "sot") > 0 Or InStr(s, "start of turn") > 0 _
            Or InStr(s, "b & t") > 0 Or InStr(s, "b&t") > 0 Then
        SeedNameFromText = "SOT": Exit Function
    End If
    If InStr(s, "eot") > 0 Or InStr(s, "end of turn") > 0 Then
        SeedNameFromText = "EOT": Exit Function
    End If
    If InStr(s, "tang") > 0 Then SeedNameFromText = "TANGENT": Exit Function
    If InStr(s, "nudge") > 0 Then SeedNameFromText = "NUDGE": Exit Function
    If InStr(s, "back to vert") > 0 Or InStr(s, "back to vertical") > 0 _
            Or InStr(s, "btv") > 0 Then
        SeedNameFromText = "VERTICAL": Exit Function
    End If
End Function

Private Sub WritePlanSecSheet(ByVal fPath As String, ByVal n As Long, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double, _
        ByRef aDls() As Double, ByRef aBld() As Double, ByRef aTrn() As Double, _
        ByRef aAnn() As String, ByRef aAuto() As String)

    Dim ws As Worksheet
    Set ws = EnsureHiddenSheet(SH_PLANSEC)
    ws.Cells(1, 1).Value = fPath
    ws.Cells(2, 1).Value = "MD"
    ws.Cells(2, 2).Value = "INC"
    ws.Cells(2, 3).Value = "AZM"
    ws.Cells(2, 4).Value = "TVD"
    ws.Cells(2, 5).Value = "NS"
    ws.Cells(2, 6).Value = "EW"
    ws.Cells(2, 7).Value = "DLS"
    ws.Cells(2, 8).Value = "BUILD"
    ws.Cells(2, 9).Value = "TURN"
    ws.Cells(2, 10).Value = "ANNOT"
    ws.Cells(2, 11).Value = "AUTONAME"
    ws.Cells(2, 12).Value = "USERNAME"
    Dim i As Long
    For i = 0 To n - 1
        ws.Cells(i + 3, 1).Value = aMd(i)
        ws.Cells(i + 3, 2).Value = aInc(i)
        ws.Cells(i + 3, 3).Value = aAzm(i)
        ws.Cells(i + 3, 4).Value = aTvd(i)
        ws.Cells(i + 3, 5).Value = ans(i)
        ws.Cells(i + 3, 6).Value = aEW(i)
        ws.Cells(i + 3, 7).Value = aDls(i)
        ws.Cells(i + 3, 8).Value = aBld(i)
        ws.Cells(i + 3, 9).Value = aTrn(i)
        ws.Cells(i + 3, 10).Value = aAnn(i)
        ws.Cells(i + 3, 11).Value = aAuto(i)
    Next i
End Sub

Private Sub WriteSurveyFromSections(ByVal fPath As String, ByVal n As Long, _
        ByRef aMd() As Double, ByRef aInc() As Double, ByRef aAzm() As Double, _
        ByRef aTvd() As Double, ByRef ans() As Double, ByRef aEW() As Double)
    Dim ws As Worksheet
    Set ws = EnsureHiddenSheet(SH_SURVEY)
    ws.Cells(1, 1).Value = fPath
    ws.Cells(2, 1).Value = "MD"
    ws.Cells(2, 2).Value = "INC"
    ws.Cells(2, 3).Value = "AZI"
    ws.Cells(2, 4).Value = "TVD"
    ws.Cells(2, 5).Value = "NS"
    ws.Cells(2, 6).Value = "EW"
    Dim i As Long
    For i = 0 To n - 1
        ws.Cells(i + 3, 1).Value = aMd(i)
        ws.Cells(i + 3, 2).Value = aInc(i)
        ws.Cells(i + 3, 3).Value = aAzm(i)
        ws.Cells(i + 3, 4).Value = aTvd(i)
        ws.Cells(i + 3, 5).Value = ans(i)
        ws.Cells(i + 3, 6).Value = aEW(i)
    Next i
End Sub

' CSV fallback: promote unique Comment / annotation stations into _OC_PlanSec.
Private Sub BuildPlanSectionsFromSurveyComments()
    If Not SheetExists(SH_SURVEY) Then Exit Sub
    Dim surv As Worksheet: Set surv = Worksheets(SH_SURVEY)
    Dim cMD As Long, cInc As Long, cAzi As Long, cTvd As Long
    Dim cNS As Long, cEW As Long, cDls As Long, cCom As Long
    Dim c As Long
    For c = 1 To 20
        Select Case LCase$(Trim$(CStr(surv.Cells(2, c).Value2 & "")))
            Case "md": cMD = c
            Case "inc": cInc = c
            Case "azi", "azm": cAzi = c
            Case "tvd": cTvd = c
            Case "ns": cNS = c
            Case "ew": cEW = c
            Case "dls": cDls = c
            Case "comment", "annot", "annotation": cCom = c
        End Select
    Next c
    If cMD = 0 Then Exit Sub

    Dim lastR As Long: lastR = surv.Cells(surv.Rows.Count, cMD).End(xlUp).Row
    If lastR < 3 Then Exit Sub

    Dim aMd() As Double, aInc() As Double, aAzm() As Double
    Dim aTvd() As Double, ans() As Double, aEW() As Double
    Dim aDls() As Double, aBld() As Double, aTrn() As Double
    Dim aAnn() As String
    ReDim aMd(0 To 80): ReDim aInc(0 To 80): ReDim aAzm(0 To 80)
    ReDim aTvd(0 To 80): ReDim ans(0 To 80): ReDim aEW(0 To 80)
    ReDim aDls(0 To 80): ReDim aBld(0 To 80): ReDim aTrn(0 To 80)
    ReDim aAnn(0 To 80)

    Dim n As Long: n = 0
    Dim r As Long
    For r = 3 To lastR
        If Not IsNumeric(surv.Cells(r, cMD).Value2) Then GoTo NextCom
        Dim ann As String: ann = ""
        If cCom > 0 Then ann = Trim$(CStr(surv.Cells(r, cCom).Value2 & ""))
        Dim takeIt As Boolean
        takeIt = (r = lastR)
        If Len(ann) > 0 Then
            If SeedNameFromText(ann) <> "" Or InStr(1, ann, "KOP", vbTextCompare) > 0 _
                    Or InStr(1, ann, "Hold", vbTextCompare) > 0 _
                    Or InStr(1, ann, "B & T", vbTextCompare) > 0 _
                    Or InStr(1, ann, "Build", vbTextCompare) > 0 Then
                takeIt = True
            End If
        End If
        If Not takeIt Then GoTo NextCom
        If n > 80 Then Exit For
        aMd(n) = CDbl(surv.Cells(r, cMD).Value2)
        If cInc > 0 Then aInc(n) = val(surv.Cells(r, cInc).Value2 & "")
        If cAzi > 0 Then aAzm(n) = val(surv.Cells(r, cAzi).Value2 & "")
        If cTvd > 0 Then aTvd(n) = val(surv.Cells(r, cTvd).Value2 & "")
        If cNS > 0 Then ans(n) = val(surv.Cells(r, cNS).Value2 & "")
        If cEW > 0 Then aEW(n) = val(surv.Cells(r, cEW).Value2 & "")
        If cDls > 0 Then aDls(n) = val(surv.Cells(r, cDls).Value2 & "")
        aAnn(n) = ann
        n = n + 1
NextCom:
    Next r
    If n < 1 Then Exit Sub

    Dim aAuto() As String
    ReDim aAuto(0 To n - 1)
    NamePlanSectionTargets n, aMd, aInc, aAzm, aDls, aBld, aTrn, aAnn, aAuto
    WritePlanSecSheet CStr(surv.Cells(1, 1).Value), n, aMd, aInc, aAzm, aTvd, ans, aEW, aDls, aBld, aTrn, aAnn, aAuto
End Sub

' ================================================================================
'  ANTI-COLLISION PDF IMPORT
'  Extracts the Summary table from a well-plan AC PDF and finds all rows
'  where Separation Factor < 2.0.  Columns: Ref MD (m), Between Centres (m), SF.
'  Results go to hidden sheet _OC_AC;  a formatted table is written to Sheet1.
' ================================================================================

Public Sub ImportAntiCollision()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select Anti-Collision PDF"
    fd.Filters.Clear
    fd.Filters.Add "PDF Files", "*.pdf", 1
    fd.Filters.Add "All Files", "*.*", 2
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub

    ImportAntiCollisionFile fd.SelectedItems(1)
End Sub

' Faded status fills (Excel's standard good / neutral / bad pastels).
Private Function cAcRed() As Long:    cAcRed = RGB(255, 199, 206):    End Function
Private Function cAcYellow() As Long: cAcYellow = RGB(255, 235, 156): End Function
Private Function cAcGreen() As Long:  cAcGreen = RGB(198, 239, 206):  End Function

' Testable core (Application.Run-able without the file dialog).
' maxSF caps which summary rows are captured at all.  It defaults high so every
' offset well is read: the Data table is then sorted by drilling priority and
' colour-banded, so wells above the concern threshold still show (in green)
' instead of leaving the table half empty.
' quiet: suppress the informational "more wells than rows" prompt (batch runs).
Public Sub ImportAntiCollisionFile(ByVal fPath As String, _
        Optional ByVal maxSF As Double = 1000000#, Optional ByVal Quiet As Boolean = False)
    Application.StatusBar = "Reading anti-collision report..."

    Dim acWs As Worksheet
    If SheetExists(SH_AC) Then
        Set acWs = Worksheets(SH_AC)
        acWs.Cells.Clear
    Else
        Set acWs = ThisWorkbook.sheets.Add( _
            After:=ThisWorkbook.sheets(ThisWorkbook.sheets.Count))
        acWs.name = SH_AC
        acWs.Visible = xlSheetVeryHidden
    End If
    acWs.Cells(1, 1).Value = fPath

    Dim pdfText As String: pdfText = ExtractPdfText(fPath)

    If pdfText = "" Then
        Application.StatusBar = "PDF extraction unavailable. Install Poppler or Adobe Acrobat."
        MsgBox "Could not read the PDF automatically." & Chr(10) & Chr(10) & _
               "For best results, re-run the import and choose Yes when " & _
               "offered the automatic Poppler download (one-time, ~16 MB)." & Chr(10) & Chr(10) & _
               "Manual alternatives:" & Chr(10) & _
               "  - Poppler for Windows (adds pdftotext.exe to PATH)" & Chr(10) & _
               "    https://github.com/oschwartz10612/poppler-windows/releases" & Chr(10) & _
               "  - Adobe Acrobat (full version, not Reader)" & Chr(10) & Chr(10) & _
               "Path stored. Try again after installing.", vbExclamation, "PDF Reader Not Available"
        UpdateImportPathDisplay SH_AC
        Exit Sub
    End If

    ' Parse Summary table rows with SF <= maxSF
    Dim nHits As Long
    Dim aWell(200) As String
    Dim aRefMD(200) As Double, aBetween(200) As Double, aSF(200) As Double
    nHits = ParseAcSummary(pdfText, aWell, aRefMD, aBetween, aSF, maxSF)

    ' Severity order (worst Separation Factor first).  The Data table is a fixed
    ' 8 rows, so this decides WHICH wells earn a slot; WriteAcConcerns then puts
    ' the ones it keeps back into depth order for display.
    SortAcRows nHits, aWell, aRefMD, aBetween, aSF, False

    ' When nothing found, save extracted text for troubleshooting (no popup:
    ' zero concerns is the normal result for most AC reports).
    If nHits = 0 And Len(pdfText) > 0 Then
        On Error Resume Next
        Dim dbgNum As Integer: dbgNum = FreeFile
        Open Environ("TEMP") & "\oc_pdf_debug.txt" For Output As #dbgNum
        Print #dbgNum, pdfText
        Close #dbgNum
        On Error GoTo 0
    End If

    ' Write to hidden sheet
    acWs.Cells(2, 1).Value = "Offset Well - Wellbore - Design"
    acWs.Cells(2, 2).Value = "Ref MD (m)"
    acWs.Cells(2, 3).Value = "Between Centres (m)"
    acWs.Cells(2, 4).Value = "Separation Factor"
    Dim i As Long
    For i = 0 To nHits - 1
        acWs.Cells(i + 3, 1).Value = aWell(i)
        acWs.Cells(i + 3, 2).Value = aRefMD(i)
        acWs.Cells(i + 3, 3).Value = aBetween(i)
        acWs.Cells(i + 3, 4).Value = aSF(i)
    Next i

    ' Render formatted table on Setup sheet
    BuildAcTable nHits, aRefMD, aBetween, aSF

    ' Fill the "AC Info & Concerns" table on the Data sheet
    WriteAcConcerns nHits, aWell, aRefMD, aBetween, aSF, Quiet

    UpdateImportPathDisplay SH_AC

    Dim nConcern As Long
    For i = 0 To nHits - 1
        If aSF(i) < AC_SF_YELLOW Then nConcern = nConcern + 1
    Next i
    Application.StatusBar = "AC import complete: " & nHits & " offset well(s), " & _
                            nConcern & " with SF < " & Format(AC_SF_YELLOW, "0.0") & "."
End Sub

' Sort the parallel AC arrays in place.  Insertion sort - nHits is capped at 200.
'   byDepth = False -> severity: lowest Separation Factor first, closest
'                      centre-to-centre breaking ties.
'   byDepth = True  -> display:  shallowest depth of closest approach (the
'                      "Closest C2C" column) first, so reading down the table
'                      follows the order the wells are met as the hole deepens.
Private Sub SortAcRows(ByVal n As Long, aWell() As String, _
        aRefMD() As Double, aBetween() As Double, aSF() As Double, _
        ByVal byDepth As Boolean)
    Dim i As Long, j As Long
    Dim kWell As String, kMD As Double, kBC As Double, kSF As Double

    For i = 1 To n - 1
        kWell = aWell(i): kMD = aRefMD(i): kBC = aBetween(i): kSF = aSF(i)
        j = i - 1
        Do While j >= 0
            If AcKeepsPlace(aSF(j), aBetween(j), aRefMD(j), kSF, kBC, kMD, byDepth) Then Exit Do
            aWell(j + 1) = aWell(j)
            aRefMD(j + 1) = aRefMD(j)
            aBetween(j + 1) = aBetween(j)
            aSF(j + 1) = aSF(j)
            j = j - 1
        Loop
        aWell(j + 1) = kWell: aRefMD(j + 1) = kMD
        aBetween(j + 1) = kBC: aSF(j + 1) = kSF
    Next i
End Sub

' True when row A keeps its place ahead of row B under the requested ordering.
Private Function AcKeepsPlace( _
        ByVal sfA As Double, ByVal bcA As Double, ByVal mdA As Double, _
        ByVal sfB As Double, ByVal bcB As Double, ByVal mdB As Double, _
        ByVal byDepth As Boolean) As Boolean
    If byDepth Then
        If mdA < mdB Then AcKeepsPlace = True: Exit Function
        If mdA > mdB Then AcKeepsPlace = False: Exit Function
        AcKeepsPlace = (sfA <= sfB)
    Else
        If sfA < sfB Then AcKeepsPlace = True: Exit Function
        If sfA > sfB Then AcKeepsPlace = False: Exit Function
        AcKeepsPlace = (bcA <= bcB)
    End If
End Function

' The table's own look is alternating unfilled / light-grey rows.  Cleared rows
' are put back to that banding so they never keep a stale status colour.
Private Sub AcResetRowFill(ws As Worksheet, ByVal r As Long, ByVal firstRow As Long)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Interior
        If (r - firstRow) Mod 2 = 0 Then
            .Pattern = xlNone
        Else
            .Pattern = xlSolid
            .Color = RGB(242, 242, 242)
        End If
    End With
End Sub

' Faded status fill for a Separation Factor.
Private Function AcBandColor(ByVal sf As Double) As Long
    If sf < AC_SF_RED Then
        AcBandColor = cAcRed()
    ElseIf sf < AC_SF_YELLOW Then
        AcBandColor = cAcYellow()
    Else
        AcBandColor = cAcGreen()
    End If
End Function

' ---- Write concerns into the 'AC Info & Concerns' table on the Data sheet ----
' Layout (located dynamically): title cell "AC Info & Concerns" in column B,
' column headers on the next row, data rows below until the row above
' "Motors On Location".  Column order per ops requirement:
'   B:C merged = Offset Well - Wellbore - Design
'   D         = Separation Factor
'   E         = Between Centres / C2C (m)
'   F         = Reference Measured Depth (m)
' Each row is filled with its faded SF status colour (red/yellow/green).
Private Sub WriteAcConcerns(nHits As Long, aWell() As String, _
        aRefMD() As Double, aBetween() As Double, aSF() As Double, _
        Optional ByVal Quiet As Boolean = False)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim titleCell As Range
    Set titleCell = ws.Range("B1:B200").Find("AC Info & Concerns", LookAt:=xlPart, LookIn:=xlValues)
    If titleCell Is Nothing Then Exit Sub

    Dim firstRow As Long: firstRow = titleCell.Row + 2   ' skip title + header row

    ' Bottom of the table = row above "Motors On Location" (fallback: 8 rows)
    Dim lastRow As Long: lastRow = firstRow + 7
    Dim botCell As Range
    Set botCell = ws.Range(ws.Cells(firstRow, 2), ws.Cells(firstRow + 40, 2)) _
                    .Find("Motors On Location", LookAt:=xlPart, LookIn:=xlValues)
    If Not botCell Is Nothing Then
        If botCell.Row > firstRow Then lastRow = botCell.Row - 1
    End If

    Dim wasProt As Boolean
    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo ReprotectFail

    ' Clear previous contents and status fill (values only; keep borders)
    Dim r As Long
    For r = firstRow To lastRow
        ws.Cells(r, 2).MergeArea.ClearContents   ' B (usually merged B:C)
        ws.Cells(r, 4).ClearContents             ' D
        ws.Cells(r, 5).ClearContents             ' E
        ws.Cells(r, 6).ClearContents             ' F
        AcResetRowFill ws, r, firstRow
    Next r

    Dim capacity As Long: capacity = lastRow - firstRow + 1
    Dim i As Long

    ' The caller hands us the wells in severity order, so the first `capacity`
    ' entries are the ones worth a slot in this fixed-height table.  Take that
    ' slice and re-sort it by depth: the table then reads in the order we drill
    ' into the wells, while a deep low-SF well can never be crowded out by
    ' shallow harmless ones.
    Dim nShow As Long: nShow = nHits
    If nShow > capacity Then nShow = capacity

    If nShow > 0 Then
        Dim sWell() As String
        Dim sMD() As Double, sBC() As Double, sSF() As Double
        ReDim sWell(nShow): ReDim sMD(nShow): ReDim sBC(nShow): ReDim sSF(nShow)
        For i = 0 To nShow - 1
            sWell(i) = aWell(i): sMD(i) = aRefMD(i)
            sBC(i) = aBetween(i): sSF(i) = aSF(i)
        Next i
        SortAcRows nShow, sWell, sMD, sBC, sSF, True

        For i = 0 To nShow - 1
            r = firstRow + i
            ws.Cells(r, 2).MergeArea.Cells(1, 1).Value = sWell(i)
            ws.Cells(r, 4).Value = sSF(i):  ws.Cells(r, 4).numberFormat = "0.000"
            ws.Cells(r, 5).Value = sBC(i):  ws.Cells(r, 5).numberFormat = "0.00"
            ws.Cells(r, 6).Value = sMD(i):  ws.Cells(r, 6).numberFormat = "0.00"
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Interior.Color = AcBandColor(sSF(i))
        Next i
    End If

    SheetReprotectAfterVba ws, wasProt

    If nHits > capacity And Not Quiet Then
        MsgBox "AC import read " & nHits & " offset well(s); this table holds " & _
               capacity & "." & vbCrLf & vbCrLf & _
               "Kept the " & capacity & " with the lowest separation factors and " & _
               "listed them by depth." & vbCrLf & _
               "The full list is on the Setup sheet AC table.", _
               vbInformation, "AC Info & Concerns"
    End If
    Exit Sub
ReprotectFail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
End Sub

' ---- Poppler (pdftotext) resolve + optional self-install ----------------------
' Finds pdftotext on PATH or in the per-user self-installed copy under
' %LOCALAPPDATA%\Poppler (pointer file written by the bootstrap).
' Returns "pdftotext" (PATH), a full exe path, or "" when unavailable.
Private Function ResolvePdftotext(Sh As Object) As String
    ResolvePdftotext = ""

    On Error Resume Next
    If Sh.Run("cmd /c where pdftotext >nul 2>nul", 0, True) = 0 Then
        ResolvePdftotext = "pdftotext"
    End If
    On Error GoTo 0
    If ResolvePdftotext <> "" Then Exit Function

    Dim ptr As String: ptr = Environ("LOCALAPPDATA") & "\Poppler\pdftotext_path.txt"
    If Dir(ptr) = "" Then Exit Function

    Dim fNum As Integer, exePath As String
    fNum = FreeFile
    On Error Resume Next
    Open ptr For Input As #fNum
    Line Input #fNum, exePath
    Close #fNum
    On Error GoTo 0

    exePath = Trim$(exePath)
    If exePath <> "" Then
        If Dir(exePath) <> "" Then ResolvePdftotext = exePath
    End If
End Function

' Offers a one-time automatic Poppler download (~16 MB) into %LOCALAPPDATA%\Poppler.
' Per-user install, no admin rights required. Asks first (never downloads silently)
' and asks at most once per Excel session. Returns the pdftotext path on success,
' "" if declined or failed - callers fall back to the built-in PowerShell parser.
Private Function OfferPopplerBootstrap(Sh As Object) As String
    OfferPopplerBootstrap = ""
    If mPopplerPrompted Then Exit Function
    mPopplerPrompted = True

    Dim ans As VbMsgBoxResult
    ans = MsgBox("For the most reliable PDF table import, this workbook can " & _
                 "download the free Poppler PDF tool (pdftotext)." & Chr(10) & Chr(10) & _
                 "One-time download, about 16 MB. Installs to your Windows " & _
                 "user folder only - no admin rights needed." & Chr(10) & Chr(10) & _
                 "Download now?" & Chr(10) & _
                 "(Choosing No continues with the built-in parser.)", _
                 vbYesNo + vbQuestion, "Improve PDF Import")
    If ans <> vbYes Then Exit Function

    Dim tmpScript As String: tmpScript = Environ("TEMP") & "\oc_poppler_install.ps1"
    Dim tmpResult As String: tmpResult = Environ("TEMP") & "\oc_poppler_path.txt"
    On Error Resume Next: Kill tmpScript: Kill tmpResult: On Error GoTo 0

    Dim fNum As Integer
    fNum = FreeFile
    Open tmpScript For Output As #fNum
    Print #fNum, BuildPopplerInstallScript()
    Close #fNum

    Application.StatusBar = "Downloading Poppler (one-time, ~16 MB)..."
    On Error Resume Next
    Sh.Run "powershell -NonInteractive -ExecutionPolicy Bypass -File """ & tmpScript & """ """ & tmpResult & """", 0, True
    On Error GoTo 0
    Application.StatusBar = False
    On Error Resume Next: Kill tmpScript: On Error GoTo 0

    Dim exePath As String: exePath = ""
    If Dir(tmpResult) <> "" Then
        fNum = FreeFile
        On Error Resume Next
        Open tmpResult For Input As #fNum
        Line Input #fNum, exePath
        Close #fNum
        Kill tmpResult
        On Error GoTo 0
        exePath = Trim$(exePath)
    End If

    Dim ok As Boolean: ok = False
    If exePath <> "" Then
        If Dir(exePath) <> "" Then ok = True
    End If

    If ok Then
        OfferPopplerBootstrap = exePath
        MsgBox "Poppler installed. PDF imports now use the highest-fidelity " & _
               "extractor automatically.", vbInformation, "Poppler Installed"
    Else
        MsgBox "The Poppler download did not complete (no internet access, or " & _
               "blocked by security policy)." & Chr(10) & Chr(10) & _
               "Continuing with the built-in parser instead.", _
               vbInformation, "Poppler Not Installed"
    End If
End Function

' ---- Returns a self-contained PowerShell Poppler installer --------------------
' Downloads the latest poppler-windows release zip (pinned fallback if the
' GitHub API is unreachable) into %LOCALAPPDATA%\Poppler, locates pdftotext.exe,
' and records its path in a pointer file plus the result file for the caller.
Private Function BuildPopplerInstallScript() As String
    Dim s As String
    s = "param([string]$resultFile)" & vbLf
    s = s & "$ErrorActionPreference = 'Stop'" & vbLf
    s = s & "try {" & vbLf
    s = s & "    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072" & vbLf
    s = s & "    $dest = Join-Path $env:LOCALAPPDATA 'Poppler'" & vbLf
    s = s & "    New-Item -ItemType Directory -Force -Path $dest | Out-Null" & vbLf
    s = s & "    $url = 'https://github.com/oschwartz10612/poppler-windows/releases/download/v26.02.0-0/Release-26.02.0-0.zip'" & vbLf
    s = s & "    try {" & vbLf
    s = s & "        $rel = Invoke-RestMethod -Uri 'https://api.github.com/repos/oschwartz10612/poppler-windows/releases/latest' -TimeoutSec 15" & vbLf
    s = s & "        $a = @($rel.assets | Where-Object { $_.name -like '*.zip' })" & vbLf
    s = s & "        if ($a.Count -gt 0) { $url = $a[0].browser_download_url }" & vbLf
    s = s & "    } catch { }" & vbLf
    s = s & "    $zip = Join-Path $env:TEMP 'oc_poppler.zip'" & vbLf
    s = s & "    Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $zip -TimeoutSec 600" & vbLf
    s = s & "    Expand-Archive -LiteralPath $zip -DestinationPath $dest -Force" & vbLf
    s = s & "    Remove-Item -LiteralPath $zip -Force -ErrorAction SilentlyContinue" & vbLf
    s = s & "    $exe = Get-ChildItem -LiteralPath $dest -Recurse -Filter 'pdftotext.exe' | Select-Object -First 1" & vbLf
    s = s & "    if (-not $exe) { exit 1 }" & vbLf
    s = s & "    Set-Content -LiteralPath (Join-Path $dest 'pdftotext_path.txt') -Value $exe.FullName -Encoding ASCII" & vbLf
    s = s & "    Set-Content -LiteralPath $resultFile -Value $exe.FullName -Encoding ASCII" & vbLf
    s = s & "} catch { exit 1 }" & vbLf
    BuildPopplerInstallScript = s
End Function

' ---- PDF text extraction (three strategies, zero required installs) ----
Private Function ExtractPdfText(pdfPath As String) As String
    ExtractPdfText = ""
    Dim tmpOut As String: tmpOut = Environ("TEMP") & "\oc_pdf_text.txt"
    Dim fNum As Integer
    Dim content As String, Ln As String

    ' ----------------------------------------------------------------
    ' Strategy 1: pdftotext (Poppler) - best layout fidelity
    '   Found on PATH or in the self-installed per-user copy; when
    '   absent, offers a one-time automatic download (prompted, no
    '   admin rights). Declining falls through to Strategy 2.
    ' ----------------------------------------------------------------
    On Error Resume Next: Kill tmpOut: On Error GoTo 0
    Dim Sh As Object
    Set Sh = CreateObject("WScript.Shell")

    Dim p2tPath As String
    p2tPath = ResolvePdftotext(Sh)
    If p2tPath = "" Then p2tPath = OfferPopplerBootstrap(Sh)

    If p2tPath <> "" Then
        On Error Resume Next
        Sh.Run """" & p2tPath & """ -layout """ & pdfPath & """ """ & tmpOut & """", 0, True
        On Error GoTo 0
    End If

    If Dir(tmpOut) <> "" Then
        fNum = FreeFile: content = ""
        Open tmpOut For Input As #fNum
        Do While Not EOF(fNum): Line Input #fNum, Ln: content = content & Ln & Chr(10): Loop
        Close #fNum
        On Error Resume Next: Kill tmpOut: On Error GoTo 0
        ExtractPdfText = content
        Exit Function
    End If

    ' ----------------------------------------------------------------
    ' Strategy 2: Built-in PowerShell PDF parser
    '   Uses only .NET Framework (Windows 7+) - no install required.
    '   Handles FlateDecode (zlib) compressed content streams.
    '   Reconstructs table rows by grouping text at the same Y position.
    ' ----------------------------------------------------------------
    Dim tmpScript As String: tmpScript = Environ("TEMP") & "\oc_pdf_parse.ps1"
    On Error Resume Next: Kill tmpScript: Kill tmpOut: On Error GoTo 0

    fNum = FreeFile
    Open tmpScript For Output As #fNum
    Print #fNum, BuildPdfExtractScript()
    Close #fNum

    Sh.Run "powershell -NonInteractive -ExecutionPolicy Bypass -File """ & tmpScript & """ """ & pdfPath & """ """ & tmpOut & """", 0, True

    If Dir(tmpOut) <> "" Then
        fNum = FreeFile: content = ""
        Open tmpOut For Input As #fNum
        Do While Not EOF(fNum): Line Input #fNum, Ln: content = content & Ln & Chr(10): Loop
        Close #fNum
        On Error Resume Next: Kill tmpOut: Kill tmpScript: On Error GoTo 0
        If Len(content) > 0 Then
            ExtractPdfText = content
            Exit Function
        End If
    End If
    On Error Resume Next: Kill tmpScript: On Error GoTo 0

    ' ----------------------------------------------------------------
    ' Strategy 3: Adobe Acrobat COM (full Acrobat only, not Reader)
    ' ----------------------------------------------------------------
    Dim acro As Object: Set acro = Nothing
    On Error Resume Next: Set acro = CreateObject("AcroExch.App"): On Error GoTo 0
    If acro Is Nothing Then Exit Function

    Dim pdDoc As Object: Set pdDoc = Nothing
    On Error Resume Next: Set pdDoc = CreateObject("AcroExch.PDDoc"): On Error GoTo 0
    If pdDoc Is Nothing Then Exit Function

    If Not pdDoc.Open(pdfPath) Then Exit Function

    Dim jsObj As Object
    On Error Resume Next: Set jsObj = pdDoc.GetJSObject: On Error GoTo 0
    If Not jsObj Is Nothing Then
        Dim pG As Long, allText As String
        For pG = 0 To pdDoc.GetNumPages - 1
            On Error Resume Next
            allText = allText & jsObj.getPageNthWord(pG, 0, True) & Chr(10)
            On Error GoTo 0
        Next pG
        ExtractPdfText = allText
    End If
    pdDoc.Close False
End Function

' ---- Returns a self-contained PowerShell PDF text extractor ----
' No external tools required. Uses .NET DeflateStream (built into Windows 7+).
' Parses PDF content streams, handles FlateDecode compression, and
' reconstructs text lines by grouping glyphs at the same Y position.
Private Function BuildPdfExtractScript() As String
    Dim s As String
    s = "param([string]$pdf, [string]$out)" & vbLf
    s = s & "$enc  = [Text.Encoding]::GetEncoding(1252)" & vbLf
    s = s & "$bytes = [IO.File]::ReadAllBytes($pdf)" & vbLf
    s = s & "$raw   = $enc.GetString($bytes)" & vbLf
    s = s & "" & vbLf
    s = s & "function Inflate([byte[]]$b) {" & vbLf
    s = s & "    try {" & vbLf
    s = s & "        $skip = if($b.Length -gt 2 -and ($b[0] -band 0x0F) -eq 8){2}else{0}" & vbLf
    s = s & "        $ms = [IO.MemoryStream]::new($b,$skip,$b.Length-$skip)" & vbLf
    s = s & "        $ds = [IO.Compression.DeflateStream]::new($ms,[IO.Compression.CompressionMode]::Decompress)" & vbLf
    s = s & "        $os = [IO.MemoryStream]::new(); $ds.CopyTo($os)" & vbLf
    s = s & "        return $enc.GetString($os.ToArray())" & vbLf
    s = s & "    } catch { return '' }" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "function Parse-Stream([string]$cs) {" & vbLf
    s = s & "    # Group text tokens by rounded Y position (Tm sets absolute coords)" & vbLf
    s = s & "    $rows = [Collections.Generic.SortedDictionary[double,string]]::new()" & vbLf
    s = s & "    $curY = 0.0" & vbLf
    s = s & "    # Match: optional leading nums + operator  OR  string literal + Tj/TJ" & vbLf
    s = s & "    $re = [regex]'(?s)(-?\d[\d.]*)\s+(-?\d[\d.]*)\s+(-?\d[\d.]*)\s+(-?\d[\d.]*)\s+(-?\d[\d.]*)\s+(-?\d[\d.]*)\s+Tm|\(([^)]*)\)\s*Tj|\[([^\]]*)\]\s*TJ'" & vbLf
    s = s & "    foreach($m in $re.Matches($cs)) {" & vbLf
    s = s & "        if($m.Groups[6].Success) {" & vbLf
    s = s & "            $curY = [math]::Round([double]$m.Groups[6].Value, 1)" & vbLf
    s = s & "        } elseif($m.Groups[7].Success) {" & vbLf
    s = s & "            if(-not $rows.ContainsKey($curY)){$rows[$curY]=''}" & vbLf
    s = s & "            $rows[$curY] += $m.Groups[7].Value + ' '" & vbLf
    s = s & "        } elseif($m.Groups[8].Success) {" & vbLf
    s = s & "            $inner = $m.Groups[8].Value" & vbLf
    s = s & "            $words = [regex]::Matches($inner,'\(([^)]*)\)')" & vbLf
    s = s & "            $piece = ($words | ForEach-Object{$_.Groups[1].Value}) -join ''" & vbLf
    s = s & "            if(-not $rows.ContainsKey($curY)){$rows[$curY]=''}" & vbLf
    s = s & "            $rows[$curY] += $piece + ' '" & vbLf
    s = s & "        }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    # Return lines sorted Y descending (PDF Y=0 is bottom of page)" & vbLf
    s = s & "    return ($rows.Keys | Sort-Object -Descending | ForEach-Object{ $rows[$_].Trim() }) -join [char]10" & vbLf
    s = s & "}" & vbLf
    s = s & "" & vbLf
    s = s & "$result = ''" & vbLf
    s = s & "$pos = 0" & vbLf
    s = s & "while($true) {" & vbLf
    s = s & "    $si = $raw.IndexOf('stream',$pos); if($si -lt 0){break}" & vbLf
    s = s & "    $di = $raw.LastIndexOf('<<',$si)" & vbLf
    s = s & "    $dict = if($di -ge 0){$raw.Substring($di,$si-$di)}else{''}" & vbLf
    s = s & "    # Skip image/font/metadata streams" & vbLf
    s = s & "    if($dict -match '/Subtype\s*/Image' -or $dict -match '/Type\s*/FontDescriptor'){$pos=$si+6;continue}" & vbLf
    s = s & "    $ss = $si+6" & vbLf
    s = s & "    if($ss -lt $raw.Length -and $raw[$ss] -eq [char]13){$ss++}" & vbLf
    s = s & "    if($ss -lt $raw.Length -and $raw[$ss] -eq [char]10){$ss++}" & vbLf
    s = s & "    $ei = $raw.IndexOf('endstream',$ss); if($ei -lt 0){break}" & vbLf
    s = s & "    $slen = $ei - $ss" & vbLf
    s = s & "    if($slen -gt 0 -and $slen -lt 5000000) {" & vbLf
    s = s & "        $cs = ''" & vbLf
    s = s & "        if($dict -match '/FlateDecode|/Fl\b') {" & vbLf
    s = s & "            $cb = $bytes[$ss..($ei-1)]" & vbLf
    s = s & "            $cs = Inflate $cb" & vbLf
    s = s & "        } else {" & vbLf
    s = s & "            $cs = $raw.Substring($ss,$slen)" & vbLf
    s = s & "        }" & vbLf
    s = s & "        if($cs){ $result += (Parse-Stream $cs) + [char]10 }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    $pos = $ei+9" & vbLf
    s = s & "}" & vbLf
    s = s & "[IO.File]::WriteAllText($out,$result,[Text.Encoding]::UTF8)" & vbLf
    BuildPdfExtractScript = s
End Function

' ---- Parse PDF text for Summary table rows with SF <= maxSF ----
'
' PRIMARY format (COMPASS "Anticollision Report" Summary section, as produced
' by the built-in PowerShell extractor which groups glyphs by Y position):
'
'   [CC, ES|SF] 100/01-10-071-12W6/00 - 100/01-10 Surveys 542.46 555.07 426.51 413.68 33.257
'
'   i.e.  <wellbore/design>  RefMD  OffsetMD  BetweenCentres  BetweenEllipses  SF
'   The whole Summary table may collapse onto ONE text line, so the regex runs
'   globally over the Summary region instead of line-by-line.
'
' LEGACY fallback (older jumbled extraction with concatenated digits):
'   "Level 4 , SF 2,604.552,604.1018.567.201.634 TOURMALINE HZ SUNDOWN H04..."
Private Function ParseAcSummary(pdfText As String, aWell() As String, _
        aRefMD() As Double, aBetween() As Double, aSF() As Double, _
        Optional ByVal maxSF As Double = 2#) As Long
    Dim nHits As Long: nHits = 0

    ' ---- Bound the search to the Summary region ----
    ' Start at "Summary"; stop at the per-well detail section (first marker found).
    Dim startPos As Long: startPos = InStr(1, pdfText, "Summary", vbTextCompare)
    If startPos = 0 Then startPos = 1
    Dim endPos As Long: endPos = Len(pdfText) + 1
    Dim marker As Variant
    For Each marker In Array("Semi Major", "Offset Design", "Highside")
        Dim p As Long: p = InStr(startPos, pdfText, CStr(marker), vbTextCompare)
        If p > 0 And p < endPos Then endPos = p
    Next marker
    Dim region As String: region = mid(pdfText, startPos, endPos - startPos)

    ' ---- PRIMARY: wellbore-anchored row with 4 distances + SF ----
    ' num2 = comma-formatted 2-decimal value; extractor may inject stray spaces
    ' inside numbers ("1, 110.00", "4. 11"), hence the optional \s? gaps.
    Dim num2 As String: num2 = "\d{1,3}(?:,\s?\d{3})*\.\s?\d{2}"
    Dim sfNum As String: sfNum = "\d{1,3}\.\s?\d{3}"
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "(\d{2,3}/\S{5,30}/\d{2}\s*-\s*.{1,60}?)\s+" & _
                 "(" & num2 & ")\s+(" & num2 & ")\s+(" & num2 & ")\s+(" & num2 & ")\s+" & _
                 "(" & sfNum & ")"

    Dim ms As Object: Set ms = re.Execute(region)
    Dim m As Object
    For Each m In ms
        Dim sfVal As Double: sfVal = CleanNum(m.SubMatches(5))
        If sfVal > 0 And sfVal <= maxSF Then
            aWell(nHits) = CleanWellName(CStr(m.SubMatches(0)))
            aRefMD(nHits) = CleanNum(m.SubMatches(1))
            aBetween(nHits) = CleanNum(m.SubMatches(3))
            aSF(nHits) = sfVal
            nHits = nHits + 1
            If nHits > 200 Then Exit For
        End If
    Next m

    If nHits > 0 Or ms.Count > 0 Then
        ' Primary pattern recognised the summary rows (even if none <= maxSF).
        ParseAcSummary = nHits
        Exit Function
    End If

    ' ---- LEGACY fallback: "Level N , SF" concatenated-digit rows ----
    Dim lines() As String: lines = Split(pdfText, Chr(10))
    Dim inSummary As Boolean: inSummary = False
    Dim i As Long

    Dim reMD As Object: Set reMD = CreateObject("VBScript.RegExp")
    reMD.Global = True
    reMD.Pattern = "\d{1,3}(?:,\d{3})+\.\d{2}"

    ' "18.567.201.634" => BC=18.56  BE=7.20  SF=1.634
    Dim reTriplet As Object: Set reTriplet = CreateObject("VBScript.RegExp")
    reTriplet.Global = False
    reTriplet.Pattern = "(\d{1,2})\.(\d{2})(\d{1,2})\.(\d{2})(\d)\.(\d{3})"

    For i = 0 To UBound(lines)
        Dim Ln As String: Ln = Trim(lines(i))
        If InStr(1, Ln, "Summary", vbTextCompare) > 0 Then inSummary = True
        If Not inSummary Then GoTo NextAcLine

        ' Must be an explicit Level N critical SF row (has both "Level" and ", SF").
        If InStr(1, Ln, "Level", vbTextCompare) = 0 Then GoTo NextAcLine
        If InStr(1, Ln, ", SF", vbTextCompare) = 0 Then GoTo NextAcLine

        Dim mdMs As Object: Set mdMs = reMD.Execute(Ln)
        If mdMs.Count < 2 Then GoTo NextAcLine
        Dim refMdVal As Double: refMdVal = CDbl(Replace(mdMs(0).Value, ",", ""))

        ' Clip to text after the SECOND comma-MD (the OffsetMD).
        Dim lastMd As Object: Set lastMd = mdMs(1)
        Dim afterMDs As String
        afterMDs = mid(Ln, lastMd.FirstIndex + lastMd.Length + 1)

        Dim tripMs As Object: Set tripMs = reTriplet.Execute(afterMDs)
        If tripMs.Count = 0 Then GoTo NextAcLine
        Dim tm As Object: Set tm = tripMs(0)
        Dim bcVal As Double: bcVal = CDbl(tm.SubMatches(0) & "." & tm.SubMatches(1))
        Dim sfVal2 As Double: sfVal2 = CDbl(tm.SubMatches(4) & "." & tm.SubMatches(5))

        If sfVal2 > 0 And sfVal2 <= maxSF Then
            ' Well name trails the numeric block in this format
            Dim tmObj As Object: Set tmObj = tripMs(0)
            Dim tail As String
            tail = Trim(mid(afterMDs, tmObj.FirstIndex + tmObj.Length + 1))
            If Len(tail) > 60 Then tail = Left(tail, 60)
            aWell(nHits) = CleanWellName(tail)
            aRefMD(nHits) = refMdVal
            aBetween(nHits) = bcVal
            aSF(nHits) = sfVal2
            nHits = nHits + 1
            If nHits > 200 Then Exit For
        End If
NextAcLine:
    Next i
    ParseAcSummary = nHits
End Function

' Strip commas and stray spaces the PDF extractor injects inside numbers.
Private Function CleanNum(ByVal s As String) As Double
    s = Replace(Replace(s, ",", ""), " ", "")
    CleanNum = CDbl(s)
End Function

' Collapse runs of whitespace the PDF extractor leaves inside well names.
Private Function CleanWellName(ByVal s As String) As String
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanWellName = Trim(s)
End Function

Private Function ExtractNums(s As String, nums() As Double) As Long
    ' Extracts numbers that have a decimal point (2 or 3 decimal places).
    ' This naturally skips integers like job numbers (34783) and Level markers (4, 5).
    ' Handles comma-formatted numbers (2,604.55) and concatenated runs (2,604.552,604.10...).
    ' Pattern: 3-decimal tried first to avoid "23.91" matching inside "23.915".
    ReDim nums(50)
    Dim n As Long: n = 0
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\d+\.\d{3}|\d{1,3}(?:,\d{3})*\.\d{2}"
    Dim ms As Object: Set ms = re.Execute(s)
    Dim m As Object
    For Each m In ms
        Dim t As String: t = Replace(m.Value, ",", "")
        nums(n) = CDbl(t): n = n + 1
        If n > 50 Then Exit For
    Next m
    ExtractNums = n
End Function

' AC Setup table: logical col 0..4 -> physical merge on row r (J=10 .. R=18)
'   0 # -> J ; 1 RefMD -> K:L ; 2 C2C -> M:N ; 3 SF -> O:P ; 4 Risk -> Q:R
Private Function AcSetupCell(ws As Worksheet, ByVal r As Long, ByVal logicalCol As Long) As Range
    Dim c0 As Long, c1 As Long
    Select Case logicalCol
        Case 0: c0 = 10: c1 = 10
        Case 1: c0 = 11: c1 = 12
        Case 2: c0 = 13: c1 = 14
        Case 3: c0 = 15: c1 = 16
        Case 4: c0 = 17: c1 = 18
        Case Else: c0 = 10: c1 = 10
    End Select
    Set AcSetupCell = ws.Range(ws.Cells(r, c0), ws.Cells(r, c1))
End Function

Private Sub AcSetupMergeRow(ws As Worksheet, ByVal r As Long)
    Dim c As Long
    For c = 0 To 4
        AcSetupCell(ws, r, c).Merge
    Next c
End Sub

' ---- Build formatted AC results table on Setup sheet (J11:R16) ----
' Headers permanent; max 5 data rows. Caller supplies severity-sorted arrays.
Public Sub BuildAcTable(nHits As Long, aRefMD() As Double, aBetween() As Double, aSF() As Double)
    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim nShow As Long
    Dim sf As Double
    Dim rowBg As Long
    Dim risk As String
    Dim riskFg As Long
    Dim vals(0 To 4) As Variant
    Dim hdrs(0 To 4) As String
    Dim cell As Range
    Const BASE_ROW As Long = 11
    Const BASE_COL As Long = 10   ' J
    Const AC_COLS  As Long = 9    ' J..R
    Const AC_MAX   As Long = 5
    Const LEGACY_ROW As Long = 1
    Const LEGACY_COL As Long = 19 ' S

    Dim wasProt As Boolean

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    wasProt = SheetUnprotectForVba(ws)
    On Error GoTo AcTableDone

    ' Clear new home + legacy S1:W block so old tables never linger
    ws.Range(ws.Cells(BASE_ROW, BASE_COL), ws.Cells(BASE_ROW + AC_MAX, BASE_COL + AC_COLS - 1)).UnMerge
    ws.Range(ws.Cells(BASE_ROW, BASE_COL), ws.Cells(BASE_ROW + AC_MAX, BASE_COL + AC_COLS - 1)).Clear
    On Error Resume Next
    ws.Range(ws.Cells(LEGACY_ROW, LEGACY_COL), ws.Cells(LEGACY_ROW + 250, LEGACY_COL + 4)).UnMerge
    ws.Range(ws.Cells(LEGACY_ROW, LEGACY_COL), ws.Cells(LEGACY_ROW + 250, LEGACY_COL + 4)).Clear
    On Error GoTo AcTableDone

    ' Keep ConfigSheet crew widths on J-R; do not override ColumnWidth here.

    hdrs(0) = "#"
    hdrs(1) = "Ref MD (m)"
    hdrs(2) = "Between Centres (m)"
    hdrs(3) = "Separation Factor"
    hdrs(4) = "Risk"

    r = BASE_ROW
    ws.Rows(r).rowHeight = 18
    AcSetupMergeRow ws, r
    For c = 0 To 4
        Set cell = AcSetupCell(ws, r, c)
        With cell
            .Value = hdrs(c)
            .Interior.Color = cTeal()
            .Font.Color = cBlk()
            .Font.bold = True
            .Font.name = "Consolas"
            .Font.Size = 8
            .VerticalAlignment = xlVAlignCenter
            If c = 0 Or c = 4 Then
                .HorizontalAlignment = xlHAlignCenter
            Else
                .HorizontalAlignment = xlHAlignRight
            End If
        End With
    Next c
    With ws.Range(ws.Cells(r, BASE_COL), ws.Cells(r, BASE_COL + AC_COLS - 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cMed()
        .Weight = xlThin
    End With
    r = r + 1

    nShow = nHits
    If nShow > AC_MAX Then nShow = AC_MAX

    If nShow <= 0 Then
        ws.Rows(r).rowHeight = 20
        With ws.Range(ws.Cells(r, BASE_COL), ws.Cells(r, BASE_COL + AC_COLS - 1))
            .Merge
            .Value = "  No data  --  use Import AC button to load"
            .Interior.Color = cBg()
            .Font.Color = cDk()
            .Font.name = "Consolas"
            .Font.Size = 8
            .VerticalAlignment = xlVAlignCenter
        End With
        r = r + 1
    Else
        For i = 0 To nShow - 1
            ws.Rows(r).rowHeight = 16
            sf = aSF(i)
            If (i Mod 2) = 0 Then
                rowBg = cWh()
            Else
                rowBg = cBg()
            End If
            If sf < 1 Then
                risk = "CRITICAL"
                riskFg = cRedTxt()
            ElseIf sf < 1.5 Then
                risk = "HIGH"
                riskFg = cRedTxt()
            Else
                risk = "CAUTION"
                riskFg = RGB(140, 90, 0)
            End If
            vals(0) = i + 1
            vals(1) = aRefMD(i)
            vals(2) = aBetween(i)
            vals(3) = sf
            vals(4) = risk
            AcSetupMergeRow ws, r
            For c = 0 To 4
                Set cell = AcSetupCell(ws, r, c)
                With cell
                    .Value = vals(c)
                    .Interior.Color = rowBg
                    If c = 3 Or c = 4 Then
                        .Font.Color = riskFg
                    Else
                        .Font.Color = cBlk()
                    End If
                    .Font.bold = (c = 4)
                    .Font.name = "Consolas"
                    .Font.Size = 8
                    .VerticalAlignment = xlVAlignCenter
                    If c = 0 Or c = 4 Then
                        .HorizontalAlignment = xlHAlignCenter
                    Else
                        .HorizontalAlignment = xlHAlignRight
                    End If
                    If c = 3 Then .numberFormat = "0.000"
                    If c = 1 Or c = 2 Then .numberFormat = "0.00"
                End With
            Next c
            With ws.Range(ws.Cells(r, BASE_COL), ws.Cells(r, BASE_COL + AC_COLS - 1)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(235, 235, 235)
                .Weight = xlHairline
            End With
            r = r + 1
        Next i
    End If

    ' Outer border covers header through last used body row; for empty state
    ' extend visual box through R16 so the reserved block is obvious.
    Dim endRow As Long
    endRow = r - 1
    If nShow <= 0 Then endRow = BASE_ROW + AC_MAX
    With ws.Range(ws.Cells(BASE_ROW, BASE_COL), ws.Cells(endRow, BASE_COL + AC_COLS - 1))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = cMed()
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = cMed()
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = cMed()
    End With

AcTableDone:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    On Error GoTo 0
End Sub

' Refresh path cells on the Setup sheet after an import
Private Sub UpdateImportPathDisplay(shName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    If ws Is Nothing Then Exit Sub
    Dim nm As String: nm = IIf(shName = SH_SURVEY, "OC_ImpPlanRow", "OC_ImpAcRow")
    Dim anchor As Range: Set anchor = ws.Range(nm)
    If Not anchor Is Nothing Then
        ' Path merge starts in the first column to the right of the button host merge
        Dim btnArea As Range: Set btnArea = anchor.MergeArea
        Dim pathCell As Range
        Set pathCell = btnArea.Cells(1, btnArea.Columns.Count).Offset(0, 1)
        pathCell.MergeArea.Cells(1, 1).Value = GetImportPath(shName)
    End If
    On Error GoTo 0
End Sub

' ================================================================================
'  AC DEMO TABLE  (Sheet1 preview with sample data)
'  Call this once after RebuildSetup to show what the AC table looks like.
' ================================================================================
Public Sub DemoAcTable()
    ' Three critical separations (SF < 2.0) from a P3 AC check -- demo data only
    '   Offset Well H -- Level 4 proximity at 2604.55 m MD
    '   Offset Well J -- Level 5 proximity at 2460.00 m MD
    '   Offset Well J -- Level 5 proximity at 2610.00 m MD  (closest approach)
    Dim nHits As Long
    Dim aRefMD(2) As Double
    Dim aBetween(2) As Double
    Dim aSepF(2) As Double
    nHits = 3
    aRefMD(0) = 2604.55: aBetween(0) = 18.56: aSepF(0) = 1.634
    aRefMD(1) = 2460#:   aBetween(1) = 18.76: aSepF(1) = 1.797
    aRefMD(2) = 2610#:   aBetween(2) = 20.81: aSepF(2) = 1.842
    BuildAcTable nHits, aRefMD, aBetween, aSepF
End Sub

' Tight crew (6 rows) + AC at J11 without a full RebuildSetup.
Public Sub RedrawCrewAndAc()
    Dim ws As Worksheet
    Dim wasProt As Boolean
    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    wasProt = SheetUnprotectForVba(ws)
    On Error Resume Next
    ws.Range(ws.Cells(R_BODY, C_CRL), ws.Cells(10, C_LAST)).UnMerge
    ws.Range(ws.Cells(R_BODY, C_CRL), ws.Cells(10, C_LAST)).Clear
    On Error GoTo Fail
    DrawCrew ws, R_BODY, 10
    ReloadAcTableFromHidden
    SheetReprotectAfterVba ws, wasProt
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
    MsgBox "RedrawCrewAndAc: " & Err.Description, vbExclamation, "Setup"
End Sub

Private Sub ReloadAcTableFromHidden()
    Dim nHits As Long
    Dim aRefMD(200) As Double, aBetween(200) As Double, aSF(200) As Double
    If Not SheetExists(SH_AC) Then
        BuildAcTable 0, aRefMD, aBetween, aSF
        Exit Sub
    End If
    Dim ac As Worksheet: Set ac = Worksheets(SH_AC)
    Dim lastR As Long: lastR = ac.Cells(ac.Rows.Count, 2).End(xlUp).Row
    Dim r As Long
    For r = 3 To lastR
        If Not IsNumeric(ac.Cells(r, 2).Value2) Then GoTo NextAc
        If nHits > 200 Then Exit For
        aRefMD(nHits) = CDbl(ac.Cells(r, 2).Value2)
        aBetween(nHits) = val(ac.Cells(r, 3).Value2 & "")
        aSF(nHits) = val(ac.Cells(r, 4).Value2 & "")
        nHits = nHits + 1
NextAc:
    Next r
    BuildAcTable nHits, aRefMD, aBetween, aSF
End Sub











