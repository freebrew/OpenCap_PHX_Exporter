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
' ================================================================================

' -- Sheet names ---------------------------------------------------------------
Private Const SH_SETUP     As String = "Setup"
Private Const SH_JOB       As String = "_OC_Job"
Private Const SH_CREW      As String = "_OC_Crew"
Private Const SH_BHA       As String = "_OC_BHA"
Private Const SH_SLIDE     As String = "_OC_Slide"
Private Const SH_INVENTORY As String = "_OC_Inventory"
Private Const SH_SURVEY    As String = "_OC_Survey"
Private Const SH_AC        As String = "_OC_AC"

' -- CSV filename tokens (case-insensitive partial match) ----------------------
Private Const TOK_JOB       As String = "job-details"
Private Const TOK_CREW      As String = "crew"
Private Const TOK_BHA       As String = "bha-equipment"
Private Const TOK_SLIDE     As String = "slide-rotate"
Private Const TOK_INVENTORY As String = "inventory"

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

' ================================================================================
'  COLOR PALETTE  (same hue as MDL_DDTools + green/red glow additions)
' ================================================================================
Private Function cWh() As Long:     cWh = RGB(255, 255, 255): End Function
Private Function cBg() As Long:     cBg = RGB(245, 245, 245): End Function
Private Function cMed() As Long:    cMed = RGB(200, 200, 200): End Function
Private Function cDk() As Long:     cDk = RGB(120, 120, 120): End Function
Private Function cBlk() As Long:    cBlk = RGB(30, 30, 30): End Function
Private Function cTeal() As Long:   cTeal = RGB(0, 79, 79): End Function
Private Function cTealLt() As Long: cTealLt = RGB(0, 110, 110): End Function
Private Function cLine() As Long:   cLine = RGB(0, 60, 60): End Function
' Green glow - detected files
Private Function cGrnBadge() As Long: cGrnBadge = RGB(0, 155, 95): End Function
Private Function cGrnRow() As Long:   cGrnRow   = RGB(228, 252, 238): End Function
Private Function cGrnTxt() As Long:   cGrnTxt   = RGB(0, 110, 60): End Function
' Red - missing files
Private Function cRedBadge() As Long: cRedBadge = RGB(176, 40, 40): End Function
Private Function cRedRow() As Long:   cRedRow   = RGB(255, 240, 240): End Function
Private Function cRedTxt() As Long:   cRedTxt   = RGB(160, 30, 30): End Function

' ================================================================================
'  PUBLIC ENTRY POINTS
' ================================================================================

Public Sub InitSetup()
    If mBusy Then Exit Sub
    mBusy = True

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    On Error GoTo ErrHandler

    SetupDataSheets
    EnsureSheet SH_SETUP, True

    Dim wbPath As String: wbPath = ThisWorkbook.Path
    If wbPath <> "" Then
        Dim fJob As String:   fJob   = FindCsvByToken(wbPath, TOK_JOB)
        Dim fCrew As String:  fCrew  = FindCsvByToken(wbPath, TOK_CREW)
        Dim fBHA As String:   fBHA   = FindCsvByToken(wbPath, TOK_BHA)
        Dim fSlide As String: fSlide = FindCsvByToken(wbPath, TOK_SLIDE)
        Dim fInv As String:   fInv   = FindCsvByToken(wbPath, TOK_INVENTORY)
        If fJob   <> "" Then LoadCsv fJob,   SH_JOB
        If fCrew  <> "" Then LoadCsv fCrew,  SH_CREW
        If fBHA   <> "" Then LoadCsv fBHA,   SH_BHA
        If fSlide <> "" Then LoadCsv fSlide, SH_SLIDE
        If fInv   <> "" Then LoadCsv fInv,   SH_INVENTORY
    End If

    BuildSetupUI

    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    mBusy = False

    Exit Sub

ErrHandler:
    Dim eN As Long: eN = Err.Number
    Dim eD As String: eD = Err.Description
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    Application.StatusBar      = False
    mBusy = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = "INIT ERR " & eN & ": " & eD
    On Error GoTo 0
End Sub

Public Sub RefreshSetup()
    If mBusy Then Exit Sub
    mBusy = True

    Dim wbPath As String: wbPath = ThisWorkbook.Path
    If wbPath = "" Then mBusy = False: Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual
    Application.StatusBar      = "OpenCap: scanning folder for CSV files..."

    On Error GoTo ErrHandler

    SetupDataSheets

    Dim fJob As String:   fJob   = FindCsvByToken(wbPath, TOK_JOB)
    Dim fCrew As String:  fCrew  = FindCsvByToken(wbPath, TOK_CREW)
    Dim fBHA As String:   fBHA   = FindCsvByToken(wbPath, TOK_BHA)
    Dim fSlide As String: fSlide = FindCsvByToken(wbPath, TOK_SLIDE)
    Dim fInv As String:   fInv   = FindCsvByToken(wbPath, TOK_INVENTORY)

    If fJob   <> "" Then LoadCsv fJob,   SH_JOB
    If fCrew  <> "" Then LoadCsv fCrew,  SH_CREW
    If fBHA   <> "" Then LoadCsv fBHA,   SH_BHA
    If fSlide <> "" Then LoadCsv fSlide, SH_SLIDE
    If fInv   <> "" Then LoadCsv fInv,   SH_INVENTORY

    Application.StatusBar = "OpenCap: building Setup sheet..."
    EnsureSheet SH_SETUP, True
    BuildSetupUI

    Application.StatusBar      = False
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    mBusy = False
    Exit Sub

ErrHandler:
    Dim eNum As Long:   eNum = Err.Number
    Dim eMsg As String: eMsg = Err.Description
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    Application.StatusBar      = False
    mBusy = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = "REFRESH ERR " & eNum & ": " & eMsg
    On Error GoTo 0
End Sub

Public Sub RebuildSetup()
    If mBusy Then Exit Sub
    mBusy = True

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    On Error GoTo ErrHandler
    EnsureSheet SH_SETUP, True
    BuildSetupUI

    Application.StatusBar      = False
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    mBusy = False
    Exit Sub

ErrHandler:
    Dim eR As Long: eR = Err.Number
    Dim eM As String: eM = Err.Description
    Application.Calculation    = xlCalculationAutomatic
    Application.EnableEvents   = True
    Application.ScreenUpdating = True
    mBusy = False
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Cells(1, 20).Value = "REBUILD ERR " & eR & ": " & eM
    On Error GoTo 0
End Sub

' ================================================================================
'  SHEET HOUSEKEEPING
' ================================================================================

Private Sub SetupDataSheets()
    Dim n(6) As String
    n(0) = SH_JOB: n(1) = SH_CREW: n(2) = SH_BHA: n(3) = SH_SLIDE: n(4) = SH_INVENTORY
    n(5) = SH_SURVEY: n(6) = SH_AC
    Dim i As Integer
    For i = 0 To 6
        If Not SheetExists(n(i)) Then
            ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets( _
                ThisWorkbook.Sheets.Count)).Name = n(i)
        End If
        Worksheets(n(i)).Visible = xlSheetVeryHidden
    Next i
End Sub

Private Sub EnsureSheet(nm As String, addBefore As Boolean)
    If Not SheetExists(nm) Then
        Dim ws As Worksheet
        If addBefore Then
            Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        Else
            Set ws = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        End If
        ws.Name = nm
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
    DrawImportFilesPanel ws, 26
    DrawNotesPanel ws, 31

    AttachButtons ws

    On Error Resume Next
    If Not Application.ActiveWorkbook Is Nothing Then
        ws.Activate
        ActiveWindow.FreezePanes = False
        ws.Cells(R_BODY, 1).Select
        ActiveWindow.FreezePanes = True
        ws.Cells(1, 1).Select
    End If
    On Error GoTo 0
End Sub

Private Sub ConfigSheet(ws As Worksheet)
    With ws
        .Cells.Interior.Color = cWh()
        .Cells.Font.Color     = cBlk()
        .Cells.Font.Name      = "Consolas"
        .Cells.Font.Size      = 9
        .Tab.Color            = cTeal()
        .DisplayPageBreaks    = False
    End With
    On Error Resume Next
    Dim owh As Object: Set owh = ActiveWindow
    If Not owh Is Nothing Then
        owh.DisplayGridlines = False
        owh.DisplayHeadings  = False
    End If
    On Error GoTo 0

    ' Column widths
    ws.Columns(C_ACCENT).ColumnWidth = 1.4
    ws.Columns(C_L1S).ColumnWidth    = 11
    ws.Columns(C_L1E).ColumnWidth    = 8
    ws.Columns(C_V1S).ColumnWidth    = 15
    ws.Columns(C_V1E).ColumnWidth    = 9
    ws.Columns(C_L2).ColumnWidth     = 14
    ws.Columns(C_V2S).ColumnWidth    = 13
    ws.Columns(C_V2E).ColumnWidth    = 8
    ws.Columns(C_GAP).ColumnWidth    = 1.8
    ws.Columns(C_CRL).ColumnWidth    = 10    ' crew role
    ws.Columns(C_CRL + 1).ColumnWidth = 4
    ws.Columns(C_CRL + 2).ColumnWidth = 17   ' crew name
    ws.Columns(C_CRL + 3).ColumnWidth = 4
    ws.Columns(C_CRR).ColumnWidth    = 18    ' crew email/phone
    ws.Columns(C_CRR + 1).ColumnWidth = 8
    ws.Columns(C_CRR + 2).ColumnWidth = 8
    ws.Columns(C_CRR + 3).ColumnWidth = 8

    ' Row heights
    Dim r As Long
    ws.Rows("1:130").RowHeight = 16
    ws.Rows(R_HDR).RowHeight = 30
    ws.Rows(R_DIV).RowHeight = 3
End Sub

' ================================================================================
'  HEADER BAND
' ================================================================================

Private Sub DrawHeader(ws As Worksheet)
    With ws.Range(ws.Cells(R_HDR, 1), ws.Cells(R_HDR, C_LAST))
        .Merge
        .Value              = "  OPENCAP  |  WELL DATABASE & SETUP"
        .Interior.Color     = cTeal()
        .Font.Color         = RGB(255, 255, 255)
        .Font.Size          = 14
        .Font.Bold          = True
        .Font.Name          = "Consolas"
        .VerticalAlignment  = xlVAlignCenter
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
    Dim J As Object: Set J = ReadJobFields()

    Dim r As Long: r = startRow

    ' -- JOB IDENTITY -------------------------------------------------------------
    SectionBar ws, r, C_ACCENT, C_V2E, "  JOB IDENTITY", cTeal(), RGB(255, 255, 255)
    r = r + 1
    Pair ws, r, "JOB ID",         GF(J, "Job ID"), _
                 "OPS STATUS",    GF(J, "opsStatus", "Ops Status")
    r = r + 1
    Pair ws, r, "JOB NAME",       GF(J, "Job Name"), _
                 "PROFILE",       GF(J, "WellProfile", "Job Type")
    r = r + 1
    Pair ws, r, "CLIENT",         GF(J, "Client"), _
                 "AFE",           GF(J, "AFE", "AFE (core)")
    r = r + 1
    Pair ws, r, "WELL / UWI",     GF(J, "UWI", "Well (core)"), _
                 "WELL LICENCE",  GF(J, "WellLicenceNumber")
    r = r + 1
    Pair ws, r, "SURFACE COORDS", GF(J, "SurfaceCoordinates"), _
                 "PROVINCE",      GF(J, "Province")
    r = r + 1
    Pair ws, r, "SPUD DATE",      FmtDate(GF(J, "SpudDate")), _
                 "START DATE",    GF(J, "Start Date", "Planned Start Date")
    r = r + 1
    Pair ws, r, "MAP SYSTEM",     GF(J, "MapSystem"), _
                 "NORTH REF",     GF(J, "NorthReference")
    r = r + 1

    ' -- WELL GEOMETRY ------------------------------------------------------------
    ws.Rows(r).RowHeight = 7: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  WELL GEOMETRY", cTealLt(), RGB(255, 255, 255)
    r = r + 1
    Pair ws, r, "GROUND MSL",     ValUnit(GF(J, "GroundMSL"), " m"), _
                 "TVD",           ValUnit(GF(J, "TVD"), " m")
    r = r + 1
    Pair ws, r, "RKB HEIGHT",     ValUnit(GF(J, "RKB"), " m"), _
                 "TOTAL KB",      ValUnit(GF(J, "TotalKB"), " m")
    r = r + 1
    Pair ws, r, "LATITUDE",       GF(J, "LatitudeDMS", "Latitude"), _
                 "LONGITUDE",     GF(J, "LongitudeDMS", "Longitude")
    r = r + 1
    Pair ws, r, "VSD",            GF(J, "VSD"), _
                 "WELL TYPE",     GF(J, "WellProfile", "Job Type")
    r = r + 1

    ' -- DIRECTIONAL PARAMETERS ---------------------------------------------------
    ws.Rows(r).RowHeight = 7: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  DIRECTIONAL PARAMETERS", cTealLt(), RGB(255, 255, 255)
    r = r + 1
    Pair ws, r, "MAG DECL",       ValUnit(GF(J, "MagneticDeclination"), Chr(176)), _
                 "REF MODEL",     GF(J, "MagneticReferenceModel")
    r = r + 1
    Pair ws, r, "DIP",            ValUnit(GF(J, "DIP"), Chr(176)), _
                 "REF DATE",      FmtDate(GF(J, "MagneticReferenceDate"))
    r = r + 1
    Pair ws, r, "TGF",            GF(J, "TGF"), _
                 "TMF",           GF(J, "TMF")
    r = r + 1
    Pair ws, r, "CONVERGENCE",    ValUnit(GF(J, "Convergence"), Chr(176)), _
                 "TOL DIP",       GF(J, "ToleranceDIP")
    r = r + 1
    Pair ws, r, "SURVEY CORR.",   GF(J, "SurveyCorrection"), _
                 "DRILL MEAS.",   GF(J, "DrillMeasuredFrom")
    r = r + 1

    ' -- EQUIPMENT & RIG ----------------------------------------------------------
    ws.Rows(r).RowHeight = 7: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  EQUIPMENT & RIG", cTeal(), RGB(255, 255, 255)
    r = r + 1
    Pair ws, r, "RIG TYPE",       GF(J, "RigType"), _
                 "TOP DRIVE",     GF(J, "TopDrive")
    r = r + 1
    Pair ws, r, "DATA RECORDER",  GF(J, "DataRecorder"), _
                 "GAMMA",         GF(J, "Gamma")
    r = r + 1
    Pair ws, r, "MWD GUIDANCE",   GF(J, "MWDGuidanceType"), _
                 "PIPE ARM",      GF(J, "PipeArm")
    r = r + 1
    Pair ws, r, "REMOTE SVC",     GF(J, "RemoteServices"), _
                 "LOADER",        GF(J, "Loader")
    r = r + 1

    ' -- CONTACTS -----------------------------------------------------------------
    ws.Rows(r).RowHeight = 7: r = r + 1
    SectionBar ws, r, C_ACCENT, C_V2E, "  CONTACTS", cTeal(), RGB(255, 255, 255)
    r = r + 1
    Pair ws, r, "COMPANY MAN",    GF(J, "CompanyMan"), _
                 "CM PHONE",      GF(J, "CompanyManPhone")
    r = r + 1
    Pair ws, r, "2ND COMP. MAN",  GF(J, "SecondCompanyMan"), _
                 "2ND CM PHONE",  GF(J, "SecondCompanyManPhone")
    r = r + 1
    Pair ws, r, "GEOLOGIST",      GF(J, "Geologist"), _
                 "GEOL. PHONE",   GF(J, "GeologistPhone")
    r = r + 1
    Pair ws, r, "DD COORDINATOR", GF(J, "DDCoordinator"), _
                 "DD COORD EMAIL", GF(J, "DDCoordinatorEmail")
    r = r + 1
    Pair ws, r, "MWD COORDINATOR", GF(J, "MWDCoordinator"), _
                 "MWD EMAIL",     GF(J, "MWDCoordinatorEmail")
    r = r + 1
    Pair ws, r, "SALES REP",      GF(J, "SalesRep"), _
                 "SALES EMAIL",   GF(J, "SalesRepEmail")
    r = r + 1

    ' Left panel outer border
    With ws.Range(ws.Cells(startRow, C_ACCENT), ws.Cells(r - 1, C_V2E))
        .Borders(xlEdgeLeft).LineStyle   = xlContinuous
        .Borders(xlEdgeLeft).Color       = cMed()
        .Borders(xlEdgeRight).LineStyle  = xlContinuous
        .Borders(xlEdgeRight).Color      = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color     = cMed()
    End With

    ' Draw crew panel alongside; cap at row 17 so files panel at row 18 has room
    DrawCrew ws, startRow, 17
End Sub

' -- Crew panel (columns J:N) --------------------------------------------------
Private Sub DrawCrew(ws As Worksheet, topRow As Long, bottomRow As Long)
    Dim r As Long: r = topRow

    SectionBar ws, r, C_CRL, C_LAST, "  CREW MANIFEST", cTeal(), RGB(255, 255, 255)
    r = r + 1

    ' Sub-header: ROLE | NAME | EMAIL | PHONE
    ws.Rows(r).RowHeight = 14
    Dim hdrBg As Long: hdrBg = cBg()
    Sgl ws, r, C_CRL, "ROLE", hdrBg, cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, C_CRL + 1, C_CRL + 2, "NAME", hdrBg, cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, C_CRL + 3, C_CRL + 5, "EMAIL", hdrBg, cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, C_CRL + 6, C_LAST, "PHONE", hdrBg, cDk(), True, 7, xlHAlignLeft
    With ws.Range(ws.Cells(r, C_CRL), ws.Cells(r, C_LAST)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = cMed(): .Weight = xlHairline
    End With
    r = r + 1

    Dim dataTop As Long: dataTop = r

    ' Rows
    Dim crew() As String: crew = ReadCrewRows()
    Dim i As Long
    For i = 0 To UBound(crew, 1)
        If r > bottomRow Then Exit For
        ws.Rows(r).RowHeight = 16

        Dim bg As Long
        If i Mod 2 = 0 Then bg = cWh() Else bg = cBg()

        Sgl ws, r, C_CRL, crew(i, 0), bg, cDk(), True, 8, xlHAlignLeft
        MrgCell ws, r, C_CRL + 1, C_CRL + 2, crew(i, 1), bg, cBlk(), (crew(i, 1) <> ""), 9, xlHAlignLeft
        MrgCell ws, r, C_CRL + 3, C_CRL + 5, crew(i, 2), bg, cDk(), False, 7, xlHAlignLeft
        MrgCell ws, r, C_CRL + 6, C_LAST, crew(i, 3), bg, cDk(), False, 7, xlHAlignLeft

        With ws.Range(ws.Cells(r, C_CRL), ws.Cells(r, C_LAST)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Color = RGB(235, 235, 235): .Weight = xlHairline
        End With
        r = r + 1
    Next i

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
        ws.Names("OC_CrewTop").Delete
        ws.Names("OC_CrewBot").Delete
        On Error GoTo 0
        ws.Names.Add Name:="OC_CrewTop", RefersToR1C1:="=R" & dataTop & "C" & C_CRL
        ws.Names.Add Name:="OC_CrewBot", RefersToR1C1:="=R" & dataBot & "C" & C_CRL

        InstallCrewSortEvent ws
    End If

    With ws.Range(ws.Cells(topRow, C_CRL), ws.Cells(dataBot, C_LAST))
        .Borders(xlEdgeLeft).LineStyle   = xlContinuous: .Borders(xlEdgeLeft).Color   = cMed()
        .Borders(xlEdgeRight).LineStyle  = xlContinuous: .Borders(xlEdgeRight).Color  = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = cMed()
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
        If InStr(1, cm.Lines(i, 1), tag, vbTextCompare) > 0 Then Exit Sub
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
        aRole(idx)  = Trim(SafeStr(ws.Cells(ro, C_CRL)))
        aName(idx)  = Trim(SafeStr(ws.Cells(ro, C_CRL + 1)))
        aEmail(idx) = Trim(SafeStr(ws.Cells(ro, C_CRL + 3)))
        aPhone(idx) = Trim(SafeStr(ws.Cells(ro, C_CRL + 6)))
    Next idx

    ' Insertion sort by (priority, name)
    Dim j As Long, k As Long
    For j = 2 To n
        Dim tR As String: tR = aRole(j)
        Dim tN As String: tN = aName(j)
        Dim tE As String: tE = aEmail(j)
        Dim tP As String: tP = aPhone(j)
        Dim pj As Long: pj = RolePriority(tR)
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
        aRole(k + 1) = tR
        aName(k + 1) = tN
        aEmail(k + 1) = tE
        aPhone(k + 1) = tP
    Next j

    ' Write back, restyle, and re-apply dropdown
    Application.ScreenUpdating = False
    For idx = 1 To n
        ro = rTop + idx - 1
        Dim bg As Long
        If (idx - 1) Mod 2 = 0 Then bg = cWh() Else bg = cBg()

        With ws.Cells(ro, C_CRL)
            .Value = aRole(idx)
            .Interior.Color = bg
            .Font.Color = cDk(): .Font.Bold = True: .Font.Size = 8
            .Font.Name = "Consolas"
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
        mrg.Font.Bold = (aName(idx) <> ""): mrg.Font.Size = 9
        mrg.Font.Name = "Consolas"

        Set mrg = ws.Range(ws.Cells(ro, C_CRL + 3), ws.Cells(ro, C_CRL + 5))
        mrg.Cells(1, 1).Value = aEmail(idx)
        mrg.Interior.Color = bg: mrg.Font.Color = cDk()
        mrg.Font.Size = 7: mrg.Font.Name = "Consolas"

        Set mrg = ws.Range(ws.Cells(ro, C_CRL + 6), ws.Cells(ro, C_LAST))
        mrg.Cells(1, 1).Value = aPhone(idx)
        mrg.Interior.Color = bg: mrg.Font.Color = cDk()
        mrg.Font.Size = 7: mrg.Font.Name = "Consolas"
    Next idx
    Application.ScreenUpdating = True
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
    SectionBar ws, r, cL, cR, "  OPENCAP EXPORT FILES", cTeal(), RGB(255, 255, 255)
    r = r + 1

    ' Column sub-headers
    ws.Rows(r).RowHeight = 14
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

    Dim tokens(4) As String, labels(4) As String, sheets(4) As String
    tokens(0) = TOK_JOB:       labels(0) = "Job Details":           sheets(0) = SH_JOB
    tokens(1) = TOK_CREW:      labels(1) = "Crew / Personnel":      sheets(1) = SH_CREW
    tokens(2) = TOK_BHA:       labels(2) = "BHA Equipment":         sheets(2) = SH_BHA
    tokens(3) = TOK_SLIDE:     labels(3) = "Slide / Rotate by Day": sheets(3) = SH_SLIDE
    tokens(4) = TOK_INVENTORY: labels(4) = "Equipment Inventory":   sheets(4) = SH_INVENTORY

    Dim fi As Long
    For fi = 0 To 4
        ws.Rows(r).RowHeight = 18

        Dim fPath As String: fPath = ""
        If wbPath <> "" Then fPath = FindCsvByToken(wbPath, tokens(fi))
        Dim found As Boolean: found = (fPath <> "")

        If found Then
            Sgl ws, r, cL, ChrW(10003), cGrnBadge(), RGB(255, 255, 255), True, 8, xlHAlignCenter
            MrgCell ws, r, cL + 1, cL + 2, labels(fi), cGrnRow(), cGrnTxt(), True, 7, xlHAlignLeft

            Dim fn As String: fn = Mid(fPath, InStrRev(fPath, Application.PathSeparator) + 1)
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

    With ws.Range(ws.Cells(panelTop, cL), ws.Cells(r - 1, cR))
        .Borders(xlEdgeLeft).LineStyle   = xlContinuous: .Borders(xlEdgeLeft).Color   = cMed()
        .Borders(xlEdgeRight).LineStyle  = xlContinuous: .Borders(xlEdgeRight).Color  = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = cMed()
    End With
End Sub

Private Sub DrawNotesPanel(ws As Worksheet, startRow As Long)
    Dim cL As Long: cL = C_CRL
    Dim cR As Long: cR = C_LAST
    Dim r As Long: r = startRow

    SectionBar ws, r, cL, cR, "  EOW / JOB NOTES", cDk(), RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 18
    r = r + 1

    Dim nr As Long
    For nr = r To r + 5: ws.Rows(nr).RowHeight = 16: Next nr
    With ws.Range(ws.Cells(r, cL), ws.Cells(r + 5, cR))
        .Merge
        .Interior.Color      = cBg()
        .Font.Color          = cDk()
        .Font.Size           = 9
        .Font.Name           = "Consolas"
        .VerticalAlignment   = xlVAlignTop
        .HorizontalAlignment = xlHAlignLeft
        .WrapText            = True
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
    Dim r As Long: r = startRow

    SectionBar ws, r, cL, cR, "  IMPORT FILES", cTealLt(), RGB(255, 255, 255)
    ws.Rows(r).RowHeight = 20
    r = r + 1

    ' Sub-header
    ws.Rows(r).RowHeight = 12
    Sgl ws, r, cL, "FILE", cBg(), cDk(), True, 7, xlHAlignLeft
    MrgCell ws, r, cL + 1, cR, "SELECTED PATH", cBg(), cDk(), True, 7, xlHAlignLeft
    With ws.Range(ws.Cells(r, cL), ws.Cells(r, cR)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = cMed(): .Weight = xlHairline
    End With
    r = r + 1

    Dim planPath As String: planPath = GetImportPath(SH_SURVEY)
    ws.Rows(r).RowHeight = 22
    Sgl ws, r, cL, "Import Plan", cBg(), cBlk(), False, 8, xlHAlignLeft
    MrgCell ws, r, cL + 1, cR, planPath, cWh(), cDk(), False, 7, xlHAlignLeft
    With ws.Range(ws.Cells(r, cL), ws.Cells(r, cR)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(230, 230, 230): .Weight = xlHairline
    End With
    On Error Resume Next: ws.Names("OC_ImpPlanRow").Delete: On Error GoTo 0
    ws.Names.Add Name:="OC_ImpPlanRow", RefersToR1C1:="=R" & r & "C" & cL
    r = r + 1

    Dim acPath As String: acPath = GetImportPath(SH_AC)
    ws.Rows(r).RowHeight = 22
    Sgl ws, r, cL, "Import AC", cBg(), cBlk(), False, 8, xlHAlignLeft
    MrgCell ws, r, cL + 1, cR, acPath, cWh(), cDk(), False, 7, xlHAlignLeft
    On Error Resume Next: ws.Names("OC_ImpAcRow").Delete: On Error GoTo 0
    ws.Names.Add Name:="OC_ImpAcRow", RefersToR1C1:="=R" & r & "C" & cL

    With ws.Range(ws.Cells(startRow, cL), ws.Cells(r, cR))
        .Borders(xlEdgeLeft).LineStyle   = xlContinuous: .Borders(xlEdgeLeft).Color   = cMed()
        .Borders(xlEdgeRight).LineStyle  = xlContinuous: .Borders(xlEdgeRight).Color  = cMed()
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
    bR.Name    = "BtnOCRefresh"
    bR.OnAction = "'" & ThisWorkbook.Name & "'!RefreshSetup"
    bR.Placement = xlMoveAndSize
    StyleBtn bR, "REFRESH CSVs", cBg(), cMed(), cBlk(), True

    ' REBUILD (a bit further left)
    Dim rb As Range: Set rb = ws.Cells(R_HDR, C_LAST - 7)
    Dim bB As Button
    Set bB = ws.Buttons.Add(rb.Left, rb.Top + 4, rb.Width * 3, rb.Height - 8)
    bB.Name    = "BtnOCRebuild"
    bB.OnAction = "'" & ThisWorkbook.Name & "'!RebuildSetup"
    bB.Placement = xlMoveAndSize
    StyleBtn bB, "REBUILD", cBg(), cMed(), cBlk(), True

    ' Import Plan button (overlays the first cell of the plan row)
    Dim planCell As Range
    On Error Resume Next: Set planCell = ws.Range("OC_ImpPlanRow"): On Error GoTo 0
    If Not planCell Is Nothing Then
        Dim bIP As Button
        Set bIP = ws.Buttons.Add(planCell.Left + 1, planCell.Top + 2, _
                                  planCell.Width * 1.9, planCell.Height - 4)
        bIP.Name    = "BtnImportPlan"
        bIP.OnAction = "'" & ThisWorkbook.Name & "'!ImportSurveyPlan"
        bIP.Placement = xlMoveAndSize
        StyleBtn bIP, "Import Plan", cBg(), cMed(), cBlk(), False
    End If

    ' Import AC button
    Dim acCell As Range
    On Error Resume Next: Set acCell = ws.Range("OC_ImpAcRow"): On Error GoTo 0
    If Not acCell Is Nothing Then
        Dim bAC As Button
        Set bAC = ws.Buttons.Add(acCell.Left + 1, acCell.Top + 2, _
                                  acCell.Width * 1.9, acCell.Height - 4)
        bAC.Name    = "BtnImportAC"
        bAC.OnAction = "'" & ThisWorkbook.Name & "'!ImportAntiCollision"
        bAC.Placement = xlMoveAndSize
        StyleBtn bAC, "Import AC", cBg(), cMed(), cBlk(), False
    End If
End Sub

Private Sub StyleBtn(btn As Button, cap As String, bg As Long, brd As Long, fg As Long, bold As Boolean)
    On Error Resume Next
    btn.Caption = cap
    btn.Font.Name  = "Consolas": btn.Font.Size = 9: btn.Font.Bold = bold: btn.Font.Color = fg
    btn.ShapeRange.Fill.ForeColor.RGB = bg
    btn.ShapeRange.Line.ForeColor.RGB = brd
    btn.ShapeRange.Line.Weight = 0.75
    On Error GoTo 0
End Sub

' ================================================================================
'  CELL / RANGE DRAWING HELPERS
' ================================================================================

Private Sub SectionBar(ws As Worksheet, r As Long, c1 As Long, c2 As Long, _
                        title As String, bg As Long, fg As Long)
    ws.Rows(r).RowHeight = 18
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Merge
        .Value              = title
        .Interior.Color     = bg
        .Font.Color         = fg
        .Font.Bold          = True
        .Font.Size          = 8
        .Font.Name          = "Consolas"
        .VerticalAlignment  = xlVAlignCenter
        .HorizontalAlignment = xlHAlignLeft
    End With
End Sub

' Pair row: left-panel label+value pair + right-panel label+value pair
Private Sub Pair(ws As Worksheet, r As Long, lbl1 As String, val1 As String, _
                                              lbl2 As String, val2 As String)
    ws.Rows(r).RowHeight = 17

    ' Accent strip
    ws.Cells(r, C_ACCENT).Interior.Color = RGB(218, 218, 218)

    ' Label 1 (B:C)
    MrgCell ws, r, C_L1S, C_L1E, lbl1, cBg(), cDk(), True, 7, xlHAlignRight

    ' Value 1 (D:E)
    With ws.Range(ws.Cells(r, C_V1S), ws.Cells(r, C_V1E))
        .Merge
        .NumberFormat       = "@"
        .Value              = val1
        .Interior.Color     = cWh()
        .Font.Color         = cBlk()
        .Font.Bold          = (val1 <> "")
        .Font.Size          = 9
        .Font.Name          = "Consolas"
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment  = xlVAlignCenter
        .IndentLevel        = 1
    End With

    ' Label 2 (F)
    Sgl ws, r, C_L2, lbl2, cBg(), cDk(), True, 7, xlHAlignRight

    ' Value 2 (G:H)
    MrgCell ws, r, C_V2S, C_V2E, val2, cWh(), cBlk(), (val2 <> ""), 9, xlHAlignLeft

    With ws.Range(ws.Cells(r, C_ACCENT), ws.Cells(r, C_V2E)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = RGB(232, 232, 232): .Weight = xlHairline
    End With
End Sub

Private Sub MrgCell(ws As Worksheet, r As Long, c1 As Long, c2 As Long, _
                    val As String, bg As Long, fg As Long, bold As Boolean, _
                    sz As Long, align As Long)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Merge
        .Value              = val
        .Interior.Color     = bg
        .Font.Color         = fg
        .Font.Bold          = bold
        .Font.Size          = sz
        .Font.Name          = "Consolas"
        .HorizontalAlignment = align
        .VerticalAlignment  = xlVAlignCenter
    End With
End Sub

Private Sub Sgl(ws As Worksheet, r As Long, c As Long, val As String, _
                bg As Long, fg As Long, bold As Boolean, sz As Long, align As Long)
    With ws.Cells(r, c)
        .Value              = val
        .Interior.Color     = bg
        .Font.Color         = fg
        .Font.Bold          = bold
        .Font.Size          = sz
        .Font.Name          = "Consolas"
        .HorizontalAlignment = align
        .VerticalAlignment  = xlVAlignCenter
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

Private Function ReadCrewRows() As String()
    Dim blankArr() As String
    ReDim blankArr(0, 3)
    ReadCrewRows = blankArr
    If Not SheetExists(SH_CREW) Then Exit Function

    Dim ws As Worksheet: Set ws = Worksheets(SH_CREW)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim cName  As Long: cName  = ColByName(ws, "Name", "FullName")
    Dim cRole  As Long: cRole  = ColByName(ws, "Role")
    Dim cWork  As Long: cWork  = ColByName(ws, "Work Type", "WorkType", "Position")
    Dim cEmail As Long: cEmail = ColByName(ws, "Email")
    Dim cPhone As Long: cPhone = ColByName(ws, "Phone", "Mobile", "MobilePhone")
    If cName = 0 Then cName = 2

    Dim buf(200, 3) As String
    Dim n As Long: n = 0
    Dim r As Long
    For r = 2 To lastRow
        Dim nm As String:  nm  = Trim(SafeStr(ws.Cells(r, cName)))
        Dim rl As String:  rl  = ""
        Dim em As String:  em  = ""
        Dim ph As String:  ph  = ""

        If cRole > 0 Then rl = Trim(SafeStr(ws.Cells(r, cRole)))
        If rl = "" And cWork > 0 Then rl = Trim(SafeStr(ws.Cells(r, cWork)))
        If cEmail > 0 Then em = Trim(SafeStr(ws.Cells(r, cEmail)))
        If cPhone > 0 Then ph = Trim(SafeStr(ws.Cells(r, cPhone)))

        If nm = "" Then GoTo SkipRow
        buf(n, 0) = rl
        buf(n, 1) = nm
        buf(n, 2) = em
        buf(n, 3) = ph
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

Private Function FindCsvByToken(folder As String, token As String) As String
    FindCsvByToken = ""
    Dim newest As Date: newest = 0
    Dim fn As String: fn = Dir(folder & Application.PathSeparator & "*.csv")
    Do While fn <> ""
        If InStr(LCase(fn), LCase(token)) > 0 Then
            Dim fp As String: fp = folder & Application.PathSeparator & fn
            On Error Resume Next
            Dim st As Date: st = FileDateTime(fp)
            If Err.Number = 0 And st >= newest Then newest = st: FindCsvByToken = fp
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
        Dim ch As String: ch = Mid(line, pos, 1)
        If ch = Chr(34) Then
            If inQ And Mid(line, pos + 1, 1) = Chr(34) Then
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
    Dim mo As String: mo = Mid(raw, 5, 2)
    Dim d As String: d = Mid(raw, 7, 2)
    If Not IsNumeric(y) Or Not IsNumeric(mo) Or Not IsNumeric(d) Then Exit Function
    FmtDate = y & "-" & mo & "-" & d
    On Error GoTo 0
End Function

' ================================================================================
'  UTILITY
' ================================================================================

Private Function SheetExists(nm As String) As Boolean
    On Error Resume Next: SheetExists = Not (ThisWorkbook.Sheets(nm) Is Nothing): On Error GoTo 0
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
    fd.Title = "Select Survey Plan CSV"
    fd.Filters.Clear
    fd.Filters.Add "CSV Files", "*.csv", 1
    fd.Filters.Add "All Files", "*.*", 2
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub

    Dim fPath As String: fPath = fd.SelectedItems(1)
    Application.StatusBar = "Importing survey plan..."

    Dim survWs As Worksheet
    If SheetExists(SH_SURVEY) Then
        Set survWs = Worksheets(SH_SURVEY)
        survWs.Cells.Clear
    Else
        Set survWs = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        survWs.Name = SH_SURVEY
        survWs.Visible = xlSheetVeryHidden
    End If

    survWs.Cells(1, 1).Value = fPath

    Dim fNum As Integer: fNum = FreeFile
    On Error GoTo ImportPlanErr
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

    UpdateImportPathDisplay SH_SURVEY
    Application.StatusBar = "Survey imported: " & (outRow - 3) & " stations from " & _
                             Mid(fPath, InStrRev(fPath, Application.PathSeparator) + 1)
    Exit Sub
ImportPlanErr:
    On Error Resume Next: Close #fNum: On Error GoTo 0
    Application.StatusBar = "Survey import failed: " & Err.Description
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
    fd.Title = "Select Anti-Collision PDF"
    fd.Filters.Clear
    fd.Filters.Add "PDF Files", "*.pdf", 1
    fd.Filters.Add "All Files", "*.*", 2
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Sub

    Dim fPath As String: fPath = fd.SelectedItems(1)
    Application.StatusBar = "Reading anti-collision report..."

    Dim acWs As Worksheet
    If SheetExists(SH_AC) Then
        Set acWs = Worksheets(SH_AC)
        acWs.Cells.Clear
    Else
        Set acWs = ThisWorkbook.Sheets.Add( _
            After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        acWs.Name = SH_AC
        acWs.Visible = xlSheetVeryHidden
    End If
    acWs.Cells(1, 1).Value = fPath

    Dim pdfText As String: pdfText = ExtractPdfText(fPath)

    If pdfText = "" Then
        Application.StatusBar = "PDF extraction unavailable. Install Poppler or Adobe Acrobat."
        MsgBox "Could not read the PDF automatically." & Chr(10) & Chr(10) & _
               "To enable automatic extraction, install one of:" & Chr(10) & _
               "  - Poppler for Windows (adds pdftotext.exe to PATH)" & Chr(10) & _
               "    https://github.com/oschwartz10612/poppler-windows/releases" & Chr(10) & Chr(10) & _
               "  - Adobe Acrobat (full version, not Reader)" & Chr(10) & Chr(10) & _
               "Path stored. Try again after installing.", vbExclamation, "PDF Reader Not Available"
        UpdateImportPathDisplay SH_AC
        Exit Sub
    End If

    ' Parse for SF < 2.0 rows from Summary table
    Dim nHits As Long
    Dim aRefMD(200) As Double, aBetween(200) As Double, aSF(200) As Double
    nHits = ParseAcSummary(pdfText, aRefMD, aBetween, aSF)

    ' Write to hidden sheet
    acWs.Cells(2, 1).Value = "Ref MD (m)"
    acWs.Cells(2, 2).Value = "Between Centres (m)"
    acWs.Cells(2, 3).Value = "Separation Factor"
    Dim i As Long
    For i = 0 To nHits - 1
        acWs.Cells(i + 3, 1).Value = aRefMD(i)
        acWs.Cells(i + 3, 2).Value = aBetween(i)
        acWs.Cells(i + 3, 3).Value = aSF(i)
    Next i

    ' Render formatted table in Sheet1
    BuildAcTable nHits, aRefMD, aBetween, aSF, fPath

    UpdateImportPathDisplay SH_AC
    Application.StatusBar = "AC import complete: " & nHits & " critical separation(s) found (SF < 2.0)."
End Sub

' ---- PDF text extraction (two strategies) ----
Private Function ExtractPdfText(pdfPath As String) As String
    ExtractPdfText = ""

    ' Strategy 1: pdftotext (Poppler) via Shell
    Dim tmpFile As String: tmpFile = Environ("TEMP") & "\oc_pdf_extract.txt"
    On Error Resume Next: Kill tmpFile: On Error GoTo 0

    Shell "cmd /c pdftotext -layout """ & pdfPath & """ """ & tmpFile & """", vbHide

    Dim deadline As Date: deadline = Now + TimeValue("00:00:06")
    Do While Dir(tmpFile) = "" And Now < deadline
        Application.Wait Now + 500
    Loop

    If Dir(tmpFile) <> "" Then
        Dim fNum As Integer: fNum = FreeFile
        On Error GoTo TryAcrobat
        Open tmpFile For Input As #fNum
        Dim content As String, ln As String
        Do While Not EOF(fNum): Line Input #fNum, ln: content = content & ln & Chr(10): Loop
        Close #fNum
        On Error Resume Next: Kill tmpFile: On Error GoTo 0
        ExtractPdfText = content
        Exit Function
    End If

TryAcrobat:
    On Error Resume Next: Close fNum: On Error GoTo 0

    ' Strategy 2: Adobe Acrobat COM
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
        Dim pg As Long, allText As String
        For pg = 0 To pdDoc.GetNumPages - 1
            On Error Resume Next
            allText = allText & jsObj.getPageNthWord(pg, 0, True) & Chr(10)
            On Error GoTo 0
        Next pg
        ExtractPdfText = allText
    End If

    pdDoc.Close False
End Function

' ---- Parse PDF text for Summary table rows with SF < 2.0 ----
Private Function ParseAcSummary(pdfText As String, _
        aRefMD() As Double, aBetween() As Double, aSF() As Double) As Long
    ParseAcSummary = 0
    Dim lines() As String: lines = Split(pdfText, Chr(10))
    Dim inSummary As Boolean: inSummary = False
    Dim nHits As Long: nHits = 0
    Dim i As Long

    For i = 0 To UBound(lines)
        Dim ln As String: ln = Trim(lines(i))
        If InStr(1, ln, "Summary", vbTextCompare) > 0 Then inSummary = True
        If Not inSummary Then GoTo NextAcLine

        ' Extract numeric tokens from the line
        Dim nums() As Double: Dim nNums As Long
        nNums = ExtractNums(ln, nums)
        ' Expect at least 3 numbers: RefMD, BetweenCentres, SepFactor (SF is typically 3rd)
        If nNums >= 3 Then
            Dim sf As Double: sf = nums(2)
            If sf > 0 And sf < 2# Then
                aRefMD(nHits)   = nums(0)
                aBetween(nHits) = nums(1)
                aSF(nHits)      = sf
                nHits = nHits + 1
                If nHits > 200 Then Exit For
            End If
        End If
NextAcLine:
    Next i
    ParseAcSummary = nHits
End Function

Private Function ExtractNums(s As String, nums() As Double) As Long
    ReDim nums(50)
    Dim parts() As String: parts = Split(s, " ")
    Dim n As Long: n = 0
    Dim p As Variant
    For Each p In parts
        Dim t As String: t = Trim(Replace(CStr(p), ",", ""))
        If IsNumeric(t) And Len(t) > 0 Then
            nums(n) = CDbl(t): n = n + 1
            If n > 50 Then Exit For
        End If
    Next p
    ExtractNums = n
End Function

' ---- Build formatted AC results table in Sheet1 ----
Public Sub BuildAcTable(nHits As Long, aRefMD() As Double, _
        aBetween() As Double, aSF() As Double, srcPath As String)

    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("Sheet1"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Tab.Color = cTeal()

    Dim r As Long: r = 1

    ' Title bar
    ws.Rows(r).RowHeight = 28
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Merge
        .Value = "  ANTI-COLLISION SUMMARY  |  Separation Factor < 2.0"
        .Interior.Color = cTeal(): .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True: .Font.Name = "Consolas": .Font.Size = 11
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    ' Source path
    ws.Rows(r).RowHeight = 14
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Merge
        .Value = "  Source: " & srcPath
        .Interior.Color = cBg(): .Font.Color = cDk()
        .Font.Name = "Consolas": .Font.Size = 7
        .VerticalAlignment = xlVAlignCenter
    End With
    r = r + 1

    ' Column widths
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 18
    ws.Columns(3).ColumnWidth = 22
    ws.Columns(4).ColumnWidth = 18
    ws.Columns(5).ColumnWidth = 12

    ' Column headers
    ws.Rows(r).RowHeight = 18
    Dim hdrs(4) As String
    hdrs(0) = "#"
    hdrs(1) = "Ref MD (m)"
    hdrs(2) = "Between Centres (m)"
    hdrs(3) = "Separation Factor"
    hdrs(4) = "Risk"
    Dim c As Long
    For c = 0 To 4
        With ws.Cells(r, c + 1)
            .Value = hdrs(c)
            .Interior.Color = cDk(): .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True: .Font.Name = "Consolas": .Font.Size = 8
            .HorizontalAlignment = IIf(c = 0 Or c = 4, xlHAlignCenter, xlHAlignRight)
            .VerticalAlignment = xlVAlignCenter
        End With
    Next c
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = cMed(): .Weight = xlThin
    End With
    r = r + 1

    If nHits = 0 Then
        ws.Rows(r).RowHeight = 20
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
            .Merge
            .Value = "  No critical separations found  (all SF >= 2.0)"
            .Interior.Color = cGrnRow(): .Font.Color = cGrnTxt()
            .Font.Name = "Consolas": .Font.Size = 9
            .VerticalAlignment = xlVAlignCenter
        End With
        r = r + 1
    Else
        Dim i As Long
        For i = 0 To nHits - 1
            ws.Rows(r).RowHeight = 18
            Dim sf As Double: sf = aSF(i)
            Dim rowBg As Long: rowBg = IIf(sf < 1.5, RGB(255, 235, 235), RGB(255, 253, 235))
            Dim rowFg As Long: rowFg = IIf(sf < 1.5, cRedTxt(), RGB(140, 90, 0))
            Dim risk As String
            Select Case True
                Case sf < 1.0: risk = "CRITICAL"
                Case sf < 1.5: risk = "HIGH"
                Case Else:     risk = "CAUTION"
            End Select

            Dim vals(4) As Variant
            vals(0) = i + 1: vals(1) = aRefMD(i)
            vals(2) = aBetween(i): vals(3) = sf: vals(4) = risk

            For c = 0 To 4
                With ws.Cells(r, c + 1)
                    .Value = vals(c)
                    .Interior.Color = rowBg: .Font.Color = rowFg
                    .Font.Name = "Consolas": .Font.Size = 9
                    .HorizontalAlignment = IIf(c = 0 Or c = 4, xlHAlignCenter, xlHAlignRight)
                    .VerticalAlignment = xlVAlignCenter
                    If c = 3 Then .NumberFormat = "0.000"
                    If c = 1 Or c = 2 Then .NumberFormat = "0.00"
                End With
            Next c
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Color = RGB(220, 220, 220): .Weight = xlHairline
            End With
            r = r + 1
        Next i
    End If

    ' Outer border
    With ws.Range(ws.Cells(3, 1), ws.Cells(r - 1, 5))
        .Borders(xlEdgeLeft).LineStyle   = xlContinuous: .Borders(xlEdgeLeft).Color   = cMed()
        .Borders(xlEdgeRight).LineStyle  = xlContinuous: .Borders(xlEdgeRight).Color  = cMed()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = cMed()
        .Borders(xlEdgeTop).LineStyle    = xlContinuous: .Borders(xlEdgeTop).Color    = cMed()
    End With

    ws.Activate
    ws.Cells(1, 1).Select
End Sub

' Refresh path cells on the Setup sheet after an import
Private Sub UpdateImportPathDisplay(shName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_SETUP)
    If ws Is Nothing Then Exit Sub
    Dim nm As String: nm = IIf(shName = SH_SURVEY, "OC_ImpPlanRow", "OC_ImpAcRow")
    Dim anchor As Range: Set anchor = ws.Range(nm)
    If Not anchor Is Nothing Then
        anchor.Offset(0, 1).MergeArea.Cells(1, 1).Value = GetImportPath(shName)
    End If
    On Error GoTo 0
End Sub

' ================================================================================
'  AC DEMO TABLE  (Sheet1 preview with sample data)
'  Call this once after RebuildSetup to show what the AC table looks like.
' ================================================================================
Public Sub DemoAcTable()
    Dim nHits As Long: nHits = 5
    Dim aR(4) As Double, aB(4) As Double, aS(4) As Double
    ' Sample critical-separation data representative of a P3 AC check
    aR(0) = 597.00:  aB(0) = 18.4:  aS(0) = 1.82
    aR(1) = 1203.50: aB(1) = 11.2:  aS(1) = 1.44
    aR(2) = 2664.50: aB(2) = 8.7:   aS(2) = 0.97
    aR(3) = 3812.00: aB(3) = 14.1:  aS(3) = 1.63
    aR(4) = 5490.25: aB(4) = 9.3:   aS(4) = 1.18
    BuildAcTable nHits, aR, aB, aS, _
        "D:\Tourmaline Sundown\34784 I Well\Well Plans\TOURMALINE HZ SUNDOWN I04-04-077-17 P3 AC.pdf (DEMO)"
End Sub
