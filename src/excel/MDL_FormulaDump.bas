Attribute VB_Name = "MDL_FormulaDump"
Option Explicit

' ================================================================================
'  MODULE: MDL_FormulaDump
'  OpenCap Field Workbook - Sheet Formula & Structure Exporter
'  v1.0.0
'
'  PURPOSE:
'   Exports every cell formula, named range, and structural note from a chosen
'   worksheet to a plain-text file so the sheet can be fully analysed or
'   reconstructed by an external tool or AI assistant.
'
'  QUICK START:
'   1. Alt+F11 > Insert > Module > Import this file
'   2. Immediate window:  DumpSheetFormulas
'      -- or --  DumpSheetFormulas "Slides"
'   3. Output file is written next to the workbook as:
'        <WorkbookName>_<SheetName>_FormulaDump.txt
'
'  OUTPUT SECTIONS:
'   [WORKBOOK]      - file name, dump timestamp
'   [SHEET]         - sheet name, used range, visibility
'   [NAMED RANGES]  - all defined names that reference this sheet
'   [MERGED CELLS]  - all merge areas on the sheet
'   [CELLS]         - every non-empty cell: address | type | formula/value | fmt
'   [DEPENDENCIES]  - cross-sheet references detected in formulas
' ================================================================================

Private Const DUMP_SECTION_SEP As String = "--------------------------------------------------------------------------------"
Private Const DUMP_HEADER_SEP  As String = "================================================================================"

' -----------------------------------------------------------------------
'  Public entry point
'  Call with no args to pick sheet interactively, or pass a sheet name.
' -----------------------------------------------------------------------
Public Sub DumpSheetFormulas(Optional ByVal sheetName As String = "")
    Dim ws      As Worksheet
    Dim outPath As String
    Dim fNum    As Integer

    ' -- Resolve the target sheet ------------------------------------------
    If sheetName = "" Then
        Dim pick As String
        pick = InputBox("Enter sheet name to dump (blank = active sheet):", _
                        "Formula Dump", ActiveSheet.Name)
        If pick = "" Then pick = ActiveSheet.Name
        sheetName = pick
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet not found: " & sheetName, vbExclamation, "Formula Dump"
        Exit Sub
    End If

    ' -- Build output path -------------------------------------------------
    Dim wbDir As String
    wbDir   = Left(ThisWorkbook.FullName, _
                   Len(ThisWorkbook.FullName) - Len(ThisWorkbook.Name))
    outPath = wbDir & _
              Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & _
              "_" & CleanFileName(ws.Name) & "_FormulaDump.txt"

    ' -- Open file ---------------------------------------------------------
    fNum = FreeFile
    Open outPath For Output As #fNum

    ' -- Write sections ----------------------------------------------------
    WriteHeader fNum, ws
    WriteNamedRanges fNum, ws
    WriteMergedCells fNum, ws
    WriteCells fNum, ws
    WriteDependencies fNum, ws

    Print #fNum, ""
    Print #fNum, DUMP_HEADER_SEP
    Print #fNum, "END OF DUMP"
    Print #fNum, DUMP_HEADER_SEP

    Close #fNum

    MsgBox "Done!  Output written to:" & vbLf & outPath, vbInformation, "Formula Dump"
    Shell "notepad.exe """ & outPath & """"
End Sub

' -----------------------------------------------------------------------
'  Section: workbook + sheet header
' -----------------------------------------------------------------------
Private Sub WriteHeader(fNum As Integer, ws As Worksheet)
    Print #fNum, DUMP_HEADER_SEP
    Print #fNum, "FORMULA DUMP  -  OpenCap Field Workbook"
    Print #fNum, DUMP_HEADER_SEP
    Print #fNum, ""
    Print #fNum, "[WORKBOOK]"
    Print #fNum, "  File      : " & ThisWorkbook.Name
    Print #fNum, "  Full path : " & ThisWorkbook.FullName
    Print #fNum, "  Dumped at : " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    Print #fNum, ""
    Print #fNum, "[SHEET]"
    Print #fNum, "  Name       : " & ws.Name
    Print #fNum, "  Code name  : " & ws.CodeName
    Print #fNum, "  Visibility : " & SheetVisibility(ws.Visible)
    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If Not ur Is Nothing Then
        Print #fNum, "  Used range : " & ur.Address(False, False)
        Print #fNum, "  Rows used  : " & ur.Rows.Count
        Print #fNum, "  Cols used  : " & ur.Columns.Count
    Else
        Print #fNum, "  Used range : (empty)"
    End If
    Print #fNum, ""
End Sub

' -----------------------------------------------------------------------
'  Section: named ranges that reference this sheet
' -----------------------------------------------------------------------
Private Sub WriteNamedRanges(fNum As Integer, ws As Worksheet)
    Print #fNum, "[NAMED RANGES]"
    Print #fNum, "  (names whose RefersTo contains this sheet)"
    Print #fNum, ""

    Dim nm      As Name
    Dim found   As Boolean
    found = False

    For Each nm In ThisWorkbook.Names
        Dim ref As String
        ref = ""
        On Error Resume Next
        ref = nm.RefersTo
        On Error GoTo 0
        If InStr(1, ref, ws.Name, vbTextCompare) > 0 Then
            Print #fNum, "  " & PadR(nm.Name, 30) & " = " & ref
            found = True
        End If
    Next nm

    If Not found Then Print #fNum, "  (none)"
    Print #fNum, ""
    Print #fNum, DUMP_SECTION_SEP
    Print #fNum, ""
End Sub

' -----------------------------------------------------------------------
'  Section: merged cell regions
' -----------------------------------------------------------------------
Private Sub WriteMergedCells(fNum As Integer, ws As Worksheet)
    Print #fNum, "[MERGED CELLS]"
    Print #fNum, ""

    Dim mergeList() As String
    Dim mergeCount  As Long
    mergeCount = 0
    ReDim mergeList(0)

    Dim cell As Range
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            Dim ma As String
            ma = cell.MergeArea.Address(False, False)
            ' deduplicate
            Dim already As Boolean
            already = False
            Dim i As Long
            For i = 0 To mergeCount - 1
                If mergeList(i) = ma Then already = True: Exit For
            Next i
            If Not already Then
                ReDim Preserve mergeList(mergeCount)
                mergeList(mergeCount) = ma
                mergeCount = mergeCount + 1
                Print #fNum, "  " & ma & "  (" & cell.MergeArea.Rows.Count & _
                              "r x " & cell.MergeArea.Columns.Count & "c)"
            End If
        End If
    Next cell

    If mergeCount = 0 Then Print #fNum, "  (none)"
    Print #fNum, ""
    Print #fNum, DUMP_SECTION_SEP
    Print #fNum, ""
End Sub

' -----------------------------------------------------------------------
'  Section: cell-by-cell dump
'  Columns: Address | DataType | Formula-or-Value | NumberFormat
' -----------------------------------------------------------------------
Private Sub WriteCells(fNum As Integer, ws As Worksheet)
    Print #fNum, "[CELLS]"
    Print #fNum, ""
    Print #fNum, "  " & PadR("Address", 10) & _
                 PadR("Type", 9) & _
                 PadR("Formula / Value", 60) & _
                 "NumberFormat"
    Print #fNum, "  " & String(10, "-") & " " & _
                 String(8, "-") & " " & _
                 String(59, "-") & " " & _
                 String(30, "-")

    Dim cell    As Range
    Dim ur      As Range

    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then
        Print #fNum, "  (sheet is empty)"
        GoTo Done
    End If

    For Each cell In ur
        ' Skip truly blank cells (no formula, no value, not part of a merge anchor)
        If cell.HasFormula = False And Len(Trim(CStr(cell.Value))) = 0 Then
            If Not cell.MergeCells Then GoTo NextCell
            If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then GoTo NextCell
        End If

        Dim cellType    As String
        Dim cellContent As String

        If cell.HasFormula Then
            cellType    = "FORMULA"
            cellContent = cell.Formula
        ElseIf cell.MergeCells Then
            cellType    = "MERGE"
            cellContent = CStr(cell.MergeArea.Cells(1, 1).Value)
        Else
            Select Case cell.Value
                Case Is = ""
                    GoTo NextCell
                Case Else
                    cellType    = VBATypeName(cell)
                    cellContent = CStr(cell.Value)
            End Select
        End If

        Print #fNum, "  " & PadR(cell.Address(False, False), 10) & _
                     PadR(cellType, 9) & _
                     PadR(Left(cellContent, 59), 60) & _
                     cell.NumberFormat

NextCell:
    Next cell

Done:
    Print #fNum, ""
    Print #fNum, DUMP_SECTION_SEP
    Print #fNum, ""
End Sub

' -----------------------------------------------------------------------
'  Section: cross-sheet dependencies found in formulas
' -----------------------------------------------------------------------
Private Sub WriteDependencies(fNum As Integer, ws As Worksheet)
    Print #fNum, "[DEPENDENCIES]"
    Print #fNum, "  (other sheet names referenced in formulas on this sheet)"
    Print #fNum, ""

    Dim deps    As New Collection
    Dim cell    As Range
    Dim ur      As Range

    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then
        Print #fNum, "  (sheet is empty)"
        GoTo Done
    End If

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.Name <> ws.Name Then
            For Each cell In ur
                If cell.HasFormula Then
                    If InStr(1, cell.Formula, "'" & sh.Name & "'", vbTextCompare) > 0 Or _
                       InStr(1, cell.Formula, sh.Name & "!", vbTextCompare) > 0 Then
                        ' Check not already in deps
                        Dim alreadyDep As Boolean
                        alreadyDep = False
                        Dim d As Variant
                        For Each d In deps
                            If d = sh.Name Then alreadyDep = True: Exit For
                        Next d
                        If Not alreadyDep Then deps.Add sh.Name
                        Exit For
                    End If
                End If
            Next cell
        End If
    Next sh

    If deps.Count = 0 Then
        Print #fNum, "  (no cross-sheet references detected)"
    Else
        For Each d In deps
            Print #fNum, "  -> " & d
        Next d
    End If

Done:
    Print #fNum, ""
End Sub

' -----------------------------------------------------------------------
'  Helpers
' -----------------------------------------------------------------------
Private Function PadR(s As String, w As Long) As String
    PadR = Left(s & Space(w), w) & " "
End Function

Private Function CleanFileName(s As String) As String
    Dim result As String
    result = s
    Dim bad As String, i As Long
    bad = "\/:*?""<>|"
    For i = 1 To Len(bad)
        result = Join(Split(result, Mid(bad, i, 1)), "_")
    Next i
    CleanFileName = result
End Function

Private Function SheetVisibility(v As Long) As String
    Select Case v
        Case xlSheetVisible:     SheetVisibility = "Visible"
        Case xlSheetHidden:      SheetVisibility = "Hidden"
        Case xlSheetVeryHidden:  SheetVisibility = "VeryHidden"
        Case Else:               SheetVisibility = "Unknown(" & v & ")"
    End Select
End Function

Private Function VBATypeName(cell As Range) As String
    Select Case cell.Value
        Case Is = ""
            VBATypeName = "EMPTY"
        Case Else
            If IsNumeric(cell.Value) Then
                If InStr(cell.NumberFormat, "%") > 0 Then
                    VBATypeName = "PERCENT"
                ElseIf InStr(cell.NumberFormat, "d") > 0 Or _
                       InStr(cell.NumberFormat, "m") > 0 Or _
                       InStr(cell.NumberFormat, "y") > 0 Then
                    VBATypeName = "DATE"
                Else
                    VBATypeName = "NUMBER"
                End If
            ElseIf cell.Value = True Or cell.Value = False Then
                VBATypeName = "BOOL"
            Else
                VBATypeName = "TEXT"
            End If
    End Select
End Function
