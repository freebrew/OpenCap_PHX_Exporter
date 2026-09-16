Attribute VB_Name = "MDL_AcPoc_Main"
Option Explicit

Public gAcPocQuiet As Boolean
Public gAcPocFailCount As Long
Public gAcPocWarnCount As Long

Private Const POC_SHEET As String = "POC"
Private Const LOG_SHEET As String = "Log"

Public Sub AcPoc_EnsureUi()
    Dim ws As Worksheet
    Dim logWs As Worksheet
    Set ws = AcPoc_EnsureSheet(POC_SHEET, False)
    Set logWs = AcPoc_EnsureSheet(LOG_SHEET, False)
    AcPoc_EnsureSheet "Corridor", False
    AcPoc_EnsureSheet "Charts", False
    AcPoc_EnsureSheet "_POC_Plan", True
    AcPoc_EnsureSheet "_POC_PlanSec", True
    AcPoc_EnsureSheet "_POC_AC_Summary", True
    AcPoc_EnsureSheet "_POC_AC_Detail", True
    AcPoc_EnsureSheet "_POC_Qualify", True

    If Len(Trim$(CStr(ws.Range("A1").Value2 & ""))) = 0 Then
        ws.Range("A1").Value2 = "AC Corridor Proof of Concept"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 16
        ws.Range("A2").Value2 = "Plan PDF"
        ws.Range("A3").Value2 = "AC PDF"
        ws.Range("A4").Value2 = "Capacity"
        ws.Range("A5").Value2 = "Output folder"
        ws.Range("B4").Value2 = 8
        ws.Range("A7").Value2 = "Standalone VBA demo. Does not write Slide Sheet Demo or Field."
        ws.Columns("A").ColumnWidth = 16
        ws.Columns("B").ColumnWidth = 90
        AcPoc_AddButton ws, "btnPickPlan", "Pick Plan PDF", 12, 80, "AcPoc_PickPlan"
        AcPoc_AddButton ws, "btnPickAc", "Pick AC PDF", 140, 80, "AcPoc_PickAC"
        AcPoc_AddButton ws, "btnRun", "Run All", 268, 80, "AcPoc_RunAll"
    End If
    If Len(Trim$(CStr(logWs.Range("A1").Value2 & ""))) = 0 Then
        logWs.Range("A1").Value2 = "Step"
        logWs.Range("B1").Value2 = "Status"
        logWs.Range("C1").Value2 = "Message"
        logWs.Range("D1").Value2 = "Timestamp"
        logWs.Range("A1:D1").Font.Bold = True
        logWs.Columns("A").ColumnWidth = 22
        logWs.Columns("B").ColumnWidth = 10
        logWs.Columns("C").ColumnWidth = 110
        logWs.Columns("D").ColumnWidth = 20
    End If
End Sub

Private Sub AcPoc_AddButton(ByVal ws As Worksheet, ByVal nm As String, _
        ByVal cap As String, ByVal leftPt As Double, ByVal topPt As Double, ByVal action As String)
    Dim btn As Button
    On Error Resume Next
    ws.Buttons(nm).Delete
    On Error GoTo 0
    Set btn = ws.Buttons.Add(leftPt, topPt, 120, 22)
    btn.name = nm
    btn.caption = cap
    btn.OnAction = action
End Sub

Public Function AcPoc_PlanPath() As String
    AcPoc_PlanPath = Trim$(CStr(AcPoc_EnsureSheet(POC_SHEET).Range("B2").Value2 & ""))
End Function

Public Function AcPoc_AcPath() As String
    AcPoc_AcPath = Trim$(CStr(AcPoc_EnsureSheet(POC_SHEET).Range("B3").Value2 & ""))
End Function

Public Function AcPoc_Capacity() As Long
    Dim v As Variant
    v = AcPoc_EnsureSheet(POC_SHEET).Range("B4").Value2
    If IsNumeric(v) Then
        AcPoc_Capacity = CLng(v)
    Else
        AcPoc_Capacity = 8
    End If
    If AcPoc_Capacity < 1 Then AcPoc_Capacity = 8
End Function

Public Function AcPoc_OutDir() As String
    Dim p As String
    p = Trim$(CStr(AcPoc_EnsureSheet(POC_SHEET).Range("B5").Value2 & ""))
    If Len(p) = 0 Then
        p = ThisWorkbook.Path
        If Right$(p, 1) <> "\" Then p = p & "\"
        AcPoc_EnsureSheet(POC_SHEET).Range("B5").Value2 = p
    Else
        If Right$(p, 1) <> "\" Then p = p & "\"
    End If
    AcPoc_OutDir = p
End Function

Public Sub AcPoc_SetQuiet(Optional ByVal quiet As Boolean = True)
    gAcPocQuiet = quiet
End Sub

Public Sub AcPoc_PickPlan()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select well-plan PDF"
    fd.Filters.Clear
    fd.Filters.Add "PDF", "*.pdf"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        AcPoc_EnsureUi
        ThisWorkbook.Worksheets(POC_SHEET).Range("B2").Value2 = fd.SelectedItems(1)
    End If
End Sub

Public Sub AcPoc_PickAC()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.title = "Select anticollision PDF"
    fd.Filters.Clear
    fd.Filters.Add "PDF", "*.pdf"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        AcPoc_EnsureUi
        ThisWorkbook.Worksheets(POC_SHEET).Range("B3").Value2 = fd.SelectedItems(1)
    End If
End Sub

Public Sub AcPoc_Log(ByVal stepName As String, ByVal status As String, ByVal msg As String)
    Dim ws As Worksheet
    Dim r As Long
    Set ws = AcPoc_EnsureSheet(LOG_SHEET, False)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value2 = stepName
    ws.Cells(r, 2).Value2 = status
    ws.Cells(r, 3).Value2 = msg
    ws.Cells(r, 4).Value2 = Format$(Now, "yyyy-mm-dd HH:nn:ss")
    If StrComp(status, "FAIL", vbTextCompare) = 0 Then gAcPocFailCount = gAcPocFailCount + 1
    If StrComp(status, "WARN", vbTextCompare) = 0 Then gAcPocWarnCount = gAcPocWarnCount + 1
End Sub

Public Sub AcPoc_StepHeader(ByVal idx As Long, ByVal total As Long, ByVal title As String)
    Application.StatusBar = "[STEP " & idx & "/" & total & "] " & title
    AcPoc_Log title, "START", "[STEP " & idx & "/" & total & "] " & title
End Sub

Public Sub AcPoc_StepDone(ByVal idx As Long, ByVal total As Long, ByVal title As String, ByVal extra As String)
    AcPoc_Log title, "OK", "[STEP " & idx & "/" & total & "] COMPLETE - " & extra
    Application.StatusBar = "[STEP " & idx & "/" & total & "] COMPLETE"
End Sub

Public Sub AcPoc_RunAll()
    Dim t0 As Double
    On Error GoTo Fail
    t0 = Timer
    gAcPocFailCount = 0
    gAcPocWarnCount = 0
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    AcPoc_EnsureUi
    AcPoc_EnsureSheet(LOG_SHEET).Range("A2:D5000").ClearContents

    AcPoc_StepHeader 1, 6, "Validate"
    AcPoc_Validate
    If gAcPocFailCount > 0 Then GoTo Done
    AcPoc_StepDone 1, 6, "Validate", "paths + pdftotext OK"

    AcPoc_StepHeader 2, 6, "Parse plan"
    AcPoc_ParsePlan
    If gAcPocFailCount > 0 Then GoTo Done
    AcPoc_StepDone 2, 6, "Parse plan", "stations + plan sections written"

    AcPoc_StepHeader 3, 6, "Parse AC summary"
    AcPoc_ParseSummary
    If gAcPocFailCount > 0 Then GoTo Done
    AcPoc_StepDone 3, 6, "Parse AC summary", "summary rows written"

    AcPoc_StepHeader 4, 6, "Parse AC detail"
    AcPoc_ParseDetail
    If gAcPocFailCount > 0 Then GoTo Done
    AcPoc_StepDone 4, 6, "Parse AC detail", "detail stations written"

    AcPoc_StepHeader 5, 6, "Qualify + self-check"
    AcPoc_Qualify
    AcPoc_SelfCheck
    If gAcPocFailCount > 0 Then GoTo Done
    AcPoc_StepDone 5, 6, "Qualify + self-check", "table matches Data-tab rule; frame proof OK"

    AcPoc_StepHeader 6, 6, "Render"
    AcPoc_Render
    If gAcPocFailCount > 0 Then GoTo Done
    AcPoc_StepDone 6, 6, "Render", "PNGs + Charts sheet"

Done:
    Application.ScreenUpdating = True
    If gAcPocFailCount = 0 Then
        AcPoc_Log "RUN", "PASS", "RUN COMPLETE | FAIL=0 WARN=" & gAcPocWarnCount & _
            " | " & Format$(Timer - t0, "0.0") & " s"
        Application.StatusBar = "AC POC COMPLETE"
        If Not gAcPocQuiet Then
            MsgBox "AC corridor POC finished with 0 FAIL." & vbCrLf & _
                   "PNGs in " & AcPoc_OutDir(), vbInformation, "AC POC"
        End If
    Else
        AcPoc_Log "RUN", "FAIL", "RUN FAILED | FAIL=" & gAcPocFailCount & " WARN=" & gAcPocWarnCount
        Application.StatusBar = "AC POC FAILED"
        If Not gAcPocQuiet Then
            MsgBox "AC corridor POC failed. See the Log sheet.", vbExclamation, "AC POC"
        End If
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    AcPoc_Log "RUN", "FAIL", "Error " & Err.Number & ": " & Err.Description
    Application.StatusBar = False
    If Not gAcPocQuiet Then MsgBox "AC POC error: " & Err.Description, vbCritical, "AC POC"
End Sub

Private Sub AcPoc_Validate()
    Dim planP As String, acP As String, outP As String
    Dim Sh As Object
    Dim p2t As String
    Dim fso As Object

    planP = AcPoc_PlanPath()
    acP = AcPoc_AcPath()
    outP = AcPoc_OutDir()

    If Len(planP) = 0 Or Dir(planP) = "" Then
        AcPoc_Log "Validate", "FAIL", "Plan PDF missing: " & planP
    Else
        AcPoc_Log "Validate", "OK", "Plan = " & planP
    End If
    If Len(acP) = 0 Or Dir(acP) = "" Then
        AcPoc_Log "Validate", "FAIL", "AC PDF missing: " & acP
    Else
        AcPoc_Log "Validate", "OK", "AC = " & acP
    End If
    If AcPoc_Capacity() < 1 Then
        AcPoc_Log "Validate", "FAIL", "Capacity must be >= 1"
    Else
        AcPoc_Log "Validate", "OK", "Capacity = " & AcPoc_Capacity()
    End If

    Set Sh = CreateObject("WScript.Shell")
    p2t = AcPoc_ResolvePdftotext(Sh)
    If p2t = "" Then
        AcPoc_Log "Validate", "FAIL", "pdftotext not found (PATH or %LOCALAPPDATA%\Poppler)"
    Else
        AcPoc_Log "Validate", "OK", "pdftotext = " & p2t
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outP) Then
        On Error Resume Next
        fso.CreateFolder outP
        On Error GoTo 0
        If Not fso.FolderExists(outP) Then
            AcPoc_Log "Validate", "FAIL", "Cannot create output folder: " & outP
        End If
    End If
End Sub
