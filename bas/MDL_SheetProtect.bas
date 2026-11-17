Attribute VB_Name = "MDL_SheetProtect"
Option Explicit

' Password used when (re)protecting sheets after VBA writes.
' Leave blank if Protect Sheet was applied with no password.
Public Const SHEET_PROTECT_PWD As String = ""

' Unprotect a sheet so VBA can write locked cells.
' Returns True if the sheet was protected before this call.
Public Function SheetUnprotectForVba(ByVal ws As Worksheet) As Boolean
    Dim wasProtected As Boolean

    wasProtected = False
    On Error Resume Next
    wasProtected = ws.ProtectContents
    On Error GoTo 0

    If Not wasProtected Then
        SheetUnprotectForVba = False
        Exit Function
    End If

    On Error Resume Next
    ws.Unprotect Password:=SHEET_PROTECT_PWD
    If ws.ProtectContents Then ws.Unprotect Password:=""
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 700, "MDL_SheetProtect", _
            "Could not unprotect sheet '" & ws.name & "'." & vbCrLf & _
            "Set SHEET_PROTECT_PWD in MDL_SheetProtect to the sheet password."
    End If

    SheetUnprotectForVba = True
End Function

' Re-apply sheet protection after VBA writes.
' UserInterfaceOnly:=True lets later VBA write locked cells in this Excel session
' while users still cannot edit locked cells.
Public Sub SheetReprotectAfterVba(ByVal ws As Worksheet, ByVal wasProtected As Boolean)
    If Not wasProtected Then Exit Sub
    If ws Is Nothing Then Exit Sub
    SheetEnsureLocked ws
End Sub

' Lock the sheet for the user even when the calc started unlocked.
' Slidesheet calculations used to re-lock only if they had unlocked a
' protected sheet, so a pass that began unlocked finished unlocked.
Public Sub SheetEnsureLocked(ByVal ws As Worksheet)
    Dim allowFmt As Boolean
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If ws.ProtectContents Then Exit Sub
    ' Slidesheet stays locked (F/G values cannot be typed over) but the
    ' DD can still change fill colors. Excel has no per-column format
    ' permission, so formatting is allowed on this sheet only.
    allowFmt = (StrComp(ws.name, "Slidesheet", vbTextCompare) = 0)
    ws.Protect Password:=SHEET_PROTECT_PWD, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFormattingCells:=allowFmt
    On Error GoTo 0
End Sub

' Keep Data!H35:M47 unlocked so formula results / mirror table stay usable when locked.
Public Sub EnsureDataBhaMirrorUnlocked()
    Dim ws As Worksheet
    Dim wasProt As Boolean

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets("Data")
    wasProt = SheetUnprotectForVba(ws)
    ws.Range("H35:M47").Locked = False
    SheetReprotectAfterVba ws, wasProt
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ws, wasProt
End Sub









