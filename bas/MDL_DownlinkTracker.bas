Attribute VB_Name = "MDL_DownlinkTracker"
Option Explicit

' CLEAR button on the Downlink Tracker tab (Field only; Demo has no tracker).
' The tab is protected with a password we do not have, so the clear touches
' only unlocked, non-formula cells in the log - exactly the user-entry area.
' Dnlk # / D/L Setting formulas, hidden helper columns, the command lookup
' tables (AA:AL), dropdowns, and PASS/FAIL formatting all stay.

Public Const SH_DLT As String = "Downlink Tracker V4.6"
Private Const DLT_LOG As String = "A3:U674"
Public Const DLT_BTN As String = "DLT_ClearBtn"

Public Sub DLT_ClearEntries()
    Dim ws As Worksheet
    Dim n As Long

    Set ws = DltSheet()
    If ws Is Nothing Then
        MsgBox "The '" & SH_DLT & "' tab is not in this workbook.", vbInformation, "Clear Downlink Tracker"
        Exit Sub
    End If
    If MsgBox("Clear every downlink entry in rows 3-674?" & vbCrLf & vbCrLf & _
              "Client / Rig / Well / Job # and all formulas stay." & vbCrLf & _
              "This cannot be undone.", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Clear Downlink Tracker") <> vbYes Then Exit Sub

    n = DLT_ClearLog(ws)
    If n < 0 Then
        MsgBox "Clear stopped: " & Err.Description, vbCritical, "Clear Downlink Tracker"
    Else
        MsgBox n & " entries cleared.", vbInformation, "Clear Downlink Tracker"
    End If
End Sub

' Returns cells cleared, or -1 on error (Err still set for the caller).
Public Function DLT_ClearLog(ByVal ws As Worksheet) As Long
    Dim c As Range
    Dim n As Long
    Dim prevScreen As Boolean, prevEvents As Boolean
    Dim prevCalc As XlCalculation

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail
    For Each c In ws.Range(DLT_LOG).Cells
        If Not c.Locked Then
            If Not c.HasFormula Then
                If Not IsEmpty(c.Value2) Then
                    c.ClearContents
                    n = n + 1
                End If
            End If
        End If
    Next c
    DLT_ClearLog = n
    GoTo Restore

Fail:
    DLT_ClearLog = -1

Restore:
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
End Function

' Places (or re-places) the CLEAR button in the frozen header row, inside the
' wide Job # cell, so it stays on screen while scrolling the log.
Public Sub DLT_EnsureClearButton()
    Dim ws As Worksheet
    Dim anchor As Range
    Dim btn As Object
    Const BTN_W As Double = 84
    Dim btnH As Double

    Set ws = DltSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Buttons(DLT_BTN).Delete
    On Error GoTo 0

    Set anchor = ws.Range("U1")
    btnH = anchor.Height - 4
    Set btn = ws.Buttons.Add(anchor.Left + anchor.Width - BTN_W - 4, anchor.Top + 2, BTN_W, btnH)
    btn.name = DLT_BTN
    btn.caption = "CLEAR"
    btn.OnAction = "DLT_ClearEntries"
    btn.Placement = xlFreeFloating
    btn.PrintObject = False
    btn.Font.bold = True
End Sub

Private Function DltSheet() As Worksheet
    On Error Resume Next
    Set DltSheet = ThisWorkbook.Worksheets(SH_DLT)
    On Error GoTo 0
End Function

