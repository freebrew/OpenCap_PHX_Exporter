Attribute VB_Name = "MDL_ScreenGuard"
Option Explicit

' ================================================================================
'  SCREEN GUARD - reference-counted freeze of screen / events / calculation
'
'  Problem this solves:
'    Refresh routines call each other (RefreshData -> RefreshAllCsvData ->
'    BuildSetupUI ...).  When each one ended with "Application.ScreenUpdating =
'    True" the screen un-froze while the OUTER routine was still working, so
'    Excel repainted half-built sheets - the flashing / blank / dark window.
'
'  Usage (always pair them, End in the error handler too):
'      ScreenBeginBusy "Refresh: importing..."
'      ... work ...
'      ScreenEndBusy
'
'  Only the outermost Begin/End pair actually touches Application state, so
'  nesting is safe to any depth.
' ================================================================================

Private mDepth As Long
Private mPrevScreen As Boolean
Private mPrevEvents As Boolean
Private mPrevCalc As Long
Private mPrevAlerts As Boolean
Private mPrevCursor As Long
Private mHaveState As Boolean

Public Sub ScreenBeginBusy(Optional ByVal statusText As String = "")
    If mDepth = 0 Then
        On Error Resume Next
        mPrevScreen = Application.ScreenUpdating
        mPrevEvents = Application.EnableEvents
        mPrevAlerts = Application.DisplayAlerts
        mPrevCursor = Application.Cursor
        mPrevCalc = Application.Calculation
        If Err.Number <> 0 Then
            mPrevCalc = xlCalculationAutomatic
            Err.Clear
        End If
        mHaveState = True

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.Cursor = xlWait
        On Error GoTo 0
    End If

    mDepth = mDepth + 1

    If Len(statusText) > 0 Then
        On Error Resume Next
        Application.StatusBar = statusText
        On Error GoTo 0
    End If
End Sub

Public Sub ScreenEndBusy()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth > 0 Then Exit Sub
    RestoreState
End Sub

' Emergency reset: use in error handlers where the nesting depth is unknown,
' or from the VBE if a crash left the screen frozen.
Public Sub ScreenForceReset()
    mDepth = 0
    RestoreState
End Sub

Public Function ScreenBusyDepth() As Long
    ScreenBusyDepth = mDepth
End Function

Private Sub RestoreState()
    On Error Resume Next
    If mHaveState Then
        Application.Calculation = mPrevCalc
        Application.EnableEvents = mPrevEvents
        Application.DisplayAlerts = mPrevAlerts
        Application.Cursor = mPrevCursor
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.Cursor = xlDefault
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    mHaveState = False
    ScreenRepaint
    On Error GoTo 0
End Sub

' Force Excel to repaint the visible grid.  Re-assigning the scroll position is
' a no-op logically but makes Excel redraw, which clears the grey/blank
' artefacts left behind after a long frozen operation.
Public Sub ScreenRepaint()
    On Error Resume Next
    Dim w As Window
    Set w = ActiveWindow
    If w Is Nothing Then Exit Sub
    w.ScrollRow = w.ScrollRow
    w.ScrollColumn = w.ScrollColumn
    On Error GoTo 0
End Sub
