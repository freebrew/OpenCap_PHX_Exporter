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
Private mPrevAnim As Boolean
Private mHaveState As Boolean
Private mSheetHold As Boolean
Private mSheetPrevCalc As Long

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
        On Error Resume Next
        mPrevAnim = Application.EnableAnimations
        Application.EnableAnimations = False
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

Public Function ScreenBusyDepth() As Long
    ScreenBusyDepth = mDepth
End Function

' While Slidesheet is active, keep Application calc Manual so committing a
' survey does not auto-recalc the whole book (and paint) before Change can
' freeze the screen and run one Calculate.
Public Sub ScreenHoldCalcForSheet()
    On Error Resume Next
    If Not mSheetHold Then
        mSheetPrevCalc = Application.Calculation
        If Err.Number <> 0 Then mSheetPrevCalc = xlCalculationAutomatic
        mSheetHold = True
    End If
    Application.Calculation = xlCalculationManual
    On Error GoTo 0
End Sub

Public Sub ScreenReleaseCalcForSheet()
    If Not mSheetHold Then Exit Sub
    On Error Resume Next
    Application.Calculation = mSheetPrevCalc
    On Error GoTo 0
    mSheetHold = False
End Sub

Public Sub ScreenForceReset()
    mDepth = 0
    If mSheetHold Then
        On Error Resume Next
        Application.Calculation = mSheetPrevCalc
        On Error GoTo 0
        mSheetHold = False
    End If
    RestoreState
End Sub

Private Sub RestoreState()
    On Error Resume Next
    If mHaveState Then
        Application.Calculation = mPrevCalc
        Application.EnableEvents = mPrevEvents
        Application.DisplayAlerts = mPrevAlerts
        Application.Cursor = mPrevCursor
        Application.EnableAnimations = mPrevAnim
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.Cursor = xlDefault
        Application.EnableAnimations = True
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
