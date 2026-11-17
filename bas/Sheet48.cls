VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' True when the previous selection was inside the Inc/Azm entry cells.
Private m_prevInFG As Boolean

' Changing a fill color raises no Change event. When the DD leaves F:G after
' recoloring, move the newest-survey bar to the new bottom green row.
' Gauge and comments wait until the selection leaves F:G, so tabbing from
' Inc to Azm does not recalculate Data, Cont DI, and the plan gauge.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim nowIn As Boolean
    On Error Resume Next
    nowIn = Not Intersect(Target, Me.Range("F13:G505")) Is Nothing
    If m_prevInFG And Not nowIn Then
        RefreshNewestSurveyIfMoved
        RefreshSlideComments
        MDL_PlanGauge.RenderPlanGauge
    End If
    m_prevInFG = nowIn
    On Error GoTo 0
End Sub

Private Sub Worksheet_Activate()
    On Error Resume Next
    ' Stay Automatic on Slidesheet — Manual here forced F9 between edits.
    ScreenReleaseCalcForSheet
    ScreenBeginBusy
    ScreenCalcStaggered
    ResizeSlidesheetClearButton
    MDL_SlidesheetClear.EnsureTargetButton
    MDL_PlanGauge.RenderPlanGauge
    HighlightNewestSurveyRow
    ScreenEndBusy
    ScreenReleaseCalcForSheet
    SheetEnsureLocked Me
    On Error GoTo 0
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next
    ScreenReleaseCalcForSheet
    On Error GoTo 0
End Sub

Private Function SurveyEntryOnly(ByVal Target As Range) As Boolean
    Dim fg As Range
    SurveyEntryOnly = False
    Set fg = Intersect(Target, Me.Range("F13:G505"))
    ' VBA does not short-circuit And: fg.Cells on Nothing would raise 91.
    If fg Is Nothing Then Exit Function
    SurveyEntryOnly = (fg.Cells.Count = Target.Cells.Count)
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Typing Inc/Azm stays on Automatic. Forcing a three-sheet calc and the
    ' gauge here is what made the sheet hitch on the tab from F to G.
    ' The newest-survey bar still has to follow the new green station.
    If SurveyEntryOnly(Target) Then
        RefreshNewestSurveyIfMoved
        CheckSurveyRowsAfterEntry Target
        Exit Sub
    End If

    ' Freeze paint for the handler. Calc stays Automatic on this tab so
    ' formulas do not wait for F9; BeginBusy is Manual only during the work.
    On Error GoTo Fail
    ScreenBeginBusy "Slidesheet: calculating..."
    ' Slidesheet first, then Data / Cont DI with the screen still frozen.
    ScreenCalcStaggered
    If Not Intersect(Target, Me.Range("AA:AB")) Is Nothing Then
        ResizeSlidesheetClearButton
    End If
    ' Name or MD / INC / AZM / TVD. Saved before the comment refresh, which
    ' rewrites T2:Y5 from the plan and would otherwise put the old TVD back.
    If Not Intersect(Target, Me.Range("U2:Y5")) Is Nothing Then
        RememberPlanTargetOverride Target
    End If
    ' Slide metres @ TF + DLR in Y, on this Change (no tab switch).
    If Not Intersect(Target, Me.Range("C13:H505,T13:U505,U2:X5")) Is Nothing Then
        RefreshSlideComments
    End If
    If Not Intersect(Target, Me.Range("C13:D505,T13:T505")) Is Nothing Then
        RefreshRotateMetres Target
    End If
    ' Gauge + sail: surveys (incl. H TVD), target window, AA/AB, waypoints.
    If Not Intersect(Target, Me.Range("D13:H520,U2:X5,AA14:AD33")) Is Nothing Then
        MDL_PlanGauge.RenderPlanGauge
    End If
    ScreenEndBusy
    ScreenReleaseCalcForSheet
    SheetEnsureLocked Me
    Exit Sub
Fail:
    On Error Resume Next
    ScreenEndBusy
    ScreenReleaseCalcForSheet
    SheetEnsureLocked Me
    On Error GoTo 0
End Sub
