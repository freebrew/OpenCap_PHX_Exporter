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

Private Sub Worksheet_Activate()
    On Error Resume Next
    ScreenHoldCalcForSheet
    ScreenBeginBusy
    ResizeSlidesheetClearButton
    MDL_PlanGauge.RenderPlanGauge
    ScreenEndBusy
    On Error GoTo 0
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next
    ScreenReleaseCalcForSheet
    On Error GoTo 0
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Hold calc + freeze paint for the whole handler. Committing a survey
    ' used to auto-recalc the book and redraw gauges/comments with the
    ' screen live — that is the stall on field laptops.
    On Error GoTo Fail
    ScreenHoldCalcForSheet
    ScreenBeginBusy "Slidesheet: calculating..."
    ' One book calc on this tab: V/W/X @ bit, sail AA/AB, waypoints, BURR.
    ' Do not leave Slidesheet — comments + gauge paint from that calc.
    Application.Calculate
    If Not Intersect(Target, Me.Range("AA:AB")) Is Nothing Then
        ResizeSlidesheetClearButton
    End If
    If Not Intersect(Target, Me.Range("Y2:Y5")) Is Nothing Then
        RememberPlanTargetOverride Target
    End If
    ' Slide metres @ TF + BURR in Y, on this Change (no tab switch).
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
    Exit Sub
Fail:
    On Error Resume Next
    ScreenEndBusy
    On Error GoTo 0
End Sub
