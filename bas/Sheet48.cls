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
    ResizeSlidesheetClearButton
    MDL_PlanGauge.RenderPlanGauge
    On Error GoTo 0
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Not Intersect(Target, Me.Range("AA:AB")) Is Nothing Then
        ResizeSlidesheetClearButton
    End If
    If Not Intersect(Target, Me.Range("Y2:Y5")) Is Nothing Then
        RememberPlanTargetOverride Target
    End If
    ' Survey / slide / target edits change Meters To Slide (X) and BURR (AR),
    ' which drive the auto slide notes in Y. Survey depth also slides the
    ' 4-row plan-target window when hidden Plan Sections exist.
    If Not Intersect(Target, Me.Range("D13:G305,T13:U305,U2:X5")) Is Nothing Then
        RefreshSlideComments
    End If
    ' Plan proximity gauge: surveys, targets, geo half-width (AB14), and
    ' sail waypoints (AC14:AD33). Without waypoints/AB14 here the red/green
    ' geo bands go stale after manual waypoint edits.
    If Not Intersect(Target, Me.Range("D13:G320,U2:X5,AB14,AC14:AD33")) Is Nothing Then
        MDL_PlanGauge.RenderPlanGauge
    End If
    On Error GoTo 0
End Sub

Private Sub Worksheet_Calculate()
    On Error Resume Next
    RefreshSlideComments
    On Error GoTo 0
End Sub
