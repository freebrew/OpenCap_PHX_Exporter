VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Workbook_Open intentionally does not call InitSetup.
' InitSetup / BuildSetupUI wipe Setup (formatting, AC table, notes).
' Use Setup buttons: REFRESH CSVs or REBUILD when a rebuild is wanted.
Private Sub Workbook_Open()
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    MDL_BrowserHandoff.BrowserHandoff_CheckOnOpen
    On Error GoTo 0
End Sub

' Cont DI tabs painted before the "Last Survey" button existed get it on first
' visit (shape only; no cells, no repaint, no course archive).
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    On Error Resume Next
    If MDL_DownlinkTracker.DLT_IsTracker(sh) Then
        MDL_DownlinkTracker.DLT_EnsureButtonsIfMissing
        Exit Sub
    End If
    If StrComp(sh.name, MDL_ContDI.SH_CONTDI, vbTextCompare) = 0 Then
        ' Button only. The target cells are references to the slidesheet,
        ' so opening this tab does not search surveys or rebuild the aim.
        MDL_ContDI.EnsureContDIPullButton
        ScreenReleaseCalcForSheet
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
    On Error Resume Next
    ' DL Tracker: typed page / command / pass-fail recomputes that row's I and J.
    If MDL_DownlinkTracker.DLT_IsTracker(sh) Then
        MDL_DownlinkTracker.DLT_OnSheetChange sh, Target
        Exit Sub
    End If
    If ScreenBusyDepth() > 0 Then Exit Sub
    If StrComp(sh.name, MDL_ContDI.SH_CONTDI, vbTextCompare) <> 0 Then Exit Sub
    MDL_ContDI.ContDI_ApplyTargetSelection sh, Target
    ScreenBeginBusy
    ScreenCalcStaggered
    ScreenEndBusy
    ScreenReleaseCalcForSheet
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetDeactivate(ByVal sh As Object)
    On Error Resume Next
    If StrComp(sh.name, "Slidesheet", vbTextCompare) = 0 Then
        ScreenReleaseCalcForSheet
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ScreenReleaseCalcForSheet
    ScreenForceReset
    On Error GoTo 0
End Sub
