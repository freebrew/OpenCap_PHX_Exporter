Attribute VB_Name = "MDL_Ribbon"
Option Explicit

' Keep a ribbon reference so Excel binds callbacks reliably.
Public gRibbon As Object

Public Sub Ribbon_OnLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub

' Workbook RibbonX callback
Public Sub RibbonShowPipeTally(control As Object)
    On Error GoTo Fail
    ShowPipeTallyForm
    Exit Sub
Fail:
    MsgBox "Pipe Tally form failed: " & Err.Description, vbCritical
End Sub
