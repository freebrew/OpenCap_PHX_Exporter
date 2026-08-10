VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCostDayBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Public dateSerial As Long
Public ParentForm As Object

Private Sub btn_Click()
    If ParentForm Is Nothing Then Exit Sub
    If dateSerial <= 0 Then Exit Sub
    ParentForm.DayButtonClicked dateSerial
End Sub
