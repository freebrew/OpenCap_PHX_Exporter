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
    ' no auto InitSetup
End Sub

' Cont DI tabs painted before the "Last Survey" button existed get it on first
' visit (shape only; no cells, no repaint, no course archive).
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If StrComp(Sh.name, MDL_ContDI.SH_CONTDI, vbTextCompare) = 0 Then
        MDL_ContDI.EnsureContDIPullButton
    End If
End Sub
