VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton2_Click()
'
' saveaspdf Macro
'
' Keyboard Shortcut: Ctrl+L
'
 Dim name As String
    
   name = ThisWorkbook.Path & "\" & Range("B5") & Range("B6") & " DATA" & ".pdf"
        
    Application.GoTo Reference:="Print_Area"

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=name, _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    Range("B7:B25").Select
    Selection.ClearContents
    Range("C7:D7").Select
    Selection.ClearContents
    Range("D8:D25").Select
    Selection.ClearContents
    Range("F7:H25").Select
    Selection.ClearContents
    Range("J7:K25").Select
    Selection.ClearContents
    Range("J3:K3").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=0
        
End Sub


Private Sub CommandButton1_Click()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    PdmConf_OnDataChange Target
    DayRoll_OnDataChange Target
    ToolHours_OnDataChange Target
    On Error GoTo 0
End Sub





