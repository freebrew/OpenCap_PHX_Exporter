Attribute VB_Name = "Module3"
Sub data_PDF()


CarryOn = MsgBox("Do you want make a copy and clear all data?", vbYesNo, "NEW DATA")
If CarryOn = vbYes Then


 sheets("Data").Select
 Dim name As String
    
   name = ThisWorkbook.Path & "\" & Range("B5") & Range("B6") & " DATA" & ".pdf"
        
    Application.GoTo Reference:="Print_Area"

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=name, _
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
    
End If
End Sub



