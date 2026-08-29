Attribute VB_Name = "Module20"
Sub Printdatapdf()
'
' Printdatapdf Macro
'

'
Dim name As String

    name = ThisWorkbook.Path & "\" & "Drilling Data" & ".pdf"
    Range("B5:F70").Select
    ActiveSheet.PageSetup.PrintArea = "$B$5:$F$70"
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=name, _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    Range("E1:V70").Select
    Range("R1").Activate
    ActiveSheet.PageSetup.PrintArea = "$B$1:$V$70"
End Sub




