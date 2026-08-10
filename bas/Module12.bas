Attribute VB_Name = "Module12"
Sub data_PDF()
Attribute data_PDF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Data_PDF Macro
'

'
    Range("A1:Y63").Select
    ActiveSheet.PageSetup.PrintArea = "$A$1:$Y$63"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        "G:\My Drive\Work\Current Wells\Axiom\30971\Paperwork.pdf", Quality:= _
        xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
End Sub
Sub CLEAR_DATA()
Attribute CLEAR_DATA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CLEAR_DATA Macro
'

'
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
Sub CLEAR_TIMETABLE()
'
' CLEAR_TIMETABLE

'
    Range("H4:J22").Select
    Selection.ClearContents
    Range("L4:N22").Select
    Selection.ClearContents
    Range("P4:Q22").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=0
End Sub
