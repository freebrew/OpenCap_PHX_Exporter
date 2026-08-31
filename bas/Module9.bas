Attribute VB_Name = "Module9"
Sub NEW_SS()
'
' NEW_SS Macro
'
' Keyboard Shortcut: Ctrl+q
  sheets("Data").Select
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
    sheets("Slidesheet").Select
    sheets("Slidesheet").Copy Before:=sheets("Slidesheet")
    sheets("Slidesheet").Select
    Range("C11:L11").Select
    Selection.ClearContents
    Range("C12:C305").Select
    Selection.ClearContents
    Range("E12:F305").Select
    Range("F305").Activate
    Selection.ClearContents
    Range("B11").Select
    Selection.ClearContents
    Range("U11:V305").Select
    Selection.ClearContents
    Range("Z197:Z305").Select
    Range("Z305").Activate
    ActiveWindow.SmallScroll Down:=129
    Range("Z305").Select
    Selection.AutoFill Destination:=Range("Z11:Z305"), Type:=xlFillDefault
    Range("Z11:Z305").Select
    Range("S12").Select
    Selection.AutoFill Destination:=Range("S11:S12"), Type:=xlFillDefault
    Range("S11:S12").Select
    Range("AF11:AH304").Select
    Selection.ClearContents
    Range("W2:X4").Select
    Selection.ClearContents
    Range("R2:T2").Select
    Selection.ClearContents
    Range("P3:R4").Select
    Selection.ClearContents
    Range("N2:N4").Select
    Selection.ClearContents
    Range("K2:L4").Select
    Selection.ClearContents
    Range("G2:H3").Select
    Selection.ClearContents
    ActiveWindow.ScrollRow = 286
    ActiveWindow.ScrollRow = 285
    ActiveWindow.ScrollRow = 239
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 211
    ActiveWindow.ScrollRow = 190
    ActiveWindow.ScrollRow = 169
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 11
    Range("G3").Select
    Selection.ClearContents
    Range("C11:L11").Select
    Selection.ClearContents
    Range("G3").Select
    Selection.ClearContents
    Range("K3").Select
    Selection.ClearContents
    Range("AB5").Select
    Selection.ClearContents
    Range("B11").Select
    
End Sub


