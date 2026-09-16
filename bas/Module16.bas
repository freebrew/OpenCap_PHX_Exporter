Attribute VB_Name = "Module16"
Sub Macro1()
'
' Macro1 Macro
'

'
    Range("C14:C505,E11:F505,C11:L11,C12:C13,U11:W505,Y11:Y505").Select
    Range("Y505").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("D12:D505,I12:L505").Select
    Range("I505").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = 19
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("S11:T505").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
End Sub
