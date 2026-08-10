Attribute VB_Name = "Module16"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("C14:C305,E11:F305,C11:L11,C12:C13,U11:W305,Y11:Y305").Select
    Range("Y305").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434828
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("D12:D305,I12:L305").Select
    Range("I305").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = 19
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("S11:T305").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
End Sub
