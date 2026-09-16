Attribute VB_Name = "Module19"
Sub Double_Stands()
'
' Double_Stands Macro
'

'
    ActiveSheet.Unprotect
    Columns("H:H").ColumnWidth = 17
    Columns("I:I").ColumnWidth = 0
    Columns("J:J").ColumnWidth = 1.43
    Columns("K:K").ColumnWidth = 14
    Columns("L:L").ColumnWidth = 0.5
    Columns("M:M").ColumnWidth = 14
    Columns("N:Z").ColumnWidth = 0
    ActiveSheet.Protect
    End Sub
Sub Triple_Stands()
'
' Triple_Stands Macro
'

'
    ActiveSheet.Unprotect
    Columns("H:H").ColumnWidth = 0
    Columns("I:I").ColumnWidth = 17
    Columns("J:J").ColumnWidth = 1.43
    Columns("K:N").ColumnWidth = 0
    Columns("O:O").ColumnWidth = 14
    Columns("P:P").ColumnWidth = 0.5
    Columns("Q:Q").ColumnWidth = 14
    Columns("R:Z").ColumnWidth = 0
    ActiveSheet.Protect
End Sub
Sub Triple_Half_Stands()
'
' Triple_Half_Stands Macro
'

'
    ActiveSheet.Unprotect
    Columns("H:H").ColumnWidth = 0
    Columns("I:I").ColumnWidth = 17
    Columns("J:J").ColumnWidth = 1.43
    Columns("K:U").ColumnWidth = 0
    Columns("V:V").ColumnWidth = 14
    Columns("W:W").ColumnWidth = 0.5
    Columns("X:X").ColumnWidth = 14
    Columns("Z:Z").ColumnWidth = 0
    ActiveSheet.Protect
End Sub
