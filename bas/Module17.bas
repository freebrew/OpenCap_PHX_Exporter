Attribute VB_Name = "Module17"
Sub nosidetrack()
'
' nosidetrack Macro
'

'
    sheets("Sidetrack Points").Visible = False
    sheets("Sidetrack").Visible = False
    sheets("Setup").Select
End Sub
Sub yessidetrack()
'
' yessidetrack Macro
'
sheets("Sidetrack Points").Visible = True
    sheets("Sidetrack").Visible = True
    sheets("Setup").Select
'
End Sub
