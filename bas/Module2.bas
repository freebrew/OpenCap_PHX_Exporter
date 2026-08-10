Attribute VB_Name = "Module2"
Sub RTP_print()

CarryOn = MsgBox("Have you filled out all Data?", vbYesNo, "COPY DATA")
If CarryOn = vbYes Then
    
    Range("B5:F48").Select
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Range("B5:F48").Select
    Selection.Copy
End If
End Sub
