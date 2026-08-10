Attribute VB_Name = "Module1"
Private Sub WriteLongTextToSheet(ByVal ws As Worksheet, ByVal textValue As String)

    Dim chunkSize As Long
    Dim pos As Long
    Dim rowNum As Long

    chunkSize = 30000
    pos = 1
    rowNum = 1

    Do While pos <= Len(textValue)
        ws.Cells(rowNum, 1).Value = mid$(textValue, pos, chunkSize)
        pos = pos + chunkSize
        rowNum = rowNum + 1
    Loop

End Sub
