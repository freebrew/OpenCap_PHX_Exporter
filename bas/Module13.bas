Attribute VB_Name = "Module13"
Sub yesOrNo()

CarryOn = MsgBox("Have you completly filled out the Data sheet for the end of the run?", vbYesNo, "NEW SLIDESHEET")
If CarryOn = vbYes Then


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
      ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 79
    Application.Width = 1388.25
    Application.Height = 626.25
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 206
    ActiveWindow.ScrollRow = 237
    ActiveWindow.ScrollRow = 240
    ActiveWindow.ScrollRow = 247
    ActiveWindow.ScrollRow = 249
    ActiveWindow.ScrollRow = 251
    ActiveWindow.ScrollRow = 255
    ActiveWindow.ScrollRow = 257
    ActiveWindow.ScrollRow = 262
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 277
    ActiveWindow.ScrollRow = 280
    ActiveWindow.ScrollRow = 282
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 285
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 294
    ActiveWindow.ScrollRow = 295
    Range("A305").Select
    Selection.AutoFill Destination:=Range("A12:A305"), Type:=xlFillDefault
    Range("A12:A305").Select
    ActiveWindow.SmallScroll Down:=288
    Range("C305").Select
    Selection.AutoFill Destination:=Range("C12:C305"), Type:=xlFillDefault
    Range("C12:C305").Select
    Range("E12:F305").Select
    Selection.ClearContents
    Range("S11:T305").Select
    Selection.ClearContents
    Range("G305:H305").Select
    Selection.AutoFill Destination:=Range("G12:H305"), Type:=xlFillDefault
    Range("G12:H305").Select
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 134
    ActiveWindow.ScrollRow = 143
    ActiveWindow.ScrollRow = 169
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 203
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 218
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 234
    ActiveWindow.ScrollRow = 245
    ActiveWindow.ScrollRow = 248
    ActiveWindow.ScrollRow = 249
    ActiveWindow.ScrollRow = 251
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 254
    ActiveWindow.ScrollRow = 255
    ActiveWindow.ScrollRow = 263
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 275
    ActiveWindow.ScrollRow = 278
    ActiveWindow.ScrollRow = 281
    ActiveWindow.ScrollRow = 288
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 295
    Range("Q305:R305").Select
    Selection.AutoFill Destination:=Range("Q11:R305"), Type:=xlFillDefault
    Range("Q11:R305").Select
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 189
    ActiveWindow.ScrollRow = 229
    ActiveWindow.ScrollRow = 242
    ActiveWindow.ScrollRow = 261
    ActiveWindow.ScrollRow = 283
    ActiveWindow.ScrollRow = 285
    ActiveWindow.ScrollRow = 288
    ActiveWindow.ScrollRow = 295
    Range("X305").Select
    Selection.AutoFill Destination:=Range("X11:X305"), Type:=xlFillDefault
    Range("X11:X305").Select
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 143
    ActiveWindow.ScrollRow = 195
    ActiveWindow.ScrollRow = 209
    ActiveWindow.ScrollRow = 222
    ActiveWindow.ScrollRow = 245
    ActiveWindow.ScrollRow = 246
    ActiveWindow.ScrollRow = 259
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 278
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 287
    ActiveWindow.ScrollRow = 288
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 292
    ActiveWindow.ScrollRow = 293
    ActiveWindow.ScrollRow = 294
    ActiveWindow.ScrollRow = 295
    Range("Z305").Select
    Selection.AutoFill Destination:=Range("Z11:Z305"), Type:=xlFillDefault
    Range("Z11:Z305").Select
    Range("AD11:AF304").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=-284
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
    
End If
End Sub



