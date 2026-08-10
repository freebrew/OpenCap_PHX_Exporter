Attribute VB_Name = "MDL_SlidesheetClear"
Option Explicit

Private Const SS_SHEET As String = "Slidesheet"
Private Const BTN_NAME As String = "btnClearSlidesheetRanges"
Private Const BTN_CAPTION As String = "Clear Ranges"
' AD1:AE2 - clear of the plan proximity gauge which owns merged AA1:AC7
Private Const BTN_ANCHOR As String = "AD1:AE2"
Private Const Y_FIRST As Long = 12
Private Const Y_LAST As Long = 305
Private Const Y_DATA_FIRST As Long = 13
Private Const Y_LABEL As String = "RKB"
Private Const TGT_FIRST As Long = 2
Private Const TGT_LAST As Long = 5
Private m_updatingY As Boolean
Private m_hlTgtRow As Long   ' last highlighted target row (0 = none)

' Clear Slidesheet survey/entry ranges:
'   D,F,G,T,U rows 13:305
'   Sail waypoints AC14:AD33 (display AA=Inc Next, AB=Geo Window; helper AE)
'   Targets U2:X5
'   Comments Y12:Y305 (then restore RKB + auto slide comments)
Public Sub ClearSlidesheetRanges()
    Dim ss As Worksheet

    On Error GoTo Fail

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ss.Range("D13:D305").ClearContents
    ss.Range("F13:F305").ClearContents
    ss.Range("G13:G305").ClearContents
    ss.Range("T13:T305").ClearContents
    ss.Range("U13:U305").ClearContents
    ss.Range("AC14:AD33").ClearContents
    ss.Range("U2:X5").ClearContents
    ss.Range("Y12:Y305").ClearContents
    ss.Range("Y12").Value2 = Y_LABEL

    Application.EnableEvents = True
    RefreshSlideComments forceAll:=True
    HighlightActiveTarget

    Application.ScreenUpdating = True

    MsgBox "Cleared:" & vbCrLf & _
           "  D/F/G/T/U 13:305" & vbCrLf & _
           "  Sail waypoints AC14:AD33" & vbCrLf & _
           "  Targets U2:X5" & vbCrLf & _
           "  Comments Y12:Y305 (auto slide text restored where inputs exist)", _
           vbInformation
    Exit Sub

Fail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ClearSlidesheetRanges failed: " & Err.Description, vbCritical
End Sub

' Write ProjSlideComment results into Y as values.
' forceAll:=True overwrites every data row (used on clear / initial convert).
' forceAll:=False only updates blank cells or prior auto "Sliding ... | BURR ..." notes.
Public Sub RefreshSlideComments(Optional ByVal forceAll As Boolean = False)
    Dim ss As Worksheet
    Dim r As Long
    Dim i As Long
    Dim autoTxt As String
    Dim curTxt As String
    Dim result As Variant
    Dim xArr As Variant
    Dim uArr As Variant
    Dim arArr As Variant
    Dim yArr As Variant
    Dim anyFormula As Boolean
    Dim writeIt As Boolean
    Dim wasProt As Boolean
    Dim leftPart As String
    Dim rightPart As String
    Dim p As Long
    Dim protoY As Range
    Dim spacePts As Double

    If m_updatingY Then Exit Sub
    m_updatingY = True

    On Error GoTo Clean

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    Application.EnableEvents = False

    ' AS (Meters To Slide) and AR are formulas. Under automatic calc Excel has
    ' already refreshed them before Worksheet_Calculate fires, so only force a
    ' pass in manual mode.
    If Application.Calculation = xlCalculationManual Then ss.Calculate

    wasProt = SheetUnprotectForVba(ss)

    ' Keep / restore label row
    If Len(Trim$(CStr(ss.Cells(Y_FIRST, "Y").text))) = 0 Then
        ss.Cells(Y_FIRST, "Y").Value2 = Y_LABEL
    End If

    ' Meters To Slide in hidden AS; required TF (to target) in hidden AT.
    ' Visible V/W/X = TVD @ Bit / INC @ Bit / AZM @ Bit. Y uses AT (not user U).
    xArr = ss.Range(ss.Cells(Y_DATA_FIRST, "AS"), ss.Cells(Y_LAST, "AS")).Value2
    uArr = ss.Range(ss.Cells(Y_DATA_FIRST, "AT"), ss.Cells(Y_LAST, "AT")).Value2
    arArr = ss.Range(ss.Cells(Y_DATA_FIRST, "AR"), ss.Cells(Y_LAST, "AR")).Value2
    yArr = ss.Range(ss.Cells(Y_DATA_FIRST, "Y"), ss.Cells(Y_LAST, "Y")).Value2

    ' Only pay for per-cell HasFormula checks if the column still holds formulas.
    anyFormula = True
    If VarType(ss.Range(ss.Cells(Y_DATA_FIRST, "Y"), ss.Cells(Y_LAST, "Y")).HasFormula) = vbBoolean Then
        anyFormula = ss.Range(ss.Cells(Y_DATA_FIRST, "Y"), ss.Cells(Y_LAST, "Y")).HasFormula
    End If

    Set protoY = ss.Cells(Y_DATA_FIRST, "Y")
    spacePts = MeasureTextWidthPts(" ", protoY)
    If spacePts < 0.25 Then spacePts = 0.25

    For i = 1 To UBound(xArr, 1)
        r = Y_DATA_FIRST + i - 1

        autoTxt = ""
        ' BURR alone is enough: blank/zero AS (tangent hold) still gets a comment.
        If IsNumberValue(arArr(i, 1)) Then
            ' uArr is required TF text from AT (may be "", R12, L20, 40M, …).
            If IsNumberValue(xArr(i, 1)) Then
                result = ProjSlideComment(xArr(i, 1), uArr(i, 1), arArr(i, 1))
            Else
                result = ProjSlideComment(0#, uArr(i, 1), arArr(i, 1))
            End If
            If Not isError(result) Then
                autoTxt = CStr(result & "")
                p = InStr(1, autoTxt, Chr$(1), vbBinaryCompare)
                If p > 0 Then
                    leftPart = Left$(autoTxt, p - 1)
                    rightPart = mid$(autoTxt, p + 1)
                    autoTxt = FitLeftRightInCell(leftPart, rightPart, ss.Cells(r, "Y"), spacePts)
                End If
            End If
        End If

        curTxt = Trim$(CStr(yArr(i, 1) & ""))

        writeIt = forceAll
        If Not writeIt Then
            writeIt = IsAutoSlideCommentText(curTxt)
            If Not writeIt And anyFormula Then
                writeIt = ss.Cells(r, "Y").HasFormula
            End If
        End If

        If writeIt Then
            If curTxt <> autoTxt Or (anyFormula And ss.Cells(r, "Y").HasFormula) Then
                If Len(autoTxt) = 0 Then
                    ss.Cells(r, "Y").ClearContents
                Else
                    With ss.Cells(r, "Y")
                        .WrapText = False
                        .IndentLevel = 0
                        .HorizontalAlignment = xlLeft
                        .Value2 = autoTxt
                    End With
                End If
            End If
        End If
    Next i

    HighlightActiveTargetOnSheet ss

Clean:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = True
    m_updatingY = False
End Sub

' Pad left/right text with spaces so the right part sits on the cell's right edge.
' Uses point widths (font-accurate), not ColumnWidth character units.
Private Function FitLeftRightInCell(ByVal leftPart As String, ByVal rightPart As String, _
                                    ByVal cell As Range, ByVal spacePts As Double) As String
    Dim avail As Double
    Dim leftPts As Double
    Dim rightPts As Double
    Dim PAD As Long
    ' Excel keeps a small built-in cell pad; 2pt leaves BURR flush to the right grid.
    Const CELL_MARGIN_PTS As Double = 2#

    leftPts = MeasureTextWidthPts(leftPart, cell)
    rightPts = MeasureTextWidthPts(rightPart, cell)
    avail = cell.Width - CELL_MARGIN_PTS
    If spacePts <= 0# Then spacePts = 1#

    PAD = CLng(Application.WorksheetFunction.Round((avail - leftPts - rightPts) / spacePts, 0))
    If PAD < 2 Then PAD = 2
    FitLeftRightInCell = leftPart & Space$(PAD) & rightPart
End Function

' Measure rendered text width in points via a temporary autosized textbox.
Private Function MeasureTextWidthPts(ByVal txt As String, ByVal proto As Range) As Double
    Dim ws As Worksheet
    Dim shp As Shape
    Dim nm As String
    Dim sz As Double

    On Error GoTo Fail
    If Len(txt) = 0 Then
        MeasureTextWidthPts = 0#
        Exit Function
    End If

    Set ws = proto.Worksheet
    nm = proto.Font.name
    sz = proto.Font.Size
    If sz <= 0# Then sz = 11#

    Set shp = ws.Shapes.AddTextbox(1, -2000, -2000, 20, 20) ' msoTextOrientationHorizontal
    With shp
        .Visible = False
        With .TextFrame
            .AutoSize = True
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .Characters.text = txt
            .Characters.Font.name = nm
            .Characters.Font.Size = sz
            If proto.Font.bold Then .Characters.Font.bold = True
        End With
        MeasureTextWidthPts = .Width
        .Delete
    End With
    Exit Function
Fail:
    On Error Resume Next
    If Not shp Is Nothing Then shp.Delete
    ' Fallback: ~half font-size points per character
    MeasureTextWidthPts = Len(txt) * sz * 0.5
End Function

' Highlight T:Y of the plan target the last surveyed row is aiming at.
' Public entrypoint for sheet events / button handlers.
Public Sub HighlightActiveTarget()
    Dim ss As Worksheet
    Dim wasProt As Boolean
    Dim prevEvents As Boolean

    On Error GoTo Fail
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    wasProt = SheetUnprotectForVba(ss)
    HighlightActiveTargetOnSheet ss
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
    Exit Sub
Fail:
    On Error Resume Next
    SheetReprotectAfterVba ss, wasProt
    Application.EnableEvents = prevEvents
End Sub

' Assumes the caller already owns EnableEvents / protection.
Private Sub HighlightActiveTargetOnSheet(ByVal ss As Worksheet)
    Dim r As Long
    Dim bitMD As Double
    Dim aimMd As Double
    Dim tgtMd As Variant
    Dim activeRow As Long
    Dim lastSurvRow As Long
    Dim hlColor As Long
    Dim rng As Range
    Dim vF As Variant
    Dim vAo As Variant
    Dim vD As Variant

    On Error GoTo Done
    hlColor = RGB(198, 239, 206)

    ' Use the last surveyed row (F filled), not the last bit-only tally row.
    ' Bit-only depths past a target MD would otherwise advance the highlight
    ' while BURR is still computed against the prior survey's aim (e.g. TAR2 / 8.37).
    lastSurvRow = 0
    bitMD = 0#
    aimMd = 0#
    For r = 12 To 305
        vF = ss.Cells(r, "F").Value2
        If IsNumeric(vF) Then
            If Len(Trim$(CStr(vF & ""))) > 0 Then
                lastSurvRow = r
                vD = ss.Cells(r, "D").Value2
                If IsNumeric(vD) Then bitMD = CDbl(vD)
                vAo = ss.Cells(r, "AO").Value2
                If IsNumeric(vAo) Then
                    aimMd = CDbl(vAo)
                Else
                    aimMd = 0#
                End If
            End If
        End If
    Next r

    activeRow = 0
    If lastSurvRow > 0 Then
        ' Prefer the row's AO (ProjActiveTargetMd) so highlight matches BURR's aim.
        If aimMd > 0# Then
            For r = TGT_FIRST To TGT_LAST
                tgtMd = ss.Cells(r, "U").Value2
                If IsNumeric(tgtMd) Then
                    If Abs(CDbl(tgtMd) - aimMd) < 0.005 Then
                        activeRow = r
                        Exit For
                    End If
                End If
            Next r
        End If
        ' Fallback: first plan MD ahead of that survey's bit depth.
        If activeRow = 0 And bitMD > 0# Then
            For r = TGT_FIRST To TGT_LAST
                tgtMd = ss.Cells(r, "U").Value2
                If IsNumeric(tgtMd) Then
                    If CDbl(tgtMd) > bitMD Then
                        activeRow = r
                        Exit For
                    End If
                End If
            Next r
        End If
    End If

    ' Skip paint work if nothing changed.
    If activeRow = m_hlTgtRow Then
        If activeRow = 0 Then GoTo Done
        If ss.Cells(activeRow, "T").Interior.Color = hlColor _
           And ss.Cells(activeRow, "Y").Interior.Color = hlColor Then GoTo Done
    End If

    ' Clear previous highlight on T2:Y5.
    Set rng = ss.Range(ss.Cells(TGT_FIRST, "T"), ss.Cells(TGT_LAST, "Y"))
    rng.Interior.Pattern = xlNone

    If activeRow >= TGT_FIRST And activeRow <= TGT_LAST Then
        ss.Range(ss.Cells(activeRow, "T"), ss.Cells(activeRow, "Y")).Interior.Color = hlColor
    End If
    m_hlTgtRow = activeRow

Done:
End Sub

' True for a real number; Empty and "" (formula blank) must not count.
Private Function IsNumberValue(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            IsNumberValue = True
        Case Else
            IsNumberValue = False
    End Select
End Function

' True if Y is empty or looks like previously generated slide text.
Private Function IsAutoSlideCommentText(ByVal t As String) As Boolean
    Dim s As String
    s = Trim$(t)
    If Len(s) = 0 Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    ' Old pipe format or new padded "Sliding ... BURR 0.00" format.
    If s Like "Sliding *| BURR *" Then
        IsAutoSlideCommentText = True
        Exit Function
    End If
    IsAutoSlideCommentText = (Left$(s, 8) = "Sliding ") And (InStr(1, s, "BURR ", vbTextCompare) > 0)
End Function

Public Sub EnsureSlidesheetClearButton()
    Dim ss As Worksheet
    Dim rng As Range
    Dim btn As Button
    Dim shp As Shape

    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    Set rng = ss.Range(BTN_ANCHOR)
    If rng.Cells.Count = 1 Then Set rng = rng.MergeArea

    On Error Resume Next
    ss.Buttons(BTN_NAME).Delete
    ss.Shapes(BTN_NAME).Delete
    On Error GoTo 0

    Set btn = ss.Buttons.Add(rng.Left, rng.Top, rng.Width, rng.Height)
    btn.name = BTN_NAME
    btn.caption = BTN_CAPTION
    btn.OnAction = "'" & ThisWorkbook.name & "'!ClearSlidesheetRanges"

    Set shp = ss.Shapes(BTN_NAME)
    shp.Placement = xlMoveAndSize
    On Error Resume Next
    btn.Font.bold = True
    btn.Font.Size = 10
    On Error GoTo 0
End Sub

' Only creates the button when it is missing.  An existing button is left exactly
' where it is: the buttons have been positioned by hand, and re-anchoring them
' would drop them back on top of the plan gauge (AA1:AE7).
Public Sub ResizeSlidesheetClearButton()
    Dim ss As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(SS_SHEET)
    If ss Is Nothing Then Exit Sub
    Set shp = ss.Shapes(BTN_NAME)
    If shp Is Nothing Then EnsureSlidesheetClearButton
    On Error GoTo 0
End Sub
