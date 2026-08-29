Attribute VB_Name = "MDL_CorridorImage"
Option Explicit

' ================================================================================
'  MDL_CorridorImage
'
'  Renders the daily email / print graphic as one tall PNG. Paper theme: white
'  ground, black type, wireframe room, thin strokes only (no flooded panels).
'  Canvas width is the former B2:F55 table (1023). Band 1 is the corridor;
'  band 2 is the ops pack (day/BHA/motor/AC/motors).
'
'  Adaptive room. VERTICAL/BUILD is a traveling-cylinder bullseye at the next
'  named T2:Y5 target (no sail UP/DN or AA14 L/R bands). LATERAL is the sail
'  tunnel with geo +/- and AA14 left/right on the walls and floor.
'
'      shaft:  concentric N/E target rings; hole + plan pierce the plane.
'      tunnel: MD along X; TVD plot exaggerated; AA14 / geo +/- on walls+floor
'
'  Framed from the last survey to the next named T2:Y5 target.
'
'  Two panels sit beside it: a true 1:1 vertical section of the whole well, and
'  the last-survey / position / requirement figures as text. Those figures are the
'  same ones the Data sheet prints in D7:F17; that block is read, never written, so
'  its formulas and Module21's morning summary are untouched.
'
'  All plan maths comes from MDL_PlanGauge through its PG_* shims, so this picture
'  and the on-screen proximity gauge can never disagree.
'
'  Geometry is a straight port of bas\_proto_tunnel_render.ps1. Coordinates below
'  are in POINTS on the same 1023 x 445 grid the prototype used as pixels; drawing
'  at points and exporting supersamples by 4/3, so the image stays crisp when the
'  mail scales it back to 1023 px wide.
'
'  Entry points:
'    RenderCorridorPng(path)              -> full report (corridor + ops), or ""
'    RenderCorridorPng(path, "corridor")  -> light band only
'    RenderCorridorPng(path, "ops")       -> ops band only
'  EMAIL attaches the full canvas as daily_report.png and inlines it under
'  the Data table (cid:daily_report.png).
' ================================================================================

Private Const CANVAS_W As Double = 1023
Private Const CORRIDOR_H As Double = 445   ' light band height; full height is mCanvasH
Private Const SHP_PREFIX As String = "OCC_"
Private Const OPS_PAD As Double = 12
Private Const OPS_GAP As Double = 8
Private Const OPS_COL_GAP As Double = 10
' Ops / metrics type is drawn larger so after Outlook fits the picture to the
' reading pane the glyphs stay near body-text size.
Private Const TYPE_SCALE_OPS As Double = 1.35
Private Const TYPE_SCALE_META As Double = 1.2

Private Const WINDOW_LEN As Double = 1220#      ' mMD of hole shown in the room
Private Const RUN_AHEAD As Double = 20#         ' mMD drawn past the next waypoint

' panel 1, the room
Private Const V_X As Double = 12
Private Const V_Y As Double = 30
Private Const V_W As Double = 490
Private Const V_H As Double = 376
Private Const KZX As Double = 3#                ' oblique axis, pt per m of lateral
Private Const KZY As Double = 2.2
Private Const SXD As Double = 0.297             ' pt per m of measured depth
Private Const DEV_SPAN_PT As Double = 220#      ' pt given to the whole up/down range

' panel 2, the 1:1 vertical section, and the reporting-day block under it
Private Const P_X As Double = 530
Private Const P_Y As Double = 30
Private Const P_W As Double = 232
Private Const P_H As Double = 250
Private Const Q_Y As Double = 290
Private Const Q_H As Double = 116

' panel 3, the metrics (wider so long position values clear the labels)
Private Const M_X As Double = 758
Private Const M_W As Double = 253
Private Const M_H As Double = 376

Private Const SS_WP_ROW_FIRST As Long = 14
Private Const SS_WP_ROW_LAST As Long = 33
Private Const SS_WP_COL_MD As Long = 29         ' AC
Private Const SS_WP_COL_TVD As Long = 30        ' AD
Private Const SS_WP_COL_INC As Long = 31        ' AE  Inc to next
Private Const SS_LAT_TOL_ADDR As String = "AA14"
Private Const SS_BIT_COL As Long = 4            ' D, bit depth on the survey row

Private Const PIE As Double = 3.14159265358979

' Live layout — set by ApplySectionLayout (vertical / build / lateral).
Private mVX As Double, mVY As Double, mVW As Double, mVH As Double
Private mPX As Double, mPY As Double, mPW As Double, mPH As Double
Private mQY As Double, mQH As Double
Private mMX As Double, mMW As Double, mMH As Double
Private mSxd As Double, mKzx As Double, mKzy As Double, mDevSpan As Double
Private mSection As String
Private mTgtName As String
Private mShaft As Boolean
Private mUseSailBands As Boolean
Private mTvdTop As Double

' ---- scene state, set by BuildScene and read by the projection ----------------
Private mMD0 As Double, mMD1 As Double
Private mDevHi As Double, mDevLo As Double
Private mLatTol As Double, mLatBox As Double, mGeoHalf As Double
Private mX0 As Double, mY0 As Double, mSV As Double

' ---- gathered survey stations inside the window -------------------------------
Private mSMD() As Double, mSDev() As Double, mSLat() As Double, mSTvd() As Double
Private mSNS() As Double, mSEW() As Double
Private mSCount As Long
Private mSvyNS As Double, mSvyEW As Double
Private mShOx As Double, mShOy As Double, mShS As Double
Private mShV0 As Double, mShNMin As Double, mShNMax As Double
Private mShEMin As Double, mShEMax As Double, mShV1 As Double
Private mTgtN As Double, mTgtE As Double, mTgtV As Double
Private mShVS As Double   ' pt per m of TVD (pierce is compressed, disk is 1:1)
' Vertical plot on the walls: plotY = mTvdDatum - TVD (positive = shallower).
' Geo corridor ribbons slope with waypoint TVD; hole uses the same frame.
Private mTvdDatum As Double

' ---- last survey and forward projection ---------------------------------------
Private mSvyMD As Double, mSvyInc As Double, mSvyAzi As Double, mSvyTvd As Double
Private mSvyDev As Double, mSvyLat As Double, mSvyPlanAzi As Double
Private mBitMD As Double
Private mWpMD As Double, mWpTvd As Double
Private mToGo As Double, mReqInc As Double, mHoldDevEnd As Double, mHoldLatEnd As Double

' ---- waypoints -----------------------------------------------------------------
Private mWMD() As Double, mWTvd() As Double, mWInc() As Double
Private mWCount As Long

' ---- shape bookkeeping ----------------------------------------------------------
Private mNames() As String
Private mNameN As Long
Private mSeq As Long
Private mWs As Worksheet
Private mCanvasH As Double
Private mOpsOrigin As Double   ' top of ops band (0 for ops-only export)
Private mTypeScale As Double   ' applied inside Tx

' ---- diagnostics ---------------------------------------------------------------
' The renderer runs unattended inside the mail build, and under automation an
' untrapped error opens the VBE and hangs the caller forever. Every entry point
' therefore traps, records where it got to, and returns empty.
Private mStage As String
Private mLastError As String

' When the renderer is driven from outside Excel, an untrapped error parks the VBE
' in break mode and the caller just sees a hang with nothing to read. The trace
' file is written as work proceeds, so the last line in it names the step that
' stalled even when nothing can be got back through COM.
Private Const TRACE_ON As Boolean = True

Private Sub Trace(ByVal s As String)
    If Not TRACE_ON Then Exit Sub
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\corridor_trace.txt" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & "  " & s
    Close #f
    On Error GoTo 0
End Sub

Public Sub CorridorTraceReset()
    On Error Resume Next
    Kill Environ$("TEMP") & "\corridor_trace.txt"
    On Error GoTo 0
End Sub


' ================================================================================
'  PUBLIC ENTRY
' ================================================================================
Public Function RenderCorridorPng(Optional ByVal outPath As String = "", _
                                  Optional ByVal band As String = "all") As String
    Dim ss As Worksheet, dt As Worksheet
    Dim wasProt As Boolean
    Dim prevUpdating As Boolean
    Dim drawn As Boolean
    Dim wantCorridor As Boolean, wantOps As Boolean

    RenderCorridorPng = ""
    mLastError = ""
    mStage = "start"
    mTypeScale = 1#
    On Error GoTo Fail

    band = LCase$(Trim$(band))
    If band = "" Then band = "all"
    wantCorridor = (band = "all" Or band = "corridor")
    wantOps = (band = "all" Or band = "ops")
    If Not wantCorridor And Not wantOps Then Err.Raise 5, , "band must be all|corridor|ops"

    mStage = "resolve sheets"
    On Error Resume Next
    Set ss = ThisWorkbook.Worksheets(MDL_PlanGauge.PG_SlidesheetName)
    Set dt = ThisWorkbook.Worksheets("Data")
    On Error GoTo Fail
    If ss Is Nothing Then Err.Raise 5, , "Slidesheet not found"
    If dt Is Nothing Then Err.Raise 5, , "Data sheet not found"

    If outPath = "" Then
        outPath = Environ$("TEMP") & "\corridor_" & band & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".png"
    End If

    mStage = "BuildScene"
    If Not BuildScene(ss) Then Err.Raise 5, , "BuildScene returned False"

    If wantCorridor And wantOps Then
        mOpsOrigin = CORRIDOR_H
        mCanvasH = CORRIDOR_H + MeasureOpsHeight(dt)
    ElseIf wantCorridor Then
        mOpsOrigin = CORRIDOR_H
        mCanvasH = CORRIDOR_H
    Else
        mOpsOrigin = 0
        mCanvasH = MeasureOpsHeight(dt)
    End If

    prevUpdating = Application.ScreenUpdating
    Set mWs = ss
    mNameN = 0
    mSeq = 0
    ReDim mNames(0 To 2047)

    mStage = "unprotect"
    wasProt = MDL_SheetProtect.SheetUnprotectForVba(ss)
    drawn = True
    Application.ScreenUpdating = False

    mStage = "DrawCanvas": DrawCanvas
    If wantCorridor Then
        mTypeScale = TYPE_SCALE_META
        mStage = "DrawRoom": DrawRoom
        mStage = "DrawSection": DrawSection ss
        mStage = "DrawDay": DrawDay dt
        mStage = "DrawMetrics": DrawMetrics ss, dt
    End If
    If wantOps Then
        mTypeScale = TYPE_SCALE_OPS
        mStage = "DrawOps": DrawOps dt
    End If
    mTypeScale = 1#

    ' Keep ScreenUpdating off through export so sheet activates / chart paste do
    ' not paint. CopyPicture still works against the shape model.
    mStage = "ExportGroup"
    RenderCorridorPng = ExportGroup(outPath)
    If RenderCorridorPng = "" Then mLastError = "ExportGroup produced no file"

    mStage = "cleanup"
    GoTo Done

Fail:
    mLastError = "stage=" & mStage & " err=" & Err.Number & " " & Err.Description
    RenderCorridorPng = ""

Done:
    On Error Resume Next
    If drawn Then
        DeleteDrawn
        MDL_SheetProtect.SheetReprotectAfterVba ss, wasProt
    End If
    Application.ScreenUpdating = prevUpdating
    mTypeScale = 1#
    Set mWs = Nothing
End Function

' Whatever went wrong on the last RenderCorridorPng call, "" if it succeeded.
Public Function CorridorLastError() As String
    CorridorLastError = mLastError
End Function

' A one-line description of the scene, for checking the picture against the sheet
' without having to open the PNG.
Public Function CorridorSceneInfo() As String
    CorridorSceneInfo = "window " & Format$(mMD0, "#,##0") & "-" & Format$(mMD1, "#,##0") & _
        " mMD " & mSection & "->" & mTgtName & IIf(mShaft, " shaft 1:1", "") & "; " & mSCount & " stations" & _
        IIf(mUseSailBands, "; geo " & ChrW(177) & Format$(mGeoHalf, "0.00") & _
        " m; lateral " & ChrW(177) & Format$(mLatTol, "0.00") & " m", "; no sail bands") & "; last svy " & _
        Format$(mSvyMD, "#,##0.00") & " inc " & Format$(mSvyInc, "0.00") & _
        " dev " & Format$(mSvyDev, "0.00") & " lat " & Format$(mSvyLat, "0.00") & _
        "; next wp " & Format$(mWpMD, "#,##0") & " to go " & Format$(mToGo, "0.0") & _
        " req " & Format$(mReqInc, "0.00") & " shapes " & mNameN
End Function

' Proves the VBE is honouring On Error rather than breaking on every error, which
' otherwise makes an unattended render look like a hang.
Public Function CorridorProbe() As String
    On Error GoTo Trapped
    CorridorProbe = "NOT TRAPPED"
    Err.Raise 5, , "probe"
    Exit Function
Trapped:
    CorridorProbe = "trapped ok: " & Err.Description
End Function

' Runs the render up to a named stage and stops, so a failure can be bisected
' without the VBE. stopAfter: scene | canvas | room | section | day | metrics | export
Public Function CorridorDiag(ByVal stopAfter As String) As String
    Dim ss As Worksheet, dt As Worksheet
    Dim wasProt As Boolean, drawn As Boolean
    Dim png As String

    mLastError = ""
    mStage = "start"
    On Error GoTo Fail

    Set ss = ThisWorkbook.Worksheets(MDL_PlanGauge.PG_SlidesheetName)
    Set dt = ThisWorkbook.Worksheets("Data")

    mStage = "scene"
    If Not BuildScene(ss) Then Err.Raise 5, , "BuildScene returned False"
    mOpsOrigin = CORRIDOR_H
    mCanvasH = CORRIDOR_H + MeasureOpsHeight(dt)
    mTypeScale = 1#
    If stopAfter = "scene" Then GoTo Done

    Set mWs = ss
    mNameN = 0: mSeq = 0
    ReDim mNames(0 To 2047)
    wasProt = MDL_SheetProtect.SheetUnprotectForVba(ss)
    drawn = True

    mStage = "canvas": DrawCanvas
    If stopAfter = "canvas" Then GoTo Done
    mStage = "room": DrawRoom
    If stopAfter = "room" Then GoTo Done
    mStage = "section": DrawSection ss
    If stopAfter = "section" Then GoTo Done
    mStage = "day": DrawDay dt
    If stopAfter = "day" Then GoTo Done
    mStage = "metrics": DrawMetrics ss, dt
    If stopAfter = "metrics" Then GoTo Done
    mStage = "ops": DrawOps dt
    If stopAfter = "ops" Then GoTo Done

    mStage = "export"
    png = ExportGroup(Environ$("TEMP") & "\corridor_diag.png")
    If png = "" Then Err.Raise 5, , "ExportGroup produced no file"

Done:
    CorridorDiag = "OK through " & mStage & " | " & CorridorSceneInfo()
    GoTo Finish

Fail:
    CorridorDiag = "FAIL at " & mStage & " | err " & Err.Number & " " & Err.Description

Finish:
    On Error Resume Next
    If drawn Then
        DeleteDrawn
        MDL_SheetProtect.SheetReprotectAfterVba ss, wasProt
    End If
    Set mWs = Nothing
End Function

' Dev helper: render once and say where it landed.
Public Sub TestRenderCorridor()
    Dim p As String: p = RenderCorridorPng()
    If p = "" Then
        MsgBox "Corridor render failed." & vbCrLf & CorridorLastError(), vbExclamation
    Else
        MsgBox "Corridor PNG written to:" & vbCrLf & p & vbCrLf & vbCrLf & CorridorSceneInfo(), vbInformation
    End If
End Sub


' ================================================================================
'  SCENE
' ================================================================================
Private Function BuildScene(ss As Worksheet) As Boolean
    BuildScene = False
    Trace "BuildScene enter"

    mGeoHalf = MDL_PlanGauge.PG_WaypointHalfWidth(ss)
    If mGeoHalf <= 0# Then mGeoHalf = 2#
    Trace "  geoHalf=" & mGeoHalf
    mLatTol = Abs(LatTolFromCell(ss.Range(SS_LAT_TOL_ADDR)))
    If mLatTol <= 0# Then mLatTol = 10#   ' last-resort default if AA14 empty
    mLatBox = mLatTol * 1.25
    Trace "  latTol(AA14)=" & mLatTol

    LoadWaypoints ss
    MergePlanTargetsAsWaypoints ss
    Trace "  waypoints=" & mWCount
    If Not WalkSurveys(ss, True) Then Exit Function      ' pass 1: find the last survey
    Trace "  pass1 done svyMD=" & mSvyMD & " bitMD=" & mBitMD

    Dim i As Long
    Dim tMd As Double, tTvd As Double, tInc As Double, tName As String
    mTgtName = "WP"
    If NextPlanTarget(ss, mSvyMD, tMd, tTvd, tInc, tName) Then
        mWpMD = tMd: mWpTvd = tTvd: mTgtName = tName
    ElseIf mWCount > 0 Then
        mWpMD = mWMD(mWCount - 1): mWpTvd = mWTvd(mWCount - 1)
        For i = 0 To mWCount - 1
            If mWMD(i) > mSvyMD Then
                mWpMD = mWMD(i): mWpTvd = mWTvd(i)
                Exit For
            End If
        Next i
    Else
        mWpMD = mSvyMD + 80#: mWpTvd = mSvyTvd
    End If
    If mWpTvd = 0# Then
        If Not MDL_PlanGauge.PG_WaypointTvdAtMd(ss, mWpMD, mWpTvd) Then
            mWpTvd = mSvyTvd
        End If
    End If

    ClassifySection mSvyInc
    FrameWindowToTarget
    ApplySectionLayout
    Trace "  section=" & mSection & " window " & mMD0 & " to " & mMD1 & " next=" & mTgtName & " " & mWpMD

    If Not WalkSurveys(ss, False) Then Exit Function      ' pass 2: collect the window
    Trace "  pass2 done stations=" & mSCount
    ' Vertical / early build sit above the sail table. If the short target
    ' window captured <2 stations, open back to MD 0 so the room has a hole.
    If mSCount < 2 Then
        mMD0 = 0#
        If mMD1 < mSvyMD + RUN_AHEAD Then mMD1 = mSvyMD + RUN_AHEAD
        ApplySectionLayout
        If Not WalkSurveys(ss, False) Then Exit Function
        Trace "  widened stations=" & mSCount
    End If
    If mSCount < 2 Then
        Trace "  abort: still <2 stations"
        Exit Function
    End If

    mShaft = (mSection <> "LATERAL")
    mUseSailBands = (mSection = "LATERAL")

    If Not MDL_PlanGauge.PG_WaypointTvdAtMd(ss, mMD0, mTvdDatum) Then
        mTvdDatum = mSTvd(0)
    End If

    Dim py As Double, gT As Double, gHi As Double, gLo As Double, mid As Double
    Dim maxLat As Double, tvdSpan As Double
    If mShaft Then
        ' True TVD-down shaft. Sail geo / AA14 bands are lateral-only.
        ' L/R is 1:1 with TVD so a few metres off plan is a hair, not a corridor.
        mDevHi = mSTvd(0): mDevLo = mSTvd(0): maxLat = 0#
        For i = 0 To mSCount - 1
            If mSTvd(i) < mDevHi Then mDevHi = mSTvd(i)
            If mSTvd(i) > mDevLo Then mDevLo = mSTvd(i)
            If Abs(mSLat(i)) > maxLat Then maxLat = Abs(mSLat(i))
        Next i
        If mWpTvd > 0# Then
            If mWpTvd < mDevHi Then mDevHi = mWpTvd
            If mWpTvd > mDevLo Then mDevLo = mWpTvd
        End If
        mDevHi = mDevHi - 8#
        mDevLo = mDevLo + 8#
        If mDevLo - mDevHi < 40# Then
            mid = (mDevHi + mDevLo) / 2#
            mDevHi = mid - 20#: mDevLo = mid + 20#
        End If
        mTvdTop = mDevHi
        tvdSpan = mDevLo - mDevHi
        ' Wide box vs a few metres of offset: 5 m off plan is a hair, not a tunnel.
        mLatBox = tvdSpan * 0.4
        If mLatBox < 50# Then mLatBox = 50#
        If mLatBox > 140# Then mLatBox = 140#
        mSV = (mVH - 80#) / tvdSpan
        mKzx = mSV
        mSxd = mSV * 0.08
        mKzy = mSV * 0.12
        mX0 = mVX + mVW * 0.52
        mY0 = mVY + 48#
        Trace "  shaft tvd " & mDevHi & "-" & mDevLo & " latBox=" & mLatBox & " maxLat=" & maxLat
    Else
        mDevHi = -1E+99: mDevLo = 1E+99
        For i = 0 To mSCount - 1
            py = PlotY(mSTvd(i))
            If py > mDevHi Then mDevHi = py
            If py < mDevLo Then mDevLo = py
        Next i
        For i = 0 To mWCount - 1
            If mWMD(i) >= mMD0 - 0.1 And mWMD(i) <= mMD1 + 0.1 Then
                gHi = PlotY(mWTvd(i) - mGeoHalf)
                gLo = PlotY(mWTvd(i) + mGeoHalf)
                If gHi > mDevHi Then mDevHi = gHi
                If gLo < mDevLo Then mDevLo = gLo
            End If
        Next i
        If MDL_PlanGauge.PG_WaypointTvdAtMd(ss, mMD0, gT) Then
            gHi = PlotY(gT - mGeoHalf): gLo = PlotY(gT + mGeoHalf)
            If gHi > mDevHi Then mDevHi = gHi
            If gLo < mDevLo Then mDevLo = gLo
        End If
        If MDL_PlanGauge.PG_WaypointTvdAtMd(ss, mMD1, gT) Then
            gHi = PlotY(gT - mGeoHalf): gLo = PlotY(gT + mGeoHalf)
            If gHi > mDevHi Then mDevHi = gHi
            If gLo < mDevLo Then mDevLo = gLo
        End If
        mDevHi = mDevHi + 1#
        mDevLo = mDevLo - 1#
        If mDevHi - mDevLo < 8# Then
            mid = (mDevHi + mDevLo) / 2#
            mDevHi = mid + 4#: mDevLo = mid - 4#
        End If
        mSV = mDevSpan / (mDevHi - mDevLo)
        mX0 = mVX + 70
        mY0 = mVY + 42.5 + mDevHi * mSV + mLatBox * mKzy
    End If

    mToGo = mWpMD - mSvyMD
    If mToGo <= 0# Then mToGo = 1#
    Dim c As Double
    c = (mWpTvd - mSvyTvd) / mToGo
    If c > 1# Then c = 1#
    If c < -1# Then c = -1#
    mReqInc = WorksheetFunction.Acos(c) * 180# / PIE
    mHoldDevEnd = mWpTvd - (mSvyTvd + mToGo * Cos(mSvyInc * PIE / 180#))
    mHoldLatEnd = mSvyLat + mToGo * Sin((mSvyAzi - mSvyPlanAzi) * PIE / 180#)

    Trace "BuildScene ok " & CorridorSceneInfo()
    BuildScene = True
End Function

Private Function LoadWaypoints(ss As Worksheet) As Boolean
    LoadWaypoints = False
    ReDim mWMD(0 To SS_WP_ROW_LAST - SS_WP_ROW_FIRST + 12)
    ReDim mWTvd(0 To SS_WP_ROW_LAST - SS_WP_ROW_FIRST + 12)
    ReDim mWInc(0 To SS_WP_ROW_LAST - SS_WP_ROW_FIRST + 12)
    mWCount = 0

    Dim r As Long, vM As Variant, vT As Variant, vI As Variant
    For r = SS_WP_ROW_FIRST To SS_WP_ROW_LAST
        vM = ss.Cells(r, SS_WP_COL_MD).Value2
        vT = ss.Cells(r, SS_WP_COL_TVD).Value2
        vI = ss.Cells(r, SS_WP_COL_INC).Value2
        If Not IsArray(vM) And Not IsArray(vT) Then
            If IsNumeric(vM) And IsNumeric(vT) Then
                If Len(Trim$(CStr(vM & ""))) > 0 And Len(Trim$(CStr(vT & ""))) > 0 Then
                    Dim keep As Boolean
                    If mWCount = 0 Then keep = True Else keep = (CDbl(vM) > mWMD(mWCount - 1))
                    If keep Then
                        mWMD(mWCount) = CDbl(vM)
                        mWTvd(mWCount) = CDbl(vT)
                        If IsNumeric(vI) And Len(Trim$(CStr(vI & ""))) > 0 Then
                            mWInc(mWCount) = CDbl(vI)
                        Else
                            mWInc(mWCount) = 0#
                        End If
                        mWCount = mWCount + 1
                    End If
                End If
            End If
        End If
    Next r
    LoadWaypoints = True
End Function

' Named stations in the T2:Y5 window (U=MD, V=INC, X=TVD, Y=name).
Private Function NextPlanTarget(ss As Worksheet, ByVal bitMD As Double, _
        ByRef tMd As Double, ByRef tTvd As Double, ByRef tInc As Double, _
        ByRef tName As String) As Boolean
    Dim r As Long, vMd As Variant, nm As String
    NextPlanTarget = False
    For r = 2 To 5
        nm = Trim$(CStr(ss.Cells(r, 25).Value2 & ""))
        vMd = ss.Cells(r, 21).Value2
        If Len(nm) = 0 Then GoTo NextTgt
        If Not IsNumeric(vMd) Then GoTo NextTgt
        If CDbl(vMd) <= bitMD + 0.005 Then GoTo NextTgt
        tMd = CDbl(vMd)
        tInc = NumCell(ss.Cells(r, 22))
        tTvd = NumCell(ss.Cells(r, 24))
        tName = nm
        NextPlanTarget = True
        Exit Function
NextTgt:
    Next r
End Function

Private Sub MergePlanTargetsAsWaypoints(ss As Worksheet)
    Dim r As Long, vMd As Variant, vT As Variant, md As Double, tvd As Double
    Dim i As Long, j As Long, keep As Boolean
    For r = 2 To 5
        vMd = ss.Cells(r, 21).Value2
        vT = ss.Cells(r, 24).Value2
        If Not (IsNumeric(vMd) And IsNumeric(vT)) Then GoTo NextMerge
        If Len(Trim$(CStr(vMd & ""))) = 0 Then GoTo NextMerge
        md = CDbl(vMd): tvd = CDbl(vT)
        keep = True
        For i = 0 To mWCount - 1
            If Abs(mWMD(i) - md) < 0.5 Then keep = False
        Next i
        If keep Then
            i = mWCount
            Do While i > 0
                If mWMD(i - 1) <= md Then Exit Do
                i = i - 1
            Loop
            For j = mWCount To i + 1 Step -1
                mWMD(j) = mWMD(j - 1)
                mWTvd(j) = mWTvd(j - 1)
                mWInc(j) = mWInc(j - 1)
            Next j
            mWMD(i) = md
            mWTvd(i) = tvd
            mWInc(i) = NumCell(ss.Cells(r, 22))
            mWCount = mWCount + 1
        End If
NextMerge:
    Next r
End Sub

Private Sub ClassifySection(ByVal incDeg As Double)
    If incDeg < 5# Then
        mSection = "VERTICAL"
    ElseIf incDeg < 80# Then
        mSection = "BUILD"
    Else
        mSection = "LATERAL"
    End If
End Sub

Private Sub FrameWindowToTarget()
    Dim lookBack As Double, maxWin As Double, minWin As Double
    Select Case mSection
        Case "VERTICAL"
            lookBack = 80#: maxWin = 400#: minWin = 60#
        Case "BUILD"
            lookBack = 250#: maxWin = 800#: minWin = 120#
        Case Else
            lookBack = 800#: maxWin = WINDOW_LEN: minWin = 200#
    End Select
    mMD1 = mWpMD + RUN_AHEAD
    If mMD1 < mSvyMD + 15# Then mMD1 = mSvyMD + 15#
    mMD0 = mSvyMD - lookBack
    If mMD0 < 0# Then mMD0 = 0#
    If mMD1 - mMD0 > maxWin Then mMD0 = mMD1 - maxWin
    If mMD1 - mMD0 < minWin Then
        mMD0 = mMD1 - minWin
        If mMD0 < 0# Then mMD0 = 0#
    End If
End Sub

Private Sub ApplySectionLayout()
    Dim span As Double
    mVY = V_Y
    Select Case mSection
        Case "VERTICAL"
            mVX = 12#: mVW = 560#: mVH = 368#
            mPX = 584#: mPY = 30#: mPW = 176#: mPH = 196#
            mQY = 234#: mQH = 164#
            mMX = 772#: mMW = 239#: mMH = 376#
            mKzx = 2.2: mKzy = 2#: mDevSpan = 280#
        Case "BUILD"
            mVX = 12#: mVW = 520#: mVH = 368#
            mPX = 544#: mPY = 30#: mPW = 196#: mPH = 214#
            mQY = 252#: mQH = 146#
            mMX = 752#: mMW = 259#: mMH = 376#
            mKzx = 2.6: mKzy = 2.1: mDevSpan = 240#
        Case Else
            mVX = V_X: mVW = V_W: mVH = V_H
            mPX = P_X: mPY = P_Y: mPW = P_W: mPH = P_H
            mQY = Q_Y: mQH = Q_H
            mMX = M_X: mMW = M_W: mMH = M_H
            mKzx = KZX: mKzy = KZY: mDevSpan = DEV_SPAN_PT
    End Select
    span = mMD1 - mMD0
    If span < 40# Then span = 40#
    mSxd = (mVW - 90#) / span
    If mSxd < 0.16 Then mSxd = 0.16
    If mSxd > 4# Then mSxd = 4#
End Sub

' Integrates the survey block once. findLastOnly stops at recording the final
' station; otherwise every station inside the window is captured as well.
Private Function WalkSurveys(ss As Worksheet, ByVal findLastOnly As Boolean) As Boolean
    Dim pMD() As Double
    Dim pInc() As Double
    Dim pAzi() As Double
    Dim pTvd() As Double
    Dim pNS() As Double
    Dim pEW() As Double
    Dim np As Long
    Dim prevMD As Double
    Dim prevInc As Double
    Dim prevAzi As Double
    Dim curN As Double
    Dim curE As Double
    Dim curV As Double
    Dim lastRow As Long
    Dim r As Long
    Dim rowFirst As Long
    Dim rowLast As Long
    Dim vMd As Variant
    Dim vInc As Variant
    Dim vAzi As Variant
    Dim mdv As Double
    Dim incv As Double
    Dim aziv As Double
    Dim dVert As Double
    Dim dNorth As Double
    Dim dEast As Double
    Dim dev As Double
    Dim lat As Double
    Dim ok As Boolean
    ' Plan position at the station. Do not shorten these to oN / oE: VBA is
    ' case-insensitive, so a variable called oN is the reserved word On and the
    ' whole procedure fails to compile with a bare "Syntax error".
    Dim planN As Double
    Dim planE As Double
    Dim planV As Double
    Dim planAzi As Double
    Dim planInc As Double
    Dim px As Double
    Dim py As Double
    Dim grav As Boolean
    Dim wt As Double

    WalkSurveys = False
    Trace "  WalkSurveys findLastOnly=" & findLastOnly

    np = MDL_PlanGauge.PG_LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
    Trace "  plan points=" & np
    If np < 2 Then Exit Function

    If Not findLastOnly Then
        ReDim mSMD(0 To 511)
        ReDim mSDev(0 To 511)
        ReDim mSTvd(0 To 511)
        ReDim mSLat(0 To 511)
        ReDim mSNS(0 To 511)
        ReDim mSEW(0 To 511)
        mSCount = 0
    End If

    rowFirst = MDL_PlanGauge.PG_SurvRowFirst
    rowLast = MDL_PlanGauge.PG_SurvRowLast

    For r = rowFirst To rowLast
        If Not MDL_PlanGauge.PG_IsSurveySummaryRow(ss, r) Then
            vMd = ss.Cells(r, 5).Value2
            vInc = ss.Cells(r, 6).Value2
            vAzi = ss.Cells(r, 7).Value2

            ok = IsNumeric(vMd) And IsNumeric(vInc) And IsNumeric(vAzi)
            If ok Then
                ok = Len(CStr(vMd & "")) > 0
            End If
            If ok Then
                ok = Len(CStr(vInc & "")) > 0
            End If
            If ok Then
                ok = Len(CStr(vAzi & "")) > 0
            End If

            If ok Then
                mdv = CDbl(vMd)
                If mdv > prevMD Then
                    incv = CDbl(vInc)
                    aziv = CDbl(vAzi)
                    MDL_PlanGauge.PG_McStep prevMD, prevInc, prevAzi, mdv, incv, aziv, dVert, dNorth, dEast
                    curN = curN + dNorth
                    curE = curE + dEast
                    curV = curV + dVert
                    prevMD = mdv
                    prevInc = incv
                    prevAzi = aziv
                    lastRow = r

                    If Not findLastOnly Then
                        If mdv >= mMD0 And mdv <= mMD1 And mSCount <= 511 Then
                            If StationOffsets(ss, mdv, incv, curN, curE, curV, np, pMD, pInc, pAzi, pTvd, pNS, pEW, dev, lat) Then
                                mSMD(mSCount) = mdv
                                mSDev(mSCount) = dev
                                mSLat(mSCount) = lat
                                mSTvd(mSCount) = curV
                                mSNS(mSCount) = curN
                                mSEW(mSCount) = curE
                                mSCount = mSCount + 1
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next r

    If prevMD <= 0# Then Exit Function

    If findLastOnly Then
        mSvyMD = prevMD
        mSvyInc = prevInc
        mSvyAzi = prevAzi
        mSvyTvd = curV
        mSvyNS = curN
        mSvyEW = curE

        MDL_PlanGauge.PG_PlanAt prevMD, np, pMD, pInc, pAzi, pTvd, pNS, pEW, planN, planE, planV, planAzi, planInc
        mSvyPlanAzi = planAzi

        grav = (prevInc >= MDL_PlanGauge.PG_CrossoverInc)
        MDL_PlanGauge.PG_FrameComponents grav, planAzi, curN - planN, curE - planE, planV - curV, px, py
        mSvyLat = px

        If MDL_PlanGauge.PG_WaypointTvdAtMd(ss, prevMD, wt) Then
            mSvyDev = wt - curV
        End If

        mBitMD = NumCell(ss.Cells(lastRow, SS_BIT_COL))
        If mBitMD < prevMD Then mBitMD = prevMD
    End If

    WalkSurveys = True
End Function

' Up/down against the geo target line and left/right against the plan, for one
' station. Returns False when the station sits outside the waypoint table, which
' is the case above the first waypoint - plotting it would draw a phantom zero.
Private Function StationOffsets(ss As Worksheet, ByVal md As Double, ByVal inc As Double, ByVal curN As Double, ByVal curE As Double, ByVal curV As Double, ByVal np As Long, pMD() As Double, pInc() As Double, pAzi() As Double, pTvd() As Double, pNS() As Double, pEW() As Double, ByRef outDev As Double, ByRef outLat As Double) As Boolean
    Dim wt As Double
    Dim planN As Double
    Dim planE As Double
    Dim planV As Double
    Dim planAzi As Double
    Dim planInc As Double
    Dim px As Double
    Dim py As Double
    Dim grav As Boolean

    StationOffsets = False

    MDL_PlanGauge.PG_PlanAt md, np, pMD, pInc, pAzi, pTvd, pNS, pEW, planN, planE, planV, planAzi, planInc
    ' Sail waypoints (AC14:AD33) start at the lateral geo window. Above that,
    ' planned vs actual still has a TVD: use the well plan.
    If Not MDL_PlanGauge.PG_WaypointTvdAtMd(ss, md, wt) Then wt = planV
    outDev = wt - curV

    grav = (inc >= MDL_PlanGauge.PG_CrossoverInc)
    MDL_PlanGauge.PG_FrameComponents grav, planAzi, curN - planN, curE - planE, planV - curV, px, py
    outLat = px
    StationOffsets = True
End Function


' ================================================================================
'  PROJECTION
' ================================================================================
Private Function ObX(ByVal md As Double, ByVal lat As Double) As Double
    If mShaft Then
        ObX = mX0 + lat * mKzx + (md - mMD0) * mSxd
    Else
        ObX = mX0 + (md - mMD0) * mSxd + lat * mKzx
    End If
End Function

Private Function ObY(ByVal plot As Double, ByVal lat As Double) As Double
    If mShaft Then
        ObY = mY0 + (plot - mTvdTop) * mSV + lat * mKzy
    Else
        ObY = mY0 - plot * mSV - lat * mKzy
    End If
End Function

' Tunnel: screen-up metres vs geo TVD at window start (positive = shallower).
' Shaft: plot coordinate is true TVD (positive down the page via ObY).
Private Function PlotY(ByVal tvd As Double) As Double
    If mShaft Then
        PlotY = tvd
    Else
        PlotY = mTvdDatum - tvd
    End If
End Function

' Shaft: plot is TVD, mDevHi = shallow, mDevLo = deep.
' Tunnel: plot is screen-up vs geo, mDevHi > mDevLo.
Private Function PlotInRoom(ByVal plot As Double) As Boolean
    If mShaft Then
        PlotInRoom = (plot >= mDevHi And plot <= mDevLo)
    Else
        PlotInRoom = (plot >= mDevLo And plot <= mDevHi)
    End If
End Function


' ================================================================================
'  DRAWING - PANEL 1, THE ROOM
' ================================================================================
Private Sub DrawCanvas()
    If mCanvasH < 40# Then mCanvasH = 40#
    ' Full report is one continuous paper canvas (corridor + ops).
    Rect 0, 0, CANVAS_W, mCanvasH, H("FFFFFF"), -1, 0#
End Sub

Private Sub DrawRoom()
    If mShaft Then
        DrawShaftProfile
        Exit Sub
    End If

    Dim i As Long
    Dim plot As Double, endGeo As Double, endHi As Double

    Tx mVX, 20, "WELLBORE CORRIDOR " & ChrW(8212) & " " & mSection & " to " & mTgtName & _
       "  " & Format$(mMD0, "#,##0") & " to " & Format$(mMD1, "#,##0") & " mMD", _
       12.5, H("FFFFFF"), "start", True
    Rect mVX, mVY, mVW, mVH, H("222222"), H("3A3A3A"), 0.75

    ' ---- back wall, at lat = +mLatBox -----------------------------------------
    Quad mMD0, mDevHi, mLatBox, mMD1, mDevHi, mLatBox, _
         mMD1, mDevLo, mLatBox, mMD0, mDevLo, mLatBox, -1, H("4A4A4A"), 0.75

    ' ---- left wall, at lat = -mLatBox : looking down the hole -----------------
    Quad mMD0, mDevHi, -mLatBox, mMD1, mDevHi, -mLatBox, _
         mMD1, mDevLo, -mLatBox, mMD0, mDevLo, -mLatBox, -1, H("4A4A4A"), 0.75

    Dim dv As Double, tick As Double, dvStep As Double
    If mShaft Then
        tick = 50#
        If (mDevLo - mDevHi) > 400# Then tick = 100#
        dvStep = tick
    Else
        tick = 2#
        dvStep = -tick
    End If
    For dv = Int(mDevHi / tick) * tick To mDevLo Step dvStep
        Dim isZero As Boolean: isZero = (Abs(dv) < 0.001)
        Ln ObX(mMD0, mLatBox), ObY(dv, mLatBox), ObX(mMD1, mLatBox), ObY(dv, mLatBox), _
           IIf(isZero, H("555555"), H("3A3A3A")), 0.7, IIf(isZero, msoLineDash, msoLineSysDot)
        Ln ObX(mMD0, -mLatBox), ObY(dv, -mLatBox), ObX(mMD1, -mLatBox), ObY(dv, -mLatBox), _
           IIf(isZero, H("444444"), H("333333")), 0.6, IIf(isZero, msoLineDash, msoLineSysDot)
        Ln ObX(mMD0, -mLatBox), ObY(dv, -mLatBox), ObX(mMD0, mLatBox), ObY(dv, mLatBox), _
           H("3A3A3A"), 0.6, msoLineSolid
        Tx ObX(mMD0, -mLatBox) - 5, ObY(dv, -mLatBox) + 3, _
           IIf(mShaft, Format$(dv, "#,##0"), SignedM(dv)), 8.5, H("A8A8A8"), "end"
    Next dv

    If mUseSailBands Then
        DrawGeoCorridorRibbon mLatBox, H("1E3A2F"), H("4CAF7A"), 1.15
        DrawGeoCorridorRibbon -mLatBox, H("1A3328"), H("3D6B54"), 1#
        If Not MDL_PlanGauge.PG_WaypointTvdAtMd(mWs, mMD1, endGeo) Then endGeo = mWpTvd
        endHi = PlotY(endGeo - mGeoHalf)
        Tx ObX(mMD1, mLatBox) - 4, ObY(endHi, mLatBox) - 4, _
           "geo " & ChrW(177) & Format$(mGeoHalf, "0.00") & " m", 8, H("4CAF7A"), "end"
    End If

    Quad mMD0, mDevLo, -mLatBox, mMD1, mDevLo, -mLatBox, _
         mMD1, mDevLo, mLatBox, mMD0, mDevLo, mLatBox, -1, H("4A4A4A"), 0.75
    If mUseSailBands Then
        Quad mMD0, mDevLo, -mLatTol, mMD1, mDevLo, -mLatTol, _
             mMD1, mDevLo, mLatTol, mMD0, mDevLo, mLatTol, H("1E3A2F"), H("4CAF7A"), 1#
    End If
    Ln ObX(mMD0, 0), ObY(mDevLo, 0), ObX(mMD1, 0), ObY(mDevLo, 0), H("4CAF7A"), 0.7, msoLineDash

    If mUseSailBands Then
        Dim latLbl As String: latLbl = Format$(mLatTol, "0.00")
        Tx ObX(mMD1, -mLatTol) + 4, ObY(mDevLo, -mLatTol) + 3, _
           latLbl & " R", 8, H("4CAF7A"), "start"
        Tx ObX(mMD1, 0) + 4, ObY(mDevLo, 0) + 3, "plan", 8, H("4CAF7A"), "start"
        Tx ObX(mMD1, mLatTol) + 4, ObY(mDevLo, mLatTol) + 3, _
           latLbl & " L", 8, H("4CAF7A"), "start"
    Else
        Tx ObX(mMD1, 0) + 4, ObY(mDevLo, 0) + 3, "plan", 8, H("A8A8A8"), "start"
    End If

    Quad mMD0, mDevHi, -mLatBox, mMD0, mDevHi, mLatBox, _
         mMD0, mDevLo, mLatBox, mMD0, mDevLo, -mLatBox, -1, H("4A4A4A"), 0.75
    If mUseSailBands Then
        Quad mMD0, mGeoHalf, -mLatTol, mMD0, mGeoHalf, mLatTol, _
             mMD0, -mGeoHalf, mLatTol, mMD0, -mGeoHalf, -mLatTol, H("1E3A2F"), H("4CAF7A"), 1.2
    End If

    ' ---- room wireframe: ceiling and the open front edges ----------------------
    Ln ObX(mMD0, -mLatBox), ObY(mDevHi, -mLatBox), ObX(mMD1, -mLatBox), ObY(mDevHi, -mLatBox), H("4A4A4A"), 0.75, msoLineSolid
    Ln ObX(mMD1, -mLatBox), ObY(mDevHi, -mLatBox), ObX(mMD1, mLatBox), ObY(mDevHi, mLatBox), H("4A4A4A"), 0.75, msoLineSolid
    Ln ObX(mMD1, -mLatBox), ObY(mDevHi, -mLatBox), ObX(mMD1, -mLatBox), ObY(mDevLo, -mLatBox), H("4A4A4A"), 0.75, msoLineSolid
    Ln ObX(mMD0, -mLatBox), ObY(mDevLo, -mLatBox), ObX(mMD1, -mLatBox), ObY(mDevLo, -mLatBox), H("4A4A4A"), 0.75, msoLineSolid

    ' ---- shadows --------------------------------------------------------------
    ' Back wall (+lat) = the only TVD path shadow. The near/front wall keeps its
    ' axis labels but no path shadow, and there are no vertical floor tracers.
    Dim xs() As Double, ys() As Double
    ReDim xs(0 To mSCount - 1): ReDim ys(0 To mSCount - 1)

    For i = 0 To mSCount - 1
        plot = PlotY(mSTvd(i))
        xs(i) = ObX(mSMD(i), mLatBox): ys(i) = ObY(plot, mLatBox)
    Next i
    PolyLine xs, ys, mSCount, H("A8CBB8"), 2.4, msoLineDash

    ' Floor shadow: left/right from plan (data sign: + = Right of plan).
    For i = 0 To mSCount - 1
        xs(i) = ObX(mSMD(i), LatScr(mSLat(i))): ys(i) = ObY(mDevLo, LatScr(mSLat(i)))
    Next i
    PolyLine xs, ys, mSCount, H("B6DBC8"), 3#, msoLineSolid

    If mSCount >= 1 Then
        Dim iLast As Long: iLast = mSCount - 1
        Dot ObX(mSMD(iLast), LatScr(mSLat(iLast))), ObY(mDevLo, LatScr(mSLat(iLast))), _
            3.4, H("B6DBC8"), H("1A1A1A"), 1
        Tx ObX(mSMD(iLast), LatScr(mSLat(iLast))), ObY(mDevLo, LatScr(mSLat(iLast))) + 12, _
           Format$(Abs(mSLat(iLast)), "0.0") & " m " & IIf(mSLat(iLast) < 0#, "L", "R"), _
           8.5, H("B6DBC8"), "middle"
    End If

    ' ---- the hole itself, floating in the middle of the room -------------------
    For i = 0 To mSCount - 1
        plot = PlotY(mSTvd(i))
        xs(i) = ObX(mSMD(i), LatScr(mSLat(i))): ys(i) = ObY(plot, LatScr(mSLat(i)))
    Next i
    PolyLine xs, ys, mSCount, H("FFFFFF"), 2.8, msoLineSolid

    ' Plan path on top of the hole so P2 reads clearly through the room.
    DrawPlanPathInRoom

    ' Geo waypoints: MD + Inc to next (AE). Attitude ticks on the left-wall shadow.
    DrawWaypointMarks

    Dim iWorstDev As Long, iWorstLat As Long
    iWorstDev = 0: iWorstLat = 0
    For i = 0 To mSCount - 1
        Dim bust As Boolean
        If mUseSailBands Then
            bust = (Abs(mSDev(i)) > mGeoHalf) Or (Abs(mSLat(i)) > mLatTol)
        Else
            bust = False
        End If
        Dot xs(i), ys(i), 2.2, IIf(bust, H("FF5555"), H("111111")), -1, 0
        If Abs(mSDev(i)) > Abs(mSDev(iWorstDev)) Then iWorstDev = i
        If Abs(mSLat(i)) > Abs(mSLat(iWorstLat)) Then iWorstLat = i
    Next i

    If mUseSailBands Then
        Dim wx As Double, wy As Double
        wx = ObX(mSMD(iWorstDev), mLatBox): wy = ObY(PlotY(mSTvd(iWorstDev)), mLatBox)
        Ln wx, wy - 5, wx + 14, wy - 18, H("FF5555"), 0.8, msoLineSolid
        Tx wx + 17, wy - 16, Format$(mSDev(iWorstDev), "0.00") & " m high at " & _
           Format$(mSMD(iWorstDev), "#,##0"), 8.5, H("FF5555"), "start"
        Tx ObX(mSMD(iWorstLat), LatScr(mSLat(iWorstLat))), ObY(mDevLo, LatScr(mSLat(iWorstLat))) + 11, _
           Format$(Abs(mSLat(iWorstLat)), "0.0") & " m " & IIf(mSLat(iWorstLat) < 0, "L", "R"), _
           8, H("8AAB9A"), "middle"
    End If

    ' ---- forward projections from the last survey ------------------------------
    Dim n As Long: n = Int(mToGo / 4#) + 2
    Dim rx() As Double, ry() As Double, hx() As Double, hy() As Double
    ReDim rx(0 To n - 1): ReDim ry(0 To n - 1)
    ReDim hx(0 To n - 1): ReDim hy(0 To n - 1)
    Dim k As Long, a As Double, md As Double, tvdProj As Double, latH As Double
    For k = 0 To n - 1
        a = k * 4#
        If a > mToGo Then a = mToGo
        md = mSvyMD + a
        tvdProj = mSvyTvd + a * Cos(mReqInc * PIE / 180#)
        rx(k) = ObX(md, LatScr(mSvyLat * (1# - a / mToGo)))
        ry(k) = ObY(PlotY(tvdProj), LatScr(mSvyLat * (1# - a / mToGo)))
        latH = mSvyLat + a * Sin((mSvyAzi - mSvyPlanAzi) * PIE / 180#)
        tvdProj = mSvyTvd + a * Cos(mSvyInc * PIE / 180#)
        hx(k) = ObX(md, LatScr(latH))
        hy(k) = ObY(PlotY(tvdProj), LatScr(latH))
    Next k
    PolyLine rx, ry, n, H("FF5555"), 2.4, msoLineSolid
    PolyLine hx, hy, n, H("E09A3D"), 2.2, msoLineDash

    Dim bx As Double, by As Double
    bx = ObX(mSvyMD, LatScr(mSvyLat)): by = ObY(PlotY(mSvyTvd), LatScr(mSvyLat))
    Dot bx, by, 5, H("FF5555"), H("1A1A1A"), 1.4
    Tx bx - 10, by - 11, "BIT", 10, H("FF5555"), "end", True

    Dim ax As Double, ay As Double
    tvdProj = mSvyTvd + mToGo * Cos(mSvyInc * PIE / 180#)
    ax = ObX(mWpMD, LatScr(mHoldLatEnd)): ay = ObY(PlotY(tvdProj), LatScr(mHoldLatEnd))
    Dot ax, ay, 4, H("1A1A1A"), H("E09A3D"), 2
    ' Keep held-arrival label clear of the BIT tag.
    If mUseSailBands Then
        Tx ax + 10, ay + 12, Format$(mHoldDevEnd, "0.00") & " m high", 8.5, H("E09A3D"), "start"
    ElseIf Abs(mHoldDevEnd) > 1# Then
        Tx ax + 10, ay + 12, Format$(Abs(mHoldDevEnd), "0.0") & " m TVD if held", 8.5, H("E09A3D"), "start"
    End If

    ' ---- reporting-day span along the front foot of the room -------------------
    ' drawn last so it sits over the floor edge rather than under it
    Dim d0 As Double, d1 As Double, dy As Double
    d0 = ClampMd(NumCell(ThisWorkbook.Worksheets("Data").Range("C4")))
    d1 = ClampMd(NumCell(ThisWorkbook.Worksheets("Data").Range("C5")))
    dy = mVY + mVH - 16
    Rect ObX(d0, -mLatBox), dy - 8, ObX(d1, -mLatBox) - ObX(d0, -mLatBox), 11, _
         H("1A3348"), H("5BA3D9"), 0.75
    Tx (ObX(d0, -mLatBox) + ObX(d1, -mLatBox)) / 2, dy, _
       "24 h " & ChrW(183) & " " & Format$(d0, "#,##0") & " " & ChrW(8594) & " " & _
       Format$(d1, "#,##0") & " m " & ChrW(183) & " " & Format$(d1 - d0, "#,##0") & " m", _
       8.5, H("5BA3D9"), "middle"

    ' ---- captions and legend ---------------------------------------------------
    If mShaft Then
        Tx mVX + 6, mVY + 16, "shaft " & ChrW(8212) & " TVD down, L/R from plan at true scale, m", 8.5, H("A8A8A8"), "start"
        Tx mVX + mVW - 6, mVY + mVH - 44, "floor " & ChrW(8212) & " displacement from plan, m", 8.5, H("A8A8A8"), "end"
    Else
        Tx mVX + 6, mVY + 16, "back wall " & ChrW(8212) & " TVD plot (0 = geo at window start), m", 8.5, H("A8A8A8"), "start"
        Tx mVX + mVW - 6, mVY + mVH - 44, "floor " & ChrW(8212) & " left / right from plan, m", 8.5, H("A8A8A8"), "end"
    End If

    Dim ly As Double: ly = mVY + mVH + 15
    Ln mVX + 2, ly, mVX + 18, ly, H("FFFFFF"), 2.6, msoLineSolid
    Tx mVX + 22, ly + 4, "hole", 10, H("FFFFFF"), "start"
    Ln mVX + 52, ly, mVX + 68, ly, H("5BA3D9"), 2.2, msoLineSolid
    Tx mVX + 72, ly + 4, "plan", 10, H("FFFFFF"), "start"
    Ln mVX + 108, ly, mVX + 124, ly, H("9EC9B0"), 2#, msoLineDash
    Tx mVX + 128, ly + 4, "shadows", 10, H("FFFFFF"), "start"
    Ln mVX + 178, ly, mVX + 194, ly, H("FF5555"), 2.4, msoLineSolid
    Tx mVX + 198, ly + 4, "required " & Format$(mReqInc, "0.00") & ChrW(176), 10, H("FFFFFF"), "start"
    Ln mVX + 292, ly, mVX + 308, ly, H("E09A3D"), 2.2, msoLineDash
    Tx mVX + 312, ly + 4, "if " & Format$(mSvyInc, "0.00") & ChrW(176) & " held", 10, H("FFFFFF"), "start"
    Dot mVX + 420, ly, 2.5, H("4CAF7A"), -1, 0
    Tx mVX + 428, ly + 4, mTgtName, 10, H("FFFFFF"), "start"
    If mShaft Then
        Tx mVX + mVW, ly + 4, "true scale 1:1 TVD / L/R  ·  MD compressed", 8.5, H("A8A8A8"), "end"
    Else
        Tx mVX + mVW, ly + 4, "MD " & ChrW(247) & Format$(1# / mSxd, "0.0") & " " & _
           ChrW(183) & " TVD " & ChrW(215) & Format$(mSV, "0") & " " & _
           ChrW(183) & " L/R " & ChrW(215) & Format$(mKzx, "0"), 8.5, H("A8A8A8"), "end"
    End If
End Sub

' Vertical / build: traveling-cylinder bullseye at the next named target.
' Rings sit in the N/E plane at target TVD, centred on the plan. Hole and
' plan pierce the plane. No sail box.
Private Sub DrawShaftProfile()
    Dim i As Long, k As Long
    Dim rMax As Double, rStep As Double, r As Double
    Dim off As Double, s As Double, tvdPad As Double
    Dim xs() As Double, ys() As Double
    Dim pMD() As Double, pInc() As Double, pAzi() As Double
    Dim pTvd() As Double, pNS() As Double, pEW() As Double
    Dim np As Long, pN As Double, pE As Double, pV As Double, pA As Double, pI As Double
    Dim fillA As Long, fillB As Long, ring As Long
    Dim ly As Double

    Tx mVX, 20, mSection & " TARGET  " & mTgtName & "  " & Format$(mWpMD, "#,##0") & " mMD", _
       12.5, H("FFFFFF"), "start", True
    Tx mVX, mVY + 14, "Traveling cylinder  ·  GN / E / S / W  ·  metres from plan", _
       8.5, H("A8A8A8"), "start"
    Rect mVX, mVY, mVW, mVH, H("FFFFFF"), H("3A3A3A"), 0.6

    mTgtN = 0#: mTgtE = 0#: mTgtV = mWpTvd
    np = MDL_PlanGauge.PG_LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
    If np >= 2 Then
        MDL_PlanGauge.PG_PlanAt mWpMD, np, pMD, pInc, pAzi, pTvd, pNS, pEW, pN, pE, pV, pA, pI
        mTgtN = pN: mTgtE = pE
        If pV > 0# Then mTgtV = pV
    End If
    If mTgtV <= 0# Then mTgtV = mSvyTvd

    ' Ring scale follows the actual miss: a 1.4 m offset gets 2 m rings, not a
    ' 40 m disk that hides it. (Hold-projection lateral is deliberately NOT
    ' included: at near-vertical inclinations the azimuth term makes it junk.)
    off = Sqr((mSvyNS - mTgtN) ^ 2 + (mSvyEW - mTgtE) ^ 2)
    rMax = 2#
    Do While off * 1.35 > rMax And rMax < 80#
        rMax = rMax * 2#
        If rMax = 16# Then rMax = 20#   ' 2/4/8 then 10-step sizes
    Loop
    rStep = rMax / 4#

    ' Disk takes ~40% of the panel width; the TVD pierce above it is
    ' compressed so the last ~toGo metres of hole fit in the headroom.
    s = (mVW * 0.15) / rMax
    If s * rMax * 0.6 > mVH * 0.3 Then s = (mVH * 0.3) / (rMax * 0.6)
    mShS = s
    mShOx = mVX + mVW * 0.5
    mShOy = mVY + mVH * 0.6

    tvdPad = mTgtV - mSvyTvd + 10#
    If tvdPad < 15# Then tvdPad = 15#
    If tvdPad > 400# Then tvdPad = 400#
    mShV0 = mTgtV - tvdPad
    mShV1 = mTgtV + tvdPad * 0.3
    mShVS = (mVH * 0.6 - 44# - rMax * s * 0.44) / tvdPad
    If mShVS < 0.05 Then mShVS = 0.05

    fillA = H("E8EEF2")
    fillB = H("F7F7F7")
    ring = 0
    For r = rMax To rStep Step -rStep
        SampleTargetRing xs, ys, r, mTgtV, 48
        If (ring Mod 2) = 0 Then
            PolyFill xs, ys, 48, fillA, H("8A9AAB"), 0.7
        Else
            PolyFill xs, ys, 48, fillB, H("8A9AAB"), 0.7
        End If
        ring = ring + 1
    Next r
    SampleTargetRing xs, ys, rMax, mTgtV, 48
    PolyLine xs, ys, 48, H("1A1A1A"), 2.2, msoLineSolid

    ' Radial to GN for orientation. Ring pitch is stated in the legend; the
    ' numbers on every ring were unreadable at small radii.
    Ln ShX(0#, 0#), ShY(0#, 0#, mTgtV), ShX(rMax, 0#), ShY(rMax, 0#, mTgtV), H("5B7A99"), 0.6, msoLineSolid

    LabelTargetCardinal rMax, 0#, "GN"
    LabelTargetCardinal 0#, rMax, "E"
    LabelTargetCardinal -rMax, 0#, "S"
    LabelTargetCardinal 0#, -rMax, "W"

    Dot ShX(0#, 0#), ShY(0#, 0#, mTgtV), 4.2, H("4A1A5C"), H("1A1A1A"), 0.9
    DrawPlanPathOnShaft3
    DrawMinCurveToTarget

    If mSCount >= 2 Then
        ReDim xs(0 To mSCount - 1): ReDim ys(0 To mSCount - 1)
        k = 0
        For i = 0 To mSCount - 1
            If mSTvd(i) >= mShV0 - 8# And mSTvd(i) <= mShV1 + 8# Then
                xs(k) = ShX(mSNS(i) - mTgtN, mSEW(i) - mTgtE)
                ys(k) = ShY(mSNS(i) - mTgtN, mSEW(i) - mTgtE, mSTvd(i))
                k = k + 1
            End If
        Next i
        If k >= 2 Then PolyLine xs, ys, k, H("6B2020"), 2.4, msoLineSolid
    End If

    Dot ShX(mSvyNS - mTgtN, mSvyEW - mTgtE), ShY(mSvyNS - mTgtN, mSvyEW - mTgtE, mSvyTvd), _
       4.5, H("6B2020"), H("1A1A1A"), 0.9
    Tx ShX(mSvyNS - mTgtN, mSvyEW - mTgtE) - 9, _
       ShY(mSvyNS - mTgtN, mSvyEW - mTgtE, mSvyTvd) + 3, "BIT", 9, H("6B2020"), "end", True

    DrawHighSideTick

    Tx mShOx, mShOy + rMax * mShS * 0.5 + 14, _
       mTgtName & "  " & Format$(mWpMD, "#,##0"), 9, H("1B4D2E"), "middle", True

    ly = mVY + mVH + 15
    Ln mVX + 2, ly, mVX + 18, ly, H("6B2020"), 2.4, msoLineSolid
    Tx mVX + 22, ly + 4, "hole", 10, H("FFFFFF"), "start"
    Ln mVX + 52, ly, mVX + 68, ly, H("5BA3D9"), 2.2, msoLineSolid
    Tx mVX + 72, ly + 4, "plan", 10, H("FFFFFF"), "start"
    Ln mVX + 104, ly, mVX + 120, ly, H("C0392B"), 1.6, msoLineDash
    Tx mVX + 124, ly + 4, "min curve " & Format$(mReqInc, "0.0") & ChrW(176), _
       10, H("FFFFFF"), "start"
    Tx mVX + 230, ly + 4, "rings every " & Format$(rStep, "General Number") & " m from plan", _
       10, H("FFFFFF"), "start"
End Sub

' Minimum-curvature correction: a smooth arc leaving the bit along its current
' attitude and arriving at the target centre along the required inclination.
' Drawn dashed so it reads apart from the plan line.
Private Sub DrawMinCurveToTarget()
    Dim n0 As Double, e0 As Double, v0 As Double
    Dim dN0 As Double, dE0 As Double, dV0 As Double
    Dim dN1 As Double, dE1 As Double, dV1 As Double
    Dim inc As Double, azi As Double, aziT As Double, incT As Double
    Dim dist As Double, t As Double
    Dim b0 As Double, b1 As Double, b2 As Double, b3 As Double
    Dim cN1 As Double, cE1 As Double, cV1 As Double
    Dim cN2 As Double, cE2 As Double, cV2 As Double
    Dim xs(0 To 24) As Double, ys(0 To 24) As Double
    Dim i As Long, pN As Double, pE As Double, pV As Double

    n0 = mSvyNS - mTgtN: e0 = mSvyEW - mTgtE: v0 = mSvyTvd
    dist = Sqr(n0 * n0 + e0 * e0 + (mTgtV - v0) * (mTgtV - v0))
    If dist < 0.5 Then Exit Sub

    inc = mSvyInc * PIE / 180#
    azi = mSvyAzi * PIE / 180#
    dN0 = Cos(azi) * Sin(inc): dE0 = Sin(azi) * Sin(inc): dV0 = Cos(inc)

    incT = mReqInc * PIE / 180#
    If Sqr(n0 * n0 + e0 * e0) > 0.01 Then
        aziT = WorksheetFunction.Atan2(-n0, -e0)   ' toward the target centre
    Else
        aziT = azi
    End If
    dN1 = Cos(aziT) * Sin(incT): dE1 = Sin(aziT) * Sin(incT): dV1 = Cos(incT)

    ' Cubic Bezier with tangents matched at both ends approximates the
    ' constant-curvature correction closely enough at plot scale.
    cN1 = n0 + dN0 * dist / 3#: cE1 = e0 + dE0 * dist / 3#: cV1 = v0 + dV0 * dist / 3#
    cN2 = -dN1 * dist / 3#: cE2 = -dE1 * dist / 3#: cV2 = mTgtV - dV1 * dist / 3#

    For i = 0 To 24
        t = CDbl(i) / 24#
        b0 = (1# - t) ^ 3
        b1 = 3# * t * (1# - t) ^ 2
        b2 = 3# * t * t * (1# - t)
        b3 = t ^ 3
        pN = b0 * n0 + b1 * cN1 + b2 * cN2
        pE = b0 * e0 + b1 * cE1 + b2 * cE2
        pV = b0 * v0 + b1 * cV1 + b2 * cV2 + b3 * mTgtV
        xs(i) = ShX(pN, pE)
        ys(i) = ShY(pN, pE, pV)
    Next i
    PolyLine xs, ys, 25, H("C0392B"), 1.6, msoLineDash
End Sub

' Target view: +North up-right (GN), +East down-right, +TVD down through the plane.
' n/e are metres from the plan at the target.
Private Function ShX(ByVal n As Double, ByVal e As Double) As Double
    ShX = mShOx + e * mShS * 0.88 + n * mShS * 0.52
End Function

Private Function ShY(ByVal n As Double, ByVal e As Double, ByVal v As Double) As Double
    ' TVD uses its own compressed scale so the pierce fits the headroom even
    ' when the rings are only a couple of metres wide.
    ShY = mShOy + (v - mTgtV) * mShVS - n * mShS * 0.4 + e * mShS * 0.18
End Function

Private Sub SampleTargetRing(xs() As Double, ys() As Double, _
        ByVal radius As Double, ByVal tvd As Double, ByVal nPts As Long)
    Dim i As Long, a As Double
    If nPts < 8 Then nPts = 8
    ReDim xs(0 To nPts - 1)
    ReDim ys(0 To nPts - 1)
    For i = 0 To nPts - 1
        a = 2# * PIE * CDbl(i) / CDbl(nPts)
        xs(i) = ShX(radius * Cos(a), radius * Sin(a))
        ys(i) = ShY(radius * Cos(a), radius * Sin(a), tvd)
    Next i
End Sub

Private Sub LabelTargetCardinal(ByVal n As Double, ByVal e As Double, ByVal lab As String)
    Dim x As Double, y As Double
    x = ShX(n * 1.12, e * 1.12)
    y = ShY(n * 1.12, e * 1.12, mTgtV)
    Tx x, y + 4, lab, 10, H("1A3A5C"), "middle", True
End Sub

Private Sub DrawHighSideTick()
    Dim azi As Double, dN As Double, dE As Double
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double
    Dim tickM As Double
    azi = mSvyAzi * PIE / 180#
    tickM = 12# / mShS          ' ~12 pt on screen, whatever the ring scale
    dN = Cos(azi) * tickM
    dE = Sin(azi) * tickM
    x0 = ShX(mSvyNS - mTgtN, mSvyEW - mTgtE)
    y0 = ShY(mSvyNS - mTgtN, mSvyEW - mTgtE, mSvyTvd)
    x1 = ShX(mSvyNS - mTgtN + dN, mSvyEW - mTgtE + dE)
    y1 = ShY(mSvyNS - mTgtN + dN, mSvyEW - mTgtE + dE, mSvyTvd)
    Ln x0, y0, x1, y1, H("8B0000"), 1.1, msoLineSolid
    Tx x1 + 3, y1, "HS", 8, H("8B0000"), "start", True
End Sub

Private Sub DrawPlanPathOnShaft3()
    Dim pMD() As Double, pInc() As Double, pAzi() As Double
    Dim pTvd() As Double, pNS() As Double, pEW() As Double
    Dim nPlan As Long
    nPlan = MDL_PlanGauge.PG_LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
    If nPlan < 2 Then Exit Sub

    Dim stepMd As Double, n As Long, i As Long, k As Long
    Dim md As Double, pN As Double, pE As Double, pV As Double, pA As Double, PI As Double
    Dim xs() As Double, ys() As Double
    stepMd = 8#
    n = Int((mMD1 - mMD0) / stepMd) + 2
    ReDim xs(0 To n - 1): ReDim ys(0 To n - 1)
    k = 0
    For i = 0 To n - 1
        md = mMD0 + CDbl(i) * stepMd
        If md > mMD1 Then md = mMD1
        MDL_PlanGauge.PG_PlanAt md, nPlan, pMD, pInc, pAzi, pTvd, pNS, pEW, _
                                pN, pE, pV, pA, PI
        If pV >= mShV0 And pV <= mShV1 Then
            xs(k) = ShX(pN - mTgtN, pE - mTgtE)
            ys(k) = ShY(pN - mTgtN, pE - mTgtE, pV)
            k = k + 1
        End If
        If md >= mMD1 Then Exit For
    Next i
    If k >= 2 Then PolyLine xs, ys, k, H("5BA3D9"), 2.1, msoLineSolid
End Sub

' Plan trajectory through the room: the true P2 well plan (Plan sheet TVD at
' each MD, lat = 0 since lateral offsets are measured from this plan line).
' A constant-inclination plan section renders as one smooth straight slope.
Private Sub DrawPlanPathInRoom()
    Dim pMD() As Double, pInc() As Double, pAzi() As Double
    Dim pTvd() As Double, pNS() As Double, pEW() As Double
    Dim nPlan As Long
    nPlan = MDL_PlanGauge.PG_LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
    If nPlan < 2 Then Exit Sub

    Dim stepMd As Double, n As Long, i As Long, k As Long
    Dim md As Double, plot As Double
    Dim pN As Double, pE As Double, pV As Double, pA As Double, PI As Double
    Dim xs() As Double, ys() As Double
    Dim planName As String

    stepMd = 8#
    n = Int((mMD1 - mMD0) / stepMd) + 2
    ReDim xs(0 To n - 1)
    ReDim ys(0 To n - 1)

    ' Clip (not clamp) to the drawn room: where the plan runs above or below the
    ' TVD window the line simply exits, instead of smearing along the ceiling.
    Dim lastX As Double, lastY As Double, drawn As Boolean
    k = 0
    For i = 0 To n - 1
        md = mMD0 + CDbl(i) * stepMd
        If md > mMD1 Then md = mMD1
        MDL_PlanGauge.PG_PlanAt md, nPlan, pMD, pInc, pAzi, pTvd, pNS, pEW, _
                                pN, pE, pV, pA, PI
        plot = PlotY(pV)
        If PlotInRoom(plot) Then
            xs(k) = ObX(md, 0#)
            ys(k) = ObY(plot, 0#)
            k = k + 1
        Else
            If k >= 2 Then
                PolyLine xs, ys, k, H("5BA3D9"), 2.3, msoLineSolid
                lastX = xs(k - 1): lastY = ys(k - 1): drawn = True
            End If
            k = 0
        End If
        If md >= mMD1 Then Exit For
    Next i
    If k >= 2 Then
        PolyLine xs, ys, k, H("5BA3D9"), 2.3, msoLineSolid
        lastX = xs(k - 1): lastY = ys(k - 1): drawn = True
    End If
    If Not drawn Then Exit Sub

    planName = Trim$(cellText(ThisWorkbook.Worksheets("Data"), "C2"))
    If Len(planName) = 0 Then planName = "Plan"
    Tx lastX + 4, lastY - 6, planName, 9, H("5BA3D9"), "start", True
End Sub

' Data lateral sign -> screen lateral. Positive data lat = RIGHT of plan.
' Walking down-hole (+MD, drawn to screen right) the traveler's right hand
' points toward the viewer, i.e. the near/front (-lat) side of the room.
Private Function LatScr(ByVal l As Double) As Double
    LatScr = -l
End Function

' Geo +/- corridor as a sloping ribbon between waypoint TVDs (plus window ends).
Private Sub DrawGeoCorridorRibbon(ByVal lat As Double, ByVal fillRgb As Long, _
                                  ByVal lineRgb As Long, ByVal lineW As Double)
    Dim mdK() As Double, tvdK() As Double, nK As Long
    Dim i As Long, gT As Double
    Dim hi() As Double, lo() As Double, midY() As Double
    Dim xs() As Double, ys() As Double

    ReDim mdK(0 To mWCount + 3)
    ReDim tvdK(0 To mWCount + 3)
    nK = 0

    If MDL_PlanGauge.PG_WaypointTvdAtMd(mWs, mMD0, gT) Then
        mdK(nK) = mMD0: tvdK(nK) = gT: nK = nK + 1
    End If
    For i = 0 To mWCount - 1
        If mWMD(i) > mMD0 + 0.05 And mWMD(i) < mMD1 - 0.05 Then
            mdK(nK) = mWMD(i): tvdK(nK) = mWTvd(i): nK = nK + 1
        End If
    Next i
    If MDL_PlanGauge.PG_WaypointTvdAtMd(mWs, mMD1, gT) Then
        mdK(nK) = mMD1: tvdK(nK) = gT: nK = nK + 1
    End If
    If nK < 2 Then Exit Sub

    ReDim hi(0 To nK - 1)
    ReDim lo(0 To nK - 1)
    ReDim midY(0 To nK - 1)
    For i = 0 To nK - 1
        hi(i) = PlotY(tvdK(i) - mGeoHalf)
        lo(i) = PlotY(tvdK(i) + mGeoHalf)
        midY(i) = PlotY(tvdK(i))
    Next i

    ' Outline only — filled geo ribbons eat toner and are lateral-only anyway.

    ReDim xs(0 To nK - 1): ReDim ys(0 To nK - 1)
    For i = 0 To nK - 1
        xs(i) = ObX(mdK(i), lat): ys(i) = ObY(hi(i), lat)
    Next i
    PolyLine xs, ys, nK, lineRgb, lineW + 0.55, msoLineSolid
    For i = 0 To nK - 1
        ys(i) = ObY(lo(i), lat)
    Next i
    PolyLine xs, ys, nK, lineRgb, lineW + 0.55, msoLineSolid
    For i = 0 To nK - 1
        ys(i) = ObY(midY(i), lat)
    Next i
    PolyLine xs, ys, nK, lineRgb, 1#, msoLineDash
End Sub

' Marks each geo waypoint on the left-wall shadow with MD + Inc-to-next, and a
' short attitude tick (Inc 90 deg = along MD; >90 tips toward drop / higher TVD).
Private Sub DrawWaypointMarks()
    Dim i As Long
    For i = 0 To mWCount - 1
        If mWMD(i) < mMD0 Or mWMD(i) > mMD1 Then GoTo NextWp

        Dim plot As Double, lat As Double
        If Not HoleAtMdPlot(mWMD(i), plot, lat) Then
            ' Ahead of / off the surveyed hole: sit on geo TVD at plan centre.
            plot = PlotY(mWTvd(i))
            lat = 0#
        End If

        Dim isNext As Boolean: isNext = (Abs(mWMD(i) - mWpMD) < 0.001)
        Dim wx As Double, wy As Double
        wx = ObX(mWMD(i), -mLatBox)
        wy = ObY(plot, -mLatBox)

        Dot wx, wy, IIf(isNext, 3.2, 2.4), H("4CAF7A"), H("1A1A1A"), 1

        ' Attitude tick in the MD / up-down plane on the left wall.
        ' Inc 90 deg = along MD; >90 tips toward drop (higher TVD / lower "high").
        If mWInc(i) > 0.1 Then
            Dim tick As Double, ang As Double, tx2 As Double, ty2 As Double
            tick = 14#
            ang = (mWInc(i) - 90#) * PIE / 180#
            tx2 = wx + tick * Cos(ang)
            ty2 = wy + tick * Sin(ang)
            Ln wx, wy, tx2, ty2, H("4CAF7A"), IIf(isNext, 1.8, 1.2), msoLineSolid
            Tx wx, wy - 10, Format$(mWInc(i), "0.00") & ChrW(176), 8, _
               IIf(isNext, H("4CAF7A"), H("8AAB9A")), "middle", isNext
        End If

        Tx wx, mVY + mVH - 30, Format$(mWMD(i), "#,##0"), 8.5, _
           IIf(isNext, H("4CAF7A"), H("A8A8A8")), "middle", isNext
NextWp:
    Next i
End Sub

' Interpolate as-drilled (plotY, lat) at a measured depth inside the window.
Private Function HoleAtMdPlot(ByVal md As Double, ByRef outPlot As Double, ByRef outLat As Double) As Boolean
    HoleAtMdPlot = False
    If mSCount < 1 Then Exit Function
    If md < mSMD(0) - 0.01 Or md > mSMD(mSCount - 1) + 0.01 Then Exit Function
    If mSCount = 1 Or md <= mSMD(0) Then
        outPlot = PlotY(mSTvd(0)): outLat = mSLat(0): HoleAtMdPlot = True: Exit Function
    End If
    If md >= mSMD(mSCount - 1) Then
        outPlot = PlotY(mSTvd(mSCount - 1)): outLat = mSLat(mSCount - 1): HoleAtMdPlot = True: Exit Function
    End If
    Dim i As Long
    For i = 0 To mSCount - 2
        If md >= mSMD(i) And md <= mSMD(i + 1) Then
            Dim f As Double, tvd As Double
            If mSMD(i + 1) > mSMD(i) Then f = (md - mSMD(i)) / (mSMD(i + 1) - mSMD(i))
            tvd = mSTvd(i) + f * (mSTvd(i + 1) - mSTvd(i))
            outPlot = PlotY(tvd)
            outLat = mSLat(i) + f * (mSLat(i + 1) - mSLat(i))
            HoleAtMdPlot = True
            Exit Function
        End If
    Next i
End Function

Private Function ClampMd(ByVal v As Double) As Double
    ClampMd = v
    If ClampMd < mMD0 Then ClampMd = mMD0
    If ClampMd > mMD1 Then ClampMd = mMD1
End Function

Private Function SignedM(ByVal v As Double) As String
    If Abs(v) < 0.001 Then
        SignedM = "0"
    ElseIf v > 0 Then
        SignedM = "+" & Format$(v, "0.0")
    Else
        SignedM = ChrW(8722) & Format$(Abs(v), "0.0")
    End If
End Function


' ================================================================================
'  DRAWING - PANEL 2, THE TRUE 1:1 VERTICAL SECTION
' ================================================================================
Private Sub DrawSection(ss As Worksheet)
    Dim pMD() As Double, pInc() As Double, pAzi() As Double
    Dim pTvd() As Double, pNS() As Double, pEW() As Double
    Dim np As Long
    np = MDL_PlanGauge.PG_LoadPlan(pMD, pInc, pAzi, pTvd, pNS, pEW)
    If np < 2 Then Exit Sub

    ' Vertical section is the horizontal run projected on the plan's own heading,
    ' taken from the final plan azimuth so the section plane matches the lateral.
    Dim vsAzi As Double: vsAzi = pAzi(np - 1)
    Dim vs() As Double
    ReDim vs(0 To np - 1)
    Dim i As Long
    Dim maxVs As Double, maxTvd As Double
    For i = 0 To np - 1
        vs(i) = pNS(i) * Cos(vsAzi * PIE / 180#) + pEW(i) * Sin(vsAzi * PIE / 180#)
        If vs(i) > maxVs Then maxVs = vs(i)
        If pTvd(i) > maxTvd Then maxTvd = pTvd(i)
    Next i
    If maxVs <= 0# Or maxTvd <= 0# Then Exit Sub

    Const PAD As Double = 20
    Dim sc As Double
    sc = maxVs / (mPW - 2 * PAD)
    If maxTvd / (mPH - 2 * PAD) > sc Then sc = maxTvd / (mPH - 2 * PAD)
    Dim gx As Double, gy As Double
    gx = mPX + PAD: gy = mPY + PAD

    Tx mPX, 20, "VERTICAL SECTION " & ChrW(183) & " 1:1", 12.5, H("FFFFFF"), "start", True
    Rect mPX, mPY, mPW, mPH, H("222222"), H("3A3A3A"), 0.75

    Ln gx - 8, gy, mPX + mPW - 8, gy, H("4A4A4A"), 1, msoLineSolid
    Dim tv As Double
    For tv = 1000 To maxTvd Step 1000
        Ln gx - 4, gy + tv / sc, mPX + mPW - 8, gy + tv / sc, H("3A3A3A"), 0.6, msoLineSysDot
        Tx gx - 6, gy + tv / sc + 3, Format$(tv, "#,##0"), 7.5, H("A8A8A8"), "end"
    Next tv

    ' the plan, split so drilled / remaining / today read differently
    Dim d0 As Double, d1 As Double
    d0 = NumCell(ThisWorkbook.Worksheets("Data").Range("C4"))
    d1 = NumCell(ThisWorkbook.Worksheets("Data").Range("C5"))
    SectionSeg pMD, pTvd, vs, np, sc, gx, gy, 0, mBitMD, H("FFFFFF"), 2.4, msoLineSolid
    SectionSeg pMD, pTvd, vs, np, sc, gx, gy, mBitMD, pMD(np - 1), H("6A6A6A"), 1.6, msoLineDash
    SectionSeg pMD, pTvd, vs, np, sc, gx, gy, d0, d1, H("5BA3D9"), 4.2, msoLineSolid

    For i = 0 To mWCount - 1
        Dim wvs As Double, wtv As Double
        If SectionAt(pMD, pTvd, vs, np, mWMD(i), wvs, wtv) Then
            Dot gx + wvs / sc, gy + wtv / sc, 2, H("4CAF7A"), -1, 0
        End If
    Next i

    ' build targets from Slidesheet!T2:Y5 - MD in U, name in Y. Repeated names
    ' (TANGENT appears twice) are only labelled once.
    Dim tRow As Long, prevName As String
    For tRow = 2 To 5
        Dim tMd As Double, tName As String
        tMd = NumCell(ss.Cells(tRow, 21))
        tName = Trim$(CStr(ss.Cells(tRow, 25).text))
        If tMd > 0# And tName <> "" Then
            Dim gvs As Double, gtv As Double
            If SectionAt(pMD, pTvd, vs, np, tMd, gvs, gtv) Then
                Dot gx + gvs / sc, gy + gtv / sc, 3, H("1A1A1A"), H("A8A8A8"), 1
                If tName <> prevName And gtv > 500# Then
                    Tx gx + gvs / sc - 7, gy + gtv / sc + 3, tName, 8, H("A8A8A8"), "end"
                End If
                prevName = tName
            End If
        End If
    Next tRow

    Dim bvs As Double, btv As Double
    If SectionAt(pMD, pTvd, vs, np, mBitMD, bvs, btv) Then
        Dot gx + bvs / sc, gy + btv / sc, 4, H("FF5555"), H("1A1A1A"), 1.2
        Tx gx + bvs / sc + 8, gy + btv / sc + 4, "BIT " & Format$(mBitMD, "#,##0"), 8.5, H("FF5555"), "start", True
    End If
    Dim tvs As Double, ttv As Double
    If SectionAt(pMD, pTvd, vs, np, pMD(np - 1), tvs, ttv) Then
        Tx gx + tvs / sc, gy + ttv / sc - 8, "TD " & Format$(pMD(np - 1), "#,##0"), 8, H("A8A8A8"), "middle"
    End If

    Tx gx - 14, gy - 7, "TVD m", 8, H("A8A8A8"), "start"
    Tx mPX + mPW - 8, mPY + mPH - 7, "vertical section m " & ChrW(8594), 8, H("A8A8A8"), "end"
End Sub

Private Function SectionAt(pMD() As Double, pTvd() As Double, vs() As Double, _
        ByVal np As Long, ByVal md As Double, _
        ByRef outVs As Double, ByRef outTvd As Double) As Boolean
    SectionAt = False
    If np < 2 Then Exit Function
    If md <= pMD(0) Then
        outVs = vs(0): outTvd = pTvd(0): SectionAt = True: Exit Function
    End If
    If md >= pMD(np - 1) Then
        outVs = vs(np - 1): outTvd = pTvd(np - 1): SectionAt = True: Exit Function
    End If
    Dim i As Long
    For i = 0 To np - 2
        If md >= pMD(i) And md <= pMD(i + 1) Then
            Dim f As Double
            If pMD(i + 1) > pMD(i) Then f = (md - pMD(i)) / (pMD(i + 1) - pMD(i))
            outVs = vs(i) + f * (vs(i + 1) - vs(i))
            outTvd = pTvd(i) + f * (pTvd(i + 1) - pTvd(i))
            SectionAt = True
            Exit Function
        End If
    Next i
End Function

Private Sub SectionSeg(pMD() As Double, pTvd() As Double, vs() As Double, _
        ByVal np As Long, ByVal sc As Double, ByVal gx As Double, ByVal gy As Double, _
        ByVal mdA As Double, ByVal mdB As Double, _
        ByVal clr As Long, ByVal wt As Double, ByVal dash As Long)
    If mdB <= mdA Then Exit Sub

    Dim xs() As Double, ys() As Double
    ReDim xs(0 To np + 1): ReDim ys(0 To np + 1)
    Dim n As Long: n = 0

    Dim avs As Double, atv As Double
    If SectionAt(pMD, pTvd, vs, np, mdA, avs, atv) Then
        xs(n) = gx + avs / sc: ys(n) = gy + atv / sc: n = n + 1
    End If
    Dim i As Long
    For i = 0 To np - 1
        If pMD(i) > mdA And pMD(i) < mdB Then
            xs(n) = gx + vs(i) / sc: ys(n) = gy + pTvd(i) / sc: n = n + 1
        End If
    Next i
    Dim bvs As Double, btv As Double
    If SectionAt(pMD, pTvd, vs, np, mdB, bvs, btv) Then
        xs(n) = gx + bvs / sc: ys(n) = gy + btv / sc: n = n + 1
    End If

    PolyLine xs, ys, n, clr, wt, dash
End Sub


' ================================================================================
'  DRAWING - REPORTING DAY AND METRICS
'
'  Every figure below is the Data sheet's own printed text, so the picture says
'  exactly what the cells say and the D7:F17 block stays read-only.
' ================================================================================
Private Sub DrawDay(dt As Worksheet)
    Rect mPX, mQY, mPW, mQH, H("222222"), H("3A3A3A"), 0.75
    Tx mPX + 11, mQY + 17, "REPORTING DAY", 9.5, H("FFFFFF"), "start", True
    Ln mPX + 11, mQY + 22, mPX + mPW - 11, mQY + 22, H("4A4A4A"), 0.75, msoLineSolid

    Dim y As Double: y = mQY + 36
    KV mPX, mPW, y, "From / to", _
       cellText(dt, "C4") & " " & ChrW(8594) & " " & cellText(dt, "C5"), 9, H("FFFFFF"), False
    y = y + 16
    KV mPX, mPW, y, "Drilled", cellText(dt, "C6"), 9, H("5BA3D9"), True
    y = y + 16
    KV mPX, mPW, y, "Slid / rot", cellText(dt, "C7") & " / " & cellText(dt, "C8"), 9, H("FFFFFF"), False
    y = y + 18
    DayStack mPX, mPW, y, "Sliding", cellText(dt, "C9")
    DayStack mPX, mPW, y, "Rotating", cellText(dt, "C10")
End Sub

Private Sub DayStack(ByVal px As Double, ByVal pw As Double, ByRef y As Double, _
                     ByVal k As String, ByVal v As String)
    Tx px + 11, y, k, 8.5, H("A8A8A8"), "start"
    y = y + 12
    Tx px + 11, y, v, 8, H("FFFFFF"), "start", False, pw - 22
    y = y + 15
End Sub

Private Sub DrawMetrics(ss As Worksheet, dt As Worksheet)
    Tx mMX, 20, "LAST SURVEY " & ChrW(183) & " POSITION " & ChrW(183) & " REQUIREMENT", _
       12.5, H("FFFFFF"), "start", True
    Rect mMX, mVY, mMW, mMH, H("222222"), H("3A3A3A"), 0.75

    Dim y As Double: y = mVY + 20

    MHead y, "LAST SURVEY"
    ' E7 and E10 carry a degree suffix in their number format, so they are
    ' formatted from the value here rather than echoed as text
    KV mMX, mMW, y, "Depth", Format$(mSvyMD, "#,##0.00") & " m", 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "INC", cellText(dt, "E8"), 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "AZM", cellText(dt, "E9"), 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "TVD", Format$(mSvyTvd, "#,##0.00") & " m", 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "DLS", cellText(dt, "E11"), 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "Bit depth", Format$(mBitMD, "#,##0.00") & " m", 10.5, H("FFFFFF"), False: y = y + 16

    MHead y, "POSITION OF WELLBORE"
    ' Short keys so values stay clear of labels in the narrow metrics column.
    KV mMX, mMW, y, "R/L from plan", cellText(dt, "E13") & " " & cellText(dt, "F13"), _
       10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "Above/Below plan", cellText(dt, "E14") & " " & cellText(dt, "F14"), _
       10.5, H("4CAF7A"), True: y = y + 16
    If mUseSailBands Then
        KV mMX, mMW, y, "Dist. from geo", cellText(dt, "E15") & " " & cellText(dt, "F15"), _
           10.5, H("FFFFFF"), False: y = y + 16
    End If
    If mUseSailBands Then
        KV mMX, mMW, y, "Corridor used", _
           Format$(100# * Abs(mSvyDev) / mGeoHalf, "0") & "% of " & ChrW(177) & Format$(mGeoHalf, "0.00") & " m", _
           10.5, H("FFFFFF"), False: y = y + 16
        KV mMX, mMW, y, "Lateral used", _
           Format$(100# * Abs(mSvyLat) / mLatTol, "0") & "% of " & ChrW(177) & Format$(mLatTol, "0.00") & " m", _
           10.5, H("FFFFFF"), False: y = y + 16
    Else
        KV mMX, mMW, y, "Offset from plan", Format$(Abs(mSvyLat), "0.00") & " m " & _
           IIf(mSvyLat < 0, "L", "R"), 10.5, H("FFFFFF"), False: y = y + 16
        KV mMX, mMW, y, "TVD vs plan", Format$(Abs(mSvyDev), "0.00") & " m", _
           10.5, H("FFFFFF"), False: y = y + 16
    End If

    MHead y, "REQUIREMENT TO " & mTgtName & " " & Format$(mWpMD, "#,##0")
    KV mMX, mMW, y, "Distance to go", Format$(mToGo, "#,##0.0") & " m", 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "Target TVD", Format$(mWpTvd, "#,##0.00") & " m", 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "Required inclination", Format$(mReqInc, "0.00") & ChrW(176), 10.5, H("FF5555"), True: y = y + 16
    KV mMX, mMW, y, "Holding", Format$(mSvyInc, "0.00") & ChrW(176), 10.5, H("FFFFFF"), False: y = y + 16
    KV mMX, mMW, y, "Correction", Format$(Abs(mReqInc - mSvyInc), "0.00") & ChrW(176) & " " & _
       IIf(mReqInc < mSvyInc, "drop", "build"), 10.5, H("FF5555"), True: y = y + 16
    KV mMX, mMW, y, "Arrival if held", Format$(Abs(mHoldDevEnd), "0.00") & " m " & _
       IIf(mHoldDevEnd >= 0, "above", "below"), 10.5, H("E09A3D"), False: y = y + 16
End Sub

Private Sub MHead(ByRef y As Double, ByVal t As String)
    y = y + 5
    Tx mMX + 11, y, t, 9.5, H("FFFFFF"), "start", True
    Ln mMX + 11, y + 4, mMX + mMW - 11, y + 4, H("4A4A4A"), 0.75, msoLineSolid
    y = y + 16
End Sub

Private Sub KV(ByVal px As Double, ByVal pw As Double, ByVal y As Double, _
               ByVal k As String, ByVal v As String, ByVal sz As Double, _
               ByVal clr As Long, ByVal bold As Boolean)
    ' Reserve a value column so long labels (e.g. Distance from Current Geo Target)
    ' cannot paint over the right-aligned value.
    Dim valW As Double, keyW As Double
    valW = 112#
    If pw < 200# Then valW = pw * 0.52
    keyW = pw - 22# - valW
    If keyW < 40# Then keyW = 40#
    Tx px + 11, y, k, sz, H("A8A8A8"), "start", False, keyW
    Tx px + pw - 11, y, v, sz, clr, "end", bold
End Sub


' ================================================================================
'  DRAWING - OPS BAND (under the corridor)
'
'  Mirrors the old B2:F55 reading order without last-survey / position (already
'  in the metrics panel) and without empty AC / motors rows.
' ================================================================================
Private Function OpsColW() As Double
    OpsColW = (CANVAS_W - 2 * OPS_PAD - OPS_COL_GAP) / 2#
End Function

Private Function CountFilledAc(dt As Worksheet) As Long
    Dim r As Long, n As Long
    n = 0
    For r = 35 To 42
        If Len(Trim$(cellText(dt, "B" & r))) > 0 Then n = n + 1
    Next r
    CountFilledAc = n
End Function

Private Function CountFilledMotors(dt As Worksheet) As Long
    Dim r As Long, n As Long
    n = 0
    For r = 45 To 55
        If Len(Trim$(cellText(dt, "B" & r))) > 0 Then n = n + 1
    Next r
    CountFilledMotors = n
End Function

Private Function MeasureOpsHeight(dt As Worksheet) As Double
    Dim nAc As Long, nMot As Long
    Dim acH As Double, motH As Double
    nAc = CountFilledAc(dt)
    nMot = CountFilledMotors(dt)
    ' Row pitch must clear TYPE_SCALE_OPS text (see DrawOps AC / motors loops).
    Const ROW_PITCH As Double = 22#
    If nAc = 0 Then acH = 44# Else acH = 32# + ROW_PITCH * CDbl(nAc)
    If nMot = 0 Then motH = 44# Else motH = 32# + ROW_PITCH * CDbl(nMot)
    ' chips + day/BHA + motor band + AC + motors + gaps + footer
    MeasureOpsHeight = OPS_GAP + 52# + OPS_GAP + 175# + OPS_GAP + 275# + _
                       OPS_GAP + acH + OPS_GAP + motH + 44#
End Function

Private Sub KVDark(ByVal px As Double, ByVal pw As Double, ByVal y As Double, _
                   ByVal k As String, ByVal v As String, ByVal sz As Double, _
                   ByVal clr As Long, ByVal bold As Boolean)
    Tx px + 11, y, k, sz, H("A8A8A8"), "start"
    Tx px + pw - 11, y, v, sz, clr, "end", bold
End Sub

Private Sub PanelHeadDark(ByVal px As Double, ByVal pw As Double, ByRef y As Double, ByVal t As String)
    Tx px + 11, y, t, 9.5, H("FFFFFF"), "start", True
    Ln px + 11, y + 4, px + pw - 11, y + 4, H("4A4A4A"), 0.75, msoLineSolid
    y = y + 16
End Sub

Private Sub SecBarDark(ByVal y As Double, ByVal t As String)
    Rect OPS_PAD, y, CANVAS_W - 2 * OPS_PAD, 26, H("2A2A2A"), H("3A3A3A"), 0.5
    Tx CANVAS_W / 2#, y + 18, t, 10, H("FFFFFF"), "middle", True
End Sub

Private Sub DrawOps(dt As Worksheet)
    Const ROW_PITCH As Double = 22#
    Dim y As Double
    Dim lx As Double, rx As Double, cw As Double
    Dim ly As Double, ry As Double
    Dim r As Long, n As Long

    cw = OpsColW()
    lx = OPS_PAD
    rx = OPS_PAD + cw + OPS_COL_GAP
    y = mOpsOrigin + OPS_GAP

    ' ---- header chips: plan / BHA / costs ------------------------------------
    Rect lx, y, CANVAS_W - 2 * OPS_PAD, 44, H("222222"), H("3A3A3A"), 0.75
    Tx lx + 14, y + 28, cellText(dt, "B2") & " " & cellText(dt, "C2"), 12, H("FFFFFF"), "start", True
    Tx lx + 220, y + 28, cellText(dt, "B3") & " " & cellText(dt, "C3"), 12, H("FFFFFF"), "start", True
    Tx lx + 400, y + 28, cellText(dt, "D3") & " " & cellText(dt, "E3"), 12, H("FFFFFF"), "start", False
    Tx CANVAS_W - OPS_PAD - 14, y + 28, cellText(dt, "D4") & " " & cellText(dt, "E4"), 12, H("FFFFFF"), "end", True
    y = y + 44 + OPS_GAP

    ' ---- day drilling | BHA totals -------------------------------------------
    Rect lx, y, cw, 167, H("222222"), H("3A3A3A"), 0.75
    Rect rx, y, cw, 167, H("222222"), H("3A3A3A"), 0.75
    ly = y + 20: ry = y + 20
    PanelHeadDark lx, cw, ly, "DAY DRILLING"
    KVDark lx, cw, ly, cellText(dt, "B4"), cellText(dt, "C4"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "B5"), cellText(dt, "C5"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "B6"), cellText(dt, "C6"), 11, H("5BA3D9"), True: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "B7"), cellText(dt, "C7"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "B8"), cellText(dt, "C8"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "B9"), cellText(dt, "C9"), 10.5, H("FFFFFF"), False: ly = ly + 16
    KVDark lx, cw, ly, cellText(dt, "B10"), cellText(dt, "C10"), 10.5, H("FFFFFF"), False: ly = ly + 16
    KVDark lx, cw, ly, cellText(dt, "B11"), cellText(dt, "C11"), 10.5, H("FFFFFF"), False

    PanelHeadDark rx, cw, ry, "TOTALS FOR BHA " & cellText(dt, "C12")
    KVDark rx, cw, ry, cellText(dt, "B13"), cellText(dt, "C13"), 11, H("FFFFFF"), False: ry = ry + 17
    KVDark rx, cw, ry, cellText(dt, "B14"), cellText(dt, "C14"), 11, H("FFFFFF"), False: ry = ry + 17
    KVDark rx, cw, ry, cellText(dt, "B15"), cellText(dt, "C15"), 11, H("FFFFFF"), False: ry = ry + 17
    KVDark rx, cw, ry, cellText(dt, "B16"), cellText(dt, "C16"), 11, H("FFFFFF"), False: ry = ry + 17
    KVDark rx, cw, ry, cellText(dt, "B17"), cellText(dt, "C17"), 11, H("FFFFFF"), False
    y = y + 167 + OPS_GAP

    ' ---- ROP + motor perf | motor info + 3rd party ---------------------------
    Rect lx, y, cw, 267, H("222222"), H("3A3A3A"), 0.75
    Rect rx, y, cw, 267, H("222222"), H("3A3A3A"), 0.75
    ly = y + 20: ry = y + 20
    PanelHeadDark lx, cw, ly, "PERFORMANCE"
    KVDark lx, cw, ly, cellText(dt, "B18"), cellText(dt, "C18"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "B19"), cellText(dt, "C19"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "B20"), cellText(dt, "C20"), 11, H("FFFFFF"), False: ly = ly + 20
    PanelHeadDark lx, cw, ly, "MOTOR PERFORMANCE"
    KVDark lx, cw, ly, cellText(dt, "D19"), cellText(dt, "E19"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "D20"), cellText(dt, "E20"), 11, H("FFFFFF"), False: ly = ly + 17
    KVDark lx, cw, ly, cellText(dt, "D21"), cellText(dt, "E21"), 11, H("FFFFFF"), False

    PanelHeadDark rx, cw, ry, "MOTOR INFORMATION"
    KVDark rx, cw, ry, cellText(dt, "B22"), cellText(dt, "C22"), 11, H("FFFFFF"), True: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B23"), cellText(dt, "C23"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B24"), cellText(dt, "C24"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B25"), cellText(dt, "C25"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B26"), cellText(dt, "C26"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B27"), cellText(dt, "C27"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B28"), cellText(dt, "C28"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B29"), cellText(dt, "C29"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B30"), cellText(dt, "C30"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B31"), cellText(dt, "C31"), 10.5, H("FFFFFF"), False: ry = ry + 16
    KVDark rx, cw, ry, cellText(dt, "B32"), cellText(dt, "C32"), 10.5, H("FFFFFF"), False: ry = ry + 18
    PanelHeadDark rx, cw, ry, cellText(dt, "D23")
    KVDark rx, cw, ry, cellText(dt, "D24"), _
           cellText(dt, "E24") & "  /  " & cellText(dt, "F24"), 10.5, H("FFFFFF"), False: ry = ry + 16
    PanelHeadDark rx, cw, ry, cellText(dt, "D29")
    KVDark rx, cw, ry, cellText(dt, "D30"), _
           cellText(dt, "E30") & "  /  " & cellText(dt, "F30"), 10.5, H("FFFFFF"), False
    y = y + 267 + OPS_GAP

    ' ---- AC Info (filled rows only) ------------------------------------------
    n = CountFilledAc(dt)
    SecBarDark y, cellText(dt, "B33")
    y = y + 36
    If n = 0 Then
        Rect OPS_PAD, y, CANVAS_W - 2 * OPS_PAD, 32, H("222222"), H("3A3A3A"), 0.5
        Tx OPS_PAD + 14, y + 20, "none", 11, H("A8A8A8"), "start"
        y = y + 32
    Else
        Rect OPS_PAD, y, CANVAS_W - 2 * OPS_PAD, 26# + ROW_PITCH * CDbl(n), H("222222"), H("3A3A3A"), 0.5
        Tx OPS_PAD + 14, y + 16, "Offset Well", 10, H("A8A8A8"), "start"
        Tx OPS_PAD + 620, y + 16, "SF", 10, H("A8A8A8"), "start"
        Tx OPS_PAD + 720, y + 16, "C2C (m)", 10, H("A8A8A8"), "start"
        Tx OPS_PAD + 860, y + 16, "Closest C2C", 10, H("A8A8A8"), "start"
        y = y + 24
        For r = 35 To 42
            If Len(Trim$(cellText(dt, "B" & r))) > 0 Then
                Tx OPS_PAD + 14, y + 2, cellText(dt, "B" & r), 10.5, H("FFFFFF"), "start"
                Tx OPS_PAD + 620, y + 2, cellText(dt, "D" & r), 10.5, H("FFFFFF"), "start"
                Tx OPS_PAD + 720, y + 2, cellText(dt, "E" & r), 10.5, H("FFFFFF"), "start"
                Tx OPS_PAD + 860, y + 2, cellText(dt, "F" & r), 10.5, H("FFFFFF"), "start"
                y = y + ROW_PITCH
            End If
        Next r
        y = y + 6
    End If
    y = y + OPS_GAP

    ' ---- Motors on location (filled rows only) -------------------------------
    n = CountFilledMotors(dt)
    SecBarDark y, cellText(dt, "B43")
    y = y + 36
    If n = 0 Then
        Rect OPS_PAD, y, CANVAS_W - 2 * OPS_PAD, 32, H("222222"), H("3A3A3A"), 0.5
        Tx OPS_PAD + 14, y + 20, "none", 11, H("A8A8A8"), "start"
        y = y + 32
    Else
        Rect OPS_PAD, y, CANVAS_W - 2 * OPS_PAD, 26# + ROW_PITCH * CDbl(n), H("222222"), H("3A3A3A"), 0.5
        Tx OPS_PAD + 14, y + 16, cellText(dt, "B44"), 10, H("A8A8A8"), "start"
        Tx OPS_PAD + 220, y + 16, cellText(dt, "C44"), 10, H("A8A8A8"), "start"
        Tx OPS_PAD + 420, y + 16, cellText(dt, "D44"), 10, H("A8A8A8"), "start"
        Tx OPS_PAD + 560, y + 16, cellText(dt, "E44"), 10, H("A8A8A8"), "start"
        Tx OPS_PAD + 720, y + 16, cellText(dt, "F44"), 10, H("A8A8A8"), "start"
        y = y + 24
        For r = 45 To 55
            If Len(Trim$(cellText(dt, "B" & r))) > 0 Then
                Tx OPS_PAD + 14, y + 2, cellText(dt, "B" & r), 11, H("FFFFFF"), "start"
                Tx OPS_PAD + 220, y + 2, cellText(dt, "C" & r), 11, H("FFFFFF"), "start"
                Tx OPS_PAD + 420, y + 2, cellText(dt, "D" & r), 11, H("FFFFFF"), "start"
                Tx OPS_PAD + 560, y + 2, cellText(dt, "E" & r), 11, H("FFFFFF"), "start"
                Tx OPS_PAD + 720, y + 2, cellText(dt, "F" & r), 11, H("FFFFFF"), "start"
                y = y + ROW_PITCH
            End If
        Next r
        y = y + 6
    End If

    If mUseSailBands Then
        Tx CANVAS_W - OPS_PAD, mCanvasH - 10, _
           "Plan " & cellText(dt, "C2") & " " & ChrW(183) & " geo " & ChrW(177) & Format$(mGeoHalf, "0.00") & _
           " m (AB14) " & ChrW(183) & " lateral " & ChrW(177) & Format$(mLatTol, "0.00") & " m (AA14)", _
           8.5, H("7A7A7A"), "end"
    Else
        Tx CANVAS_W - OPS_PAD, mCanvasH - 10, _
           "Plan " & cellText(dt, "C2") & " " & ChrW(183) & " shaft 1:1 (sail bands are lateral-only)", _
           8.5, H("7A7A7A"), "end"
    End If
End Sub


' ================================================================================
'  SHAPE PRIMITIVES
' ================================================================================
Private Function NextName() As String
    mSeq = mSeq + 1
    NextName = SHP_PREFIX & Format$(mSeq, "0000")
End Function

Private Sub Remember(ByVal nm As String)
    If mNameN > UBound(mNames) Then ReDim Preserve mNames(0 To mNameN + 512)
    mNames(mNameN) = nm
    mNameN = mNameN + 1
End Sub

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function H(ByVal hex6 As String) As Long
    H = RGB(CLng("&H" & mid$(hex6, 1, 2)), _
            CLng("&H" & mid$(hex6, 3, 2)), _
            CLng("&H" & mid$(hex6, 5, 2)))
End Function

' Paper theme: no flooded dark/colour panels. Thin dark strokes still identify
' hole / plan / required / held. Call sites keep the old hex names.
Private Function PrintFill(ByVal clr As Long) As Long
    If clr < 0 Then PrintFill = clr: Exit Function
    If clr = H("1A1A1A") Or clr = H("222222") Or clr = H("2A2A2A") _
            Or clr = H("242424") Or clr = H("252825") Then
        PrintFill = H("FFFFFF"): Exit Function
    End If
    If clr = H("1E3A2F") Or clr = H("1A3328") Or clr = H("1A3348") Then
        PrintFill = -1: Exit Function
    End If
    If clr = H("B6DBC8") Then PrintFill = H("888888"): Exit Function
    PrintFill = clr
End Function

Private Function PrintLine(ByVal clr As Long) As Long
    If clr < 0 Then PrintLine = clr: Exit Function
    If clr = H("FFFFFF") Then PrintLine = H("1A1A1A"): Exit Function
    If clr = H("A8CBB8") Or clr = H("B6DBC8") Or clr = H("9EC9B0") Then
        PrintLine = H("777777"): Exit Function
    End If
    If clr = H("3A3A3A") Or clr = H("333333") Or clr = H("444444") Then
        PrintLine = H("B0B0B0"): Exit Function
    End If
    If clr = H("5BA3D9") Then PrintLine = H("1A3A5C"): Exit Function
    If clr = H("4CAF7A") Then PrintLine = H("1B4D2E"): Exit Function
    If clr = H("3D6B54") Then PrintLine = H("1B4D2E"): Exit Function
    If clr = H("FF5555") Then PrintLine = H("8B0000"): Exit Function
    If clr = H("E09A3D") Then PrintLine = H("6B4A00"): Exit Function
    PrintLine = clr
End Function

Private Function PrintText(ByVal clr As Long) As Long
    If clr = H("FFFFFF") Then PrintText = H("1A1A1A"): Exit Function
    If clr = H("A8A8A8") Then PrintText = H("4A4A4A"): Exit Function
    If clr = H("7A7A7A") Then PrintText = H("555555"): Exit Function
    If clr = H("8AAB9A") Then PrintText = H("4A4A4A"): Exit Function
    If clr = H("4CAF7A") Then PrintText = H("1B4D2E"): Exit Function
    If clr = H("5BA3D9") Then PrintText = H("1A3A5C"): Exit Function
    If clr = H("FF5555") Then PrintText = H("8B0000"): Exit Function
    If clr = H("E09A3D") Then PrintText = H("6B4A00"): Exit Function
    If clr = H("B6DBC8") Then PrintText = H("555555"): Exit Function
    PrintText = clr
End Function

Private Sub StyleLine(shp As Shape, ByVal clr As Long, ByVal wt As Double, ByVal dash As Long)
    clr = PrintLine(clr)
    If clr < 0 Then
        shp.line.Visible = msoFalse
    Else
        shp.line.Visible = msoTrue
        shp.line.ForeColor.RGB = clr
        shp.line.Weight = wt
        shp.line.DashStyle = dash
    End If
End Sub

Private Sub StyleFill(shp As Shape, ByVal clr As Long)
    clr = PrintFill(clr)
    If clr < 0 Then
        shp.fill.Visible = msoFalse
    Else
        shp.fill.Visible = msoTrue
        shp.fill.Solid
        shp.fill.ForeColor.RGB = clr
    End If
End Sub

Private Sub Rect(ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal hgt As Double, _
                 ByVal fillClr As Long, ByVal lineClr As Long, ByVal wt As Double)
    If w < 0.1 Then w = 0.1
    If hgt < 0.1 Then hgt = 0.1
    Dim shp As Shape
    Set shp = mWs.Shapes.AddShape(msoShapeRectangle, l, t, w, hgt)
    shp.name = NextName(): Remember shp.name
    StyleFill shp, fillClr
    StyleLine shp, lineClr, wt, msoLineSolid
    shp.Shadow.Visible = msoFalse
End Sub

Private Sub Ln(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
               ByVal clr As Long, ByVal wt As Double, ByVal dash As Long)
    Dim shp As Shape
    Set shp = mWs.Shapes.AddLine(x1, y1, x2, y2)
    shp.name = NextName(): Remember shp.name
    StyleLine shp, clr, wt, dash
End Sub

Private Sub Dot(ByVal cx As Double, ByVal cy As Double, ByVal r As Double, _
                ByVal fillClr As Long, ByVal lineClr As Long, ByVal wt As Double)
    Dim shp As Shape
    Set shp = mWs.Shapes.AddShape(msoShapeOval, cx - r, cy - r, 2 * r, 2 * r)
    shp.name = NextName(): Remember shp.name
    StyleFill shp, fillClr
    StyleLine shp, lineClr, wt, msoLineSolid
    shp.Shadow.Visible = msoFalse
End Sub

' A closed four-sided face of the room, given in world coordinates.
Private Sub Quad(ByVal m1 As Double, ByVal d1 As Double, ByVal l1 As Double, _
                 ByVal m2 As Double, ByVal d2 As Double, ByVal l2 As Double, _
                 ByVal m3 As Double, ByVal d3 As Double, ByVal l3 As Double, _
                 ByVal m4 As Double, ByVal d4 As Double, ByVal l4 As Double, _
                 ByVal fillClr As Long, ByVal lineClr As Long, ByVal wt As Double)
    Dim fb As FreeformBuilder
    Set fb = mWs.Shapes.BuildFreeform(msoEditingCorner, ObX(m1, l1), ObY(d1, l1))
    fb.AddNodes msoSegmentLine, msoEditingAuto, ObX(m2, l2), ObY(d2, l2)
    fb.AddNodes msoSegmentLine, msoEditingAuto, ObX(m3, l3), ObY(d3, l3)
    fb.AddNodes msoSegmentLine, msoEditingAuto, ObX(m4, l4), ObY(d4, l4)
    fb.AddNodes msoSegmentLine, msoEditingAuto, ObX(m1, l1), ObY(d1, l1)

    Dim shp As Shape
    Set shp = fb.ConvertToShape
    shp.name = NextName(): Remember shp.name
    StyleFill shp, fillClr
    StyleLine shp, lineClr, wt, msoLineSolid
    shp.Shadow.Visible = msoFalse
End Sub

Private Sub PolyLine(xs() As Double, ys() As Double, ByVal n As Long, _
                     ByVal clr As Long, ByVal wt As Double, ByVal dash As Long)
    If n < 2 Then Exit Sub
    Dim fb As FreeformBuilder
    Set fb = mWs.Shapes.BuildFreeform(msoEditingCorner, xs(0), ys(0))
    Dim i As Long
    For i = 1 To n - 1
        fb.AddNodes msoSegmentLine, msoEditingAuto, xs(i), ys(i)
    Next i

    Dim shp As Shape
    Set shp = fb.ConvertToShape
    shp.name = NextName(): Remember shp.name
    shp.fill.Visible = msoFalse
    StyleLine shp, clr, wt, dash
    shp.Shadow.Visible = msoFalse
End Sub

Private Sub PolyFill(xs() As Double, ys() As Double, ByVal n As Long, _
                     ByVal fillClr As Long, ByVal lineClr As Long, ByVal wt As Double)
    If n < 3 Then Exit Sub
    Dim fb As FreeformBuilder
    Dim i As Long
    Dim shp As Shape
    Set fb = mWs.Shapes.BuildFreeform(msoEditingCorner, xs(0), ys(0))
    For i = 1 To n - 1
        fb.AddNodes msoSegmentLine, msoEditingAuto, xs(i), ys(i)
    Next i
    fb.AddNodes msoSegmentLine, msoEditingAuto, xs(0), ys(0)
    Set shp = fb.ConvertToShape
    shp.name = NextName(): Remember shp.name
    StyleFill shp, fillClr
    StyleLine shp, lineClr, wt, msoLineSolid
    shp.Shadow.Visible = msoFalse
End Sub

' Text placed the way SVG places it: x is the anchor, y is the baseline.
Private Sub Tx(ByVal x As Double, ByVal y As Double, ByVal s As String, _
               ByVal sz As Double, ByVal clr As Long, ByVal anchor As String, _
               Optional ByVal bold As Boolean = False, _
               Optional ByVal maxW As Double = 0#)
    ' The box is only a container for the alignment; it is invisible. It must not
    ' run past the canvas edge, because the group's bounding box is what gets
    ' exported and a box hanging over the right edge silently widens the image.
    Dim l As Double, w As Double, align As Long
    clr = PrintText(clr)
    If mTypeScale > 0.01 Then sz = sz * mTypeScale
    Select Case LCase$(anchor)
        Case "end"
            If maxW > 0# Then
                w = maxW
                l = x - w
                If l < 0# Then l = 0#: w = x
            Else
                l = 0: w = x
            End If
            align = msoAlignRight
        Case "middle"
            w = 2 * MinD(x, CANVAS_W - x)
            If maxW > 0# And maxW < w Then w = maxW
            l = x - w / 2: align = msoAlignCenter
        Case Else
            l = x
            If maxW > 0# Then
                w = maxW
            Else
                w = CANVAS_W - x
            End If
            align = msoAlignLeft
    End Select
    If w < 8 Then w = 8
    If l + w > CANVAS_W Then w = CANVAS_W - l
    If w < 8 Then w = 8

    Dim shp As Shape
    Set shp = mWs.Shapes.AddTextbox(msoTextOrientationHorizontal, l, y - sz * 1.25, w, sz * 1.8)
    shp.name = NextName(): Remember shp.name
    shp.fill.Visible = msoFalse
    shp.line.Visible = msoFalse
    shp.Shadow.Visible = msoFalse
    With shp.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
        .VerticalAnchor = msoAnchorTop
        With .TextRange
            .text = s
            .ParagraphFormat.Alignment = align
            .Font.name = "Calibri"
            .Font.Size = sz
            .Font.bold = IIf(bold, msoTrue, msoFalse)
            .Font.fill.ForeColor.RGB = clr
        End With
    End With
End Sub


' ================================================================================
'  EXPORT
' ================================================================================
' Groups everything drawn and exports it through a throwaway chart. Shapes.Range
' gets unreliable with very large name arrays, so the group is built in chunks and
' the chunks are then grouped together.
Private Function ExportGroup(ByVal outPath As String) As String
    ExportGroup = ""
    Trace "  ExportGroup shapes=" & mNameN
    If mNameN = 0 Then Exit Function

    Dim current As Collection
    Set current = New Collection
    Dim i As Long
    For i = 0 To mNameN - 1
        current.Add mNames(i)
    Next i

    Dim nxt As Collection
    Do While current.Count > 1
        Set nxt = New Collection          ' must be a fresh one each pass
        Dim idx As Long: idx = 1
        Do While idx <= current.Count
            Dim take As Long
            take = current.Count - idx + 1
            If take > 60 Then take = 60
            If take = 1 Then
                nxt.Add current(idx)
                idx = idx + 1
            Else
                Dim arr() As Variant
                ReDim arr(0 To take - 1)
                Dim k As Long
                For k = 0 To take - 1
                    arr(k) = current(idx + k)
                Next k
                Dim g As Shape
                Set g = mWs.Shapes.Range(arr).Group
                g.name = NextName(): Remember g.name
                nxt.Add g.name
                idx = idx + take
            End If
        Loop
        Set current = nxt
    Loop

    Dim grp As Shape
    Set grp = mWs.Shapes(current(1))
    Trace "  grouped into " & grp.name & " " & Format$(grp.Width, "0") & "x" & Format$(grp.Height, "0") & " pt"

    ' Activate the sheet/chart under ScreenUpdating=False so nothing paints.
    ' Only force the application visible when it was started invisible (automation);
    ' never flash a session the user already has open.
    Dim prevVisible As Boolean
    Dim prevSheet As Object
    Dim forcedVisible As Boolean
    Dim exportW As Double, exportH As Double
    Dim pic As Shape
    prevVisible = Application.Visible
    Set prevSheet = mWs.Parent.ActiveSheet
    forcedVisible = False
    If Not prevVisible Then
        Application.Visible = True
        forcedVisible = True
    End If
    mWs.Activate

    ' Always export the paper canvas (0,0)-(CANVAS_W, mCanvasH). Off-canvas
    ' strokes must not inflate the EMAIL PNG to sheet-sized.
    exportW = CANVAS_W
    exportH = mCanvasH
    If exportH < 40# Then exportH = 40#

    Dim co As ChartObject
    Set co = mWs.ChartObjects.Add(0, 0, exportW, exportH)
    co.name = SHP_PREFIX & "CHART"
    Trace "  chart added " & Format$(exportW, "0") & "x" & Format$(exportH, "0") & " pt"

    On Error GoTo Fail
    co.Chart.ChartArea.Border.LineStyle = xlNone
    co.Chart.ChartArea.fill.Visible = msoFalse

    grp.CopyPicture xlScreen, xlPicture
    Trace "  copied to clipboard"

    ' Chart.Paste only takes a picture when the chart is the active object; called
    ' on an inactive chart it reports success and pastes nothing.
    co.Activate
    co.Chart.Paste
    Trace "  pasted, chart shapes=" & co.Chart.Shapes.Count

    If co.Chart.Shapes.Count = 0 Then
        grp.CopyPicture xlScreen, xlBitmap
        co.Chart.Paste
        Trace "  bitmap retry, chart shapes=" & co.Chart.Shapes.Count
    End If
    If co.Chart.Shapes.Count = 0 Then Err.Raise 5, , "nothing pasted into the chart"

    Set pic = co.Chart.Shapes(1)
    pic.LockAspectRatio = msoFalse
    pic.Left = -grp.Left
    pic.Top = -grp.Top
    pic.Width = grp.Width
    pic.Height = grp.Height

    co.Chart.Export Filename:=outPath, FilterName:="PNG"
    Trace "  exported " & outPath

    co.Delete
    prevSheet.Activate
    If forcedVisible Then Application.Visible = prevVisible
    ExportGroup = outPath
    Exit Function

Fail:
    mLastError = "ExportGroup err " & Err.Number & " " & Err.Description
    Trace "  EXPORT FAIL " & mLastError
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Not prevSheet Is Nothing Then prevSheet.Activate
    If forcedVisible Then Application.Visible = prevVisible
End Function

Private Sub DeleteDrawn()
    Dim i As Long
    For i = mWs.Shapes.Count To 1 Step -1
        If Left$(mWs.Shapes(i).name, Len(SHP_PREFIX)) = SHP_PREFIX Then mWs.Shapes(i).Delete
    Next i
    Dim co As ChartObject
    For Each co In mWs.ChartObjects
        If Left$(co.name, Len(SHP_PREFIX)) = SHP_PREFIX Then co.Delete
    Next co
    mNameN = 0
End Sub


' ================================================================================
'  CELL HELPERS
' ================================================================================
Private Function NumCell(rg As Range) As Double
    Dim v As Variant
    On Error Resume Next
    v = rg.Value2
    On Error GoTo 0
    If IsArray(v) Then Exit Function
    If IsNumeric(v) Then NumCell = CDbl(v)
End Function

' Lateral half-width from Slidesheet AA14 (numeric, or text like "+/-10.00 m").
Private Function LatTolFromCell(rg As Range) As Double
    Dim v As Variant
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim buf As String
    Dim started As Boolean

    LatTolFromCell = 0#
    On Error Resume Next
    v = rg.Value2
    On Error GoTo 0
    If Not IsArray(v) Then
        If IsNumeric(v) Then
            LatTolFromCell = Abs(CDbl(v))
            Exit Function
        End If
    End If

    On Error Resume Next
    s = Trim$(CStr(rg.text))
    On Error GoTo 0
    If Len(s) = 0 Then Exit Function

    buf = ""
    started = False
    For i = 1 To Len(s)
        ch = mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            buf = buf & ch
            started = True
        ElseIf started Then
            Exit For
        End If
    Next i
    If Len(buf) > 0 And IsNumeric(buf) Then LatTolFromCell = Abs(CDbl(buf))
End Function

Private Function cellText(ws As Worksheet, ByVal addr As String) As String
    On Error Resume Next
    cellText = Trim$(CStr(ws.Range(addr).text))
    On Error GoTo 0
    ' the sheet prints its own trailing colons on labels; the panel adds its own layout
    If right$(cellText, 1) = ":" Then cellText = Left$(cellText, Len(cellText) - 1)
End Function





