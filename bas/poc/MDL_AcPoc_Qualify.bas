Attribute VB_Name = "MDL_AcPoc_Qualify"
Option Explicit

' Data-tab rule: sort all summary rows by SF then C2C, keep `capacity`,
' re-sort that slice by Ref MD. Plot set = unique wells in the kept slice.

Public Sub AcPoc_Qualify()
    Dim src As Worksheet
    Dim dst As Worksheet
    Dim lastR As Long
    Dim n As Long
    Dim i As Long
    Dim cap As Long
    Dim nShow As Long
    Dim well() As String
    Dim md() As Double, bc() As Double, sf() As Double, offMd() As Double
    Dim warn() As String
    Dim seen As Object
    Dim key As String
    Dim nUniq As Long

    Set src = AcPoc_EnsureSheet("_POC_AC_Summary", True)
    lastR = AcPoc_LastDataRow(src, 1)
    If lastR < 2 Then
        AcPoc_Log "Qualify", "FAIL", "No AC summary rows to qualify"
        Exit Sub
    End If

    n = lastR - 1
    ReDim well(0 To n - 1)
    ReDim md(0 To n - 1)
    ReDim offMd(0 To n - 1)
    ReDim bc(0 To n - 1)
    ReDim sf(0 To n - 1)
    ReDim warn(0 To n - 1)
    For i = 0 To n - 1
        well(i) = CStr(src.Cells(i + 2, 1).Value2 & "")
        md(i) = CDbl(src.Cells(i + 2, 2).Value2)
        offMd(i) = CDbl(src.Cells(i + 2, 3).Value2)
        bc(i) = CDbl(src.Cells(i + 2, 4).Value2)
        sf(i) = CDbl(src.Cells(i + 2, 6).Value2)
        warn(i) = CStr(src.Cells(i + 2, 7).Value2 & "")
    Next i

    AcPoc_SortAcRows n, well, md, offMd, bc, sf, warn, False

    cap = AcPoc_Capacity()
    nShow = n
    If nShow > cap Then nShow = cap

    AcPoc_SortAcRows nShow, well, md, offMd, bc, sf, warn, True

    Set dst = AcPoc_EnsureSheet("_POC_Qualify", True)
    AcPoc_ClearSheet dst
    dst.Range("A1").Value2 = "Well"
    dst.Range("B1").Value2 = "RefMD"
    dst.Range("C1").Value2 = "OffMD"
    dst.Range("D1").Value2 = "SF"
    dst.Range("E1").Value2 = "C2C"
    dst.Range("F1").Value2 = "Warning"
    dst.Range("G1").Value2 = "Plot"
    dst.Range("A1:G1").Font.Bold = True

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    nUniq = 0
    For i = 0 To nShow - 1
        dst.Cells(i + 2, 1).Value2 = well(i)
        dst.Cells(i + 2, 2).Value2 = md(i)
        dst.Cells(i + 2, 3).Value2 = offMd(i)
        dst.Cells(i + 2, 4).Value2 = sf(i)
        dst.Cells(i + 2, 4).NumberFormat = "0.000"
        dst.Cells(i + 2, 5).Value2 = bc(i)
        dst.Cells(i + 2, 5).NumberFormat = "0.00"
        dst.Cells(i + 2, 6).Value2 = warn(i)
        key = AcPoc_WellKey(well(i))
        If Not seen.Exists(key) Then
            seen.Add key, well(i)
            nUniq = nUniq + 1
            dst.Cells(i + 2, 7).Value2 = "Y"
        Else
            dst.Cells(i + 2, 7).Value2 = ""
        End If
        dst.Range(dst.Cells(i + 2, 1), dst.Cells(i + 2, 7)).Interior.Color = AcPoc_BandColor(sf(i))
    Next i

    dst.Range("I1").Value2 = "Unique wells"
    dst.Range("J1").Value2 = nUniq

    AcPoc_Log "Qualify", "OK", "Kept " & nShow & " of " & n & " summary rows; " & _
        nUniq & " unique wells for the plot"
End Sub

Public Function AcPoc_BandColor(ByVal sf As Double) As Long
    If sf < 1.5 Then
        AcPoc_BandColor = RGB(255, 199, 206)
    ElseIf sf < 2# Then
        AcPoc_BandColor = RGB(255, 235, 156)
    Else
        AcPoc_BandColor = RGB(198, 239, 206)
    End If
End Function

Public Function AcPoc_InkColor(ByVal sf As Double) As Long
    If sf < 1.5 Then
        AcPoc_InkColor = RGB(198, 40, 40)
    ElseIf sf < 2# Then
        AcPoc_InkColor = RGB(249, 168, 37)
    Else
        AcPoc_InkColor = RGB(46, 125, 50)
    End If
End Function

Private Sub AcPoc_SortAcRows(ByVal n As Long, well() As String, md() As Double, _
        offMd() As Double, bc() As Double, sf() As Double, warn() As String, _
        ByVal byDepth As Boolean)
    Dim i As Long, j As Long
    Dim kWell As String, kWarn As String
    Dim kMd As Double, kOff As Double, kBc As Double, kSf As Double

    For i = 1 To n - 1
        kWell = well(i): kMd = md(i): kOff = offMd(i)
        kBc = bc(i): kSf = sf(i): kWarn = warn(i)
        j = i - 1
        Do While j >= 0
            If AcPoc_KeepsPlace(sf(j), bc(j), md(j), kSf, kBc, kMd, byDepth) Then Exit Do
            well(j + 1) = well(j)
            md(j + 1) = md(j)
            offMd(j + 1) = offMd(j)
            bc(j + 1) = bc(j)
            sf(j + 1) = sf(j)
            warn(j + 1) = warn(j)
            j = j - 1
        Loop
        well(j + 1) = kWell: md(j + 1) = kMd: offMd(j + 1) = kOff
        bc(j + 1) = kBc: sf(j + 1) = kSf: warn(j + 1) = kWarn
    Next i
End Sub

Private Function AcPoc_KeepsPlace( _
        ByVal sfA As Double, ByVal bcA As Double, ByVal mdA As Double, _
        ByVal sfB As Double, ByVal bcB As Double, ByVal mdB As Double, _
        ByVal byDepth As Boolean) As Boolean
    If byDepth Then
        If mdA < mdB Then AcPoc_KeepsPlace = True: Exit Function
        If mdA > mdB Then AcPoc_KeepsPlace = False: Exit Function
        AcPoc_KeepsPlace = (sfA <= sfB)
    Else
        If sfA < sfB Then AcPoc_KeepsPlace = True: Exit Function
        If sfA > sfB Then AcPoc_KeepsPlace = False: Exit Function
        AcPoc_KeepsPlace = (bcA <= bcB)
    End If
End Function
