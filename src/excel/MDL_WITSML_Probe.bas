Option Explicit

' ==============================================================================
' WITSML NETWORK PROBE (VBA)
' ------------------------------------------------------------------------------
' Purpose:
'   - Probe a list/subnet for likely WITSML SOAP (WMLS 1.4.1.1) endpoints
'   - Call WMLS_GetCap
'   - Attempt WMLS_GetFromStore for:
'       * trajectory (surveys)
'       * tubular   (pipe tally candidate object)
'
' Import:
'   VBA Editor -> File -> Import File... -> select this .bas
'
' Notes:
'   - Use only on networks/systems you are authorized to test.
'   - Many terminals require auth/TLS/client cert; this is a test harness.
' ==============================================================================

Private Type ProbeResult
    Host As String
    Url As String
    Reachable As Boolean
    HttpOk As Boolean
    GetCapOk As Boolean
    SurveyOk As Boolean
    TubularOk As Boolean
    Notes As String
End Type

Private Const SHEET_NAME As String = "WITSML_PROBE"
Private Const SHEET_PAYLOAD As String = "WITSML_PAYLOAD"
Private Const ACTION_GETCAP As String = "http://www.witsml.org/action/120/WMLS_GetCap"
Private Const ACTION_GETFROMSTORE As String = "http://www.witsml.org/action/120/WMLS_GetFromStore"
Private Const TARGET_HOST As String = "10.100.14.60"
Private Const SINGLE_TARGET_MODE As Boolean = True
Private Const LOW_NOISE_MODE As Boolean = True

' Optional auth settings (leave blank for unauthenticated test)
Private Const WITSML_USER As String = ""
Private Const WITSML_PASS As String = ""

Private mPayloadWs As Worksheet
Private mPayloadRow As Long

Public Sub WITSML_RunProbe()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_NAME)
    WriteHeader ws

    Set mPayloadWs = EnsureSheet(SHEET_PAYLOAD)
    WritePayloadHeader mPayloadWs
    mPayloadRow = 2

    ' -------------------------------
    ' Configure scan targets
    ' -------------------------------
    Dim hosts As Collection
    Set hosts = New Collection

    ' Preferred: explicit list
    hosts.Add TARGET_HOST
    If Not SINGLE_TARGET_MODE Then
        hosts.Add "witsml-terminal.local"
        hosts.Add "10.10.20.11"
        hosts.Add "10.10.20.12"
    End If

    ' If you want a subnet sweep instead, comment out explicit hosts above
    ' and uncomment this line:
    ' Set hosts = BuildHostsFromSubnet("10.10.20.", 1, 254)

    Dim endpointPaths As Variant
    endpointPaths = Array("/WITSML/WMLS.asmx")

    Dim ports As Variant
    ports = Array(80)

    Dim schemes As Variant
    schemes = Array("http")

    If Not LOW_NOISE_MODE Then
        endpointPaths = Array( _
            "/WITSML/WMLS.asmx", _
            "/witsml/wmls.asmx", _
            "/WMLS/WMLS.asmx", _
            "/wmls/wmls.asmx" _
        )
        ports = Array(80, 443)
        schemes = Array("http", "https")
    End If

    Dim outRow As Long
    outRow = 2

    Dim h As Variant, p As Variant, s As Variant, ep As Variant
    For Each h In hosts
        Dim host As String
        host = CStr(h)

        If IsHostReachable(host) = False Then
            Dim rr As ProbeResult
            rr.Host = host
            rr.Reachable = False
            rr.Notes = "Ping unreachable or blocked."
            WriteResult ws, outRow, rr
            outRow = outRow + 1
            GoTo NextHost
        End If

        Dim endpointFound As Boolean
        endpointFound = False

        For Each s In schemes
            If endpointFound Then Exit For
            For Each p In ports
                If endpointFound Then Exit For
                For Each ep In endpointPaths
                    Dim url As String
                    url = BuildUrl(CStr(s), host, CLng(p), CStr(ep))

                    Dim r As ProbeResult
                    r = ProbeSingleEndpoint(host, url)
                    WriteResult ws, outRow, r
                    outRow = outRow + 1

                    ' Stop scanning once we find a valid WITSML endpoint.
                    If r.GetCapOk Then
                        endpointFound = True
                        Exit For
                    End If

                    ' In low-noise mode, do not iterate aggressively.
                    If LOW_NOISE_MODE Then
                        endpointFound = True
                        Exit For
                    End If

                    SleepMs 500
                Next ep
            Next p
        Next s

NextHost:
    Next h

    ws.Columns("A:I").AutoFit
    mPayloadWs.Columns("A:G").AutoFit
    MsgBox "WITSML probe complete. See sheet '" & SHEET_NAME & "'.", vbInformation
    Exit Sub

EH:
    MsgBox "WITSML_RunProbe error: " & Err.Description, vbExclamation
End Sub

Private Function ProbeSingleEndpoint(ByVal host As String, ByVal url As String) As ProbeResult
    Dim pr As ProbeResult
    pr.Host = host
    pr.Url = url
    pr.Reachable = True

    Dim capResp As String
    capResp = HttpPostSoap(url, ACTION_GETCAP, BuildGetCapEnvelope())
    WritePayload host, url, "WMLS_GetCap", capResp

    If Len(capResp) = 0 Then
        pr.Notes = "No HTTP response."
        ProbeSingleEndpoint = pr
        Exit Function
    End If
    pr.HttpOk = True

    If InStr(1, capResp, "WMLS_GetCapResult", vbTextCompare) > 0 Or _
       InStr(1, capResp, "<capServers", vbTextCompare) > 0 Then
        pr.GetCapOk = True
    Else
        pr.Notes = "HTTP OK but not recognized as WITSML GetCap endpoint."
        ProbeSingleEndpoint = pr
        Exit Function
    End If

    Dim surveyResp As String
    surveyResp = HttpPostSoap( _
        url, _
        ACTION_GETFROMSTORE, _
        BuildGetFromStoreEnvelope("trajectory", BuildTrajectoryQuery()) _
    )
    WritePayload host, url, "WMLS_GetFromStore:trajectory", surveyResp
    If InStr(1, surveyResp, "WMLS_GetFromStoreResult", vbTextCompare) > 0 Or _
       InStr(1, surveyResp, "<trajectory", vbTextCompare) > 0 Then
        pr.SurveyOk = True
    End If

    Dim tubularResp As String
    tubularResp = HttpPostSoap( _
        url, _
        ACTION_GETFROMSTORE, _
        BuildGetFromStoreEnvelope("tubular", BuildTubularQuery()) _
    )
    WritePayload host, url, "WMLS_GetFromStore:tubular", tubularResp
    If InStr(1, tubularResp, "WMLS_GetFromStoreResult", vbTextCompare) > 0 Or _
       InStr(1, tubularResp, "<tubular", vbTextCompare) > 0 Then
        pr.TubularOk = True
    End If

    If pr.SurveyOk Or pr.TubularOk Then
        pr.Notes = "Success."
    Else
        pr.Notes = "GetCap OK; trajectory/tubular not returned (auth/filter/object mismatch possible)."
    End If

    ProbeSingleEndpoint = pr
End Function

Private Function HttpPostSoap(ByVal url As String, ByVal soapAction As String, ByVal body As String) As String
    On Error GoTo EH

    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Keep probe responsive; avoid "endless" waits on dead endpoints.
    xhr.setTimeouts 1200, 1200, 2500, 2500
    xhr.Open "POST", url, False
    xhr.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    xhr.setRequestHeader "SOAPAction", soapAction

    If Len(WITSML_USER) > 0 Then
        xhr.setRequestHeader "Authorization", "Basic " & Base64Encode(WITSML_USER & ":" & WITSML_PASS)
    End If

    xhr.send body

    If xhr.Status >= 200 And xhr.Status < 500 Then
        HttpPostSoap = CStr(xhr.responseText)
    Else
        HttpPostSoap = ""
    End If
    Exit Function

EH:
    HttpPostSoap = ""
End Function

Private Sub SleepMs(ByVal ms As Long)
    Dim t As Single
    t = Timer + (ms / 1000#)
    Do While Timer < t
        DoEvents
    Loop
End Sub

Private Function BuildGetCapEnvelope() As String
    BuildGetCapEnvelope = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
        "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
        "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<WMLS_GetCap xmlns=""http://www.witsml.org/message/120"">" & _
              "<OptionsIn></OptionsIn>" & _
            "</WMLS_GetCap>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"
End Function

Private Function BuildGetFromStoreEnvelope(ByVal wmlType As String, ByVal queryXml As String) As String
    BuildGetFromStoreEnvelope = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " & _
        "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
        "xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<WMLS_GetFromStore xmlns=""http://www.witsml.org/message/120"">" & _
              "<WMLtypeIn>" & wmlType & "</WMLtypeIn>" & _
              "<QueryIn><![CDATA[" & queryXml & "]]></QueryIn>" & _
              "<OptionsIn>returnElements=all</OptionsIn>" & _
              "<CapabilitiesIn></CapabilitiesIn>" & _
            "</WMLS_GetFromStore>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"
End Function

Private Function BuildTrajectoryQuery() As String
    BuildTrajectoryQuery = _
        "<trajectorys version=""1.4.1.1"">" & _
          "<trajectory uidWell="""" uidWellbore="""" uid="""">" & _
            "<nameWell></nameWell><nameWellbore></nameWellbore>" & _
          "</trajectory>" & _
        "</trajectorys>"
End Function

Private Function BuildTubularQuery() As String
    BuildTubularQuery = _
        "<tubulars version=""1.4.1.1"">" & _
          "<tubular uidWell="""" uidWellbore="""" uid="""">" & _
            "<nameWell></nameWell><nameWellbore></nameWellbore>" & _
          "</tubular>" & _
        "</tubulars>"
End Function

Private Function IsHostReachable(ByVal host As String) As Boolean
    On Error GoTo EH
    Dim sh As Object, execObj As Object, outText As String
    Set sh = CreateObject("WScript.Shell")
    Set execObj = sh.Exec("cmd /c ping -n 1 -w 400 " & host)
    outText = execObj.StdOut.ReadAll
    IsHostReachable = (InStr(1, outText, "TTL=", vbTextCompare) > 0)
    Exit Function
EH:
    IsHostReachable = False
End Function

Private Function BuildHostsFromSubnet(ByVal subnetPrefix As String, ByVal startHost As Long, ByVal endHost As Long) As Collection
    Dim c As New Collection
    Dim i As Long
    For i = startHost To endHost
        c.Add subnetPrefix & CStr(i)
    Next i
    Set BuildHostsFromSubnet = c
End Function

Private Function BuildUrl(ByVal scheme As String, ByVal host As String, ByVal port As Long, ByVal path As String) As String
    BuildUrl = scheme & "://" & host & ":" & CStr(port) & path
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add
        EnsureSheet.Name = sheetName
    End If
End Function

Private Sub WriteHeader(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:I1").Value = Array( _
        "Host", "URL", "Reachable", "HTTP", "GetCap", "Survey(trajectory)", "Pipe tally(tubular)", "Notes", "Timestamp" _
    )
    ws.Rows(1).Font.Bold = True
End Sub

Private Sub WriteResult(ByVal ws As Worksheet, ByVal rowNum As Long, ByRef r As ProbeResult)
    ws.Cells(rowNum, 1).Value = r.Host
    ws.Cells(rowNum, 2).Value = r.Url
    ws.Cells(rowNum, 3).Value = IIf(r.Reachable, "Y", "N")
    ws.Cells(rowNum, 4).Value = IIf(r.HttpOk, "Y", "N")
    ws.Cells(rowNum, 5).Value = IIf(r.GetCapOk, "Y", "N")
    ws.Cells(rowNum, 6).Value = IIf(r.SurveyOk, "Y", "N")
    ws.Cells(rowNum, 7).Value = IIf(r.TubularOk, "Y", "N")
    ws.Cells(rowNum, 8).Value = r.Notes
    ws.Cells(rowNum, 9).Value = Now
End Sub

Private Sub WritePayloadHeader(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Host", "URL", "Action", "Has WITSML Tag", "Length", "Snippet", "Timestamp")
    ws.Rows(1).Font.Bold = True
End Sub

Private Sub WritePayload(ByVal host As String, ByVal url As String, ByVal actionName As String, ByVal payload As String)
    If mPayloadWs Is Nothing Then Exit Sub
    mPayloadWs.Cells(mPayloadRow, 1).Value = host
    mPayloadWs.Cells(mPayloadRow, 2).Value = url
    mPayloadWs.Cells(mPayloadRow, 3).Value = actionName
    mPayloadWs.Cells(mPayloadRow, 4).Value = IIf(HasWitsmlTag(payload), "Y", "N")
    mPayloadWs.Cells(mPayloadRow, 5).Value = Len(payload)
    mPayloadWs.Cells(mPayloadRow, 6).Value = PayloadSnippet(payload, 1800)
    mPayloadWs.Cells(mPayloadRow, 7).Value = Now
    mPayloadRow = mPayloadRow + 1
End Sub

Private Function HasWitsmlTag(ByVal payload As String) As Boolean
    Dim s As String
    s = LCase$(payload)
    HasWitsmlTag = (InStr(1, s, "<trajectory", vbTextCompare) > 0) Or _
                   (InStr(1, s, "<tubular", vbTextCompare) > 0) Or _
                   (InStr(1, s, "<capservers", vbTextCompare) > 0) Or _
                   (InStr(1, s, "wmls_getfromstoreresult", vbTextCompare) > 0) Or _
                   (InStr(1, s, "wmls_getcapresult", vbTextCompare) > 0)
End Function

Private Function PayloadSnippet(ByVal payload As String, ByVal maxLen As Long) As String
    If Len(payload) <= maxLen Then
        PayloadSnippet = payload
    Else
        PayloadSnippet = Left$(payload, maxLen) & "...[truncated]"
    End If
End Function

' Basic Base64 helper for optional auth header.
' Uses MSXML DOM typed-value conversion to avoid external dependencies.
Private Function Base64Encode(ByVal plainText As String) As String
    Dim xmlObj As Object
    Dim nodeObj As Object
    Set xmlObj = CreateObject("MSXML2.DOMDocument.6.0")
    Set nodeObj = xmlObj.createElement("b64")
    nodeObj.DataType = "bin.base64"
    nodeObj.nodeTypedValue = StrConv(plainText, vbFromUnicode)
    Base64Encode = Replace(nodeObj.Text, vbLf, "")
End Function

