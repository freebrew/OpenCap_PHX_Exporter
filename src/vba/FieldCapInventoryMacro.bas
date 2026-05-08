Attribute VB_Name = "FieldCapInventoryMacro"
Option Explicit

Private Const FIELD_CAP_QUERY_NAME As String = "FieldCapJob20786Inventory"
Private Const FIELD_CAP_SHEET_NAME As String = "FieldCap Inventory"
Private Const FIELD_CAP_TABLE_NAME As String = "FieldCapInventoryTable"

Public Sub FetchFieldCapInventory()
    Dim ws As Worksheet
    Dim listObject As ListObject
    Dim connectionString As String

    Set ws = GetOrCreateWorksheet(FIELD_CAP_SHEET_NAME)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo CleanFail

    ClearInventorySheet ws
    DeleteWorkbookQuery FIELD_CAP_QUERY_NAME

    ThisWorkbook.Queries.Add Name:=FIELD_CAP_QUERY_NAME, Formula:=BuildFieldCapInventoryQuery()

    connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & _
        FIELD_CAP_QUERY_NAME & """;Extended Properties="""""

    Set listObject = ws.ListObjects.Add(SourceType:=0, Source:=connectionString, Destination:=ws.Range("A1"))

    With listObject
        .Name = FIELD_CAP_TABLE_NAME
        .TableStyle = "TableStyleMedium2"
    End With

    With listObject.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & FIELD_CAP_QUERY_NAME & "]")
        .Refresh BackgroundQuery:=False
    End With

    ws.Columns("A:G").AutoFit
    ws.Activate
    ws.Range("A1").Select

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "FieldCap inventory refresh failed." & vbCrLf & vbCrLf & _
        "Excel may ask you to sign in to FieldCap/Microsoft the first time this runs." & vbCrLf & _
        "Error: " & Err.Description, vbExclamation, "FieldCap Inventory"
End Sub

Public Sub RefreshFieldCapInventory()
    FetchFieldCapInventory
End Sub

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

Private Sub ClearInventorySheet(ByVal ws As Worksheet)
    Dim listObject As ListObject

    For Each listObject In ws.ListObjects
        listObject.Delete
    Next listObject

    ws.Cells.Clear
End Sub

Private Sub DeleteWorkbookQuery(ByVal queryName As String)
    On Error Resume Next
    ThisWorkbook.Queries(queryName).Delete
    On Error GoTo 0
End Sub

Private Function BuildFieldCapInventoryQuery() As String
    Dim lines As Collection
    Set lines = New Collection

    AddQueryLine lines, "let"
    AddQueryLine lines, "    ODataUrl = ""https://fieldcap-cdn.phxtech.com/odata/ToolAssemblyItems?$expand=JobTool,JobTool($expand=Item),JobTool($expand=ItemSerial),ToolAssembly&$filter=((null%20eq%20DeletedBy)%20and%20(ToolAssembly/ClientJobId%20eq%2020786))"","
    AddQueryLine lines, "    FirstText = (values as list) as nullable text =>"
    AddQueryLine lines, "        let"
    AddQueryLine lines, "            CleanValues = List.RemoveNulls(values),"
    AddQueryLine lines, "            FirstValue = if List.Count(CleanValues) = 0 then null else CleanValues{0}"
    AddQueryLine lines, "        in"
    AddQueryLine lines, "            if FirstValue = null then null else Text.From(FirstValue),"
    AddQueryLine lines, "    FirstNumber = (values as list) as nullable number =>"
    AddQueryLine lines, "        let"
    AddQueryLine lines, "            Numbers = List.RemoveNulls(List.Transform(values, each try Number.From(_) otherwise null))"
    AddQueryLine lines, "        in"
    AddQueryLine lines, "            if List.Count(Numbers) = 0 then null else Numbers{0},"
    AddQueryLine lines, "    FormatMinutes = (value as any) as nullable text =>"
    AddQueryLine lines, "        let"
    AddQueryLine lines, "            RawMinutes = try Number.RoundDown(Number.From(value)) otherwise null,"
    AddQueryLine lines, "            Hours = if RawMinutes = null then null else Number.IntegerDivide(RawMinutes, 60),"
    AddQueryLine lines, "            Minutes = if RawMinutes = null then null else Number.Mod(RawMinutes, 60)"
    AddQueryLine lines, "        in"
    AddQueryLine lines, "            if RawMinutes = null then null else Text.From(Hours) & "":"" & Text.PadStart(Text.From(Minutes), 2, ""0""),"
    AddQueryLine lines, "    RecordFields = (tableValue as table, columnName as text, desiredFields as list) as list =>"
    AddQueryLine lines, "        if Table.HasColumns(tableValue, columnName) then"
    AddQueryLine lines, "            let"
    AddQueryLine lines, "                Records = List.RemoveNulls(Table.Column(tableValue, columnName)),"
    AddQueryLine lines, "                AvailableFields = if List.Count(Records) = 0 then {} else Record.FieldNames(Records{0})"
    AddQueryLine lines, "            in"
    AddQueryLine lines, "                List.Intersect({desiredFields, AvailableFields})"
    AddQueryLine lines, "        else"
    AddQueryLine lines, "            {},"
    AddQueryLine lines, "    FieldNamesContaining = (row as record, tokens as list) as list =>"
    AddQueryLine lines, "        List.Select(Record.FieldNames(row), (fieldName) => List.AllTrue(List.Transform(tokens, (token) => Text.Contains(Text.Lower(fieldName), Text.Lower(token))))),"
    AddQueryLine lines, "    FirstNumberByNames = (row as record, names as list) as nullable number =>"
    AddQueryLine lines, "        FirstNumber(List.Transform(names, each Record.FieldOrDefault(row, _, null))),"
    AddQueryLine lines, "    FirstNumberByTokenGroups = (row as record, tokenGroups as list) as nullable number =>"
    AddQueryLine lines, "        let"
    AddQueryLine lines, "            Matches = List.RemoveNulls(List.Transform(tokenGroups, each FirstNumberByNames(row, FieldNamesContaining(row, _))))"
    AddQueryLine lines, "        in"
    AddQueryLine lines, "            if List.Count(Matches) = 0 then null else Matches{0},"
    AddQueryLine lines, "    Source = OData.Feed(ODataUrl, null, [Implementation=""2.0""]),"
    AddQueryLine lines, "    JobToolFields = RecordFields(Source, ""JobTool"", {""Item"", ""ItemSerial"", ""ShippingStatus"", ""ShippingStatusName"", ""Status"", ""TransferInDate"", ""TransferOutDate"", ""JobHours"", ""JobHours1"", ""JobHours2"", ""JobHours3"", ""JobHours4"", ""JobHours5"", ""JobHours6"", ""JobHours7"", ""JobHours8"", ""JobHours9"", ""JobHours10"", ""HSLS"", ""Hsls"", ""HslsHours"", ""TotalHsls"", ""TotalHSLS""}),"
    AddQueryLine lines, "    ExpandedJobTool = if List.Count(JobToolFields) = 0 then Source else Table.ExpandRecordColumn(Source, ""JobTool"", JobToolFields, List.Transform(JobToolFields, each ""JobTool."" & _)),"
    AddQueryLine lines, "    ItemFields = RecordFields(ExpandedJobTool, ""JobTool.Item"", {""ItemName"", ""Name"", ""Description""}),"
    AddQueryLine lines, "    ExpandedItem = if List.Count(ItemFields) = 0 then ExpandedJobTool else Table.ExpandRecordColumn(ExpandedJobTool, ""JobTool.Item"", ItemFields, List.Transform(ItemFields, each ""Item."" & _)),"
    AddQueryLine lines, "    SerialFields = RecordFields(ExpandedItem, ""JobTool.ItemSerial"", {""SerialNumber"", ""SerialNo"", ""Serial"", ""ItemSerialNumber"", ""Name""}),"
    AddQueryLine lines, "    ExpandedSerial = if List.Count(SerialFields) = 0 then ExpandedItem else Table.ExpandRecordColumn(ExpandedItem, ""JobTool.ItemSerial"", SerialFields, List.Transform(SerialFields, each ""ItemSerial."" & _)),"
    AddQueryLine lines, "    WithItem = Table.AddColumn(ExpandedSerial, ""Item"", each FirstText({Record.FieldOrDefault(_, ""Item.ItemName"", null), Record.FieldOrDefault(_, ""Item.Name"", null), Record.FieldOrDefault(_, ""Item.Description"", null), Record.FieldOrDefault(_, ""ItemName"", null), Record.FieldOrDefault(_, ""Name"", null)}), type nullable text),"
    AddQueryLine lines, "    WithSerial = Table.AddColumn(WithItem, ""Serial #"", each FirstText({Record.FieldOrDefault(_, ""ItemSerial.SerialNumber"", null), Record.FieldOrDefault(_, ""ItemSerial.SerialNo"", null), Record.FieldOrDefault(_, ""ItemSerial.Serial"", null), Record.FieldOrDefault(_, ""ItemSerial.ItemSerialNumber"", null), Record.FieldOrDefault(_, ""ItemSerial.Name"", null), Record.FieldOrDefault(_, ""SerialNumber"", null), Record.FieldOrDefault(_, ""SerialNo"", null), Record.FieldOrDefault(_, ""Serial"", null)}), type nullable text),"
    AddQueryLine lines, "    WithShippingStatus = Table.AddColumn(WithSerial, ""Shipping Status"", each let ExplicitStatus = FirstText({Record.FieldOrDefault(_, ""JobTool.ShippingStatus"", null), Record.FieldOrDefault(_, ""JobTool.ShippingStatusName"", null), Record.FieldOrDefault(_, ""JobTool.Status"", null), Record.FieldOrDefault(_, ""ShippingStatus"", null), Record.FieldOrDefault(_, ""ShippingStatusName"", null), Record.FieldOrDefault(_, ""Status"", null)}), TransferIn = FirstText({Record.FieldOrDefault(_, ""JobTool.TransferInDate"", null), Record.FieldOrDefault(_, ""TransferInDate"", null)}), TransferOut = FirstText({Record.FieldOrDefault(_, ""JobTool.TransferOutDate"", null), Record.FieldOrDefault(_, ""TransferOutDate"", null)}) in if ExplicitStatus <> null then ExplicitStatus else if TransferIn <> null and TransferOut = null then ""On Location"" else if TransferOut <> null then ""Transferred Out"" else null, type nullable text),"
    AddQueryLine lines, "    WithRawJobMinutes = Table.AddColumn(WithShippingStatus, ""Job Hours Raw Minutes"", each let Exact = FirstNumberByNames(_, {""JobHours"", ""JobHours1"", ""JobHours2"", ""JobHours3"", ""JobHours4"", ""JobHours5"", ""JobHours6"", ""JobHours7"", ""JobHours8"", ""JobHours9"", ""JobHours10"", ""JobTool.JobHours"", ""JobTool.JobHours1"", ""JobTool.JobHours2"", ""JobTool.JobHours3"", ""JobTool.JobHours4"", ""JobTool.JobHours5"", ""JobTool.JobHours6"", ""JobTool.JobHours7"", ""JobTool.JobHours8"", ""JobTool.JobHours9"", ""JobTool.JobHours10""}) in if Exact <> null then Exact else FirstNumberByTokenGroups(_, {{""job"", ""hour""}, {""total"", ""hour""}, {""tool"", ""hour""}}), type nullable number),"
    AddQueryLine lines, "    WithRawHslsMinutes = Table.AddColumn(WithRawJobMinutes, ""HSLS Raw Minutes"", each let Exact = FirstNumberByNames(_, {""HSLS"", ""Hsls"", ""HslsHours"", ""TotalHsls"", ""TotalHSLS"", ""JobTool.HSLS"", ""JobTool.Hsls"", ""JobTool.HslsHours"", ""JobTool.TotalHsls"", ""JobTool.TotalHSLS""}) in if Exact <> null then Exact else FirstNumberByTokenGroups(_, {{""hsls""}, {""hsl""}}), type nullable number),"
    AddQueryLine lines, "    WithJobHours = Table.AddColumn(WithRawHslsMinutes, ""Job Hours"", each FormatMinutes(Record.FieldOrDefault(_, ""Job Hours Raw Minutes"", null)), type nullable text),"
    AddQueryLine lines, "    WithHsls = Table.AddColumn(WithJobHours, ""HSLS Display"", each FormatMinutes(Record.FieldOrDefault(_, ""HSLS Raw Minutes"", null)), type nullable text),"
    AddQueryLine lines, "    Selected = Table.SelectColumns(WithHsls, {""Item"", ""Serial #"", ""Shipping Status"", ""Job Hours"", ""HSLS Display"", ""Job Hours Raw Minutes"", ""HSLS Raw Minutes""}),"
    AddQueryLine lines, "    Output = Table.RenameColumns(Selected, {{""HSLS Display"", ""HSLS""}}),"
    AddQueryLine lines, "    Sorted = Table.Sort(Output, {{""Item"", Order.Ascending}, {""Serial #"", Order.Ascending}})"
    AddQueryLine lines, "in"
    AddQueryLine lines, "    Sorted"

    BuildFieldCapInventoryQuery = JoinQueryLines(lines)
End Function

Private Sub AddQueryLine(ByVal lines As Collection, ByVal lineText As String)
    lines.Add lineText
End Sub

Private Function JoinQueryLines(ByVal lines As Collection) As String
    Dim values() As String
    Dim index As Long

    ReDim values(1 To lines.Count)
    For index = 1 To lines.Count
        values(index) = CStr(lines(index))
    Next index

    JoinQueryLines = Join(values, vbCrLf)
End Function
