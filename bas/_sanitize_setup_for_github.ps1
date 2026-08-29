# Sanitize Setup / Job / Crew PII in the live workbook for a public GitHub push.
# REAL data must already be backed up (bak_REAL_SENSITIVE_*). Does NOT push.
# Does NOT run InitSetup. Edits _OC/_FC job+crew sheets and Setup display cells.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$src = Join-Path $root "Slide Sheet - Demo.xlsm"
$review = Join-Path $root "docs\SANITIZE_REVIEW.md"

# ---- fictional values --------------------------------------------------------
$job = [ordered]@{
    "Job ID"                 = "99001"
    "Job Name"               = "NWE HZ WILLOWBEND 04-22-68-08 - demo-rig -"
    "Job Code"               = "99001"
    "Client"                 = "Northwind Energy Ltd."
    "Job Type"               = "Hz"
    "Description"            = "Demo well package for repository review (fictional)."
    "AFE"                    = "DEMO-AFE-01"
    "CCEmail"                = "ops@demo-dd.example;updates@demo-dd.example"
    "ClientEmail"            = "field@northwind.example"
    "CompanyMan"             = "Jordan Hale"
    "CompanyManEmail"        = "jordan.hale@northwind.example"
    "CompanyManPhone"        = "403-555-0142"
    "DDCoordinator"          = "Sam Rivera"
    "DDCoordinatorEmail"     = "sam.rivera@demo-dd.example"
    "Directions"             = "From Demo City, take Hwy 1 west 40 km to Twp Rd 680. South 8 km to RR 85. West 2 km to lease road; follow signs to location."
    "Geologist"              = "Avery Brooks"
    "GeologistEmail"         = "avery.brooks@demo-geo.example"
    "Latitude"               = "54.208333"
    "LatitudeDMS"            = "54 12 30.000 N"
    "Longitude"              = "-115.500000"
    "LongitudeDMS"           = "115 30 00.000 W"
    "MWDCoordinator"         = "Riley Chen"
    "MWDCoordinatorEmail"    = "riley.chen@demo-mwd.example"
    "opsStatus"              = "Field Active"
    "OtherPersonnel"         = '[{"String1":"Morgan Lee, Casey Quinn","String3":"","String4":"Second Company Man"},{"String1":"Avery Brooks","String3":"avery.brooks@demo-geo.example","String4":"Geologist"}]'
    "PreferredTrucking"      = "Prairie Haul Ltd."
    "PrimaryZone"            = "Demo Zone"
    "Province"               = "AB"
    "SalesRep"               = "Taylor Brooks"
    "SalesRepEmail"          = "taylor.brooks@demo-sales.example"
    "SecondCompanyMan"       = "Morgan Lee, Casey Quinn"
    "SurfaceCoordinates"     = "04-22-068-08 W5"
    "UWI"                    = "100/04-22-068-08W5/00"
    "WellLicenceNumber"      = "990001"
    "WellPlanner"            = "Alex Morgan"
    "WellPlannerEmail"       = "alex.morgan@demo-plan.example"
    "WellProfile"            = "Horizontal"
    "WellType"               = "Gas"
    "Rig Name"               = "Apex-214"
}

$crew = @(
    @{ Name = "Alex Rivera";   Email = "alex.rivera@demo-crew.example";   Phone = "780-555-0101"; Role = "DD";  Line = "1" },
    @{ Name = "Jamie Cole";    Email = "jamie.cole@demo-crew.example";    Phone = "403-555-0102"; Role = "MWD"; Line = "2" },
    @{ Name = "Morgan Blake";  Email = "morgan.blake@demo-crew.example";  Phone = "780-555-0103"; Role = "DD";  Line = "3" }
)

function Set-JobField($ws, [hashtable]$map) {
    $lastCol = $ws.Cells.Item(1, $ws.Columns.Count).End(-4159).Column
    $byHeader = @{}
    for ($c = 1; $c -le $lastCol; $c++) {
        $h = ([string]$ws.Cells.Item(1, $c).Text).Trim()
        if ($h -ne "") { $byHeader[$h] = $c }
    }
    foreach ($k in $map.Keys) {
        if (-not $byHeader.ContainsKey($k)) { continue }
        $c = $byHeader[$k]
        $cell = $ws.Cells.Item(2, $c)
        $cell.NumberFormat = "@"
        $cell.Value2 = [string]$map[$k]
    }
}

function Clear-Hyperlinks($rng) {
    try { if ($rng.Hyperlinks.Count -gt 0) { $rng.Hyperlinks.Delete() } } catch {}
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$xl.AutomationSecurity = 3
$wb = $null
try {
    $miss = [Type]::Missing
    $wb = $xl.Workbooks.Open($src, 0, $false, $miss, $miss, $miss, $true)
    if ($wb.ReadOnly) { throw "workbook opened read-only - close it in Excel and re-run" }

    foreach ($name in @("_OC_Job", "_FC_Job")) {
        $ws = $wb.Worksheets.Item($name)
        Set-JobField $ws $job
        "updated $name"
    }

    foreach ($name in @("_OC_Crew", "_FC_Crew")) {
        $ws = $wb.Worksheets.Item($name)
        # keep header row; rewrite data rows
        $lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
        if ($lastRow -ge 2) { $ws.Range("A2:J$lastRow").ClearContents() }
        for ($i = 0; $i -lt $crew.Count; $i++) {
            $r = $i + 2
            $ws.Cells.Item($r, 1).NumberFormat = "@"; $ws.Cells.Item($r, 1).Value2 = $job["Job ID"]
            $ws.Cells.Item($r, 2).NumberFormat = "@"; $ws.Cells.Item($r, 2).Value2 = $crew[$i].Name
            $ws.Cells.Item($r, 4).NumberFormat = "@"; $ws.Cells.Item($r, 4).Value2 = $crew[$i].Email
            $ws.Cells.Item($r, 5).NumberFormat = "@"; $ws.Cells.Item($r, 5).Value2 = $crew[$i].Phone
            $ws.Cells.Item($r, 6).NumberFormat = "@"; $ws.Cells.Item($r, 6).Value2 = $crew[$i].Role
            $ws.Cells.Item($r, 10).NumberFormat = "@"; $ws.Cells.Item($r, 10).Value2 = $crew[$i].Line
        }
        "updated $name ($($crew.Count) people)"
    }

    $setup = $wb.Worksheets.Item("Setup")
    # Unprotect briefly if needed
    $wasProtected = $false
    try {
        if ($setup.ProtectContents) {
            $setup.Unprotect("")
            $wasProtected = $true
        }
    } catch {
        try { $setup.Unprotect(); $wasProtected = $true } catch {}
    }

    # JOB IDENTITY + land (rows from dump)
    $setup.Range("D4").NumberFormat = "@"; $setup.Range("D4").Value2 = $job["Job ID"]
    $setup.Range("D5").NumberFormat = "@"; $setup.Range("D5").Value2 = $job["Job Name"]
    $setup.Range("D6").NumberFormat = "@"; $setup.Range("D6").Value2 = $job["Client"]
    $setup.Range("D7").NumberFormat = "@"; $setup.Range("D7").Value2 = $job["UWI"]
    $setup.Range("G7").NumberFormat = "@"; $setup.Range("G7").Value2 = $job["WellLicenceNumber"]
    $setup.Range("D8").NumberFormat = "@"; $setup.Range("D8").Value2 = $job["SurfaceCoordinates"]
    $setup.Range("G6").NumberFormat = "@"; $setup.Range("G6").Value2 = $job["AFE"]
    $setup.Range("D15").NumberFormat = "@"; $setup.Range("D15").Value2 = $job["LatitudeDMS"]
    $setup.Range("G15").NumberFormat = "@"; $setup.Range("G15").Value2 = $job["LongitudeDMS"]
    $setup.Range("D29").NumberFormat = "@"; $setup.Range("D29").Value2 = $job["Rig Name"]

    # CONTACTS
    $setup.Range("D32").NumberFormat = "@"; $setup.Range("D32").Value2 = $job["CompanyMan"]
    $setup.Range("G32").NumberFormat = "@"; $setup.Range("G32").Value2 = $job["CompanyManPhone"]
    $setup.Range("D33").NumberFormat = "@"; $setup.Range("D33").Value2 = $job["SecondCompanyMan"]
    $setup.Range("D34").NumberFormat = "@"; $setup.Range("D34").Value2 = $job["Geologist"]
    $setup.Range("D35").NumberFormat = "@"; $setup.Range("D35").Value2 = $job["DDCoordinator"]
    $setup.Range("G35").NumberFormat = "@"; $setup.Range("G35").Value2 = $job["DDCoordinatorEmail"]
    Clear-Hyperlinks $setup.Range("G35")
    try { $setup.Hyperlinks.Add($setup.Range("G35"), "mailto:" + $job["DDCoordinatorEmail"]) | Out-Null } catch {}
    $setup.Range("D36").NumberFormat = "@"; $setup.Range("D36").Value2 = $job["MWDCoordinator"]
    $setup.Range("G36").NumberFormat = "@"; $setup.Range("G36").Value2 = $job["MWDCoordinatorEmail"]
    Clear-Hyperlinks $setup.Range("G36")
    try { $setup.Hyperlinks.Add($setup.Range("G36"), "mailto:" + $job["MWDCoordinatorEmail"]) | Out-Null } catch {}
    $setup.Range("D37").NumberFormat = "@"; $setup.Range("D37").Value2 = $job["SalesRep"]
    $setup.Range("G37").NumberFormat = "@"; $setup.Range("G37").Value2 = $job["SalesRepEmail"]
    Clear-Hyperlinks $setup.Range("G37")
    try { $setup.Hyperlinks.Add($setup.Range("G37"), "mailto:" + $job["SalesRepEmail"]) | Out-Null } catch {}

    # CREW MANIFEST (J/K/M/P from dump: role empty, name col11, email 13, phone 16)
    for ($i = 0; $i -lt $crew.Count; $i++) {
        $r = 5 + $i
        $setup.Range("J$r").NumberFormat = "@"; $setup.Range("J$r").Value2 = $crew[$i].Role
        $setup.Range("K$r").NumberFormat = "@"; $setup.Range("K$r").Value2 = $crew[$i].Name
        $setup.Range("M$r").NumberFormat = "@"; $setup.Range("M$r").Value2 = $crew[$i].Email
        Clear-Hyperlinks $setup.Range("M$r")
        try { $setup.Hyperlinks.Add($setup.Range("M$r"), "mailto:" + $crew[$i].Email) | Out-Null } catch {}
        $setup.Range("P$r").NumberFormat = "@"; $setup.Range("P$r").Value2 = $crew[$i].Phone
    }

    # Scrub local file paths / real well names on Setup import list
    $demoFiles = @{
        "M20" = "demo-job-99001-job-details.csv"
        "M21" = "demo-job-99001-crew.csv"
        "M22" = "demo-job-99001-bha-equipment.csv"
        "M23" = "demo-job-99001-slide-rotate.csv"
        "M24" = "demo-job-99001-inventory.csv"
        "M25" = "demo-job-99001-ticket-costs.csv"
        "M29" = "demo-job-99001-well-plan-p2.csv"
        "M30" = "demo-job-99001-well-plan-p1-ac.pdf"
    }
    foreach ($addr in $demoFiles.Keys) {
        try {
            $setup.Range($addr).NumberFormat = "@"
            $setup.Range($addr).Value2 = $demoFiles[$addr]
        } catch {}
    }
    foreach ($addr in @("Q20", "Q21", "Q22", "Q23", "Q24", "Q25")) {
        try {
            $setup.Range($addr).NumberFormat = "@"
            $setup.Range($addr).Value2 = "C:\Demo\Well\99001\"
        } catch {}
    }

    if ($wasProtected) {
        try { $setup.Protect("") } catch {}
    }
    "updated Setup display"

    # Data sheet labels that expose real UWIs / paths
    $data = $wb.Worksheets.Item("Data")
    try {
        if ($data.ProtectContents) { try { $data.Unprotect("") } catch { $data.Unprotect() } }
    } catch {}
    foreach ($addr in @("B35", "B36", "A35", "A36")) {
        $v = [string]$data.Range($addr).Text
        if ($v -match "071-12|ELMWORTH|PARAMOUNT|Akita|sinclair|13-14-71") {
            $data.Range($addr).NumberFormat = "@"
            $data.Range($addr).Value2 = "100/04-22-068-08W5/00 - Demo Offset Surveys"
        }
    }
    # Attachment path cells I29:I33 often hold local paths
    foreach ($r in 29..33) {
        $cell = $data.Range("I$r")
        $v = [string]$cell.Text
        if ($v -match "Phoenix|ELMWORTH|34801|35780|Users\\\\|Desktop") {
            $cell.NumberFormat = "@"
            $cell.Value2 = ""
        }
    }
    "updated Data labels/paths"

    $wb.Save()
    "saved $src"
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

$md = @"
# Sanitize review â€” fictional Setup data

**Do not push until you open the workbook and confirm.**

Real (sensitive) workbook backup:
``backups/Slide Sheet - Demo.xlsm.bak_REAL_SENSITIVE_*``

Restore field copy (Excel closed):
``````powershell
Copy-Item -LiteralPath (Get-ChildItem 'backups\Slide Sheet - Demo.xlsm.bak_REAL_SENSITIVE_*' | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName -Destination 'Slide Sheet - Demo.xlsm' -Force
``````

## Fictional values now in the workbook

| Field | Value |
|------|-------|
| Client | Northwind Energy Ltd. |
| Job ID / Code | 99001 |
| Job Name | NWE HZ WILLOWBEND 04-22-68-08 - demo-rig - |
| Well / UWI | 100/04-22-068-08W5/00 |
| Surface / land | 04-22-068-08 W5 |
| Lat / Long | 54 12 30.000 N / 115 30 00.000 W |
| Rig Name | Apex-214 |
| Well licence / AFE | 990001 / DEMO-AFE-01 |
| Company Man | Jordan Hale / 403-555-0142 / jordan.hale@northwind.example |
| 2nd Company Man | Morgan Lee, Casey Quinn |
| Geologist | Avery Brooks / avery.brooks@demo-geo.example |
| DD Coordinator | Sam Rivera / sam.rivera@demo-dd.example |
| MWD Coordinator | Riley Chen / riley.chen@demo-mwd.example |
| Sales Rep | Taylor Brooks / taylor.brooks@demo-sales.example |
| Well Planner | Alex Morgan / alex.morgan@demo-plan.example |
| Crew | Alex Rivera, Jamie Cole, Morgan Blake (demo emails / 555 phones) |
| Local paths on Setup | scrubbed to ``C:\Demo\Well\99001\`` |

Also update EMAIL defaults in ``Module11`` (To/CC/Subject) to demo addresses before push.

## Still may contain operational data

Surveys, costs, motor S/Ns, BHA details, and OpenCap CSV folders under the repo were **not** fully scrubbed. Review those separately if the repo is public.
"@
Set-Content -LiteralPath $review -Value $md -Encoding UTF8
"review notes: $review"
