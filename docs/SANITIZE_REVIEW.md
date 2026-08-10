# Sanitize review â€” fictional Setup data

**Public/demo workbook:** `Slide Sheet - Demo.xlsm` (fictional identity applied).

**Do not push until you open the workbook and confirm.**

Real (sensitive) workbook backup:
`backups/Slide Sheet - 34801.xlsm.bak_REAL_SENSITIVE_*`

Restore field copy (Excel closed):
```powershell
Copy-Item -LiteralPath (Get-ChildItem 'backups\Slide Sheet - 34801.xlsm.bak_REAL_SENSITIVE_*' | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName -Destination 'Slide Sheet - Demo.xlsm' -Force
```

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
| Local paths on Setup | scrubbed to `C:\Demo\Well\99001\` |

Also update EMAIL defaults in `Module11` (To/CC/Subject) to demo addresses before push.

## Still may contain operational data

Surveys, costs, motor S/Ns, BHA details, and OpenCap CSV folders under the repo were **not** fully scrubbed. Review those separately if the repo is public.
