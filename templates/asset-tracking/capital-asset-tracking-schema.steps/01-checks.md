---
step: 01-checks
title: "Validating the build"
platform_facts: []
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

## Validating the build

Per `_template-schema.md` §4.2 — the structural baseline instantiated against this schema's
fourteen tables.

| # | Check |
|---|---|
| 1 | An ordinary insert succeeds on each of the fourteen tables (the eight entities and the six lookups), supplying every `Required` field. |
| 2 | Every `Required` field on every table refuses a missing value; no `AllowZeroLength` fields are declared in this schema, so that half of the check doesn't apply here. |
| 3 | Each declared unique index refuses its duplicate: `tblAsset.AssetBarcode`; `tblSite.SiteCode`; `tblRoom` on (`SiteID`, `RoomNumber`); `tblDepartment.DepartmentName`. |
| 4 | Sample the seventeen relationships in `## Relationships` by their two declared behaviors, not one at a time: a **restrict** relationship (e.g. `tblRoom → tblAsset`) refuses deleting the parent while a child row exists; a **cascade** relationship (`tblAsset → tblAssetHistory`, `tblInventoryAuditSession → tblInventoryAuditScan`) deletes the children when the parent goes. Confirm at least one of each kind actually behaves as declared, not only that `validate` resolved the table names. |
| 5 | Every seed row in the six lookup tables (`## Entities → Lookup tables`) is present exactly as listed — all seven `tlkpAssetCategory` rows, all five `tlkpAssetStatus` rows, and so on. Confirm `tlkpAssetCategory.IsDepreciable` specifically: `Land` is `False`, every other seeded row is `True`. |
| 6 | The audit columns the active standards layer supplies (see Standards Layer below) stamp correctly on every `tbl`/`tlkp` table: who created the row and when, on insert; who last changed it and when, on each update; and the created pair left frozen on a later update. |
| 7 | An insert on the child side of a **restrict** relationship (e.g. a `tblAsset` row citing a `RoomID` that doesn't exist) is refused — the same sample used for check 4, tested from the other direction. |
| 8 | An insert with `AssetDescription`, `SiteName`, or another `Text(n)` field longer than its declared width is refused, not silently truncated — no field in this schema documents truncation as intended behavior. Test it on the route this build actually inserts by: a bound form or recordset refuses an over-long value, while a SQL `INSERT` stores the first n characters without raising an error, so where the build inserts by SQL the build itself must refuse the value, or the entry records that it does not. |
| 9 | `Description` is present on every field of every built table, matching this template's own Purpose & rules text. |
| 10 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming what already exists — never a bare "duplicate object" error. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
