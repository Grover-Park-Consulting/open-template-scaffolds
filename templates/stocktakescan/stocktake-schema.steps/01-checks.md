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

Per `_template-schema.md` §4.2 — the structural baseline instantiated against this schema's eight
new tables, plus the two fields it grafts onto the existing `Products` table (check 7).

| # | Check |
|---|---|
| 1 | An ordinary insert succeeds on each of the eight new tables, supplying every `Required` field. |
| 2 | Every `Required` field on every new table refuses a missing value; no `AllowZeroLength` fields are declared in this schema, so that half of the check doesn't apply here. |
| 3 | `StockTakeCount`'s unique index on (`StockTakeSessionID`, `ProductID`) refuses the duplicate it names — a second count line for the same product in the same session. |
| 4 | `StockTakeSession → StockTakeCount` and `StockTakeCount → StockTakeScan` cascade-delete as declared. `Products → StockTakeCount` does **not** — deleting a product with count history is refused, never silently dropping the history — and `Products → ProductVarianceAllowance` **does** cascade, removing the tolerance row when its product goes. Confirm both the cascading and the non-cascading case directly; they sit on opposite sides of the same host table. |
| 5 | The three `SystemSettings` seed rows (`DefaultAllowableShortageRate`, `DefaultAllowableOverageRate`, `DuplicateScanWindowSeconds`) are present with the values the developer chose, in the host's existing `[percent*1000]`/plain-integer conventions as documented. |
| 6 | The audit columns the active standards layer supplies (see Standards Layer below) stamp correctly on every new table: who created the row and when, on insert; who last changed it and when, on each update; and the created pair left frozen on a later update. |
| 7 | **Graft-specific, not part of the generic baseline:** `Products.SKUBarCode` and `Products.QuantityInPackage` exist (created fresh, or confirmed against fields the developer already had), `SKUBarCode` carries its index and the index is **unique** — inserting or updating a second product to the same barcode as an existing one is refused — and, where the host is split, every front end's linked-table definition of `Products` was refreshed and shows both new fields. A front end that wasn't relinked is the specific, silent failure this template's own Prerequisites section warns about: code referencing either field fails with "item not found in this collection." |
| 8 | A `StockTakeCount` insert citing a `StockTakeSessionID` or `ProductID` that doesn't exist is refused; likewise a `StockTakeScan` insert citing a `StockTakeCountID` that doesn't exist. |
| 9 | An insert with `ScanCode` or another `Text(n)` field longer than its declared width is refused, not silently truncated. Test it on the route this build actually inserts by: a bound form or recordset refuses an over-long value, while a SQL `INSERT` stores the first n characters without raising an error, so where the build inserts by SQL the build itself must refuse the value, or the entry records that it does not. |
| 10 | `Description` is present on every field of every new table, matching this template's own Purpose & rules text — and on `Products.SKUBarCode`/`QuantityInPackage`, where the two grafted fields carry their own descriptions. |
| 11 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming what already exists — never a bare "duplicate object" error, and never a second attempt to add `SKUBarCode`/`QuantityInPackage` to `Products` if they're already there. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
