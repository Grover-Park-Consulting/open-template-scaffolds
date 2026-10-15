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

Per `_template-schema.md` §4.2 — the structural baseline instantiated against `tblErrorLog` alone.
This table declares no relationships, no seed rows, and no audit columns by design, so those
baseline checks are marked not applicable rather than skipped silently.

| # | Check |
|---|---|
| 1 | An ordinary insert into `tblErrorLog`, supplying every Required column, succeeds. |
| 2 | An insert with `ErrorDescription = ""` (a zero-length string, not Null) succeeds — confirms `AllowZeroLength = True` actually took on the built table, not only that the field table says so. |
| 3 | An insert omitting any one Required column in turn is refused. |
| 4 | Not applicable — no unique index beyond the PK is declared. |
| 5 | Not applicable — `## Relationships` declares none, by design (Business Rule 2). |
| 6 | Not applicable — no `seeds` are declared. |
| 7 | Not applicable — audit columns are deliberately not applied to this table (see Standards Layer). |
| 8 | Not applicable — no FK column is declared on `tblErrorLog` (see Relationships: "deliberately unrelated"), so there is no insert-side reference to test. |
| 9 | A value longer than `ModuleName`, `ProcedureName` or `ErrorUser` (each Text(100)) is refused on every route the logger writes by, not cut short; `ErrorDescription` is the one field the logger shortens to fit, by Business Rule 4. Confirm both halves of that distinction on the logger's own write routes. A SQL `INSERT` the logger does not make is not tested (`_template-schema.md` §4.2 item 8). |
| 10 | `Description` is present on every field of the built table, matching this template's own Purpose & rules text — confirms the second-pass `Description` set (`_materialization.md` rule 2) actually ran. |
| 11 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming `tblErrorLog` as already existing — never a bare "duplicate object" error. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
