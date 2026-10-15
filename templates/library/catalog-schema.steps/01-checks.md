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

Per `_template-schema.md` §4.2 — the structural baseline instantiated against this schema's ten
tables, plus one behavior specific to this template (check 7).

| # | Check |
|---|---|
| 1 | An ordinary insert succeeds on each of the ten tables, supplying every `Required` field. |
| 2 | Every `Required` field on every table refuses a missing value; no `AllowZeroLength` fields are declared in this schema, so that half of the check doesn't apply here. |
| 3 | `tblPublicationCreator`'s unique index on (`PublicationID`, `CreatorID`) and `tblPublicationGenre`'s on (`PublicationID`, `GenreID`) each refuse the duplicate pairing they name. |
| 4 | `tblPublication → tblPublicationCreator` and `tblPublication → tblPublicationGenre` cascade-delete as declared (deleting a publication removes its creator and genre links); every other relationship in `## Relationships` does **not** cascade — deleting a looked-up creator, publisher, genre, or shelf location while it's referenced is refused, never silently orphaning or cascading into `tblPublication`. |
| 5 | Not applicable — no `seeds` are declared for the five lookup tables; they ship empty. |
| 6 | The house audit columns (`CreatedDate`/`CreatedBy`/`ModifiedDate`/`ModifiedBy`/`AccessTS`) stamp correctly on every table, per Standards Layer below. |
| 7 | `PublicationSortTitle` is populated correctly on insert and refreshed on a title edit, through whichever mechanism was chosen for Business Rule 2 (data macro or front-end code) — and, if front-end code was chosen, confirm directly that a title changed by another route (direct table edit, import) does **not** update the sort title, since that gap is the documented cost of that choice. |
| 8 | A `tblPublicationCreator` or `tblPublicationGenre` insert citing a `PublicationID`, `CreatorID`, or `GenreID` that doesn't exist is refused. |
| 9 | An insert with `PublisherName`, `CreatorLastName`, or another `Text(n)` field longer than its declared width is refused, not silently truncated. `PublicationTitle` is Memo and has no such limit, by design. Test it on the route this build actually inserts by: a bound form or recordset refuses an over-long value, while a SQL `INSERT` stores the first n characters without raising an error, so where the build inserts by SQL the build itself must refuse the value, or the entry records that it does not. |
| 10 | `Description` is present on every field of every built table, matching this template's own Purpose & rules text. |
| 11 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming what already exists — never a bare "duplicate object" error. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
