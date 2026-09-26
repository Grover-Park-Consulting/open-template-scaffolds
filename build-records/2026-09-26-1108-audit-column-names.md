# Build Record: Stop Hardcoding a Source Database's Audit-Column Names

**Who reads this:** George (and any future AI assistant working in this library). The fifth
implementation task, per `build-records/2026-09-26-task-brief-audit-column-names.md`.

**Author:** the AI assistant (deepseek-v4-pro), via OpenCode.
**Date/time:** 2026-09-26 11:08.

**Frozen files (`FREEZE.md`, the 18): no.** I touched none of them. `standards/audit-columns.md` and
`templates/_template-schema.md` were both read, not written.

Two templates no longer assert one source database's audit-column names as if universal. Four spots,
before/after below. No file content beyond these four spots changed.

---

## 1. `templates/stocktakescan/stocktake-schema.md` — checklist item 6

**Before**
> | 6 | The house audit columns (`AddedBy`/`AddedOn`/`ModifiedBy`/`ModifiedOn`, per the Northwind data-macro pattern) stamp correctly on every new table, per Standards Layer below. |

**After**
> | 6 | The audit columns the active standards layer supplies (see Standards Layer below) stamp correctly on every new table: who created the row and when, on insert; who last changed it and when, on each update; and the created pair left frozen on a later update. |

No column names at all now — the check validates whatever the active standards layer supplies.

## 2. `templates/stocktakescan/stocktake-schema.md` — Standards Layer, "Audit columns"

**Before**
> - **Audit columns** — `AddedBy`, `AddedOn`, `ModifiedBy`, `ModifiedOn` on every new table,
>   maintained by the Northwind data-macro audit pattern (NorthwindFeatures #30).

**After**
> - **Audit columns** — the house default set (`CreatedDate`, `CreatedBy`, `ModifiedDate`,
>   `ModifiedBy`) on every new table, maintained by the active standards layer's mechanism
>   (`standards/audit-columns.md`). A host this template is grafted onto may already stamp
>   who-and-when its own way — a Northwind-derived database, for one, commonly uses
>   `AddedBy`/`AddedOn`/`ModifiedBy`/`ModifiedOn` via data macros. Check the target database's
>   existing convention before naming any new audit column and match it; use the house names only
>   where no existing convention is present.

House default stated; the Northwind names demoted to an example; checking the target database
instructed.

## 3. `templates/asset-tracking/capital-asset-tracking-schema.md` — checklist item 6

**Before**
> | 6 | The house audit columns (`AddedBy`/`AddedOn`/`ModifiedBy`/`ModifiedOn`, from the host's audit convention) stamp correctly on every `tbl`/`tlkp` table, per Standards Layer below. |

**After**
> | 6 | The audit columns the active standards layer supplies (see Standards Layer below) stamp correctly on every `tbl`/`tlkp` table: who created the row and when, on insert; who last changed it and when, on each update; and the created pair left frozen on a later update. |

## 4. `templates/asset-tracking/capital-asset-tracking-schema.md` — Standards Layer, "Audit columns"

**Before**
> - **Audit columns** — `AddedBy` / `AddedOn` / `ModifiedBy` / `ModifiedOn` on every `tbl`/`tlkp`
>   table, supplied by the host's audit convention. Not present in the field tables above; see
>   Business Rule 3 for the separate, template-defined movement/change history, which is not a
>   substitute for these.

**After**
> - **Audit columns** — the house default set (`CreatedDate` / `CreatedBy` / `ModifiedDate` /
>   `ModifiedBy`) on every `tbl`/`tlkp` table, maintained by the active standards layer's mechanism
>   (`standards/audit-columns.md`). The contributing host's own database stamped who-and-when its
>   own way (`AddedBy`/`AddedOn`/`ModifiedBy`/`ModifiedOn`, for one), and a target database may do
>   the same. Check the target database's existing convention before naming any new audit column and
>   match it, falling back to the house names only where none is present. Not present in the field
>   tables above; see Business Rule 3 for the separate, template-defined movement/change history,
>   which is not a substitute for these.

The closing "Not present in the field tables above; see Business Rule 3…" sentence was kept, since
it distinguishes the standard audit columns from the template's own movement/change history.

## What I left alone, and why

- `templates/audit/audit-logging-lite-scaffold.md` — the reference, not touched.
- `examples/northwind-stocktake/` — a Northwind worked example, where Northwind names are correct.
  Frozen; not touched.
- **Naming-conventions bullets** in both files — checked for the identical audit-column-style
  defect and did not find it. Each states which house style the template follows and acknowledges the
  standard is replaceable, which the brief allows. Left alone.

## Verification

- `python mcp-server/run_validate.py` → **21/21 well-formed**.
- `python -m unittest discover -s mcp-server/tests` → **31 OK, 1 skipped**.
- Re-read both edited sections: neither still names a fixed audit-column set as the
  expected/validated one. Source-database names appear only as an example ("for one"), never as the
  standard.
- `git status` → only the two intended files modified.
