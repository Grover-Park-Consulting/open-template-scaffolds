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

Per `_template-schema.md` §4.2 — the structural baseline instantiated against this schema's six
tables. Run every row that applies to what this build actually created: the three system tables
always, the three sample tables only on Path A. On Path B, substitute the developer's own audited
tables wherever a sample-table row is marked Path A only.

| # | Check |
|---|---|
| 1 | An ordinary insert succeeds on each of `tblAuditLog`, `tblLongTextBackup`, `tblAuditLogConfig`, and — Path A only — `tblClient`, `tblSupportTicket`, `tlkpTicketPriority`. |
| 2 | Every `Required` field on each table refuses a missing value; no `AllowZeroLength` fields are declared in this schema, so that half of the check doesn't apply here. |
| 3 | Each declared unique index refuses its duplicate: `tblClient.ClientName`; `tlkpTicketPriority.TicketPriorityName`; `tblAuditLogConfig` on (`TableName`, `FieldName`); `tblLongTextBackup` on (`TableName`, `PrimaryKey`, `FieldName`). |
| 4 | `tblClient → tblSupportTicket` and `tlkpTicketPriority → tblSupportTicket` are enforced with no cascade delete, exactly as declared — a delete attempt on a referenced `tblClient`/`tlkpTicketPriority` row with dependent tickets is refused. **And the negative case**: the three system tables have no enforced relationship to each other, to the sample tables, or to whatever real tables Path B audits — deleting an audited row must never be blocked by, or cascade into, `tblAuditLog`. |
| 5 | `tlkpTicketPriority`'s four seed rows are present with the exact `SortOrder` values specified (Low 10, Normal 20, High 30, Urgent 40). |
| 6 | On the sample tables (Path A) or the developer's own audited tables (Path B): `CreatedDate`/`CreatedBy` stamp on insert, `ModifiedDate`/`ModifiedBy` stamp on update, `Created*` stays frozen on a later update. **And the negative case**: the three system tables carry none of the house audit columns — confirm they're genuinely absent, not merely unpopulated. |
| 7 | A `tblSupportTicket` insert with a `ClientID` or `TicketPriorityID` that doesn't exist in `tblClient`/`tlkpTicketPriority` is refused. |
| 8 | An insert with `ClientName`, `TicketSubject`, or any other `Text(n)` field longer than its declared width is refused, not silently truncated — no field in this schema documents truncation as intended behavior. Test it on the route this build actually inserts by: a bound form or recordset refuses an over-long value, while a SQL `INSERT` stores the first n characters without raising an error, so where the build inserts by SQL the build itself must refuse the value, or the entry records that it does not. |
| 9 | `Description` is present on every field of every built table, matching this template's own Purpose & rules text. |
| 10 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming what already exists — never a bare "duplicate object" error. On Path A specifically, also confirm it doesn't create a second copy of the three sample tables. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`. Note beside each
entry which build path (A or B) it was run against, since several entries read differently depending
on which tables actually exist.
