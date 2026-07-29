---
template: audit-logging-lite-schema
title: Access Audit Logging (Lite) — Table Schema
domain: audit
type: table-schema
version: 0.1.0
status: draft
standards_layer: [audit-columns, naming-conventions, error-handling]
new_tables:
  - tblAuditLog
  - tblLongTextBackup
  - tblAuditLogConfig
  - tblClient
  - tblSupportTicket
  - tlkpTicketPriority
seeds:
  - tlkpTicketPriority.Low
  - tlkpTicketPriority.Normal
  - tlkpTicketPriority.High
  - tlkpTicketPriority.Urgent
build_paths:
  - "Path A — Demo build: try the system out on made-up tables the generator creates for you.
    Nothing real is touched."
  - "Path B — Add it to a database you already have: point the generator at your own existing
    tables instead of the made-up ones. Back up the file first (see warnings)."
warnings:
  - Data Macros cannot audit Long Text (Memo) fields on their own. Before building, list every
    Long Text field in the tables to be audited and confirm the list with the developer — any
    table carrying one needs the hybrid VBA path (tblLongTextBackup + the BackupLongTextFieldsDM
    helper + five data macros); a table without one needs only the three After macros.
  - The macro generator must run in the same accdb as the audited tables — in a split design,
    that is the back end. The three system tables are created in the back end and linked to the
    front end.
  - Every audited table is expected to have a single-column AutoNumber primary key. If any table
    to be audited has a different key design (composite, text, no PK), stop and tell the
    developer this template will not work for that table out of the box — they are free to adapt
    it, but the adaptation is theirs. The paired scaffold's CheckAuditReadiness procedure checks
    for this automatically.
  - Path B (adding this to a database you already use) is much less forgiving than the demo.
    Make a copy of the .accdb file before running any of the setup steps against it — Data
    Macros get attached directly to your live tables.
  - If a table already carries its own Data Macros before this system is added to it (for example,
    the standards/audit-columns.md Before Change stamping macro), they are silently replaced —
    loading a table's macros from text replaces the whole set, it does not merge. The paired
    scaffold backs up a table's existing macros automatically before replacing them, but does not
    restore that stamping logic afterward — that is the developer's call.
house_assumptions:
  - tblAuditLogConfig.IsPrimaryKey — every audited table is assumed to have a single-column
    numeric (AutoNumber/Long) primary key; the Long Text backup plumbing and the generated macro
    XML key on one numeric PK, so composite or text keys are not supported
  - tblAuditLog — audited rows are referenced by name and key value (TableName + PrimaryKey),
    deliberately without enforced relationships, so audit history survives deletion of the rows
    it describes
  - tblAuditLogConfig — the schema scan selects candidate tables by the tbl/tlkp prefix naming
    convention, excluding tmp (one code filter in the paired scaffold, kept in sync between
    Two_PopulateConfigTable and CheckAuditReadiness); everything finer-grained is decided in
    data via IsAuditable. Adopters on other naming conventions adjust that one filter
  - tblClient — the sample tables (tblClient, tblSupportTicket, tlkpTicketPriority) are
    demonstration stand-ins showing both macro paths; a real build applies the system to the
    adopter's own tables
---

# Access Audit Logging (Lite) — Table Schema

## Intent

A **lite, self-contained change-audit system for Access** built on table-level **Data Macros** —
the Access engine's equivalent of a database trigger. Every insert, update, and delete on an
audited table writes rows to a central **audit log**, one row per affected field, recording the
old value, the new value, who made the change, and when. Because the macros live on the tables
themselves, the trail is complete no matter how the data is touched: forms, queries, VBA, imports,
or direct table edits.

Three system tables do the work: **`tblAuditLog`** (the trail), **`tblAuditLogConfig`** (the
control panel — populated by scanning the schema, with auditing scope then decided in its
`IsAuditable` flags, as data, not code), and **`tblLongTextBackup`** (staging that solves the
platform's hard limit — Data Macros cannot read a Long Text field's old value, so a small VBA
helper backs it up *before* the change and the after-macro retrieves it; see Business Rules 2–3).

The design is drawn from a system running in production: the same three tables, macro set, and
Long Text hybrid path stand behind a live multi-user Access application whose audit trail —
including captured old values of Long Text fields — validates the mechanism end to end.

"Lite" is deliberate: this template audits changes and preserves old values. It does **not**
attempt restore tooling, retention automation, or a review UI — those are named in
`## Parked / future considerations`.

### Two ways to build this

**Path A — try it out first.** The paired scaffold can create two made-up tables (`tblClient`,
`tblSupportTicket`) plus a short pick-list (`tlkpTicketPriority`) so you can watch the audit
trail work before you touch anything real: one table with no Long Text field (three macros) and
one with a Long Text field (five macros), so both kinds of tracking show up right away.

**Path B — add this to a database you already use.** The same three system tables and the same
scan-and-generate steps apply directly to your own existing tables — the made-up tables above
are never created; they exist only to make Path A a complete, working demo on their own. Because
this changes real tables, back up the .accdb file first (see Warnings).

Either way, the three system tables and the generator steps are identical — only whether the two
made-up tables get created differs.

## Entities

### tblAuditLog

Grain: one row per **field affected** by one insert, update, or delete on an audited table.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `AuditLogID` | AutoNumber | PK | Surrogate key |
| `TableName` | Text(50) | Required | Audited table the change happened in |
| `PrimaryKey` | Long | Required | Key value of the changed row in that table (Business Rule 4) |
| `FieldName` | Text(50) | Required | The field this row records |
| `OperationType` | Text(25) | Required | `Insert`, `Update`, or `Delete` — stamped by the macro |
| `OldValue` | Memo | Nullable | Value before the change; Null on insert |
| `NewValue` | Memo | Nullable | Value after the change; Null on delete |
| `DateChanged` | Date/Time | Required | Stamped `Now()` by the macro |
| `ChangedBy` | Text(50) | Required | Stamped by the macro (Business Rule 9) |

Indexes: PK on `AuditLogID`; index on (`TableName`, `DateChanged`) for trail queries; index on
(`TableName`, `PrimaryKey`) for per-record history.

### tblLongTextBackup

Grain: one row per Long Text field value backed up immediately **before** an update or delete —
transient staging, not history (Business Rule 8).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `BackupID` | AutoNumber | PK | Surrogate key |
| `TableName` | Text(50) | Required | Source table |
| `PrimaryKey` | Long | Required | Key value of the row being changed |
| `FieldName` | Text(50) | Required | The Long Text field backed up |
| `OldValue` | Memo | Nullable | The pre-change Long Text content |
| `DateChanged` | Date/Time | Required | When the backup was taken |
| `ChangedBy` | Text(50) | Required | Who triggered it |

Indexes: PK on `BackupID`; unique on (`TableName`, `PrimaryKey`, `FieldName`) — the helper
replaces any earlier backup for the same field of the same row before writing a new one.

### tblAuditLogConfig

Grain: one row per **scanned field** — the system's control panel. The scan writes every field
of every candidate table; what is and isn't audited is then decided **in this data** by flipping
`IsAuditable`, not by editing code (Business Rule 5).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ConfigID` | AutoNumber | PK | Surrogate key |
| `TableName` | Text(50) | Required | Table the field belongs to |
| `FieldName` | Text(50) | Required | The scanned field |
| `FieldPosition` | Long | Required | The field's ordinal position in its table |
| `DataType` | Long | Required | DAO type code; `dbMemo` (12) routes the field down the Long Text path |
| `IsPrimaryKey` | Boolean | Required | Default False; exactly one True row per table (Business Rule 4) |
| `IsAuditable` | Boolean | Required | Default True; flip to False to exclude a field from auditing (Business Rule 5) |

Indexes: PK on `ConfigID`; unique on (`TableName`, `FieldName`).

### tblClient *(sample)*

Grain: one row per client. **No Long Text field — this table gets the standard three After
macros.**

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ClientID` | AutoNumber | PK | Surrogate key |
| `ClientName` | Text(100) | Required | |
| `EmailAddress` | Text(100) | Nullable | |
| `CellPhone` | Text(25) | Nullable | |

Indexes: PK on `ClientID`; unique on `ClientName`.

### tblSupportTicket *(sample)*

Grain: one row per support ticket. **`TicketDetail` is Long Text — this table gets all five
macros** (Business Rules 2–3).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `SupportTicketID` | AutoNumber | PK | Surrogate key |
| `ClientID` | Long | FK → tblClient, Required | |
| `TicketPriorityID` | Long | FK → tlkpTicketPriority, Required | |
| `TicketSubject` | Text(255) | Required | |
| `TicketDetail` | Memo | Nullable | The Long Text demonstration field — audited via the hybrid path |
| `TicketOpenedDate` | Date/Time | Required | |
| `TicketClosedDate` | Date/Time | Nullable | Open tickets have no close date |

Indexes: PK on `SupportTicketID`; FK indexes on `ClientID`, `TicketPriorityID`.

### tlkpTicketPriority *(sample lookup)*

Grain: one row per priority level.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `TicketPriorityID` | AutoNumber | PK | Surrogate key |
| `TicketPriorityName` | Text(30) | Required | |
| `SortOrder` | Long | Required | Display order |

Indexes: PK on `TicketPriorityID`; unique on `TicketPriorityName`.

Seed rows: Low (10), Normal (20), High (30), Urgent (40).

**Deliberate teaching point:** the paired scaffold's schema scan takes `tbl…` **and** `tlkp…`
tables (never `tmp…`), so this lookup reaches the config table right alongside the business
tables — a real Path B build against a live schema turned up a table shaped like this one and
confirmed lookups belong in scope by default, not outside it. The boundary is still a visible,
editable one-line test, not an accident; it's just wider than an earlier draft of this template
drew it (everything inside the boundary is then decided by `IsAuditable` flags, as data).

## Relationships

- `tblClient (1) → (∞) tblSupportTicket` on `ClientID` — enforced, no cascade delete (tickets
  outlive nothing; delete clients only after their tickets are resolved or reassigned).
- `tlkpTicketPriority (1) → (∞) tblSupportTicket` on `TicketPriorityID` — enforced, no cascade.
- **The three system tables are deliberately unrelated** — to each other and to the audited
  tables. `tblAuditLog` and `tblLongTextBackup` reference audited rows by `TableName` +
  `PrimaryKey` value so the trail survives deletion of the rows it describes, and so one log
  serves every audited table without a web of enforced keys.

## Business Rules

1. **Same-accdb rule.** Data Macros attach to tables in the accdb they live in, so the generator
   and the three system tables belong in the **back end** of a split design. All three system
   tables are linked to the front end (the audit log for viewing; the backup table because
   front-end-triggered macros must reach it).
2. **Three-or-five branch.** A table with **no** Long Text fields gets three macros —
   AfterInsert, AfterUpdate, AfterDelete. A table **with** at least one Long Text field also
   gets BeforeChange and BeforeDelete, which back the Long Text values up before the change.
3. **Long Text flow.** BeforeChange/BeforeDelete call the `BackupLongTextFieldsDM` VBA helper →
   old value lands in `tblLongTextBackup` → the After macro retrieves it with `LookupRecord`
   and writes it to `tblAuditLog.OldValue`. This is the workaround for the platform limit: a
   Data Macro cannot read `[Old].[LongTextField]`.
4. **Single AutoNumber PK, always.** Every audited table is expected to have a single-column
   AutoNumber primary key, recorded in `tblAuditLogConfig.IsPrimaryKey`. A table with any other
   key design (composite, text, no PK) is called out at build time: the template will not work
   for it out of the box, and adapting it is the adopter's own project. The paired scaffold's
   `CheckAuditReadiness` procedure checks every candidate table against this rule and lists any
   that fail it, before macros are generated.
5. **Audit scope is data, not code.** The config scan writes every field of every candidate
   table; excluding a field or a whole table means flipping its `IsAuditable` flag, not editing
   code. What the flag starts as depends on which of the two build paths you're on: **Path A**
   (try-it-out build) starts every field switched ON, so you switch OFF what you don't want;
   **Path B** (a database you already use) starts every field switched OFF, so you switch ON —
   table by table — only what you actually want tracked, which is the safer default on tables
   this system wasn't designed around. Three hard exceptions apply either way: the three system
   tables are **never** given macros (auditing the audit trail would loop — their config rows,
   if present, stay OFF); noisy always-changing fields (row-version/timestamp columns, house
   audit columns) are always seeded OFF by the scan; and the audited table's own primary-key
   field is seeded OFF as well. Its value is not lost — every log row carries it in
   `tblAuditLog.PrimaryKey`, which is what identifies the record and what gap analysis reads;
   a field row for the key would only store the same value a second time, once per insert and
   once per delete.
6. **Every log row names its operation; only real changes are logged on update.** The macro
   stamps `OperationType` (`Insert` / `Update` / `Delete`); an insert row leaves `OldValue`
   Null and a delete row leaves `NewValue` Null. On update, the macro compares old and new
   values (`StrComp` on `Nz`-wrapped values) and logs only fields that actually changed. Long
   Text fields are always logged on update — the comparison cannot be done in the macro.
   **Practical effect, so this doesn't read as a bug:** every update to a row that contains a
   Long Text field writes an audit row for that field, even when the field did not change — the
   macro cannot compare Long Text values, so it logs the value as it stands. Change one ordinary
   field and expect two rows: the field you changed, and the Long Text field with the same
   content in `OldValue` and `NewValue` (both empty if the field is empty). Those rows are
   correct, not an error.
7. **Regenerate after schema change — and regeneration replaces, it never merges.** Adding a
   table or field, or changing a field's type to or from Long Text, requires re-running the
   config scan and regenerating the macros. The macros are point-in-time artifacts of the
   schema. This cuts both ways: generating for a table **replaces its entire Data Macro set**,
   including any macros the table already had for reasons unrelated to this system (most
   notably the house audit-column stamping macro in `standards/audit-columns.md`). The paired
   scaffold detects an existing macro set before overwriting it and backs it up automatically,
   but it does not merge the old logic into the new macros — re-adding any lost stamping logic
   is the developer's call, and worth checking for specifically on Path B. The risk also runs
   the other way, which is what a real Path B target actually looked like: every table carried
   the house audit columns but **no** table had a stamping macro yet. Adding the stamping macro
   *after* the audit macros are in place replaces them, exactly as generating audit macros
   replaces a stamping macro — `LoadFromText` swaps a table's whole macro set in either
   direction, and there is no partial or merge path. Where a table needs both, generate them
   together. If a table doesn't need the house audit columns at all, leaving them off removes
   this collision entirely: the table then needs one macro set, written once, with nothing for a
   later regeneration to overwrite.
8. **The backup table is staging, not history.** `tblLongTextBackup` may be cleared at any time;
   the durable record is `tblAuditLog`, which is append-only. Retention/archival policy for the
   log is the adopter's call.
9. **`ChangedBy` identity.** The macros stamp `CurrentUser()` — dependency-free, fires from any
   client, but returns `Admin` unless workgroup security is in use. A real-username upgrade
   (a public VBA helper returning `Environ("USERNAME")`, present in both back end and front
   end) is a named Extra Option; it trades a VBA dependency for a real name. The production
   system this template is drawn from runs the upgrade — its trail shows real Windows
   usernames — so the option is proven, not speculative.
10. **Samples are stand-ins, Path A only.** `tblClient`, `tblSupportTicket`, and
    `tlkpTicketPriority` exist only to prove both macro paths in a try-it-out build — the
    paired scaffold's `Zero_CreateSampleTables` procedure creates them, and that procedure is
    skipped entirely on Path B. A build against a database you already use (Path B) never
    creates these three tables; it applies the system directly to your own.

## Standards Layer

- **Audit columns** — a deliberate boundary: the three **system tables do not receive** the house
  audit columns (`DateChanged`/`ChangedBy` here are functional log fields, not the standards
  convention — an append-only log needs no self-stamp). The **sample/business tables** follow
  the house audit-columns convention as usual; note that self-stamp and timestamp fields
  (e.g. a row-version column) are seeded `IsAuditable = False` by the config scan so they don't
  drown the log (Business Rule 5). The paired scaffold's exclusion list names this repo's actual
  house columns (`CreatedDate`/`CreatedBy`/`ModifiedDate`/`ModifiedBy`/`AccessTS`, per
  `standards/audit-columns.md`); a fork with different audit-column names updates that list to
  match, or the fork's own columns will sweep in as tracked fields instead of being excluded.
  **Worth flagging, not deciding for you:** once a table has this Lite audit system turned on,
  its own `CreatedDate`/`CreatedBy`/`ModifiedDate`/`ModifiedBy` columns tell you less than they
  used to — `tblAuditLog` now holds a full history of who changed what and when, for every field,
  which is strictly more than a single "last modified by/when" pair can show. Some adopters keep
  the house columns anyway, for a quick single-row answer without a join to the log; others drop
  them once the log is in place. This template doesn't remove them for you — that's your call.
- **Naming conventions** — this template is written in `tbl`/`tlkp` prefix style, and the paired
  scaffold's config scan keys on the `tbl`/`tlkp` prefixes (excluding `tmp`); a practice on
  another convention regenerates under its own names and adjusts that one filter.
- **Error handling** — the house errHandler pattern for the paired scaffold's VBA.

## Extra Options

*Empty in the base template. Filled per client engagement.*

- **Real-username identity** — a public `AuditUser()` helper in both back end and front end,
  replacing `CurrentUser()` in the generated macros and the Long Text helper (Business Rule 9).
  Take the helper as written in `standards/audit-columns.md` and `templates/_materialization.md`
  — including its fallback for an empty `Environ$` — rather than writing a one-liner here.
  Running in the production system this template is drawn from.
- **Audit trail viewer** — a read-only form/report over `tblAuditLog` filtered by table, record,
  date range, or user.
- **Log retention** — a periodic archival job moving aged `tblAuditLog` rows to an archive table
  or file.

## Parked / future considerations

- **Full audit system** — config-driven restore/undo from the trail, retention automation, and a
  managed review UI are the "grown-up" version this Lite template deliberately stops short of.
  A note for whoever builds the restore half: it reads each `OldValue` back into its own field
  and addresses the row by `tblAuditLog.PrimaryKey`. The key field is deliberately not logged as
  a field row of its own (Business Rule 5), which keeps it where a restore needs it — in the
  `WHERE`, never in the `SET`.
- **Composite/text primary keys** — would require reworking the backup plumbing and macro XML
  (Business Rule 4).
