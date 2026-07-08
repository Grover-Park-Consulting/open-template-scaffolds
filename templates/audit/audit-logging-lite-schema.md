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
    it, but the adaptation is theirs.
house_assumptions:
  - tblAuditLogConfig.IsPrimaryKey — every audited table is assumed to have a single-column
    numeric (AutoNumber/Long) primary key; the Long Text backup plumbing and the generated macro
    XML key on one numeric PK, so composite or text keys are not supported
  - tblAuditLog — audited rows are referenced by name and key value (TableName + PrimaryKey),
    deliberately without enforced relationships, so audit history survives deletion of the rows
    it describes
  - tblAuditLogConfig — the schema scan selects candidate tables by the tbl prefix naming
    convention (one code filter in the paired scaffold); everything finer-grained is decided in
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

Two sample business tables and one lookup are included so a generated build demonstrates both
macro paths immediately: `tblClient` (no Long Text — three macros) and `tblSupportTicket` (a Long
Text field — five macros), with `tlkpTicketPriority` showing what the default scan boundary
leaves *out* of auditing. Replace them with — or simply apply the system to — your real tables.

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

**Deliberate teaching point:** the paired scaffold's schema scan takes `tbl…` tables only, so
this `tlkp…` lookup never even reaches the config table unless the adopter widens that one
prefix test — the outer boundary of the audit net is a visible, editable decision, not an
accident (everything inside the boundary is then decided by `IsAuditable` flags, as data).

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
   for it out of the box, and adapting it is the adopter's own project.
5. **Audit scope is data, not code.** The config scan writes every field of every candidate
   table with `IsAuditable` defaulting True; excluding a field or a whole table means flipping
   its flags, not editing code. Two hard exceptions live above the flags: the three system
   tables are **never** given macros (auditing the audit trail would loop — their config rows,
   if present, stay False), and noisy always-changing fields (row-version/timestamp columns,
   house audit columns) are seeded False by the scan.
6. **Every log row names its operation; only real changes are logged on update.** The macro
   stamps `OperationType` (`Insert` / `Update` / `Delete`); an insert row leaves `OldValue`
   Null and a delete row leaves `NewValue` Null. On update, the macro compares old and new
   values (`StrComp` on `Nz`-wrapped values) and logs only fields that actually changed. Long
   Text fields are always logged on update — the comparison cannot be done in the macro.
7. **Regenerate after schema change.** Adding a table or field, or changing a field's type to
   or from Long Text, requires re-running the config scan and regenerating the macros. The
   macros are point-in-time artifacts of the schema.
8. **The backup table is staging, not history.** `tblLongTextBackup` may be cleared at any time;
   the durable record is `tblAuditLog`, which is append-only. Retention/archival policy for the
   log is the adopter's call.
9. **`ChangedBy` identity.** The macros stamp `CurrentUser()` — dependency-free, fires from any
   client, but returns `Admin` unless workgroup security is in use. A real-username upgrade
   (a public VBA helper returning `Environ("USERNAME")`, present in both back end and front
   end) is a named Extra Option; it trades a VBA dependency for a real name. The production
   system this template is drawn from runs the upgrade — its trail shows real Windows
   usernames — so the option is proven, not speculative.
10. **Samples are stand-ins.** `tblClient`, `tblSupportTicket`, and `tlkpTicketPriority` exist
    to prove both macro paths in a fresh build. Real deployments apply the system to the
    adopter's own tables and may omit the samples entirely.

## Standards Layer

- **Audit columns** — a deliberate boundary: the three **system tables do not receive** the house
  audit columns (`DateChanged`/`ChangedBy` here are functional log fields, not the standards
  convention — an append-only log needs no self-stamp). The **sample/business tables** follow
  the house audit-columns convention as usual; note that self-stamp and timestamp fields
  (e.g. a row-version column) are seeded `IsAuditable = False` by the config scan so they don't
  drown the log (Business Rule 5).
- **Naming conventions** — this template is written in `tbl`/`tlkp` prefix style, and the paired
  scaffold's config scan keys on the `tbl` prefix; a practice on another convention regenerates
  under its own names and adjusts that one filter.
- **Error handling** — the house errHandler pattern for the paired scaffold's VBA.

## Extra Options

*Empty in the base template. Filled per client engagement.*

- **Real-username identity** — a public `AuditUser()` helper (`Environ("USERNAME")`) in both back
  end and front end, replacing `CurrentUser()` in the generated macros and the Long Text helper
  (Business Rule 9). Running in the production system this template is drawn from.
- **Audit trail viewer** — a read-only form/report over `tblAuditLog` filtered by table, record,
  date range, or user.
- **Log retention** — a periodic archival job moving aged `tblAuditLog` rows to an archive table
  or file.

## Parked / future considerations

- **Full audit system** — config-driven restore/undo from the trail, retention automation, and a
  managed review UI are the "grown-up" version this Lite template deliberately stops short of.
- **Composite/text primary keys** — would require reworking the backup plumbing and macro XML
  (Business Rule 4).
