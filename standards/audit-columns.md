# Audit Columns — GPC Default Standards Layer

> **GPC default; fork-and-replace.** Templates defer audit columns to this file — audit columns
> never appear in a template body (per `templates/_template-schema.md` §6). Applies to Microsoft
> Access and SQL Server.

## The GPC audit set

Every `tbl` and `tlkp` table carries these five columns, **always last** in column order
(see `naming-conventions.md` §5.4):

| Column | Type (SQL Server) | Type (Access local) | Rule |
|---|---|---|---|
| `CreatedDate` | `DATETIME NOT NULL DEFAULT GETDATE()` | Date/Time, required | Stamped on INSERT; never updated after |
| `CreatedBy` | `NVARCHAR(100) NOT NULL DEFAULT SUSER_SNAME()` | Text, required | Stamped on INSERT; never updated after |
| `ModifiedDate` | `DATETIME NULL` | Date/Time, nullable | `NULL` until first update; stamped on every change |
| `ModifiedBy` | `NVARCHAR(100) NULL` | Text, nullable | `NULL` until first update; stamped on every change |
| `AccessTS` | `TIMESTAMP NOT NULL` | *(not applicable on local tables)* | Rowversion enabling optimistic concurrency on Access-linked tables |

## How they're maintained

- **SQL Server:** `CreatedDate` / `CreatedBy` are populated by column **defaults** on INSERT.
  `ModifiedDate` / `ModifiedBy` are stamped by an **AFTER UPDATE trigger** on every subsequent
  change — *never by application code*, so audit stamping cannot be bypassed through any interface.
  `AccessTS` is the SQL Server `TIMESTAMP` (rowversion).
- **Access (local tables):** stamped by a **Before Change data macro** on the table — the
  interface-independent equivalent of the SQL Server trigger (it fires no matter how the row is
  written: form, query, direct edit, VBA, import). A default alone can't do it: `CreatedBy` needs the
  current user, and the ACE engine can't evaluate `Environ()` in a default (see `_materialization.md`
  rule 5), so a `Required` `CreatedBy` with no macro **blocks every insert**. The macro:
  - **INSERT** (`IsNull([Old].[<PK>])`) → `CreatedDate = Now()`, `CreatedBy = AuditUser()`.
  - **UPDATE** (else) → `ModifiedDate = Now()`, `ModifiedBy = AuditUser()`; `Created*` stay frozen.
  - Before Change runs **before** Required validation, so it satisfies a `Required` `CreatedBy`.
  - **User identity:** a Public VBA function `AuditUser()` returning `Environ$("USERNAME")`, which the
    macro calls as `=AuditUser()`. A data macro *can* call a public function in the same accdb. Prefer
    this to `CurrentUser()`, which returns `"Admin"` without workgroup security.
  - **Never make an audit field Long Text (Memo)** — data macros can't set Long Text at all; the audit
    set is Short Text / Date-Time by design.
  `AccessTS` does **not** apply — Access has no native rowversion type; it appears only on tables
  linked from SQL Server. *(How the macro is built — DAO can't create data macros — is in
  `templates/_materialization.md`.)*

## Notes

- **`USys` configuration tables are exempt** from the audit-column requirement.
- **Host conventions differ, and that's expected.** A forked practice supplies its own audit names
  here; e.g. a Northwind-derived database uses `AddedBy` / `AddedOn` / `ModifiedBy` / `ModifiedOn`
  via data macros. A template that declares `standards_layer: [audit-columns]` inherits whatever
  *this* file specifies — swap the file, regenerate, and the audit columns follow your house.
