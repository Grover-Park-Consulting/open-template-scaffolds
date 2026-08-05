# Audit Columns — OTS Default Standards Layer

**Who reads this:** the AI assistant, applying these rules to what it generates or a shop deciding
what to replace with their own. If you are building something from a template, only the first two
paragraphs below are for your benefit; the rest of the file is not written for you.

**Building from a template?** What this file decides is the five extra columns every table gets,
recording who created each record and when, and who changed it last.

You do not have to do anything with this information. The rules reach you in the code you receive
whether or not you read this file. Using them as they are is the normal choice. If you came here
because you were asked whether you want to use these rules as they are or make them your own, you
don't answer the question here. After you've read about the choice here, go back and answer it where
it was asked. [`README.md`](README.md) lists all seven files if you want to see the others first.

> **This is the OTS default standards layer.** When you fork the library, replace this file with
> your own house audit-column rules, or add to them. Templates defer audit columns to this file —
> audit columns never appear in a template body (per `templates/_template-schema.md` §6). Applies to
> Microsoft Access and SQL Server.

## The OTS audit set

Every `tbl` and `tlkp` table carries these five columns, **always last** in column order
(see `naming-conventions.md` §6.4):

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
    macro calls as `=AuditUser()`. A data macro *can* call a public function in the accdb the edit is
    happening in. Prefer this to `CurrentUser()`, which returns `"Admin"` without workgroup security.
  - **In a split database, `AuditUser()` must exist in every front end as well as the back end.** A
    split database is the normal shape for a multi-user Access application: the tables live in one
    file (the **back end**) and each person runs their own copy of a second file holding the forms and
    code (the **front end**), whose tables are links to the back end. The stamping macro is attached to
    the table, in the back end — but when it fires because someone edited through a *link*, it looks
    for `AuditUser()` in **that person's front end**. If the function isn't there, the macro can't
    stamp `CreatedBy`, and since `CreatedBy` is `Required` the save fails outright: that front end
    cannot insert a row at all. Put the same module in the back end and in every front end, and
    re-import it everywhere whenever it changes — nothing keeps the copies in step for you. On a
    single-file database there is only one place for it, and this does not arise.
  - **`AuditUser()` must never return an empty string.** `Environ$("USERNAME")` comes back empty in
    some contexts — a scheduled task, a service account, a locked-down profile — and because
    `CreatedBy` is `Required`, an empty result blocks the insert outright. The helper falls back to
    `"Unknown"` when the environment gives it nothing: a stamped row naming an unknown user is a
    record; a refused write is not. (Implementation in `templates/_materialization.md`.)
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
  *this* file specifies — replace the file, and tables built from then on carry your house's audit
  columns.
