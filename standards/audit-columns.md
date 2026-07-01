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
- **Access (local tables):** maintained by a **Before Change data macro** or the application's save
  path. `AccessTS` does **not** apply — Access has no native rowversion type; that column appears
  only on tables linked from SQL Server.

## Notes

- **`USys` configuration tables are exempt** from the audit-column requirement.
- **Host conventions differ, and that's expected.** A forked practice supplies its own audit names
  here; e.g. a Northwind-derived database uses `AddedBy` / `AddedOn` / `ModifiedBy` / `ModifiedOn`
  via data macros. A template that declares `standards_layer: [audit-columns]` inherits whatever
  *this* file specifies — swap the file, regenerate, and the audit columns follow your house.
