---
template: error-log-schema
title: Error Log Table
domain: errors
type: table-schema
version: 0.1.0
status: draft
standards_layer:
  - naming-conventions
  - audit-columns
  - design-principles
new_tables: [tblErrorLog]
house_assumptions:
  - "tblErrorLog — the house audit columns are deliberately NOT applied to this table. It is
    written from inside a failing error handler, and the OTS audit set is maintained by a Before
    Change data macro calling a VBA function, with a Required CreatedBy. Every one of those is
    another thing that has to be present and working at the exact moment something already is
    not. ErrorOccurredOn and ErrorUser record the same facts as ordinary columns the logger
    writes itself."
warnings:
  - Nothing about this table may be able to refuse a write. No Required column the logger does
    not supply itself, no engine-evaluated default, no data macro, no relationship to another
    table. A log that rejects the row records nothing, and the error it was called about is lost
    with it.
  - In a split database, a log table in the back end cannot be written when the back end is the
    thing that failed — which is the error you most want kept. The paired scaffold's wizard
    offers a text-file fallback for exactly this case.
---

# Error Log Table

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

The one table an application writes to when something goes wrong: which error, in which module and
procedure, on which line, when, and for whom. It is the destination for `LogError`, the shared
logger defined in the paired scaffold (`error-logging-scaffold`), and it exists so that a failure
leaves a record somebody can look at later instead of a dialog somebody clicked away.

It is a deliberately small table with one unusual property: **it must never be able to refuse a
row.** Everything below follows from that. A log table with a required column the logger doesn't
fill, a default the database engine has to work out, or a rule tying it to another table, will one
day reject the write — and it will do it at the worst possible moment, because the moment it is
being written is a moment something has already gone wrong.

Whether this table is used at all, and which file it lives in, are answered by the wizard in the
paired scaffold. This template describes the table those answers produce.

## Prerequisites

**Which file it goes in.** These templates are oriented toward a **split database** — the normal
shape for an Access application, especially with more than one user: one file holds the tables (the
**back end**), and each person runs their own copy of a second file holding the forms, reports, and
code (the **front end**), whose tables are links pointing at the back end. Tables go in the back
end.

For this one table there is a real reason to consider the other answer, and the paired scaffold
asks about it directly (its wizard, Step 4). A shared log in the back end puts everyone's errors in
one place; a log in each front end is always writable, including when the back end is unreachable,
but you have to collect the copies to see the whole picture. **The table definition is identical
either way** — only its home changes.

A single-file database — one .accdb holding everything — is a perfectly good choice for one person
and works the same way; there is only one file, so the question does not arise.

## Entities

### tblErrorLog

Grain: one row per error reported to `LogError`.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ErrorLogID` | AutoNumber | PK | Surrogate key. Also serves as the **reference number** shown to the person at the keyboard, where the wizard's Step 6 includes one — so a support call can name the exact row. |
| `ErrorNumber` | Long | Required | `Err.Number`, read on the logger's first line before anything can clear it. |
| `ErrorDescription` | Text(255) | Required | `Err.Description`, **truncated** to fit rather than allowed to fail on length. A description longer than the field is a shortened record; a refused write is no record at all. |
| `ModuleName` | Text(100) | Required | The module the failing line sits in, supplied by whichever reporting method the wizard's Step 2 chose. Both methods produce the same bare module name, so this column carries one format. |
| `ProcedureName` | Text(100) | Required | The procedure the failing line sits in. |
| `ErrorLineNumber` | Long | Required | `Erl` — the last numbered line executed before the failure. **`0` where the procedure carries no line numbers**, which is a legitimate answer and not a missing value; the column is never Null. |
| `ErrorOccurredOn` | Date/Time | Required | When the logger recorded it, supplied by the logger as a resolved value, never as an engine-evaluated default. |
| `ErrorUser` | Text(100) | Required | Who was running the code. Falls back to `"Unknown"` rather than empty — see Business Rule 5. |

Indexes: PK on `ErrorLogID`; index on `ErrorOccurredOn` (what went wrong recently); index on
`ModuleName` (everything that has ever failed in one module).

Derived (not stored): none. Nothing in this table is computed from anything else, deliberately —
a computation is one more thing that can fail during a write that must not fail.

## Relationships

**None, and that is the design.** `tblErrorLog` is deliberately unrelated to every other table in
the database.

A relationship with referential integrity gives the engine grounds to refuse an insert — a
`ModuleName` that has to match a row somewhere, a user who has to exist in a people table. Any such
rule turns a logging failure into a silent loss of the record. `ModuleName`, `ProcedureName`, and
`ErrorUser` are therefore stored as plain text, not as foreign keys, even though a tidier schema
would relate them.

## Business Rules

1. **Append-only.** Rows are inserted and never updated or deleted by the application. Clearing
   old entries is a housekeeping job somebody runs on purpose (see Extra Options), not something
   the application does on its own.
2. **Nothing about this table may refuse a row.** No Required column the logger does not supply
   itself, no default the database engine has to evaluate, no data macro attached to the table, and
   no relationship to another table. This is the rule the other rules serve.
3. **The logger supplies every column.** Values are resolved in VBA and written as literals. The
   ACE database engine — not VBA — evaluates anything handed to it as text, so `Now()` and
   `Environ()` are not available to it (see `templates/_materialization.md` rule 5); a default
   depending on either would block the insert outright.
4. **`ErrorDescription` is truncated to fit, never allowed to overflow.** A shortened description
   still names the problem; a failed write names nothing.
5. **`ErrorUser` is never empty.** `Environ$("USERNAME")` comes back empty in some contexts — a
   scheduled task, a service account, a locked-down profile — and the column is Required, so an
   empty result would block the insert. It falls back to `"Unknown"`, matching the rule
   `standards/audit-columns.md` states for `AuditUser()`. A row naming an unknown user is a
   record; a refused write is not.
6. **`ErrorLineNumber` is `0`, not Null, when line numbering is off.** `Erl` returns 0 in an
   unnumbered procedure. Storing that plainly means "we know there was no line number" rather than
   "we don't know" — and it keeps the column Required, so the logger can never leave it out.
7. **The reference number is `ErrorLogID`.** Where the wizard's Step 6 shows the person a
   reference, that is the number it shows. Where the error went to a text file instead, there is no
   `ErrorLogID`, and the timestamp on the line is the reference — the paired scaffold says so at
   the point it matters.
8. **The table's home does not change its definition.** Back end or front end (wizard Step 4), the
   fields, types, and indexes above are the same.

## Standards Layer

- **Naming conventions** — OTS house style (`tbl` prefix, PascalCase, no underscores, every field
  name qualified so it means something on its own). Regenerates under a forked practice's own
  conventions without editing this template.
- **Audit columns** — **deliberately not applied here**, and declared as an assumption in this
  template's front-matter rather than left implicit. The OTS default maintains `CreatedDate` /
  `CreatedBy` with a Before Change data macro that calls a VBA function, and makes `CreatedBy`
  Required. On an ordinary table that is exactly right. On this one it would mean an insert made
  from inside a failing error handler depends on a data macro being attached, a function being
  present — and in a split database, present in *that person's front end* as well — and all of it
  working. If any part is missing, the insert fails and the error disappears. `ErrorOccurredOn` and
  `ErrorUser` carry the same two facts as ordinary columns the logger fills itself. A forked
  practice that wants its house audit set here anyway should read Business Rule 2 first.
- **Design principles** — the paired scaffold's procedures follow the house rules for structure;
  this template defines only the table they write to.

## Extra Options

*Empty in the base template. Filled per client engagement.*

- **`ErrorComputerName`** — the machine the error happened on, alongside `ErrorUser`. Worth adding
  where several people share a login, or where you need to know which front end produced a run of
  failures.
- **A context note the caller passes in** — an optional extra argument on `LogError` recording what
  the code was doing ("saving invoice 4471"), stored in a qualified free-text column.
- **Retention** — a query or scheduled routine that deletes rows older than an agreed age, so the
  log does not grow without limit. Deliberately not automatic; see Business Rule 1.
- **A browsing form** — a read-only list of recent errors with a filter by module and date, for a
  developer or support person rather than an end user.
- **Forwarding** — copying new rows to a central SQL Server table or emailing a daily summary,
  where several installations are supported by one team.
