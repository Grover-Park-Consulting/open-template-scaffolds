---
template: audit-logging-lite-scaffold
title: Access Audit Logging (Lite) — VBA Scaffold
domain: audit
type: vba-scaffold
version: 0.19.1
status: stable
implements: audit-logging-lite-schema
requires_tables:
  - tblAuditLog
  - tblLongTextBackup
  - tblAuditLogConfig
platform_facts: [target-file, sql-insert-truncation, data-macro-rules, domain-function-transaction]
steps: [01-wizard-questions, 02-before-the-modules, 03-sample-tables-and-readiness, 04-audit-tables-and-config, 05-generator, 06-create-all-data-macros, 07-after-macro-builders, 08-before-macro-builders, 09-long-text-admin-verify, 10-run-the-sequence, 11-checks]
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
  - design-principles
  - audit-columns
target_module: modAddDataMacros
new_procedures:
  - Zero_CreateSampleTables (Path A only)
  - AddAuditColumns (Path A only)
  - Two_CreateAuditTables
  - Three_PopulateConfigTable
  - IsAuditCandidateTable
  - IsNamedInScopeList
  - IsAuditableKeyType
  - IsUnauditableFieldType
  - ListOpenObjects
  - One_CheckAuditReadiness
  - Four_GenerateAllAuditDataMacros
  - CreateAllDataMacros
  - MacroBackupIsOurs
  - RemoveDataMacrosForTable
  - BuildAfterInsertMacro
  - BuildAfterUpdateMacro
  - BuildAfterDeleteMacro
  - BuildBeforeChangeMacro
  - BuildBeforeDeleteMacro
  - AuditSetField
  - AuditKeyExpression
  - AuditValueExpression
  - AuditKeyLocalVar
  - GetComparisonExpression
  - AuditUser (not built when the host's own identity function is used)
  - BackupLongTextFieldsDM
  - BackupAndRemoveAllDataMacros
  - DumpTableMacros
  - ListMacroEvents
build_paths:
  - "Path A — Demo build: try the system out on three made-up tables the generator creates for
    you (Zero_CreateSampleTables). Nothing real is touched."
  - "Path B — Add it to a real database you already have: skip Zero_CreateSampleTables and point
    the generator at your own existing tables instead. Back up the file first (see warnings)."
house_assumptions:
  - "Audit scope (AUDIT_SCOPE_MODE) — chosen by the developer at wizard Step 4, never inferred from
    what the tables are called without being told to. The Standard answer matches tbl/tlkp naming,
    but only because the developer picked that answer over List or All, which are offered equally."
  - "Identity source (AuditUser) — the Windows account name is the preferred choice for the name
    recorded against a change, and what this template ships with. The developer can override it
    before generating — with the database's own identity function, where one already tracks the
    signed-in person by their real name, or with CurrentUser() for the Access-session identity
    instead. Both alternatives are asked for at 'Before you write the modules' and named under
    Extra Options — this is a default offered, never a choice made for the developer."
warnings:
  - Data Macros cannot audit Long Text (Memo) fields on their own. Before building, list every
    Long Text field in the tables to be audited and confirm the list with the developer — any
    table carrying one takes the hybrid VBA path (BeforeChange/BeforeDelete backing values up
    through BackupLongTextFieldsDM); a table without one needs only the three After macros.
    Build that list from the same test the generator itself uses — the field type
    Three_PopulateConfigTable records as DataType 12 — and never from a separate pass over the
    tables. A list assembled independently can differ from the one the build acts on, and then
    the developer has confirmed something other than what gets built. That same test also
    catches Hyperlink fields, which Access stores as Long Text and reports as the same type.
    They are handled correctly, but name them to the developer as Hyperlink, which is what they
    see in the table designer.
  - Two other field types are never audited at all, and the developer is told which ones they have
    rather than left to notice the gap later. An Attachment field cannot be read or written by a
    Data Macro, so a macro referencing one does not work. A calculated field is never edited by
    anyone — its value comes from other fields in the same row, and those fields are audited
    themselves — so a log row for it would record a change nobody made. Three_PopulateConfigTable
    seeds both switched off, and Four_GenerateAllAuditDataMacros ignores the switch if either is
    turned back on afterwards. Before generating, name both types to the developer along with the
    tables and fields concerned.
  - This module must run in the same accdb as the audited tables — the back end of a split
    design. modAuditLongText (BOTH AuditUser and BackupLongTextFieldsDM) must additionally exist in
    every front end, because a data macro fired by a front-end edit resolves the function there.
    AuditUser is needed on every build, Long Text or not, because the stamping macro calls it on
    every table — without it in the front end, front-end inserts fail outright. Where a build uses
    the host database's own identity function instead of AuditUser, that requirement moves with it to
    whichever module holds that function — the same rule, a different file — and modAuditLongText is
    then needed only where there is a Long Text field. The copies must be kept identical by hand;
    nothing enforces that.
  - Close every object in the database before generating — tables, forms, reports and queries.
    Four_GenerateAllAuditDataMacros opens each table in design view to attach its macros, and it
    cannot do that while anything is using the table. This applies whether or not another copy of
    the database is open, and it is the commonest cause of a partial run. Close every other copy as
    well — in a split design, the back end and every front end. A table held open stops the run on
    that table only; the rest still finish. Re-running is safe, so the fix is to close everything
    and run it again.
  - DAO cannot create Data Macros. The only build path is writing UTF-16 XML to a file and
    loading it with Application.LoadFromText acTableDataMacro — exactly what this module does.
  - Every audited table is expected to have a single-field primary key of a kind whose values fit
    in a Long Integer, which means AutoNumber, Long Integer, Integer or Byte. If any table to be
    audited has a different key design (composite, text, no PK, or a number that does not fit a
    Long Integer), stop and tell the developer this template will not work for that table out of
    the box — they are free to adapt it, but the adaptation is theirs. One_CheckAuditReadiness checks
    for this automatically. A Replication ID key is the one exception and is the developer's
    choice rather than a flat exclusion — AUDIT_GUID_KEYS_AUDITED carries the answer, and it
    must be settled before Two_CreateAuditTables runs, since that builds the log's key column.
  - Never name a Replication ID field inside a Data Macro LookUpRecord WhereCondition. Doing so
    was measured to stop the whole condition filtering — every clause, not just that one — so
    the lookup returns the first row of the table and the log records another row's data as
    this row's history. No string conversion inside the condition fixes it. Carry the key in a
    local variable set with the conversion forced ([Key] & "" or CStr), declared before the
    LookUpRecord, and compare against that variable.
  - Path B (an existing accdb with real tables and real data) is much less forgiving than the
    demo. This scaffold attaches Data Macros to the tables the developer already has data in, and it
    changes those tables rather than copies of them; a macro attached wrongly can start refusing
    saves on a table people are using. A build against a database in real use is preceded by a
    backup copy of the file, and the developer is asked for one before anything is changed. If they
    say there is no copy, stop and build nothing.
  - A linked table cannot carry a Data Macro. Auditing is attached to the table itself, in the
    file where the table really lives, so a table that appears in this file only as a link is never
    in scope - whatever the scope setting says, and whether or not the developer named it.
    IsAuditCandidateTable excludes linked tables ahead of every other test, and
    Three_PopulateConfigTable names any the developer asked for by name, so the omission is visible
    rather than silent.
  - Application.LoadFromText replaces a table's ENTIRE macro set — it never merges. This generator
    therefore emits the house audit-column stamping (standards/audit-columns.md) and the change
    auditing TOGETHER, in one Before Change macro, so that one macro carries both and neither can
    replace the other. Any
    OTHER Data Macro a table already carries — business logic of your own, written for reasons
    unrelated to this system — is still replaced. And where a table already stamps who-and-when, that
    macro is replaced by this generator's version, which may not behave identically — an existing one
    may stamp the modified pair on insert as well as on update, or put back a cleared value, neither
    of which this one does. Both write the same columns, so the change is easy to miss. The generator
    backs up a table's existing macros automatically before replacing them, but nothing restores that
    logic afterward; re-adding it is the developer's call. Check for this specifically on Path B.
---

# Access Audit Logging (Lite) — VBA Scaffold

**Status last determined:** 2026-09-26.

## Contents

- Intent
- Prerequisites
- Standards Gate
- Standards Layer
- Extra Options
- Parked / future considerations

The wizard, the module declarations, the procedures, the order the developer runs them in, and the
checks are in this template's eleven step files, delivered one at a time (`steps` in the front
matter lists them in order).

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

The working half of the Lite audit system: the VBA that **creates the three system tables**,
**scans the schema into the config table**, and **generates + attaches the Data Macros** the
paired table template (`audit-logging-lite-schema`) describes. Unlike most scaffolds in this
library, the procedures here are **complete, working code**, not skeletons — the same design
(tables, macro set, Long Text hybrid path) stands behind a live production Access application
whose audit trail validates it end to end. Audit **scope** is decided in data: the scan writes
every candidate field to the config table and you flip `IsAuditable` flags — the `[BUSINESS
LOGIC]` markers land on that review step and on the one code setting that decides which tables are
in scope at all; `[STANDARDS]` markers cover the usual deferred house style.

### Two ways to use this

**Path A** builds three made-up tables and switches auditing on for them, so you can watch the
audit trail work without touching anything real. **Path B** switches auditing on for tables you
already have.

**Which one you are doing, and every other decision this build involves, is asked by the wizard
(step 1, and Steps 6 to 9 in step 10)** — one question at a time, with the reasoning and the warnings available at the step each
belongs to rather than all at once here. The sequence that follows shows what those answers
produce.

```vba
' ---------- Path A — try it out first (nothing real is touched) ----------
Zero_CreateSampleTables         ' 0. create the two made-up tables and a short pick-list
One_CheckAuditReadiness         ' 1. optional here: these tables were just built to the shape
                                '    this system needs, so the check has nothing to find
Two_CreateAuditTables           ' 2. create the 3 tables the audit trail itself lives in
Three_PopulateConfigTable       ' 3. make a list of every field in every table that could be
                                '    audited, switched ON to start
'    ... open the list (tblAuditLogConfig) and switch OFF anything you don't want tracked ...
Four_GenerateAllAuditDataMacros ' 4. turn on tracking for everything still switched ON

' ---------- Path B — add this to a database you already use ----------
' >>> back up the .accdb file first — this step changes real, live tables <<<
One_CheckAuditReadiness         ' 1. FIRST: tells you which of your tables can't be tracked
                                '    as-is (see Business Rule 4 below), before anything at
                                '    all is created
Two_CreateAuditTables           ' 2. create the 3 tables the audit trail itself lives in
Three_PopulateConfigTable False ' 3. make a list of every field in every table that could be
                                '    audited, switched OFF to start
'    ... open the list (tblAuditLogConfig) and switch ON tracking, table by table, for whatever
'        you actually want a history of ...
Four_GenerateAllAuditDataMacros ' 4. turn on tracking for everything switched ON
```

`One_CheckAuditReadiness` and `Four_GenerateAllAuditDataMacros` are called above as bare statements
(the ordinary, interactive way — each pops its own `MsgBox`). Both also take an optional
`bSilent` argument: call them as `sResult = One_CheckAuditReadiness(True)` /
`sResult = Four_GenerateAllAuditDataMacros(True)` to read the same report back as a `String`
with no dialog at all — the way a script, test harness, or an AI assistant facilitating the build
should read the result, rather than adding a throwaway diagnostic just to see what happened.
**Passing `True` is what suppresses the dialog**, not assigning the return value: VBA gives a
procedure no way to detect whether it was called as a function or as a statement.

Then link the three system tables into the front end and import the Long Text helper module
there (see `BackupLongTextFieldsDM`).

**Module homes** (four modules, one job each):

| Module | Procedures | Lives in |
|---|---|---|
| `modAddDataMacros` | `Zero_CreateSampleTables`, `AddAuditColumns`, the four numbered procedures, `CreateAllDataMacros`, the five `Build*` XML builders, `AuditSetField`, `GetComparisonExpression`, `IsAuditCandidateTable`, `IsNamedInScopeList` | Back end only |
| `modAuditLongText` | `AuditUser`, `BackupLongTextFieldsDM`, `BackupKeyLiteral`, `SourceKeyLiteral` | **Back end AND every front end** |
| `modAuditAdmin` | `BackupAndRemoveAllDataMacros` | Back end only |
| `modAuditVerify` | `DumpTableMacros`, `ListMacroEvents` | Back end only |

**Front-matter `target_module` names `modAddDataMacros` because the format allows only one, and
that is where the build itself runs. This table is the authoritative placement** — four modules,
one job each. `modAuditLongText` is the one that must exist in more than one file.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — the working structure provided here.
- **`[STANDARDS]`** — house style, deferred to the standards layer (`error-handling.md`,
  `query-style.md`, `naming-conventions.md`). The error blocks below use the dependency-free
  `MsgBox` default; substitute your house logger per `error-handling.md`.
- **`[BUSINESS LOGIC]`** — the audit-scope decisions you must make: the scan boundary, set in
  `AUDIT_SCOPE_MODE` and read by `IsAuditCandidateTable`, and the `IsAuditable` flag review in
  `tblAuditLogConfig` after the scan.

## Prerequisites

| Object | Role |
|---|---|
| `audit-logging-lite-schema` system tables | `tblAuditLog` / `tblLongTextBackup` / `tblAuditLogConfig` — created by `Two_CreateAuditTables`, described in the paired template |
| The audited tables | Each with a single-column numeric PK (schema Business Rule 4) |
| A Trusted Location | The generator and the macros' VBA calls run only with code enabled |
| `Microsoft Scripting Runtime` (late-bound) | `FileSystemObject` writes the UTF-16 macro XML; `CreateObject` is used, no reference needed |
| **Every object in the database closed** | Close every table, form, report and query before you run **any** of these procedures, and leave them closed until the whole run is finished. `Four_GenerateAllAuditDataMacros` opens each table in **design view** to attach its macros and cannot do that while anything is using the table, so it is the step that fails outright; the earlier steps do not fail, they rebuild the settings the already-attached macros read while someone could still be editing. This applies whether or not another copy of the database is open. |
| **Every other copy of the database closed** | The same requirement, one file further out: another Access instance holding one of those tables blocks it too. In a split design that means the back end *and* every front end — see below. |
| **The database file writable** | Windows can mark a file read-only, and Access opens it anyway — in read-only mode, with no warning until something tries to write. `Four_GenerateAllAuditDataMacros` fails there, *after* the modules are imported and the tables are built. Check the file's properties before you start, and clear it on the back end and on every front end. |

### Ask before building

**Where this is a database already in use, ask for a backup copy before changing anything.** This
template alters the file it is built into. Two questions, in this order:

1. *"Is this a database you already use, or a new one you're trying this out on?"*
2. Where it is one they already use: *"Make a copy of the file before I start. Say when it's done and
   the build continues; say there's no copy and it stops."*

**Stop and build nothing if they say there is no copy.** A copy made before anything changes is the
only way back.

A new database, or one they are trying this out on, needs no copy — the question ends at step 1.

### Where each module goes in a split database

A **split database** is the normal shape for a multi-user Access application: one file holds the
tables (the **back end**, usually on a shared drive), and each person gets their own copy of a
second file holding the forms, reports, and code (the **front end**), whose tables are *links*
pointing at the back end. These templates are designed for that shape. A single-file database — one
.accdb holding everything — is still a legitimate choice for one user, and everything here works
there too; put every module in that one file and ignore the column below.

| Module | Back end | Front end | Why |
|---|---|---|---|
| `modAddDataMacros` (this generator) | **Yes** | No | Data Macros attach to tables in the file the tables actually live in. Run the generator where the tables are. |
| `modAuditLongText` (`AuditUser`, `BackupLongTextFieldsDM`) | **Yes** | **Yes — both** | A Data Macro that fires because someone edited through a *linked* table looks for the function in **that person's front end**. If it isn't there, the save fails. This is the one placement mistake a single-file test can never catch. |
| `modAuditAdmin` (`BackupAndRemoveAllDataMacros`) | **Yes** | No | Maintenance on the tables themselves. |
| `modAuditVerify` (`DumpTableMacros`, `ListMacroEvents`) | **Yes** | No | Reads the macros attached to the tables. |

The three system tables live in the back end and are **linked** into each front end — the audit log
so it can be viewed, the backup table because a front-end-triggered macro has to reach it.

> **`modAuditLongText` is required on every build, not only where there is Long Text.** Its name
> undersells it: it also holds `AuditUser()`, which the stamping macro calls on **every** table. A
> database with no Long Text field anywhere still needs this module in the back end and in every
> front end, or nothing can be inserted at all.
>
> **The copies must stay identical, and nothing enforces that.** It is the same source in two or more
> files, kept in step by hand. Edit it once and you have to apply that edit everywhere — a front end
> running an older copy fails only on the edits that reach the changed line, which is the kind of bug
> that takes a day to find. Treat the back end's copy as the original and re-import it to each front
> end after any change.

## Standards Gate

**Before anything below, run the standards gate** from the `standards-gate` part, in full. It is
one question in the ordinary case and it settles whose rules govern this build. It is a separate
wizard from the one below, and it is asked first: the disclosure line, then the gate, then this
template's house assumptions and warnings, then its entry question.

Both versions of this template run the gate — this one and the outcome-first method
(`audit-logging-lite-outcome-first`), which states the same result and leaves the route open.
They differ in method, not in whose rules govern them. Between them they are the gate's pilot,
and no other template runs it.

## Standards Layer

- **Error handling** — the blocks above ship the dependency-free `MsgBox` default; substitute
  your house pattern per `error-handling.md`. **This is a question, not a default** — see
  *Before you write the modules*, which is where it gets asked. One deliberate exception is
  annotated in place: `BackupLongTextFieldsDM` stays quiet (it runs inside every save).
- **Query style** — the inline SQL kept here is from the proven source; rewrite per
  `query-style.md` if your house centralizes SQL differently.
- **Naming conventions** — the `Standard` scope setting is the naming convention made
  executable: `tbl` and `tlkp` tables, never `tmp`. It is one of the three answers Step 4 offers,
  not the only one. `AUDIT_SCOPE_MODE` holds the answer, `IsAuditCandidateTable` reads it, and both
  the config scan and `One_CheckAuditReadiness` call that — so scope is one setting in one place
  whichever answer was given. A practice on another convention answers `List` or `All` rather than
  editing a test.
- **Design principles** — one job per procedure throughout: one sample-data setup (Path A only),
  three numbered entry points, one safety check, five single-macro builders, one comparison
  helper, one staging function, one admin reset.

## Extra Options

*Named optional extensions, none of them filled in for an engagement.*

- **The identity function this database already has** — on Path B, a database that has been in use
  often has its own, returning the signed-in person's name rather than their Windows account. Where
  you find one, offer it by name and let the developer decide. *Before you write the modules* is
  where that question is asked, and finding a function is not the same as being told to use it.
  **If the developer chooses it**, use it in **all six** sites: the four in the table below, plus
  the two in `BuildBeforeChangeMacro` that stamp `CreatedBy` and `ModifiedBy` on the record.
  `AuditUser()` is then not built at all, and the front-end placement requirement moves with the
  choice.
- **`CurrentUser()` identity instead of `AuditUser()`** — the Access-session user rather than the
  Windows user, and no VBA dependency for identity. `AuditUser()` is the **preferred choice** (schema Business
  Rule 9); this option swaps it back. Choose it if your shop wants the Access user, or wants no VBA
  in the identity path at all.

  **Four sites to change**, all of them in this scaffold:

  | Module | Procedure | What to change |
  |---|---|---|
  | `modAddDataMacros` | `BuildAfterInsertMacro` | the `NewAudit.ChangedBy` value |
  | `modAddDataMacros` | `BuildAfterUpdateMacro` | the `ChangedBy` value (both branches) |
  | `modAddDataMacros` | `BuildAfterDeleteMacro` | the `ChangedBy` value (both branches) |
  | `modAuditLongText` | `BackupLongTextFieldsDM` | `rs!ChangedBy` |

  **Change all four or none.** Each writes to the same `tblAuditLog.ChangedBy`, and the Long Text
  path writes a row the After macro later reads back — so a partial change puts two different names
  on one edit.

  **One thing to decide with it.** The house `standards/audit-columns.md` calls `AuditUser()` for the
  stamped `CreatedBy`/`ModifiedBy`, and this scaffold's stamping macro follows that file. If you
  change the four sites above and leave the standards layer alone, the **log** will say `Admin` while
  the **record** says the real user, for the same change — a trail that disagrees with the row it
  describes. Two coherent ways to hold it:

  - **`CurrentUser()` everywhere** — also point your forked `audit-columns.md` at `CurrentUser()`, so
    stamping and logging agree. Fully dependency-free for identity.
  - **`AuditUser()` everywhere** — the preferred choice; change nothing.

  Mixed is the one to avoid, and it's the one you get by editing only half.
- **Scheduled backup-table cleanup** — a maintenance routine clearing aged `tblLongTextBackup`
  rows (schema Business Rule 8).
- **Audit trail viewer** — a read-only form or report over `tblAuditLog`, filtered by table and
  date. Nothing in this template surfaces the trail to a user; it only writes it.

## Parked / future considerations

- **Restore/undo tooling** — reconstructing a record from its `tblAuditLog` trail; the full
  (non-Lite) system's headline feature.
- **Composite/text primary keys** — `One_CheckAuditReadiness` detects these and tells you to fix or
  exclude the table; the staging plumbing and macro XML themselves still assume one numeric PK
  and don't support them (schema Business Rule 4).