---
template: audit-logging-lite-scaffold
title: Access Audit Logging (Lite) — Data Macro Generator VBA Scaffold
domain: audit
type: vba-scaffold
version: 0.7.0
status: draft
wizard: true
implements: audit-logging-lite-schema
requires_tables:
  - tblAuditLog
  - tblLongTextBackup
  - tblAuditLogConfig
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
  - One_CreateAuditTables
  - Two_PopulateConfigTable
  - CheckAuditReadiness
  - Three_GenerateAllAuditDataMacros
  - CreateAllDataMacros
  - BuildAfterInsertMacro
  - BuildAfterUpdateMacro
  - BuildAfterDeleteMacro
  - BuildBeforeChangeMacro
  - BuildBeforeDeleteMacro
  - AuditSetField
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
warnings:
  - Data Macros cannot audit Long Text (Memo) fields on their own. Before building, list every
    Long Text field in the tables to be audited and confirm the list with the developer — any
    table carrying one takes the hybrid VBA path (BeforeChange/BeforeDelete backing values up
    through BackupLongTextFieldsDM); a table without one needs only the three After macros.
  - This module must run in the same accdb as the audited tables — the back end of a split
    design. modAuditLongText (BOTH AuditUser and BackupLongTextFieldsDM) must additionally exist in
    every front end, because a data macro fired by a front-end edit resolves the function there.
    AuditUser is needed on every build, Long Text or not, because the stamping macro calls it on
    every table — without it in the front end, front-end inserts fail outright. Where a build uses
    the host database's own identity function instead of AuditUser, that requirement moves with it to
    whichever module holds that function — the same rule, a different file — and modAuditLongText is
    then needed only where there is a Long Text field. The copies must be kept identical by hand;
    nothing enforces that.
  - Close every copy of the database before generating. Three_GenerateAllAuditDataMacros opens each
    table in design view, and
    a table held open by another Access instance stops the run part-way, leaving some tables done
    and some not. Re-running is safe, so a partial run is fixed by closing everything and repeating.
  - DAO cannot create Data Macros. The only build path is writing UTF-16 XML to a file and
    loading it with Application.LoadFromText acTableDataMacro — exactly what this module does.
  - Every audited table is expected to have a single-column AutoNumber primary key. If any table
    to be audited has a different key design (composite, text, no PK), stop and tell the
    developer this template will not work for that table out of the box — they are free to adapt
    it, but the adaptation is theirs. CheckAuditReadiness checks for this automatically.
  - Path B (an existing accdb with real tables and real data) is much less forgiving than the
    demo. Make a copy of the .accdb file before running any of these steps against it — Data
    Macros get attached directly to your live tables, and this is not a step to redo casually.
  - Application.LoadFromText replaces a table's ENTIRE macro set — it never merges. This generator
    therefore emits the house audit-column stamping (standards/audit-columns.md) and the change
    auditing TOGETHER, in one Before Change macro, so the two no longer destroy each other. Any
    OTHER Data Macro a table already carries — business logic of your own, written for reasons
    unrelated to this system — is still replaced. The generator backs up a table's existing macros
    automatically before replacing them, but nothing restores that logic afterward; re-adding it is
    the developer's call. Check for this specifically on Path B.
  - If any importer un-escapes HTML/XML entities in VBA source on the way in (this repo's own
    Access Explorer MCP code-import tools do), a literal &lt;&gt; in GetComparisonExpression
    becomes a raw <> and breaks the generated macro XML (error 3870). The function builds the
    entity from Chr(38) at runtime for exactly this reason — don't revert it to a literal.
---

# Access Audit Logging (Lite) — Data Macro Generator VBA Scaffold

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
LOGIC]` markers land on that review step and on the one code filter (which table prefix to
scan); `[STANDARDS]` markers cover the usual deferred house style.

### Two ways to use this

**Path A** builds three made-up tables and switches auditing on for them, so you can watch the
audit trail work without touching anything real. **Path B** switches auditing on for tables you
already have.

**Which one you are doing, and every other decision this build involves, is asked by the wizard
below** — one question at a time, with the reasoning and the warnings available at the step each
belongs to rather than all at once here. The sequence that follows shows what those answers
produce.

```vba
' ---------- Path A — try it out first (nothing real is touched) ----------
Zero_CreateSampleTables          ' 0. create the two made-up tables and a short pick-list
One_CreateAuditTables            ' 1. create the 3 tables the audit trail itself lives in
Two_PopulateConfigTable          ' 2. make a list of every field in every table that could be
                                  '    audited, switched ON to start
'    ... open the list (tblAuditLogConfig) and switch OFF anything you don't want tracked ...
Three_GenerateAllAuditDataMacros ' 3. turn on tracking for everything still switched ON

' ---------- Path B — add this to a database you already use ----------
' >>> back up the .accdb file first — this step changes real, live tables <<<
One_CreateAuditTables            ' 1. create the 3 tables the audit trail itself lives in
Two_PopulateConfigTable False    ' 2. make a list of every field in every table that could be
                                  '    audited, switched OFF to start
CheckAuditReadiness              '    a safety check: tells you if any of your tables can't be
                                  '    tracked as-is (see Business Rule 4 below), before anything
                                  '    is changed
'    ... open the list (tblAuditLogConfig) and switch ON tracking, table by table, for whatever
'        you actually want a history of ...
Three_GenerateAllAuditDataMacros ' 3. turn on tracking for everything switched ON
```

`CheckAuditReadiness` and `Three_GenerateAllAuditDataMacros` are called above as bare statements
(the ordinary, interactive way — each pops its own `MsgBox`). Both also take an optional
`bSilent` argument: call them as `sResult = CheckAuditReadiness(True)` /
`sResult = Three_GenerateAllAuditDataMacros(True)` to read the same report back as a `String`
with no dialog at all — the way a script, test harness, or an AI assistant facilitating the build
should read the result, rather than adding a throwaway diagnostic just to see what happened.
**Passing `True` is what suppresses the dialog**, not assigning the return value: VBA gives a
procedure no way to detect whether it was called as a function or as a statement.

Then link the three system tables into the front end and import the Long Text helper module
there (see `BackupLongTextFieldsDM`).

**Module homes** (four modules, one job each):

| Module | Procedures | Lives in |
|---|---|---|
| `modAddDataMacros` | `Zero_CreateSampleTables`, `AddAuditColumns`, the numbered procedures, `CheckAuditReadiness`, `CreateAllDataMacros`, the five `Build*` XML builders, `AuditSetField`, `GetComparisonExpression` | Back end only |
| `modAuditLongText` | `AuditUser`, `BackupLongTextFieldsDM` | **Back end AND every front end** |
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
- **`[BUSINESS LOGIC]`** — the audit-scope decisions you must make: the scan-boundary prefix
  test in `Two_PopulateConfigTable`, and the `IsAuditable` flag review in `tblAuditLogConfig`
  after the scan.

## Prerequisites

| Object | Role |
|---|---|
| `audit-logging-lite-schema` system tables | `tblAuditLog` / `tblLongTextBackup` / `tblAuditLogConfig` — created by `One_CreateAuditTables`, described in the paired template |
| The audited tables | Each with a single-column numeric PK (schema Business Rule 4) |
| A Trusted Location | The generator and the macros' VBA calls run only with code enabled |
| `Microsoft Scripting Runtime` (late-bound) | `FileSystemObject` writes the UTF-16 macro XML; `CreateObject` is used, no reference needed |
| **Every copy of the database closed** | `Three_GenerateAllAuditDataMacros` opens each table in **design view** to attach its macros. Another Access instance holding one of those tables breaks the run part-way through, leaving some tables done and some not. In a split design that means the back end *and* every front end — see below. |
| **The database file writable** | Windows can mark a file read-only, and Access opens it anyway — in read-only mode, with no warning until something tries to write. `Three_GenerateAllAuditDataMacros` fails there, *after* the modules are imported and the tables are built. Check the file's properties before you start, and clear it on the back end and on every front end. |

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

### Before you write the modules

**To the AI assistant.** `standards/error-handling.md` requires you to ask the developer how errors
are reported, and to ask it even when the answer looks settled. **That question is not one of this
template's nine wizard steps and is not meant to be** — the standards file owns it. Ask it through
the selection control before you write any of the four modules, and name it as a question this build
added rather than folding it into the template's count. The answer changes the handler in every
procedure you emit, so asking afterwards means writing them twice. The blocks below ship option 3
inline; anything else means replacing the handler in all four modules, not some of them.

**Look before you ask.** The question is the same on every build; the answers are not, and two of
them exist only if you go and see. Open the file wizard Step 2 named and find out what is there
before you compose the question — asking first offers three answers when there were four.

**What the answers are here.** The three in `error-handling.md` — and on Path B usually a fourth:
**the error handler this database already has.** A database that has been in use generally has one,
and it is usually the right answer, because the code you are adding should fail the way the rest of
the database already fails. Offer it by name if you find one. If it records the failure without
telling anyone, `error-handling.md` says what to do about that.

**When there is no logger at all** — none built here, none of the developer's own — options 1 and 2
are not available, because generated code has to compile on the machine it lands on. Offer the two
real choices: install a logger first, or use the message box now. Never quietly pick the message box
and report it afterwards.

**One procedure keeps its own handler whatever is chosen.** `BackupLongTextFieldsDM` runs inside
every save, called by the data macro itself. A handler that shows a message there interrupts someone
who was only editing a record, over a failure in bookkeeping they never asked for — and a handler
that blocks costs them the edit outright. It stays quiet: log silently if the chosen pattern can, and
never block. Anything else you add to that path follows the same rule.

**Whose name is recorded — ask this the same way, and look first.** The template's answer is
`AuditUser()`, which returns the Windows account name (schema Business Rule 9, and
`standards/audit-columns.md` names it). `CurrentUser()` is the named alternative in Extra Options.
**On Path B there is usually a third, and it is often the best one: the identity function this
database already has.** A database built by a team frequently has one, and it may return the
signed-in person's actual name rather than their Windows account. Find whatever fills the host's
existing tracking columns, offer it by name, and say what it returns that the template's answer does
not.

**If the host's function is chosen, use it in all six places, not four.** Extra Options lists four
sites and those four are the audit *log*. `BuildBeforeChangeMacro` calls the identity function twice
more, for the stamped `CreatedBy` and `ModifiedBy` on the *record*. Change the log's four, leave the
record's two, and you get two different names on one edit — the failure Extra Options already
describes. Choosing the host's function departs from `audit-columns.md`, which names `AuditUser()`:
it holds for this build and is not written back to `standards/`.

**The front-end placement requirement belongs to whichever function is chosen, not to
`AuditUser()`.** The stamping macro calls it on every table, and a macro firing because somebody
edited through a link resolves the function in **that person's front end** — absent, and that front
end cannot insert a row at all. Everything this template says about putting `modAuditLongText` in
every front end is that requirement wearing `AuditUser()`'s name. Choose the host's function and the
requirement does not go away: it moves to whichever module holds it, which is a module this template
does not own and will not place for you. Say which module that is, and tell the developer it now has
to be in every front end.

**And one thing stops being true when that happens.** `modAuditLongText` is described here as
required on every build, which holds because `AuditUser()` lives in it. Replace `AuditUser()` with
the host's function and it holds only `BackupLongTextFieldsDM` — needed where there is a Long Text
field, and nowhere else.

### Before you run the generator

Close **every** copy of the database — the back end and each front end.
`Three_GenerateAllAuditDataMacros` opens each table in
design view to attach its macros, and a table held open elsewhere stops the run part-way, leaving
some tables generated and others not. Re-running is safe (generation replaces), so the fix for a
partial run is simply to close everything and run it again.

## Standards Gate

**Before anything below, run the standards gate** — `templates/_standards-gate.md`, in full. It is
one question in the ordinary case and it settles whose rules govern this build. It is a separate
wizard from the one below, and it is asked first: the disclosure line, then the gate, then this
template's house assumptions and warnings, then its entry question.

This template is the gate's pilot and the only template that runs it.

## Wizard

Nine questions, asked one at a time, preceded by the entry question
(`templates/_template-schema.md` §10.6) that asks whether you want to answer them at all. The
standards gate above is asked before all of them and is not counted among the nine.

**One of the nine — Step 3 — is only asked where you are adding auditing to a database you already
use**, so the try-it-out build asks eight and the other asks all nine.

**If the entry question is answered `Just build it`:** Steps 1, 4, 5 and 6 use their preferred
choices, and the rest are still asked — they have no preferred choice, because each needs something
only you can supply: which file the tables are in, whether you have a backup, whether a list is
right, whether the switches say what you meant, and permission to change your tables. On the
try-it-out demo that is Steps 2, 7, 8 and 9 — four questions instead of eight. On a database you
already use it is those four and Step 3 as well — five instead of nine, because only you can say
whether you have a copy to go back to. State the preferred choices being used before acting on them.

This is a **presentation device, not a second build path** — the same decisions, the same
generated result, met one at a time instead of all at once. Nothing is installed to run it and no
form is built; the AI assistant asks the questions in conversation. See
`templates/_template-schema.md` §10.

> **If an AI assistant is running this for someone:** ask each step and wait for the answer. Never
> infer one — not from what the database looks like, not from reasoning that makes an answer seem
> obvious ("it already has real data, so it must be Path B"). Never work out a check procedure's
> answer yourself by reading the tables: run the procedure at the step that calls for it, show what
> it said, and stop there. Don't collapse the sequence into a single upfront report, even where
> every fact in it turns out correct. Offer to go back at every step after the first, and when an
> earlier answer changes, discard the answers after it and resume forward from there.

**Before step 1**, three things that apply to the whole build whatever the answers are, and that
you can act on right now:

- **Close every copy of the database** — the back end and every front end. The last step opens each
  table in design view, and a table held open by another Access instance stops the run part-way,
  leaving some tables done and some not. Re-running is safe, so the fix for a partial run is to
  close everything and repeat.
- **This module runs in the same file as the tables it audits** — the back end of a split design.
  Data Macros attach to tables in the file the tables actually live in.
- **`modAuditLongText` goes in every front end as well**, because it holds `AuditUser()`, which the
  stamping macro calls on every table. Without it there, a front end cannot insert a row at all. If
  this build uses a function your own database already has for that instead, the same requirement
  applies to whichever module holds that one.

Nothing else is said here. Every other warning this template carries is raised at the step where
you can do something about it.

### Step 1 — Which database are you building this into?

**Ask:** Are you trying this out on made-up tables, or adding it to a database you already use?

| Option | Short description |
|---|---|
| `A try-it-out demo` | Three made-up tables are created and audited. Nothing you already have is touched. |
| `A database you already use` | Auditing is switched on for your own tables. No tables are created for you. |

**Preferred:** `A try-it-out demo` — this template's own. The standards layer does not speak to
this choice.

**Skip when:** never. This is the first question after the entry question.

<details>
<summary>Tell me more about the two builds</summary>

**The try-it-out demo** creates three made-up tables — a client list, a support ticket list, and a
short pick-list of ticket priorities — and switches auditing on for them, so you can watch the
audit trail work before you touch anything real. Good for a first look, a demo, or learning what
this system does. Nothing in your own database is affected, because your own tables are not
involved at all.

**A database you already use** works directly against the tables you have. It does not create the
made-up tables. Because it changes something real, it is much less forgiving: Data Macros get
attached to your live tables, and that is not a step to redo casually.

Both run the same numbered procedures. What differs is `Zero_CreateSampleTables` (demo only), the
starting point for the tracking flags, and two safety steps that only the second path needs.

</details>

### Step 2 — Which file holds the tables?

**Ask:** Which file holds the tables you want audited?

| Option | Short description |
|---|---|
| *one row per database file found* | The tables live in this one. |
| `They're all in one file` | Everything is in a single file, so there is nowhere else for them to be. |

**Preferred:** none. Only you know which file is which, and working it out from a file name is
exactly the kind of guess this wizard is built to avoid.

**Skip when:** never.

**To the AI assistant: the option list holds three answers at most, and `They're all in one file`
is one of them — so two database files fill it exactly.** An ordinary folder holds more than two:
the file with the real tables, a copy of the file people run, and a backup of either. Narrow the
list before offering it rather than truncating it afterwards. A file holding only links to tables
elsewhere cannot be the answer, and neither can one with no tables in it at all, so neither is
offered. Drop `They're all in one file` as soon as a second database file has been found — the
folder has already answered it. If more than three candidates still survive, offer three and say in
the question itself that a different file can be named instead. There is no room on this step for
`Go back to the previous question`; that is expected here, not an omission.

<details>
<summary>Tell me more about why the file matters</summary>

The tracking gets attached to the tables themselves, so it has to be built in the file the tables
actually live in. Build it in the wrong one and nothing errors — you simply get a database with
some code in it and no tracking anywhere.

Most Access applications with more than one user are split across two files: one holds the tables
(usually on a shared drive), and each person runs their own copy of a second file holding the
forms, reports and code. The second file doesn't really contain the tables — it contains links
pointing at them. Tracking has to go where the real tables are.

If the two files are named alike, or one is a copy, open each and look: the one with real tables in
it rather than linked ones is the answer.

</details>

### Step 3 — Have you made a backup copy of this database?

**Ask:** Have you made a backup copy of this database?

| Option | Short description |
|---|---|
| `Yes, I have a copy I can go back to` | Carry on to the next question. |
| `No — stop so I can make one` | Everything stops here. Nothing has been changed. |

**Preferred:** none. Only you can say whether you have one.

**Skip when:** Step 1 chose the try-it-out demo — a demo touches nothing you would want back.

<details>
<summary>Tell me more about why a copy matters here</summary>

This build attaches Data Macros directly to your live tables, and `Application.LoadFromText`
**replaces a table's entire macro set** — it never merges. Any other Data Macro a table already
carries, written by you for reasons unrelated to auditing, is replaced along with everything else.

Existing macros are backed up to a file before they are replaced, and the last step tells you which
tables that happened to. But nothing puts that logic back afterward: re-adding it is your call,
from the backup.

Copying the database file is the same precaution you would take before any change you cannot easily
undo, and it is the only one that covers everything at once.

</details>

### Step 4 — Which tables should be audited?

**Ask:** Which tables should be audited?

| Option | Short description |
|---|---|
| `Tables named tbl… or tlkp…, but not tmp…` | Your data and lookup tables. Temporary tables are left out. |
| `A different set — I'll tell you which` | You say which tables, and I use that instead. |

**Preferred:** `Tables named tbl… or tlkp…, but not tmp…` — the naming style these templates
follow, which is where those prefixes are defined.

**Skip when:** never.

<details>
<summary>Tell me more about which tables get audited</summary>

This answer is the **only** one in the whole system that ends up written into code. Every
finer-grained
choice — which tables, which individual fields — is a flag you set in the config table afterward,
as data. That is deliberate: changing your mind about a field should not mean editing VBA.

Access's own system tables (the ones whose names begin `MSys`) never match either prefix, so they
are never in scope regardless of what you choose here.

One thing to know if you change it: the same test appears in **two** procedures,
`Two_PopulateConfigTable` and `CheckAuditReadiness`. Change one and you must change the other, or
the readiness check will be reporting on a different set of tables than the scan wrote.

</details>

### Step 5 — Start by tracking everything, or nothing?

**Ask:** Should every field start out tracked, or should none of them?

| Option | Short description |
|---|---|
| `Track everything to start` | Everything is switched on, and you switch off what you don't want. |
| `Track nothing to start` | Everything is switched off, and you switch on what you do want. |

**Preferred:** follows Step 1 — `Track everything to start` on the demo, `Track nothing to start`
on a database you already use. This template's own.

**Skip when:** never.

<details>
<summary>Tell me more about the starting point</summary>

Either way you get one line per field and you decide the rest by flipping switches, so neither
answer locks anything in. Under the covers this sets a single argument on
`Two_PopulateConfigTable` — nothing after it starts everything on, `False` starts everything off.

**Starting on** suits the demo, where there are nine or ten fields and you want to see the trail
working immediately.

**Starting off** is the safer footing on tables this system was not designed around, where "track
everything" can sweep in more than you meant — long free-text notes, columns another process
rewrites constantly, fields you would rather not have a second copy of.

Some rows are switched **off no matter which you choose**: the three system tables (auditing the
audit trail would loop), each table's own primary key (its value is already on every log row), and
the house audit columns and other always-changing system fields. Nothing is hidden — those rows
are written and shown, just switched off.

**What you will see, so it does not look wrong:** most of the rows belong to the three system
tables, and on a small database that can be well over half of them. Sort by table name and review
only the tables you recognise as your own.

</details>

### Step 6 — Check the tables first?

**Ask:** Should I read your table definitions and report any table this system can't track as it
stands?

| Option | Short description |
|---|---|
| `Yes, check them first` | I look at every table and tell you what I find. Nothing is changed. |
| `No, skip the check` | Go straight on. |

**Preferred:** `Yes, check them first` — this template's own. On a database you already use it is
expected rather than optional, for the reason in the note below.

**Skip when:** never, though it is genuinely optional on the demo, where the tables were built by
this template and are already known to be the right shape.

<details>
<summary>Tell me more about the check</summary>

It reads your table definitions and reports whether each one will work. It changes nothing at all,
so there is no risk in running it, and it can be run again at any time. The procedure behind it is
`CheckAuditReadiness`.

**What it is looking for:** every table needs **one single number field as its primary key, set to
auto-number**. Most tables you designed yourself already look like this. Older or inherited tables
sometimes don't — a table with no primary key set, one that uses two or more fields together as its
key, or one keyed on a text code will not work with this system as it stands.

That matters most on a database you already use, which is exactly where such a table is likely to
turn up.

If a table isn't ready you have two ways out: fix that table's primary key, or leave the table out
by switching its fields off. Adapting the template to a different key design is possible, but it is
your adaptation, not something this template supports.

</details>

### Step 7 — Are these the long-text fields?

**Ask:** These are the long-text fields I found. Is that right?

| Option | Short description |
|---|---|
| `Yes, that's right` | Tables with one of these get some extra handling. |
| `No — let me look first` | Nothing happens until you say so. |

**Preferred:** none. It is your database, and the build acts on this list.

**Skip when:** never — but where no long-text field is found anywhere, this is an empty list to
confirm rather than a decision.

<details>
<summary>Tell me more about Long Text fields</summary>

**A Data Macro cannot read or write a Long Text field at all.** That is a limit of the Access
engine, not a choice this template made, and it cannot be worked around inside the macro.

So a table carrying one takes a different route: a Before Change and a Before Delete macro call a
VBA function (`BackupLongTextFieldsDM`) that copies the old value into a staging table first, and
the audit macro reads it back from there. A table with no Long Text field needs only the three
simpler After macros.

That is why the list matters: it decides which macros each table gets. Getting it wrong doesn't
produce an error — it produces an audit trail that quietly records nothing for that field.

The function this depends on lives in `modAuditLongText`, which must exist in the back end **and in
every front end**, for the reason given before step 1.

</details>

### Step 8 — Have you set the tracking switches the way you want them?

**Ask:** Have you set the tracking switches the way you want them?

| Option | Short description |
|---|---|
| `Yes, they're how I want them` | Carry on. |
| `Not yet — I'll set them now` | I close the database so you can open `tblAuditLogConfig` and set them, then wait until you say you're done. |

**Preferred:** none. Only you know whether the list says what you meant.

**Skip when:** never. This is the review the whole design is built around.

**To the AI assistant: if you are building the database yourself rather than handing over a script,
you are holding the file open and the developer cannot open it.** Close it, say that you have, and
wait for them. Then reopen it to carry on. This applies at every step that asks the developer to go
and look at a table, not only this one — an instruction to open something they are locked out of
reads as the wizard being broken.

<details>
<summary>Tell me more about the switches</summary>

This is where you actually decide what gets tracked, and you decide it as data rather than in code
— one line per field, with a Yes/No switch. Open `tblAuditLogConfig`, sort by table name, and set
them.

It is worth spending a minute on, because the next step acts on exactly what is in that table. It
is also the cheapest thing here to change your mind about later: flip a switch, run the last step
again, and the macros are rebuilt.

The rows for the three system tables are switched off and must stay off — auditing the audit trail
would loop.

</details>

### Step 9 — Ready to switch auditing on?

**Ask:** Ready to switch auditing on?

| Option | Short description |
|---|---|
| `Yes, switch it on` | Attaches the tracking to your tables and reports what happened to each one. Any Data Macros they already carry are replaced. |
| `Not yet` | Nothing is changed. |

**Preferred:** none. This is the step that changes your tables.

**Skip when:** never.

<details>
<summary>Tell me more about what this step does</summary>

For each table still switched on, the generator writes the macro definitions out as XML and loads
them onto the table. It has to be done this way: **DAO cannot create Data Macros**, and writing
UTF-16 XML and loading it with `Application.LoadFromText` is the only build path there is.

`LoadFromText` replaces a table's entire macro set. It never merges. This is why the house
audit-column stamping and the change auditing are generated together, in one Before Change
macro — so the two no longer destroy each other. But any *other* Data Macro a table carries,
***business logic of your own written for unrelated reasons is replaced too.*** The generator backs a
table's existing macros up to `DataMacroBackups\` before replacing them and names the affected
tables in its report; putting that logic back is your call.

**It is safe to run again.** Generation replaces rather than accumulates, so a run that stopped
part-way — because a table was held open somewhere — is fixed by closing everything and running it
again.

**A table with auditing switched off still gets a stamping macro.** That is not an oversight: the
house audit columns are `Required`, and a table with no macro at all has no way to fill them, so it
would refuse every insert.

**One thing to check about the code you imported, now that it is about to run.** If your import
tool un-escapes XML entities — this repo's own Access Explorer MCP code-import tools do — then the
literal `&lt;&gt;` inside `GetComparisonExpression` has become a raw `<>`, and the macro XML this
step generates will fail with error 3870. The function builds that entity from `Chr(38)` at run
time for exactly this reason; confirm it survived the import, and don't revert it to a literal.

</details>

## Module-level declarations

Put these at the top of `modAddDataMacros`, below `Option Explicit`.

```vba
Option Compare Database
Option Explicit

' [STANDARDS — audit-columns.md] The house audit column names, in ONE place.
' Three procedures below use them: Two_PopulateConfigTable (to seed them
' not-auditable), BuildBeforeChangeMacro (to stamp them), and
' Zero_CreateSampleTables (to create them on the sample tables).
' A shop that forks this library and renames its audit columns changes these
' four lines and nothing else. AccessTS is deliberately absent — it is a SQL
' Server rowversion and does not apply to a local Access table.
Private Const AUDIT_CREATED_DATE  As String = "CreatedDate"
Private Const AUDIT_CREATED_BY    As String = "CreatedBy"
Private Const AUDIT_MODIFIED_DATE As String = "ModifiedDate"
Private Const AUDIT_MODIFIED_BY   As String = "ModifiedBy"
```

**To the AI assistant — Path B, before you import the module.** These four values name the columns
the generator stamps. A host database that already carries tracking columns may name them something
else — `AddedOn`/`AddedBy`/`ModifiedOn`/`ModifiedBy` is one form you will meet, and there are others.

**First find out how the host fills those columns today, because it changes what every answer below
costs.** There are three ways and you cannot tell them apart without looking:

- **A data macro on the table.** The generator replaces a table's entire macro set, so the host's
  stamping is overwritten by the generated one automatically, and the original is written to
  `DataMacroBackups\` first. Nothing extra for anyone to do — but read the old macro before you
  replace it and tell the developer if it did anything the generated one will not.
- **A column default.** Common on a date column; it cannot fill a user-name column, because the
  engine cannot reach the current user from a default. The generator never touches defaults, so it
  stays and keeps firing. A default and the generated macro both writing the same column is
  harmless — Before Change runs at the engine and its value is the one that lands — but say that you
  found it.
- **VBA — a form's Before Update, a save routine, anywhere else.** The generator does not touch VBA.
  That code keeps running after the build, alongside the generated macro, both writing the same
  columns. Search the database's own modules and form code for the four column names so you know
  whether this is the case, and tell the developer what you found. Retiring their code is their
  call, not this template's.

**When the host's names differ from the four above, ask. Never settle it yourself.** Put it through
the selection control as its own question, and name it as a question this database added rather than
folding it into the template's own count. The two answers, and both belong in the question:

- **Use the names this database already has.** No columns are added. Set the four values above to
  the host's names. `Two_PopulateConfigTable` then seeds those columns not-auditable, and
  `BuildBeforeChangeMacro` stamps them, exactly as it would the library's own names.
- **Use the library's names.** The host's four columns are replaced by the library's four. Leave the
  values above as shipped; they seed and stamp the library's columns. **Say what this costs before
  they choose it, because nothing here does it for them:** the four columns have to be added, the
  values already in the host's columns carried across, the old four dropped, and every form, query,
  report and line of the database's own code that names them corrected. **What it costs depends on
  what you found above** — a data macro is replaced anyway and a default goes with the column it sits
  on, but **their own VBA stops working the moment the old columns are dropped, because the columns
  it names are gone.** Check whether the tables hold data and say so when you ask: on an empty table
  this is cheap; on a table with history it is not.

**When no table carries tracking columns**, there is nothing to reconcile and nothing to ask. Leave
the four values above as shipped, and **tell the developer that these four columns will be added to
each table they chose to audit.** Two things belong in what you tell them: `CreatedDate` and
`CreatedBy` are required, so a table that already holds rows needs the column added, filled, and only
then marked required; and every row that was already there will carry the date the columns were added
rather than the date the record was really created, because that history was never kept.

**When some tables carry them and others don't, ask what happens to the ones that don't** — the split
is usually a decision somebody already made, and overriding it silently is the wrong move. Two
answers: add the four columns so every audited table stamps who and when, or leave those tables as
they are, in which case they still get their audit trail and simply have nothing to stamp. Both work.
Say which tables you're asking about, by name.

Left as shipped on a host that names them differently, with nobody asked, they match no column: the
host's tracking columns get seeded as ordinary tracked fields, `BuildBeforeChangeMacro` emits no
stamping for any table, and because loading a macro set replaces the whole set, a table that arrived
with stamping loses it. Nothing errors.

## Procedures

### Zero_CreateSampleTables — `Public Function` → `String` (Path A only — setup step 0)

**Skip this procedure entirely if you're doing Path B** (adding audit tracking to a database
you already use) — it only exists to build the made-up tables for trying the system out.

Creates the two made-up tables (`tblClient`, `tblSupportTicket`) and a short pick-list
(`tlkpTicketPriority`) described in the paired schema template, with the four starter pick-list
rows (Low, Normal, High, Urgent) and the two links between the tables. Same idempotent style as
`One_CreateAuditTables` — an existing table is reported and skipped, so it's safe to re-run.

**All three carry the house audit columns** (`standards/audit-columns.md`), because a demo that
leaves them off doesn't demonstrate the thing most likely to bite on real tables: stamping and
change-auditing both wanting the Before Change macro. With them present, the Path A build shows the
two being generated together, which is what a real table needs. It also means the four seed rows
have to supply `CreatedDate` and `CreatedBy` themselves — they're inserted at step 0, before any
stamping macro exists, and those columns are `Required`. See the comment on the `INSERT` statements
below, and `templates/_materialization.md` rule 5 for why the values must be resolved in VBA first.

It returns the same wording it shows on screen. Run it from the Immediate window as
`Zero_CreateSampleTables` and you get the message box; call it as
`sResult = Zero_CreateSampleTables(True)` and you get the text back with **no** message box —
which is what an automated caller needs, because nothing is there to click a dialog away.

```vba
Public Function Zero_CreateSampleTables(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Creates tblClient, tlkpTicketPriority, tblSupportTicket (schema template
    '            entities) and seeds tlkpTicketPriority. Path A (try-it-out build) only —
    '            skip this procedure for Path B (an existing accdb's own tables). Idempotent:
    '            each block is skipped if its table already exists.
    '            Passing bSilent:=True suppresses the message box and returns the same text,
    '            so a caller with no one at the keyboard does not hang on a dialog.
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim rel As DAO.Relation
    Dim sReport As String
    Dim bFailed As Boolean
    Dim sAuditUser As String
    Dim sAuditNow As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' ========== tblClient ==========
    On Error Resume Next
    Set tdf = db.TableDefs("tblClient")
    If Not tdf Is Nothing Then
        Debug.Print "tblClient already exists"
        GoTo CreateTicketPriority
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblClient")

    Set fld = tdf.CreateField("ClientID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ClientName", dbText, 100)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("EmailAddress", dbText, 100)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("CellPhone", dbText, 25)
    tdf.Fields.Append fld

    AddAuditColumns tdf

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ClientID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("ClientName")
    idx.Unique = True
    Set fld = idx.CreateField("ClientName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblClient created"

CreateTicketPriority:
    ' ========== tlkpTicketPriority ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tlkpTicketPriority")
    If Not tdf Is Nothing Then
        Debug.Print "tlkpTicketPriority already exists"
        GoTo CreateSupportTicket
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tlkpTicketPriority")

    Set fld = tdf.CreateField("TicketPriorityID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketPriorityName", dbText, 30)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("SortOrder", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    AddAuditColumns tdf

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("TicketPriorityID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("TicketPriorityName")
    idx.Unique = True
    Set fld = idx.CreateField("TicketPriorityName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    db.TableDefs.Refresh

    ' [SCAFFOLD + _materialization.md rule 5] Starter pick-list rows (schema:
    '            Low/Normal/High/Urgent). These are inserted at step 0, BEFORE any
    '            stamping macro exists — and CreatedDate/CreatedBy are Required, so the
    '            values have to be supplied here. They cannot be supplied as functions:
    '            the ACE engine, not VBA, evaluates the text of a db.Execute INSERT, and
    '            Environ() and AuditUser() are unknown to it ("Undefined function").
    '            Resolve each value in VBA first, then concatenate it in as a literal.
    sAuditUser = Environ$("USERNAME")
    If Len(sAuditUser) = 0 Then sAuditUser = "Unknown"
    sAuditNow = "#" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "#"

    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('Low', 10, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('Normal', 20, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('High', 30, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('Urgent', 40, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError

    Debug.Print "tlkpTicketPriority created and seeded"

CreateSupportTicket:
    ' ========== tblSupportTicket ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblSupportTicket")
    If Not tdf Is Nothing Then
        Debug.Print "tblSupportTicket already exists"
        GoTo CreateRelationships
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblSupportTicket")

    Set fld = tdf.CreateField("SupportTicketID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ClientID", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketPriorityID", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketSubject", dbText, 255)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketDetail", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketOpenedDate", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketClosedDate", dbDate)
    tdf.Fields.Append fld

    AddAuditColumns tdf

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("SupportTicketID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("ClientID")
    Set fld = idx.CreateField("ClientID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("TicketPriorityID")
    Set fld = idx.CreateField("TicketPriorityID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblSupportTicket created"

CreateRelationships:
    ' ========== Relationships (schema: enforced, no cascade) ==========
    On Error Resume Next
    db.Relations.Delete "tblClient_tblSupportTicket"
    db.Relations.Delete "tlkpTicketPriority_tblSupportTicket"
    On Error GoTo errHandler

    Set rel = db.CreateRelation("tblClient_tblSupportTicket", "tblClient", "tblSupportTicket", 0)
    Set fld = rel.CreateField("ClientID")
    fld.ForeignName = "ClientID"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("tlkpTicketPriority_tblSupportTicket", "tlkpTicketPriority", "tblSupportTicket", 0)
    Set fld = rel.CreateField("TicketPriorityID")
    fld.ForeignName = "TicketPriorityID"
    rel.Fields.Append fld
    db.Relations.Append rel

    Debug.Print "Relationships created"

    sReport = "Sample tables created. You're on Path A (try-it-out build) — nothing in your " & _
        "own database was touched."

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set rel = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Zero_CreateSampleTables = sReport
    ' [SCAFFOLD] One message, whatever happened — the success text or the error text, never
    '            both. Building the report first and showing it here is what prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    bFailed = True
    sReport = "ERROR creating sample tables: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### AddAuditColumns — `Private Sub` (Path A only — helper to step 0)

Appends the house audit set to a table being built, always last in column order (per
`standards/naming-conventions.md` §6.4). Called once per sample table by
`Zero_CreateSampleTables`, before the table is appended to the database.

`AccessTS` from the standards file is deliberately **not** created: it's a SQL Server rowversion
column, and a local Access table has no equivalent. None of the four is Long Text — a Data Macro
cannot write a Long Text field at all, which is why the audit set is Short Text and Date/Time by
design.

**On Path B this procedure is not used.** Your own tables already have whatever audit columns your
standards give them; the generator reads what's there rather than adding anything.

```vba
Private Sub AddAuditColumns(tdf As DAO.TableDef)
    ' [STANDARDS — audit-columns.md] The house audit set, always last in column order.
    '            Filled at run time by the Before Change stamping macro that
    '            BuildBeforeChangeMacro emits together with the audit macros.
    '            Column NAMES come from the constants at the top of this module — change
    '            them there, not here, if your house uses different names.
    Dim fld As DAO.Field

    Set fld = tdf.CreateField(AUDIT_CREATED_DATE, dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField(AUDIT_CREATED_BY, dbText, 100)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField(AUDIT_MODIFIED_DATE, dbDate)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField(AUDIT_MODIFIED_BY, dbText, 100)
    tdf.Fields.Append fld

    Set fld = Nothing
End Sub
```

### One_CreateAuditTables — `Public Function` → `String` (setup step 1)

Creates the three system tables via DAO, idempotently — an existing table is reported and
skipped, so it is safe to re-run. Field-by-field DAO `CreateField` (never `CREATE TABLE`
DDL — see `templates/_materialization.md`). Each table is built with the indexes the schema
template declares for it, not the primary key alone.

**The skip is whole-table.** A table that already exists is left exactly as it is, indexes
included — re-running this procedure will not add a missing index to a table built by an earlier
version of this scaffold. If you have such a build, add the secondary and unique indexes by
hand, or start fresh in a copy.

```vba
Public Function One_CreateAuditTables(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Creates tblAuditLog, tblLongTextBackup, tblAuditLogConfig (schema template
    '            entities). Idempotent: each block is skipped if its table already exists.
    '            Passing bSilent:=True suppresses the message box and returns the same text,
    '            so a caller with no one at the keyboard does not hang on a dialog.
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim sReport As String
    Dim bFailed As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [SCAFFOLD] The outcome text is set here, not at the end, because an "already exists"
    '            branch jumps straight to Cleanup. The errHandler replaces it on failure.
    sReport = "Audit tables are in place. Any that already existed were left exactly as " & _
        "they are, indexes included."

    ' ========== tblAuditLog ==========
    On Error Resume Next
    Set tdf = db.TableDefs("tblAuditLog")
    If Not tdf Is Nothing Then
        Debug.Print "tblAuditLog already exists"
        GoTo CreateLongTextBackup
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblAuditLog")

    Set fld = tdf.CreateField("AuditLogID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("PrimaryKey", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OperationType", dbText, 25)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OldValue", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("NewValue", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DateChanged", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ChangedBy", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("AuditLogID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCHEMA] Secondary indexes the schema template declares: trail queries by table and
    '          date, per-record history by table and key value.
    Set idx = tdf.CreateIndex("TableNameDateChanged")
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("DateChanged")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("TableNamePrimaryKey")
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("PrimaryKey")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblAuditLog created"

CreateLongTextBackup:
    ' ========== tblLongTextBackup ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblLongTextBackup")
    If Not tdf Is Nothing Then
        Debug.Print "tblLongTextBackup already exists"
        GoTo CreateConfig
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblLongTextBackup")

    Set fld = tdf.CreateField("LongTextBackupID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("PrimaryKey", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OldValue", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DateChanged", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ChangedBy", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("LongTextBackupID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCHEMA] Unique per table/record/field — BackupLongTextFieldsDM replaces any earlier
    '          backup for the same field of the same row, and this enforces that one-row rule.
    Set idx = tdf.CreateIndex("TableNamePrimaryKeyFieldName")
    idx.Unique = True
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("PrimaryKey")
    idx.Fields.Append fld
    Set fld = idx.CreateField("FieldName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblLongTextBackup created"

CreateConfig:
    ' ========== tblAuditLogConfig ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblAuditLogConfig")
    If Not tdf Is Nothing Then
        Debug.Print "tblAuditLogConfig already exists"
        GoTo Cleanup
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblAuditLogConfig")

    Set fld = tdf.CreateField("AuditLogConfigID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldPosition", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DataType", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("IsPrimaryKey", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"     ' [SCAFFOLD] set before Append — never as DDL DEFAULT
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("IsAuditable", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "True"      ' [SCAFFOLD] audit scope is decided in this flag, as data
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("AuditLogConfigID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCHEMA] Unique per table/field — one config row per scanned field, so a re-scan or a
    '          hand edit cannot leave two rows disagreeing about the same field.
    Set idx = tdf.CreateIndex("TableNameFieldName")
    idx.Unique = True
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("FieldName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblAuditLogConfig created"

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing
    One_CreateAuditTables = sReport
    ' [SCAFFOLD] One message, whatever happened — the success text or the error text, never
    '            both. Building the report first and showing it here is what prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    bFailed = True
    sReport = "ERROR creating tables: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### Two_PopulateConfigTable — `Public Function` → `String` (setup step 2)

Scans the schema into `tblAuditLogConfig`: **every field of every candidate table**, with its
ordinal position, DAO type code, a flag on the table's PK field, and `IsAuditable`. Nothing is
silently dropped — exclusions are *seeded* as `IsAuditable = False` rows: the three system
tables, the audited table's own primary-key field (its value is already on every log row, in
`tblAuditLog.PrimaryKey`), plus fields that would just add noise — this repo's house audit columns
(`CreatedDate`/`CreatedBy`/`ModifiedDate`/`ModifiedBy`/`AccessTS`, per `standards/audit-columns.md`)
and a few other always-changing system columns (`SSMA_TimeStamp`, `ValidFrom`, `ValidTo`). **After
running, open the config table and review the flags** — that review, in data, is where the audit
net is drawn (schema Business Rule 5).

**What you will see, so it doesn't look wrong:** most of the rows belong to the three system
tables — `tblAuditLog`, `tblLongTextBackup`, `tblAuditLogConfig` — and on a small database that
can be well over half of them. They are all switched OFF and must stay OFF; auditing the audit
trail would loop. Sort by `TableName` and review only the rows for tables you recognise as your
own. The system rows are shown rather than hidden on purpose — nothing the scan did is invisible
to you.

Takes one optional Yes/No setting that decides the starting point for everything else:

- **Path A (try-it-out build):** run `Two_PopulateConfigTable` with nothing after it. Every
  field starts switched ON, and you switch OFF the few you don't want tracked.
- **Path B (a database you already use):** run `Two_PopulateConfigTable False`. Every field
  starts switched OFF, and you switch ON — table by table — only what you actually want a
  history of. This is the safer starting point on tables this system wasn't designed around,
  where "track everything" could sweep in more than you meant.

```vba
Public Function Two_PopulateConfigTable(Optional bDefaultAuditable As Boolean = True, _
                                        Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Rebuild the audit configuration from the live schema. Scope decisions
    '            live in the IsAuditable flags afterward, not in this code.
    '            Passing bSilent:=True suppresses the message box and returns the same text,
    '            so a caller with no one at the keyboard does not hang on a dialog.
    '            bDefaultAuditable sets the starting point for ordinary fields only:
    '            True  (preferred; demo build) — everything starts switched ON, you switch OFF
    '                  what you don't want tracked.
    '            False (Path B — call as Two_PopulateConfigTable(False)) — everything
    '                  starts switched OFF, you switch ON what you do want tracked.
    '            The three system tables and the noisy always-changing fields below are
    '            always switched OFF, no matter which way this is called.
    Dim db As DAO.Database
    Dim tdef As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim pkField As DAO.Field
    Dim sSql As String
    Dim isPK As Boolean
    Dim isAuditable As Boolean
    Dim pkFieldName As String
    Dim lRowCount As Long
    Dim sReport As String
    Dim bFailed As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb
    lRowCount = 0

    ' Clear existing config
    db.Execute "DELETE * FROM tblAuditLogConfig", dbFailOnError

    For Each tdef In db.TableDefs
        ' [BUSINESS LOGIC — scan boundary] Which tables are candidates at all. This default
        ' scans tables prefixed tbl or tlkp (your data and lookup tables under
        ' naming-conventions.md) but never tmp (temporary/working tables); system (MSys)
        ' tables never match either prefix. On another naming policy, change this one test —
        ' every finer-grained decision is a flag in the config table, not code. Keep
        ' CheckAuditReadiness's copy of this test in sync if you change it here.
        ' >>> adjust the prefix test to your naming convention <<<
        If (Left(tdef.Name, 3) = "tbl" Or Left(tdef.Name, 4) = "tlkp") _
            And Left(tdef.Name, 3) <> "tmp" Then

            ' Get the primary key field name for this table
            pkFieldName = ""
            For Each idx In tdef.Indexes
                If idx.Primary Then
                    For Each pkField In idx.Fields
                        pkFieldName = pkField.Name
                        Exit For
                    Next pkField
                    Exit For
                End If
            Next idx

            For Each fld In tdef.Fields
                isPK = (fld.Name = pkFieldName)

                ' [SCAFFOLD] Seed IsAuditable: False for the system tables themselves
                '            (schema Business Rule 5 — never audit the audit trail) and for
                '            noisy always-changing fields; the starting point set by
                '            bDefaultAuditable for everything else. Review and flip flags in
                '            tblAuditLogConfig after the scan.
                ' [STANDARDS — audit-columns.md] The four house audit columns come from the
                '            constants at the top of this module — one place to change if your
                '            standards/audit-columns.md names them differently. They are a
                '            VBA-side mirror of that file, not a live read of it.
                '            AccessTS is named here as a literal because it is a SQL Server
                '            rowversion that only appears on linked tables, so it is never one
                '            of the columns this generator creates or stamps.
                '            SSMA_TimeStamp/ValidFrom/ValidTo are not house audit columns; they
                '            are left here because a table carrying them already has its own
                '            change-tracking mechanism (e.g. SQL Server temporal system-versioning)
                '            that this scan would otherwise log as noisy, always-changing values.
                Select Case True
                    Case tdef.Name = "tblAuditLog", _
                         tdef.Name = "tblLongTextBackup", _
                         tdef.Name = "tblAuditLogConfig"
                        isAuditable = False
                    Case fld.Name = AUDIT_CREATED_DATE, fld.Name = AUDIT_CREATED_BY, _
                         fld.Name = AUDIT_MODIFIED_DATE, fld.Name = AUDIT_MODIFIED_BY, _
                         fld.Name = "AccessTS", fld.Name = "SSMA_TimeStamp", _
                         fld.Name = "ValidFrom", fld.Name = "ValidTo"
                        isAuditable = False
                    ' [SCAFFOLD] The table's own primary key. Its value is already on every
                    '            log row in tblAuditLog.PrimaryKey, which is what identifies
                    '            the record; a field row here would store the same value a
                    '            second time (schema Business Rule 5). A developer can flip
                    '            it on afterward if they want that.
                    Case isPK
                        isAuditable = False
                    Case Else
                        isAuditable = bDefaultAuditable
                End Select

                ' [STANDARDS — query-style.md] inline INSERT kept from the working source
                sSql = "INSERT INTO tblAuditLogConfig " & _
                    "(TableName, FieldName, FieldPosition, DataType, IsPrimaryKey, IsAuditable) " & _
                    "VALUES ('" & tdef.Name & "', '" & fld.Name & "', " & fld.OrdinalPosition & _
                    ", " & fld.Type & ", " & isPK & ", " & isAuditable & ")"
                db.Execute sSql, dbFailOnError
                lRowCount = lRowCount + 1
            Next fld
        End If
    Next tdef

    sReport = "Table list built: " & lRowCount & " field row(s) written to tblAuditLogConfig." & _
        vbCrLf & "Open tblAuditLogConfig and check the IsAuditable switches before you run " & _
        "the next step."

Cleanup:
    Set pkField = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set tdef = Nothing
    Set db = Nothing
    Two_PopulateConfigTable = sReport
    ' [SCAFFOLD] One message, whatever happened — the success text or the error text, never
    '            both. Building the report first and showing it here is what prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    bFailed = True
    sReport = "ERROR populating config: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### CheckAuditReadiness — `Public Function` → `String` (safety check — run before step 3, required for Path B)

A read-only check you can run any time, at no risk — it doesn't change anything. It looks at each
table you might track and tells you whether this system will actually work on it: **every table
needs one, single number field as its primary key, set to auto-number** (schema Business Rule 4).
Most tables you design yourself already look like this. Older or borrowed tables sometimes
don't — a table with no primary key set, one that uses two or more fields together as its key,
or one that uses a text code instead of a number, will not work with this system as-is.

Run this after `Two_PopulateConfigTable` and before `Three_GenerateAllAuditDataMacros`. This
step is **required for Path B**, since a database you didn't design the audit system around is
far more likely to have a table shaped this way. It's optional on Path A, where the sample
tables are already known to be shaped correctly.

If a table isn't ready, you have two choices: fix that table's primary key, or leave it out —
open `tblAuditLogConfig` and switch `IsAuditable` to No for every row belonging to that table.

**Returns the same report it shows in the message box, as text.** Called as a bare statement
(`CheckAuditReadiness`) it behaves exactly as before — it pops the `MsgBox` for a person sitting
at the keyboard. Called as `sResult = CheckAuditReadiness(True)`, the same text comes back as a
`String` and no dialog is shown — for a script, a test harness, or an AI assistant facilitating a
build to read directly, rather than needing to add its own throwaway diagnostic to see the
verdict. **The `True` is what suppresses the dialog**, not the assignment: VBA cannot tell whether
a procedure was called as a function or as a statement.

```vba
Public Function CheckAuditReadiness(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Read-only pre-flight check. Looks at the real table definitions (not the
    '            config table) so a multi-field primary key is never missed. Run after
    '            Two_PopulateConfigTable and before Three_GenerateAllAuditDataMacros —
    '            required on Path B, where tables were not designed around this system.
    '            Returns the report as a String. Pass bSilent:=True to suppress the MsgBox so
    '            an automated caller is never left waiting on a dialog nobody can dismiss.
    Dim db As DAO.Database
    Dim tdef As DAO.TableDef
    Dim idx As DAO.Index
    Dim lPkFieldCount As Long
    Dim lPkFieldType As Long
    Dim lProblemCount As Long
    Dim sMsg As String
    Dim sReport As String

    On Error GoTo errHandler
    Set db = CurrentDb
    lProblemCount = 0
    sMsg = ""

    For Each tdef In db.TableDefs
        ' [BUSINESS LOGIC — scan boundary] Same tbl-or-tlkp-but-not-tmp test as
        ' Two_PopulateConfigTable; keep the two in sync if you change one.
        If (Left(tdef.Name, 3) = "tbl" Or Left(tdef.Name, 4) = "tlkp") _
            And Left(tdef.Name, 3) <> "tmp" _
            And tdef.Name <> "tblAuditLog" _
            And tdef.Name <> "tblLongTextBackup" _
            And tdef.Name <> "tblAuditLogConfig" Then

            lPkFieldCount = 0
            lPkFieldType = -1
            For Each idx In tdef.Indexes
                If idx.Primary Then
                    lPkFieldCount = idx.Fields.Count
                    If lPkFieldCount = 1 Then
                        lPkFieldType = tdef.Fields(idx.Fields(0).Name).Type
                    End If
                End If
            Next idx

            If lPkFieldCount = 0 Then
                lProblemCount = lProblemCount + 1
                sMsg = sMsg & tdef.Name & " — no primary key is set" & vbCrLf
                Debug.Print tdef.Name & ": NOT READY — no primary key"
            ElseIf lPkFieldCount > 1 Then
                lProblemCount = lProblemCount + 1
                sMsg = sMsg & tdef.Name & " — primary key uses more than one field" & vbCrLf
                Debug.Print tdef.Name & ": NOT READY — primary key has " & lPkFieldCount & " fields"
            ElseIf lPkFieldType <> dbLong Then
                lProblemCount = lProblemCount + 1
                sMsg = sMsg & tdef.Name & " — primary key is not an AutoNumber/Long Number field" & vbCrLf
                Debug.Print tdef.Name & ": NOT READY — primary key type is " & lPkFieldType
            Else
                Debug.Print tdef.Name & ": ready"
            End If
        End If
    Next tdef

    If lProblemCount = 0 Then
        sReport = "Every table checked is ready — each has one auto-number primary key. " & _
            "Safe to run Three_GenerateAllAuditDataMacros."
    Else
        sReport = lProblemCount & " table(s) are NOT ready yet:" & vbCrLf & vbCrLf & sMsg & vbCrLf & _
            "This system only works on tables with one auto-number (or plain number) " & _
            "primary key field. Either fix that table's primary key, or leave it out — set " & _
            "IsAuditable to No for all of that table's rows in tblAuditLogConfig — before you " & _
            "run Three_GenerateAllAuditDataMacros."
    End If

    If Not bSilent Then MsgBox sReport, IIf(lProblemCount = 0, vbInformation, vbExclamation)
    CheckAuditReadiness = sReport

Cleanup:
    Set idx = Nothing
    Set tdef = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    CheckAuditReadiness = "Error checking audit readiness: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox CheckAuditReadiness, vbCritical
    Resume Cleanup
    Resume
End Function
```

### Three_GenerateAllAuditDataMacros — `Public Function` → `String` (setup step 3)

Reads the (reviewed) config, groups fields by table, and calls `CreateAllDataMacros` for each.
Re-runnable: reloading a table's macro XML replaces what was there (schema Business Rule 7).

**Returns a per-table report as text**, in addition to the summary `MsgBox` — one line per table
(`OK`, `SKIPPED`, or `ERROR: ...`), the same detail `CreateAllDataMacros` used to only send to
`Debug.Print`. Call it as `sResult = Three_GenerateAllAuditDataMacros(True)` to read that report
directly with no dialog — a script or an AI assistant facilitating the build no longer has to add
its own diagnostic wrapper to see which tables actually succeeded. `bSilent` is passed down into
`CreateAllDataMacros`, so a per-table error can't strand an automated caller either.

```vba
Public Function Three_GenerateAllAuditDataMacros(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Generate and attach audit Data Macros for every configured table.
    '            Returns a per-table report so a caller — human or automated — can see exactly
    '            what happened to each table. Pass bSilent:=True to suppress every MsgBox,
    '            including the per-table ones raised inside CreateAllDataMacros.
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dictTables As Object          ' Scripting.Dictionary, late-bound
    Dim sTableName As String
    Dim sFieldName As String
    Dim lFieldDataType As Long
    Dim bFieldIsPK As Boolean
    Dim bFieldIsAuditable As Boolean
    Dim fieldList As Collection
    Dim fieldInfo As Variant
    Dim sTempPath As String
    Dim lTableCount As Long
    Dim vCurrentTable As Variant
    Dim sTableResult As String
    Dim sReport As String

    On Error GoTo errHandler
    Set db = CurrentDb
    Set dictTables = CreateObject("Scripting.Dictionary")
    sReport = ""

    ' Read configuration and group fields by table, in field order. All rows come along —
    ' the PK row is needed for plumbing even if its IsAuditable flag was flipped; the
    ' builders skip non-auditable fields when emitting audit actions.
    Set rs = db.OpenRecordset( _
        "SELECT TableName, FieldName, DataType, IsPrimaryKey, IsAuditable " & _
        "FROM tblAuditLogConfig ORDER BY TableName, FieldPosition", dbOpenSnapshot)

    Do While Not rs.EOF
        sTableName = Nz(rs!TableName, "")
        sFieldName = Nz(rs!FieldName, "")
        lFieldDataType = Nz(rs!DataType, 0)
        bFieldIsPK = Nz(rs!IsPrimaryKey, False)
        bFieldIsAuditable = Nz(rs!IsAuditable, False)

        ' [SCAFFOLD] Hard guard above the flags: the system tables never get macros
        '            (schema Business Rule 5), whatever their config rows say.
        If sTableName <> "tblAuditLog" _
            And sTableName <> "tblLongTextBackup" _
            And sTableName <> "tblAuditLogConfig" _
            And sTableName <> "" And sFieldName <> "" Then

            If Not dictTables.Exists(sTableName) Then
                Set fieldList = New Collection
                dictTables.Add sTableName, fieldList
            Else
                Set fieldList = dictTables(sTableName)
            End If
            ' Field info as array: (FieldName, DataType, IsPrimaryKey, IsAuditable)
            fieldInfo = Array(sFieldName, lFieldDataType, bFieldIsPK, bFieldIsAuditable)
            fieldList.Add fieldInfo
        End If
        rs.MoveNext
    Loop
    rs.Close

    sTempPath = Environ("TEMP") & "\"

    lTableCount = 0
    For Each vCurrentTable In dictTables.Keys
        sTableName = CStr(vCurrentTable)
        Set fieldList = dictTables(sTableName)
        sTableResult = CreateAllDataMacros(sTableName, fieldList, sTempPath, bSilent)
        sReport = sReport & sTableName & ": " & sTableResult & vbCrLf
        lTableCount = lTableCount + 1
    Next vCurrentTable

    If Not bSilent Then MsgBox "Generated audit data macros for " & lTableCount & " table(s). " & _
        "See the returned report for the per-table result.", vbInformation
    Three_GenerateAllAuditDataMacros = sReport

Cleanup:
    Set rs = Nothing
    Set db = Nothing
    Set dictTables = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Three_GenerateAllAuditDataMacros = sReport & "ERROR: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
    Resume
End Function
```

### CreateAllDataMacros — `Private Function` → `String`

Builds one table's macro XML into a single document, writes it UTF-16, and loads it with
`LoadFromText acTableDataMacro`. **How many macros a table gets varies** (schema Business Rule 2) —
the three After macros whenever anything is being audited, BeforeChange whenever the table carries
the house audit columns or an auditable Long Text field, BeforeDelete only for Long Text.

**The insert-blocking trap this guards against.** Before this revision, a table with no auditable
fields was skipped entirely, which was harmless when the generator only ever wrote audit macros —
the table simply had no audit trail. Now that the same pass also writes the **audit-column stamping**
macro, skipping the table would leave it with no stamping macro at all. `CreatedDate` and `CreatedBy`
are `Required`, and no column default can reach the Windows username, so **that table would reject
every insert** — nobody could add a record to it through any interface. Switching auditing off for a
table is the ordinary Path B workflow, so it must never break writing to that table. The audit
actions are skipped; the stamping macro is still written. Only a table needing neither is skipped.

**Backs up a table's existing macros before replacing them.** `LoadFromText` replaces a table's
entire Data Macro set — it does not merge. The commonest case of that used to be this generator
destroying the house stamping macro, which is exactly why the two are now generated together. What
remains is any **other** Data Macro a table carries — business logic written for reasons unrelated to
this system, which a Path B table may well have. Before loading, this checks `MSysObjects` for an
existing macro set and, if it finds one, exports it to a timestamped backup file first — the same
technique `BackupAndRemoveAllDataMacros` uses to detect a table's macros. It does **not** merge the
old logic into the new macros; it only makes sure nothing is destroyed without a copy and a
plain-language warning first. Re-implementing anything lost is the developer's call.

**`DataMacroBackups\` grows every time you regenerate, and nothing prunes it.** Once a table carries
macros, *every* subsequent run backs it up again — so a schema you regenerate ten times leaves ten
files per table beside the database. That is deliberate: throwing away the only copy of a macro set
to save a few kilobytes is the wrong trade. But it does mean the folder is yours to clear out, and
the newest file for a table is the one that matters.

```vba
Private Function CreateAllDataMacros(sTableName As String, fieldList As Collection, sTempPath As String, Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Generate one table's Data Macros as a SINGLE XML document and attach it.
    '            How many macros that is depends on two independent things (schema
    '            Business Rule 2):
    '              - the 3 After macros, whenever the table has any auditable field;
    '              - BeforeChange, whenever the table carries the house audit columns OR
    '                has an auditable Long Text field (that one macro does both jobs);
    '              - BeforeDelete, only for an auditable Long Text field.
    '            Returns a one-line status ("OK ...", "SKIPPED ...", or "ERROR: ...") so the
    '            caller can report per-table results without relying on Debug.Print alone.
    Dim db As DAO.Database
    Dim rsCheck As DAO.Recordset
    Dim sXmlContent As String
    Dim sBeforeChange As String
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim txtFile As Object
    Dim sFilePath As String
    Dim sPrimaryKeyField As String
    Dim fieldInfo As Variant
    Dim bHasLongText As Boolean
    Dim lAuditableCount As Long
    Dim lMacroCount As Long
    Dim sWhatWasBuilt As String
    Dim bHadExistingMacros As Boolean
    Dim sBackupFolder As String
    Dim sBackupFile As String
    Dim sBackupNote As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' Find the PK field, count auditable fields, and detect auditable Long Text
    ' (schema Business Rules 2, 4, 5)
    bHasLongText = False
    lAuditableCount = 0
    For Each fieldInfo In fieldList
        If fieldInfo(2) = True Then sPrimaryKeyField = fieldInfo(0)
        If fieldInfo(3) = True Then
            lAuditableCount = lAuditableCount + 1
            If fieldInfo(1) = dbMemo Then bHasLongText = True
        End If
    Next fieldInfo

    ' BeforeChange carries the audit-column stamping as well as Long Text staging, so it is
    ' built for every table; it comes back "" only when neither job applies to this one.
    sBeforeChange = BuildBeforeChangeMacro(sTableName, fieldList, sPrimaryKeyField)

    ' [SCAFFOLD] THE TRAP THIS GUARDS AGAINST. A table with nothing auditable used to be
    '            skipped outright, which was harmless when this generator only wrote audit
    '            macros. Now that the same pass also writes the audit-column stamping,
    '            skipping the table would leave it with NO stamping macro — and CreatedDate
    '            and CreatedBy are Required, with no default able to reach the username. That
    '            table would then reject EVERY insert. Turning auditing off for a table is the
    '            normal Path B workflow, so it must never break writing to that table: the
    '            audit actions are skipped, the stamping macro is still emitted. Only a table
    '            that needs neither is skipped entirely.
    If lAuditableCount = 0 And Len(sBeforeChange) = 0 Then
        Debug.Print "  - Skipped (nothing to audit, and no audit columns to stamp)"
        CreateAllDataMacros = "SKIPPED (nothing to audit, no audit columns to stamp)"
        Exit Function
    End If

    ' [SCAFFOLD] Safety net for Path B: LoadFromText replaces a table's WHOLE macro set. If
    '            this table already has one (e.g. the house audit-column stamping macro),
    '            back it up before it's overwritten — see the note above this code block.
    bHadExistingMacros = False
    Set rsCheck = db.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Name='" & sTableName & "' AND Type=1 AND Not IsNull(LvExtra)", _
        dbOpenSnapshot)
    bHadExistingMacros = Not rsCheck.EOF
    rsCheck.Close
    Set rsCheck = Nothing

    sBackupNote = ""
    If bHadExistingMacros Then
        sBackupFolder = CurrentProject.Path & "\DataMacroBackups\"
        If Dir(sBackupFolder, vbDirectory) = "" Then MkDir sBackupFolder
        sBackupFile = sBackupFolder & sTableName & "_PreAuditBackup_" & Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Application.SaveAsText acTableDataMacro, sTableName, sBackupFile
        sBackupNote = " (existing macros backed up to " & sBackupFile & " before replacing)"
        Debug.Print "  - " & sTableName & " already had Data Macros — backed up to " & sBackupFile
    End If

    ' One XML document carrying all of this table's macros
    sXmlContent = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    sXmlContent = sXmlContent & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application"">"

    ' The three After macros are the audit trail itself — no auditable fields, none needed.
    lMacroCount = 0
    If lAuditableCount > 0 Then
        sXmlContent = sXmlContent & BuildAfterInsertMacro(sTableName, fieldList, sPrimaryKeyField)
        sXmlContent = sXmlContent & BuildAfterUpdateMacro(sTableName, fieldList, sPrimaryKeyField)
        sXmlContent = sXmlContent & BuildAfterDeleteMacro(sTableName, fieldList, sPrimaryKeyField)
        lMacroCount = 3
    End If

    If Len(sBeforeChange) > 0 Then
        sXmlContent = sXmlContent & sBeforeChange
        lMacroCount = lMacroCount + 1
    End If

    ' BeforeDelete stages Long Text values for the AfterDelete log row, so it is only
    ' wanted when that log row is actually going to be written.
    If bHasLongText Then
        sXmlContent = sXmlContent & BuildBeforeDeleteMacro(sTableName, fieldList, sPrimaryKeyField)
        lMacroCount = lMacroCount + 1
    End If

    sXmlContent = sXmlContent & "</DataMacros>"

    ' [SCAFFOLD] Write UTF-16 (CreateTextFile third argument True) — LoadFromText requires it —
    '            then load with the table held open in design view so the save sticks.
    sFilePath = sTempPath & sTableName & "_DataMacros.xml"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(sFilePath, True, True)
    txtFile.Write sXmlContent
    txtFile.Close
    Set txtFile = Nothing

    DoCmd.OpenTable sTableName, acViewDesign, acHidden
    Application.LoadFromText acTableDataMacro, sTableName, sFilePath
    DoCmd.Close acTable, sTableName, acSaveYes

    fso.DeleteFile sFilePath

    ' [SCAFFOLD] Report what was actually built rather than a fixed count — the count varies
    '            by table now, and a reader comparing tables needs to see why.
    If lAuditableCount = 0 Then
        sWhatWasBuilt = "BeforeChange only — audit-column stamping, no audit trail " & _
            "(this table has no auditable fields; inserts still work)"
    ElseIf bHasLongText Then
        sWhatWasBuilt = "3 After + BeforeChange (stamping and Long Text staging) + BeforeDelete"
    ElseIf Len(sBeforeChange) > 0 Then
        sWhatWasBuilt = "3 After + BeforeChange (stamping)"
    Else
        sWhatWasBuilt = "3 After (this table carries no house audit columns, so nothing to stamp)"
    End If

    Debug.Print "  - " & lMacroCount & " data macro(s) created: " & sWhatWasBuilt
    CreateAllDataMacros = "OK — " & lMacroCount & " macro(s): " & sWhatWasBuilt & sBackupNote

Cleanup:
    Set rsCheck = Nothing
    Set txtFile = Nothing
    Set fso = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    CreateAllDataMacros = "ERROR: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox "Error creating macros for " & sTableName & ": " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
    Resume
End Function
```

### BuildAfterInsertMacro — `Private Function` → `String`

Emits the AfterInsert `<DataMacro>` fragment: one `tblAuditLog` row per configured field, marking
`OldValue` as `[NEW RECORD]`.

```vba
Private Function BuildAfterInsertMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterInsert: log every configured field of the new row.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String

    sXml = "<DataMacro Event=""AfterInsert""><Statements>"

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        sFieldName = fieldInfo(0)

        sXml = sXml & "<CreateRecord>"
        sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        sXml = sXml & "<Statements>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sXml = sXml & "</Action>"

        If sPrimaryKeyField <> "" Then
            sXml = sXml & "<Action Name=""SetField"">"
            sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sPrimaryKeyField & "]</Argument>"
            sXml = sXml & "</Action>"
        End If

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
        sXml = sXml & "</Action>"

        ' OperationType (schema Business Rule 6); OldValue stays Null on an insert
        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""Insert""</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
        sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sFieldName & "]</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
        sXml = sXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] Identity — AuditUser() is the PREFERRED choice.
        '            The BeforeChange stamping macro calls it because audit-columns.md does,
        '            so the log names the same person the stamped row does.
        '            CurrentUser() is a named Extra Option: switch this and the three other
        '            ChangedBy sites together (AfterUpdate, AfterDelete, and
        '            BackupLongTextFieldsDM), and switch the stamping in your standards
        '            layer too — half a change puts two names on one edit.
        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        sXml = sXml & "<Argument Name=""Value"">AuditUser()</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "</Statements></CreateRecord>"
      End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterInsertMacro = sXml
End Function
```

### BuildAfterUpdateMacro — `Private Function` → `String`

Emits the AfterUpdate fragment. Ordinary fields: log only when the value actually changed
(`GetComparisonExpression`), reading the old value from `[Old]`. Long Text fields: always log
(schema Business Rule 6), retrieving the old value from `tblLongTextBackup` via `LookUpRecord` —
the BeforeChange macro put it there (schema Business Rule 3).

```vba
Private Function BuildAfterUpdateMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterUpdate: one conditional block per non-PK field; Long Text goes
    '            through the LookUpRecord retrieval path.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim lFldType As Long
    Dim bIsLongText As Boolean

    sXml = "<DataMacro Event=""AfterUpdate""><Statements>"

    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        lFldType = fieldInfo(1)
        bIsLongText = (lFldType = dbMemo)

        ' Auditable fields only (schema Business Rule 5); PK never changes, skip it
        If fieldInfo(3) = True And sFieldName <> sPrimaryKeyField Then
            sXml = sXml & "<ConditionalBlock><If>"
            sXml = sXml & "<Condition>" & GetComparisonExpression(sTableName, sFieldName, lFldType) & "</Condition>"
            sXml = sXml & "<Statements>"

            ' Long Text: fetch the backed-up old value (schema Business Rule 3)
            If bIsLongText Then
                sXml = sXml & "<LookUpRecord>"
                sXml = sXml & "<Data Alias=""BackupRec"">"
                sXml = sXml & "<Reference>tblLongTextBackup</Reference>"
                sXml = sXml & "<WhereCondition>"
                sXml = sXml & "[tblLongTextBackup].[TableName]=""" & sTableName & """ And "
                sXml = sXml & "[tblLongTextBackup].[PrimaryKey]=[" & sTableName & "].[" & sPrimaryKeyField & "] And "
                sXml = sXml & "[tblLongTextBackup].[FieldName]=""" & sFieldName & """"
                sXml = sXml & "</WhereCondition>"
                sXml = sXml & "</Data>"
                sXml = sXml & "<Statements>"
            End If

            sXml = sXml & "<CreateRecord>"
            If bIsLongText Then
                sXml = sXml & "<Data><Reference>tblAuditLog</Reference></Data>"
            Else
                sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
            End If
            sXml = sXml & "<Statements>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
            sXml = sXml & "</Action>"

            If sPrimaryKeyField <> "" Then
                sXml = sXml & "<Action Name=""SetField"">"
                If bIsLongText Then
                    sXml = sXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
                Else
                    sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
                End If
                sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sPrimaryKeyField & "]</Argument>"
                sXml = sXml & "</Action>"
            End If

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
            sXml = sXml & "</Action>"

            ' OperationType (schema Business Rule 6)
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""Update""</Argument>"
            sXml = sXml & "</Action>"

            ' OldValue — from the backup for Long Text, from [Old] otherwise
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
                sXml = sXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
                sXml = sXml & "<Argument Name=""Value"">[Old].[" & sFieldName & "]</Argument>"
            End If
            sXml = sXml & "</Action>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sFieldName & "]</Argument>"
            sXml = sXml & "</Action>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
            sXml = sXml & "</Action>"

            ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
            ' [STANDARDS / schema Business Rule 9] AuditUser() preferred choice — same person as the
            '            stamp. CurrentUser() is an Extra Option; change all four ChangedBy
            '            sites together, never just one.
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">AuditUser()</Argument>"
            sXml = sXml & "</Action>"

            sXml = sXml & "</Statements></CreateRecord>"

            If bIsLongText Then
                sXml = sXml & "</Statements></LookUpRecord>"
            End If

            sXml = sXml & "</Statements></If></ConditionalBlock>"
        End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterUpdateMacro = sXml
End Function
```

### BuildAfterDeleteMacro — `Private Function` → `String`

Emits the AfterDelete fragment: one `tblAuditLog` row per field, `NewValue` marked `[DELETED]`,
old values read from `[Old]` — except Long Text, retrieved from the backup the BeforeDelete macro
staged.

```vba
Private Function BuildAfterDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterDelete: log every configured field of the deleted row.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim lFldType As Long
    Dim bIsLongText As Boolean

    sXml = "<DataMacro Event=""AfterDelete""><Statements>"

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        sFieldName = fieldInfo(0)
        lFldType = fieldInfo(1)
        bIsLongText = (lFldType = dbMemo)

        If bIsLongText Then
            sXml = sXml & "<LookUpRecord>"
            sXml = sXml & "<Data Alias=""BackupRec"">"
            sXml = sXml & "<Reference>tblLongTextBackup</Reference>"
            sXml = sXml & "<WhereCondition>"
            sXml = sXml & "[tblLongTextBackup].[TableName]=""" & sTableName & """ And "
            sXml = sXml & "[tblLongTextBackup].[PrimaryKey]=[Old].[" & sPrimaryKeyField & "] And "
            sXml = sXml & "[tblLongTextBackup].[FieldName]=""" & sFieldName & """"
            sXml = sXml & "</WhereCondition>"
            sXml = sXml & "</Data>"
            sXml = sXml & "<Statements>"
        End If

        sXml = sXml & "<CreateRecord>"
        If bIsLongText Then
            sXml = sXml & "<Data><Reference>tblAuditLog</Reference></Data>"
        Else
            sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        End If
        sXml = sXml & "<Statements>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sXml = sXml & "</Action>"

        If sPrimaryKeyField <> "" Then
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">[Old].[" & sPrimaryKeyField & "]</Argument>"
            sXml = sXml & "</Action>"
        End If

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
        sXml = sXml & "</Action>"

        ' OperationType (schema Business Rule 6); NewValue stays Null on a delete
        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""Delete""</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[Old].[" & sFieldName & "]</Argument>"
        End If
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
        sXml = sXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        End If
        ' [STANDARDS / schema Business Rule 9] AuditUser() — same person as the stamp.
        sXml = sXml & "<Argument Name=""Value"">AuditUser()</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "</Statements></CreateRecord>"

        If bIsLongText Then
            sXml = sXml & "</Statements></LookUpRecord>"
        End If
      End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterDeleteMacro = sXml
End Function
```

### BuildBeforeChangeMacro — `Private Function` → `String`

**This one macro does two jobs, and it has to.** `LoadFromText` replaces a table's *entire* Data
Macro set rather than adding to it, so the two things that both need to happen on a Before Change —
stamping the house audit columns, and staging Long Text values before they're overwritten — cannot
be loaded as separate macros. The second load would silently delete the first. They are generated
together instead:

1. **Stamp the house audit columns** (`standards/audit-columns.md`), on any table that carries them.
   Without this, a `Required` `CreatedBy` with nothing able to fill it **rejects every insert**.
2. **Stage Long Text old values** ahead of an update, by calling `BackupLongTextFieldsDM` for each
   Long Text field — the workaround for a Data Macro being unable to read `[Old].[LongTextField]`.

Both key on the same test, `IsNull([Old].[PK])` — true on an insert, false on an update — so they
merge into one conditional block with no conflict.

The function looks at what the table actually has and emits only what applies. A table carrying
neither the audit columns nor a Long Text field gets an **empty string back**, meaning no Before
Change macro at all — that table's set is just the three After macros.

**Why the branches are assembled before the XML:** a branch with no actions in it would produce an
empty `<Statements></Statements>` block, and we have no evidence Access accepts one. Rather than find
out on someone's live table, each branch's actions are built first; a branch with nothing to do is
left out, and if only the update branch has work the condition is inverted (`Not IsNull(...)`) so an
ordinary If-without-Else is emitted instead.

```vba
Private Function BuildBeforeChangeMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD + STANDARDS — audit-columns.md] BeforeChange carries TWO jobs, because
    '            LoadFromText replaces a table's whole macro set and they cannot be
    '            loaded separately without one destroying the other:
    '              1. Stamp the house audit columns (any table that carries them).
    '              2. Stage Long Text old values ahead of an update (Long Text tables).
    '            Both key on the same discriminator, IsNull([Old].[PK]).
    '            Returns "" when the table needs neither — no BeforeChange macro at all.
    Dim sXml As String
    Dim sInsertActions As String
    Dim sUpdateActions As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim bHasLongText As Boolean
    Dim bHasCreatedDate As Boolean
    Dim bHasCreatedBy As Boolean
    Dim bHasModifiedDate As Boolean
    Dim bHasModifiedBy As Boolean

    ' What does this table actually carry? The audit columns arrive with IsAuditable
    ' False (they are always seeded off) — presence is what decides stamping, not the flag.
    bHasLongText = False
    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        If StrComp(sFieldName, AUDIT_CREATED_DATE, vbTextCompare) = 0 Then bHasCreatedDate = True
        If StrComp(sFieldName, AUDIT_CREATED_BY, vbTextCompare) = 0 Then bHasCreatedBy = True
        If StrComp(sFieldName, AUDIT_MODIFIED_DATE, vbTextCompare) = 0 Then bHasModifiedDate = True
        If StrComp(sFieldName, AUDIT_MODIFIED_BY, vbTextCompare) = 0 Then bHasModifiedBy = True
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then bHasLongText = True
    Next fieldInfo

    ' ---- INSERT branch: stamp Created*, never touch Modified* ----
    If bHasCreatedDate Then sInsertActions = sInsertActions & AuditSetField(AUDIT_CREATED_DATE, "Now()")
    If bHasCreatedBy Then sInsertActions = sInsertActions & AuditSetField(AUDIT_CREATED_BY, "AuditUser()")

    If bHasLongText Then
        ' Nothing to back up on an insert — there is no prior value. The marker is set
        ' anyway so lngPKValue is defined on both paths.
        sInsertActions = sInsertActions & "<Action Name=""SetLocalVar"">"
        sInsertActions = sInsertActions & "<Argument Name=""Name"">lngPKValue</Argument>"
        sInsertActions = sInsertActions & "<Argument Name=""Value"">0</Argument>"
        sInsertActions = sInsertActions & "</Action>"
    End If

    ' ---- UPDATE branch: stamp Modified*, leave Created* frozen ----
    If bHasModifiedDate Then sUpdateActions = sUpdateActions & AuditSetField(AUDIT_MODIFIED_DATE, "Now()")
    If bHasModifiedBy Then sUpdateActions = sUpdateActions & AuditSetField(AUDIT_MODIFIED_BY, "AuditUser()")

    If bHasLongText Then
        sUpdateActions = sUpdateActions & "<Action Name=""SetLocalVar"">"
        sUpdateActions = sUpdateActions & "<Argument Name=""Name"">lngPKValue</Argument>"
        sUpdateActions = sUpdateActions & "<Argument Name=""Value"">=[" & sPrimaryKeyField & "]</Argument>"
        sUpdateActions = sUpdateActions & "</Action>"

        sUpdateActions = sUpdateActions & "<Action Name=""SetLocalVar"">"
        sUpdateActions = sUpdateActions & "<Argument Name=""Name"">strTableName</Argument>"
        sUpdateActions = sUpdateActions & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sUpdateActions = sUpdateActions & "</Action>"

        ' One backup call per Long Text field — a data macro CAN call a public VBA
        ' function in the same accdb; that is the hinge of the whole hybrid method.
        For Each fieldInfo In fieldList
            sFieldName = fieldInfo(0)
            If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
                sUpdateActions = sUpdateActions & "<Action Name=""SetLocalVar"">"
                sUpdateActions = sUpdateActions & "<Argument Name=""Name"">varLongTextBackup</Argument>"
                sUpdateActions = sUpdateActions & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & sFieldName & """)</Argument>"
                sUpdateActions = sUpdateActions & "</Action>"
            End If
        Next fieldInfo
    End If

    ' Neither job applies to this table
    If Len(sInsertActions) = 0 And Len(sUpdateActions) = 0 Then
        BuildBeforeChangeMacro = ""
        Exit Function
    End If

    ' [SCAFFOLD] Emit only branches that have actions. An empty <Statements></Statements>
    '            is untested against Access, and a live table is the wrong place to find
    '            out — so when only the update branch has work, invert the condition and
    '            emit a plain If with no Else.
    sXml = "<DataMacro Event=""BeforeChange""><Statements><ConditionalBlock>"

    If Len(sInsertActions) > 0 Then
        sXml = sXml & "<If><Condition>IsNull([Old].[" & sPrimaryKeyField & "])</Condition>"
        sXml = sXml & "<Statements>" & sInsertActions & "</Statements></If>"
        If Len(sUpdateActions) > 0 Then
            sXml = sXml & "<Else><Statements>" & sUpdateActions & "</Statements></Else>"
        End If
    Else
        sXml = sXml & "<If><Condition>Not IsNull([Old].[" & sPrimaryKeyField & "])</Condition>"
        sXml = sXml & "<Statements>" & sUpdateActions & "</Statements></If>"
    End If

    sXml = sXml & "</ConditionalBlock></Statements></DataMacro>"

    BuildBeforeChangeMacro = sXml
End Function
```

### BuildBeforeDeleteMacro — `Private Function` → `String`

Emits the BeforeDelete fragment for a Long Text table: captures the PK, then stages each Long
Text value through `BackupLongTextFieldsDM` so the AfterDelete macro can log it. Returns an empty
string when the table has no Long Text fields.

```vba
Private Function BuildBeforeDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] BeforeDelete: stage Long Text values ahead of a delete
    '            (schema Business Rules 2-3).
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim bHasLongText As Boolean

    bHasLongText = False
    For Each fieldInfo In fieldList
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            bHasLongText = True
            Exit For
        End If
    Next fieldInfo
    If Not bHasLongText Then
        BuildBeforeDeleteMacro = ""
        Exit Function
    End If

    sXml = "<DataMacro Event=""BeforeDelete""><Statements>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">lngPKValue</Argument>"
    sXml = sXml & "<Argument Name=""Value"">=[" & sPrimaryKeyField & "]</Argument>"
    sXml = sXml & "</Action>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">strTableName</Argument>"
    sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
    sXml = sXml & "</Action>"

    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            sXml = sXml & "<Action Name=""SetLocalVar"">"
            sXml = sXml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
            sXml = sXml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & sFieldName & """)</Argument>"
            sXml = sXml & "</Action>"
        End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"

    BuildBeforeDeleteMacro = sXml
End Function
```

### AuditSetField — `Private Function` → `String`

One `SetField` action, in macro XML. `BuildBeforeChangeMacro` calls it four times — twice on the
insert branch and twice on the update branch — to stamp the house audit columns. It exists so the
XML for a stamp appears once rather than four times, which is what makes the stamping easy to read
against `standards/audit-columns.md`.

```vba
Private Function AuditSetField(ByVal sField As String, ByVal sValue As String) As String
    ' [STANDARDS — audit-columns.md] One SetField action for a stamped audit column.
    AuditSetField = "<Action Name=""SetField""><Argument Name=""Field"">" & sField & _
        "</Argument><Argument Name=""Value"">" & sValue & "</Argument></Action>"
End Function
```

### GetComparisonExpression — `Private Function` → `String`

The change test the AfterUpdate macro embeds per field. Ordinary fields compare old and new with
`StrComp` over `Nz`-wrapped values; Long Text returns `True` — always log — because the macro
cannot read the old value to compare (schema Business Rule 6).

**The not-equal operator is built from `Chr(38)`, not typed as a literal `&lt;&gt;`.** A real
build hit error 3870 ("Microsoft Access cannot interpret the text you are pasting as a data
macro") on every table because the VBA source was imported through a tool that **un-escapes HTML
entities on the way in** — a literal `&lt;&gt;` in this function's source became a raw `<>` once
imported, which breaks the macro XML the moment it's loaded (the `<` is read as an opening tag).
Assembling the entity from character codes at runtime means no literal `&`-entity ever exists in
the source for an importer to touch — the string is only ever `&lt;&gt;` in memory, never in text
anyone's tool re-reads. See `templates/_materialization.md` for the general rule this follows.

```vba
Private Function GetComparisonExpression(sTableName As String, sFieldName As String, lFldType As Long) As String
    ' [SCAFFOLD] Per-type change test for the AfterUpdate conditional block.
    Dim sNotEqual As String
    Select Case lFldType
        Case dbMemo
            ' Long Text: always log (cannot compare the old value in-macro)
            GetComparisonExpression = "True"
        Case Else
            ' Built from Chr(38), not typed as a literal &lt;&gt; — see the note above this
            ' block. This is XML for "<>"; the expression lives inside <Condition>...</Condition>.
            sNotEqual = Chr(38) & "lt;" & Chr(38) & "gt;"
            GetComparisonExpression = "StrComp(NZ([" & sTableName & "].[" & sFieldName & "],""""),NZ([Old].[" & sFieldName & "],""""),0)" & sNotEqual & "0"
    End Select
End Function
```

### AuditUser — `Public Function` → `String` (module `modAuditLongText` — back end AND front end)

**The single most important four lines in this template, and the easiest to leave out.** Every
generated stamping macro calls `AuditUser()` on **every** table — Long Text or not — to fill
`CreatedBy` and `ModifiedBy`. Those columns are `Required`, so a table whose macro cannot resolve
this function **rejects every insert**, with an error that names the function rather than the cause.

`modAuditLongText`'s name undersells it. A database with no Long Text field anywhere still needs
this module, in the back end **and in every front end**, because a macro fired by an edit through a
linked table looks for the function in *that person's* front end.

Returns `"Unknown"` rather than an empty string (schema Business Rule 9): `Environ$("USERNAME")`
comes back empty in some contexts — a scheduled task, a service account, a locked-down profile —
and because `CreatedBy` is `Required`, an empty result would block the insert outright. A row naming
an unknown user is a record; a refused write is not.

```vba
Public Function AuditUser() As String
    ' [STANDARDS — audit-columns.md] The identity every generated stamping macro calls, as
    '            =AuditUser(). Never returns an empty string: CreatedBy is Required, and an
    '            empty value would block the insert (schema Business Rule 9).
    '            CurrentUser() is a named alternative — see Extra Options — but it returns
    '            "Admin" for everyone unless workgroup security is in use.
    AuditUser = Environ$("USERNAME")
    If Len(AuditUser) = 0 Then AuditUser = "Unknown"
End Function
```

### BackupLongTextFieldsDM — `Public Function` (module `modAuditLongText` — back end AND front end)

The VBA half of the hybrid Long Text method, called *by the Before macros themselves*. Replaces
any earlier backup for the same table/field/record, reads the current Long Text value, and writes
it to `tblLongTextBackup` for the After macro to retrieve.

**Placement is the trap:** it must exist in the **back end** (so `LoadFromText` resolves the name
when the macros are created) **and in every front end** (a macro fired by a front-end edit
resolves the function in the front end's VBA project). Missing it in either place surfaces as a
data-macro execution error at save time.

```vba
Public Function BackupLongTextFieldsDM(strTableName As String, lngPKValue As Long, strFieldName As String)
    ' [SCAFFOLD] Stage one Long Text field's current value before an update or delete.
    '            Called by the generated BeforeChange / BeforeDelete Data Macros.
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsOldValue As DAO.Recordset
    Dim strPKField As String
    Dim strOldValue As Variant

    On Error GoTo errHandler
    Set db = CurrentDb

    If strTableName = "" Then Exit Function

    ' Replace any earlier backup for this table/field/record
    db.Execute "DELETE FROM tblLongTextBackup WHERE TableName='" & strTableName & _
        "' AND FieldName='" & strFieldName & "' AND PrimaryKey=" & lngPKValue, dbFailOnError

    ' The audited table's PK field name comes from the config
    strPKField = DLookup("FieldName", "tblAuditLogConfig", _
        "TableName='" & strTableName & "' AND IsPrimaryKey=" & True)

    If lngPKValue > 0 Then    ' updates and deletes only — a new record has no old value
        Set rsOldValue = db.OpenRecordset("SELECT " & strFieldName & " FROM " & strTableName & _
            " WHERE " & strPKField & "=" & lngPKValue)
        strOldValue = rsOldValue.Fields(strFieldName).Value
        rsOldValue.Close

        Set rs = db.OpenRecordset("tblLongTextBackup", dbOpenDynaset)
        rs.AddNew
        rs!TableName = strTableName
        rs!PrimaryKey = lngPKValue
        rs!FieldName = strFieldName
        rs!OldValue = strOldValue
        rs!DateChanged = Now()
        ' [STANDARDS / schema Business Rule 9] AuditUser() preferred choice, same as the macros. This
        '            row is the one the After macro reads back, so it is the easiest of the
        '            four sites to miss when applying the CurrentUser() Extra Option — and
        '            missing it puts two names on one edit.
        rs!ChangedBy = AuditUser()
        rs.Update
        rs.Close
    End If

Cleanup:
    On Error Resume Next
    Set rsOldValue = Nothing
    Set rs = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] deliberately quiet: this runs inside a data-macro save;
    '            a MsgBox here would interrupt every user's save. Log if your house pattern
    '            has a silent logger; never block.
    Resume Cleanup
    Resume
End Function
```

### BackupAndRemoveAllDataMacros — `Public Function` → `Boolean` (module `modAuditAdmin` — back end only)

The reset tool for regeneration (schema Business Rule 7): exports every table's current data
macros to timestamped XML backups, then strips them by loading an empty macro document. Run it
before re-running `Three_GenerateAllAuditDataMacros` when the audit scope changes — the backups
double as your archive of prior macro states.

```vba
Public Function BackupAndRemoveAllDataMacros(Optional strBackupPath As String = "", _
                                             Optional bSilent As Boolean = False) As Boolean
    ' [SCAFFOLD] Back up, then remove, the data macros on every table that has them.
    '            Passing bSilent:=True suppresses both message boxes, so a caller with no one
    '            at the keyboard does not hang on a dialog. The Boolean return is the result.
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strTempFile As String
    Dim strBackupFile As String
    Dim intFileNum As Integer
    Dim intMacrosRemoved As Integer

    On Error GoTo errHandler
    Set db = CurrentDb
    intMacrosRemoved = 0

    If strBackupPath = "" Then
        strBackupPath = CurrentProject.Path & "\DataMacroBackups\"
    End If
    If Dir(strBackupPath, vbDirectory) = "" Then
        MkDir strBackupPath
    End If

    ' An empty macro document: loading it replaces (removes) a table's data macros
    strTempFile = Environ("TEMP") & "\BlankDataMacro.xml"
    intFileNum = FreeFile
    Open strTempFile For Output As intFileNum
    Print #intFileNum, "<?xml version=""1.0"" encoding=""UTF-16""?>"
    Print #intFileNum, "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/04/application"">"
    Print #intFileNum, "</DataMacros>"
    Close #intFileNum

    ' Tables with data macros: MSysObjects.LvExtra is non-null for them
    strSQL = "SELECT [Name] FROM MSysObjects " & _
             "WHERE Not IsNull(LvExtra) AND Type = 1 " & _
             "ORDER BY [Name]"
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)

    Do While Not rst.EOF
        Debug.Print "Processing: " & rst!Name

        strBackupFile = strBackupPath & rst!Name & "_DataMacro_" & _
                        Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Application.SaveAsText acTableDataMacro, rst!Name, strBackupFile

        Application.LoadFromText acTableDataMacro, rst!Name, strTempFile
        intMacrosRemoved = intMacrosRemoved + 1
        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
    Kill strTempFile

    If Not bSilent Then MsgBox "Successfully backed up and removed data macros from " & _
           intMacrosRemoved & " tables." & vbCrLf & _
           "Backups saved to: " & strBackupPath, vbInformation, "Data Macros Removed"

    BackupAndRemoveAllDataMacros = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    '            One exit for every error: the recordset above already selects only tables
    '            that have data macros, so there is no "nothing to export" case to skip past,
    '            and skipping one item mid-loop is not something this pattern does.
    If Not bSilent Then MsgBox Err.Number & " Error: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    If Dir(strTempFile) <> "" Then Kill strTempFile
    BackupAndRemoveAllDataMacros = False
    Resume Cleanup
    Resume
End Function
```

### DumpTableMacros — `Public Function` → `String` (module `modAuditVerify` — back end only)

**Check the artifact, not the report.** Every procedure above returns a status line saying what it
did. That line is the generator's own account of itself — if the generator is wrong, the line is
wrong in exactly the same way, and everything still looks fine. This function and `ListMacroEvents`
below read the macros back **off the table**, and show what is actually attached. That is the only
answer capable of contradicting the generator.

`? DumpTableMacros("tblSupportTicket")` in the Immediate window returns the full macro XML — use it
when you need to see the actions *inside* a macro rather than just which macros exist.

It is read-only: it exports the macro set to a temporary file, reads it, and deletes the file.
Nothing is changed. `SaveAsText` writes UTF-16, which is why the file is opened with the `-1`
(TristateTrue) argument — reading it as ANSI returns unusable text.

```vba
Option Compare Database
Option Explicit

Public Function DumpTableMacros(ByVal sTable As String) As String
    ' [SCAFFOLD] Read a table's attached Data Macro set back out, so a build can be
    '            verified against the table itself rather than the generator's own report.
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim txt As Object
    Dim sPath As String
    Dim sOut As String

    On Error GoTo errHandler

    sPath = Environ$("TEMP") & "\" & sTable & "_verify.xml"
    If Dir(sPath) <> "" Then Kill sPath

    Application.SaveAsText acTableDataMacro, sTable, sPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.OpenTextFile(sPath, 1, False, -1)   ' -1 = TristateTrue, i.e. UTF-16
    sOut = txt.ReadAll
    txt.Close

    fso.DeleteFile sPath
    DumpTableMacros = sOut

Cleanup:
    On Error Resume Next
    Set txt = Nothing
    Set fso = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    DumpTableMacros = "ERROR: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### ListMacroEvents — `Public Function` → `String` (module `modAuditVerify` — back end only)

The compact check, and the one to run first. It lists just the macro events attached to a table, in
order, with the size of the set:

```text
? ListMacroEvents("tblSupportTicket")
tblSupportTicket: AfterInsert | AfterUpdate | AfterDelete | BeforeChange | BeforeDelete | (24063 chars)
```

Read it against Business Rule 2 — four macros on an ordinary table carrying the house audit columns,
five when it also has an audited Long Text field. **A missing event means the generator did not do
what its status line said it did**, and that is exactly the discrepancy worth catching before anyone
trusts the audit trail.

```vba
Public Function ListMacroEvents(ByVal sTable As String) As String
    ' [SCAFFOLD] Just the Event= names, in order — a one-line structural check that a table
    '            carries the macros the generator said it created.
    Dim sXml As String
    Dim lPos As Long
    Dim lEnd As Long
    Dim sOut As String

    On Error GoTo errHandler

    sXml = DumpTableMacros(sTable)
    If Left(sXml, 6) = "ERROR:" Then
        ListMacroEvents = sXml
        Exit Function
    End If

    lPos = InStr(1, sXml, "Event=", vbTextCompare)
    Do While lPos > 0
        lEnd = InStr(lPos + 7, sXml, Chr(34))
        sOut = sOut & Mid(sXml, lPos + 7, lEnd - lPos - 7) & " | "
        lPos = InStr(lEnd, sXml, "Event=", vbTextCompare)
    Loop

    ListMacroEvents = sTable & ": " & sOut & "(" & Len(sXml) & " chars)"

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    ListMacroEvents = "ERROR: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

## Standards Layer

- **Error handling** — the blocks above ship the dependency-free `MsgBox` default; substitute
  your house pattern per `error-handling.md`. **This is a question, not a default** — see
  *Before you write the modules*, which is where it gets asked. One deliberate exception is
  annotated in place: `BackupLongTextFieldsDM` stays quiet (it runs inside every save).
- **Query style** — the inline SQL kept here is from the proven source; rewrite per
  `query-style.md` if your house centralizes SQL differently.
- **Naming conventions** — the config scan's `tbl`/`tlkp` (excluding `tmp`) prefix filter is the
  naming convention made executable, and `CheckAuditReadiness` repeats the same filter so the two
  stay in step; adjust both with your prefix policy.
- **Design principles** — one job per procedure throughout: one sample-data setup (Path A only),
  three numbered entry points, one safety check, five single-macro builders, one comparison
  helper, one staging function, one admin reset.

## Extra Options

*Empty in the base template. Filled per client engagement.*

- **The identity function this database already has** — on Path B, a database that has been in use
  often has its own, returning the signed-in person's name rather than their Windows account. Where
  it does, that is usually the better answer than either choice below. Use it in **all six** sites:
  the four in the table below, plus the two in `BuildBeforeChangeMacro` that stamp `CreatedBy` and
  `ModifiedBy` on the record. `AuditUser()` is then not built at all, and the front-end placement
  requirement moves with the choice — see *Before you write the modules*, which is also where this
  gets asked.
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
  describes. A live trial hit exactly that. Two coherent ways to hold it:

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
- **Composite/text primary keys** — `CheckAuditReadiness` detects these and tells you to fix or
  exclude the table; the staging plumbing and macro XML themselves still assume one numeric PK
  and don't support them (schema Business Rule 4).
