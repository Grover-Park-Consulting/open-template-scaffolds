---
template: officiating-assignment-scaffold
title: Officiating Assignment — Assignment & Pay VBA Scaffold
domain: scheduling-assignment
type: vba-scaffold
version: 0.10.0
status: stable
implements: officiating-assignment-schema
requires_tables:
  - tblGame
  - tblOfficial
  - tblGameOfficial
  - tlkpOfficialPosition
  - tblPositionRate
  - tblTeam
  - tlkpPlayLevel
  - tblAppSetting
platform_facts: [target-file, sql-insert-truncation, data-macro-rules, recordset-append-crash, app-startup-autoexec]
steps: [01-open-and-confirm-tables, 02-procedures-1, 03-procedures-2, 04-checks]
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
target_module: modOfficiatingAssignment
new_procedures:
  - AssignOfficial
  - ValidateAssignment
  - GetApplicablePayRate
  - GetAppSetting
  - EnsureGameValidationRule
  - GameLevelID
  - OfficialAge
  - EnsurePhotoFolder
  - SetOfficialPhoto
  - ListOpenObjects
warnings:
  - This build's active-official check (Business Rule 3, inside ValidateAssignment) only covers
    assignments made through AssignOfficial. An assignment inserted directly into tblGameOfficial, or
    by an import, is not checked. This is the VBA route named in
    officiating-assignment-outcome-first — that template's Data Macro route closes this gap,
    at greater build cost. Choosing this scaffold is choosing this trade-off.
  - EnsureGameValidationRule writes a validation rule onto tblGame itself, replacing any rule text
    already there, and Access checks that rule against the rows already in the table.
    SetOfficialPhoto writes to tblOfficial. A build against a database in real use is preceded by a
    backup copy of the file, and the developer is asked for one before anything is changed.
---

# Officiating Assignment — Assignment & Pay VBA Scaffold

**Status last determined:** 2026-09-26.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Realize seven of the officiating **table** template's (`officiating-assignment-schema`)
nine Business Rules in code: assigning an official to a game position with friendly validation
and an active-official check (Rules 2, 3), resolving the effective-dated pay rate (Rule 6),
keeping a game internally consistent — two different teams, an end time after the start
(Rules 4, 7) — deriving a game's play level and an official's age rather than storing either
(Rules 5, 8), and handling a photo as a file name plus one shared, confirmed folder (Rule 9).
The other two — the crew being a junction, and assignment uniqueness — are already complete the
moment the table template's tables are built; nothing here touches them. As with every scaffold,
this supplies **procedure skeletons** — signatures, recordset plumbing, control flow, error
structure — with the domain logic marked against the table template's numbered **Business
Rules**, and house style deferred to the standards layer.

**The other version of this template — the outcome-first method,
`officiating-assignment-outcome-first` — produces the same result from a specification
rather than working code, and offers a genuine choice for Rule 3** (a Data Macro that closes the
coverage gap named in the warning above, or this scaffold's own VBA route). Either can be built
against your own database, and they can be built one after the other, against separate copies, to
compare.

A query-style note worth keeping: the source database this domain was shaped from stored the
crew as two hardcoded columns and needed a `UNION ALL` query to un-pivot them into
one-row-per-assignment. Against the junction, that whole query collapses to
`SELECT ... FROM tblGameOfficial` with joins — the before/after argument for the
normalization, written per `query-style.md`.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here (signature, plumbing, control flow, error structure).
- **`[STANDARDS]`** — house style, deferred to the standards layer (`error-handling.md`, `query-style.md`, `naming-conventions.md`).
- **`[BUSINESS LOGIC]`** — the domain rule you fill in, sourced from the table template's Business Rules.

## Prerequisites

| Object | Role |
|---|---|
| `officiating-assignment-schema` tables | The scaffold runs against the tables that template creates (`tblGame`/`tblOfficial`/`tblGameOfficial`, the lookups, `tblPositionRate`, `tblAppSetting`) |
| `tblAppSetting.OfficialPhotoFolder` seed row | Read by `GetAppSetting`, `EnsurePhotoFolder`, and `SetOfficialPhoto` (Business Rule 9). **In a split database this must be an absolute shared path**, e.g. `\\server\share\OfficialPhotos\` — the schema's own seeded value is relative, which only suits a single-file database. |
| A photo picker | The screen that calls `SetOfficialPhoto` with the file the person chose — a form concern, deferred to a `form-spec` template. |
| A central error logger | `error-handling.md` |

### Ask before building

**Where this is a database already in use, ask for a backup copy before changing anything.** This
template alters the file it is built into. Two questions, in this order:

1. *"Is this a database you already use, or a new one you're trying this out on?"*
2. Where it is one they already use: *"Make a copy of the file before I start. Say when it's done and
   the build continues; say there's no copy and it stops."*

**Stop and build nothing if they say there is no copy.** A copy made before anything changes is the
only way back.

A new database, or one they are trying this out on, needs no copy — the question ends at step 1.

### Where this module goes in a split database

A **split database** is the normal shape for Access applications, especially those in multi-user
environments: one file holds the tables (the **back end**, usually on a shared network drive — never
on OneDrive, Dropbox, or any other file-syncing cloud folder, which corrupts a shared Access back
end), and each person runs their own copy of a second file holding the forms, reports, and code (the
**front end**), whose tables are *links* pointing at the back end. A single-file database — one
.accdb holding everything — is an acceptable choice for one user, and everything here works there
too; put everything in that one file and ignore the distinction.

| Module | Back end | Front end | Why |
|---|---|---|---|
| `modOfficiatingAssignment` (this scaffold) | No | **Yes** | Every procedure here is called by the assignment form — it validates what the user picked, resolves a rate to display, and reads a setting the form needs. That is front-end work, so the code belongs beside the form. |

**Assigning officials is a shared job.** An assignor at the desk and a second person working from
home are both making assignments into the one back end, and that changes two things in this module:

1. **`ValidateAssignment` cannot be relied on alone.** It checks that the position is free, then
   `AssignOfficial` inserts. Between those two moments somebody else can insert the same position
   for the same game — both people's checks passed, and the junction's unique index on
   (`GameID`, `OfficialPositionID`) refuses the second insert with an engine error. Checking first does not
   remove the race; it makes it rare and keeps the *ordinary* refusal friendly, which is its real
   job. `AssignOfficial` must still trap the duplicate-key error and turn it into the same plain
   message, or the second assignor sees a raw Access error instead of "that position is already
   filled."
2. **`GetAppSetting` reads a shared table.** `tblAppSetting` lives in the back end, so a setting
   changed once reaches everyone — which is exactly why the photo folder is a setting (schema
   Business Rule 9) rather than a constant compiled into each front end.

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure plus the error-reporting call and
  line-number policy come from `error-handling.md`, which ranks three options and says when each
  fits. The `MsgBox` block shown is option 3, which needs nothing installed; a practice with its
  own logger substitutes it at the call site, and may number lines or not.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house
  query standard (aliasing, where querydefs live, formatting, safe criteria).
- **Naming** — procedure, variable, and parameter names follow `naming-conventions.md`.

## Extra Options

*Named optional extensions, none of them filled in for an engagement; the filled copy is saved to the
developer's own library, not committed here.*

- **RemoveAssignment / ReassignOfficial** — the un-assign and swap counterparts of
  `AssignOfficial`, with the same friendly validation.
- **Crew-completeness check** — a per-game function reporting unfilled positions against the
  engagement's required crew.
- **Compensation rollup** — per-official or per-game pay summaries built on
  `GetApplicablePayRate` (pairs with the schema's compensation-ledger Extra Option).

## Parked / future considerations (not in this design)

- **Availability / conflict checking** — refusing assignments that overlap an official's other
  games by date/time; needs an engagement's travel-time and doubleheader policies first.
- **Rate-snapshot on assignment** — if pay must survive later rate edits, the assignment row
  stores the resolved rate at assignment time (the schema's compensation-ledger Extra Option
  carries the shape).
