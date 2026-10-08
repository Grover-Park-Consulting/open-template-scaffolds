---
template: time-off-ledger-scaffold
title: Time Off Ledger — VBA Scaffold
domain: time-off
type: vba-scaffold
version: 0.3.0
status: draft
implements: time-off-ledger-schema
requires_tables:
  - tlkpTimeOffCategory
  - tlkpEntryReason
  - tblEmployee
  - tblAccrualSchedule
  - tblTimeOffEntry
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
  - design-principles
  - startup-conventions
target_module: modTimeOffRules, modTimeOffPosting
new_procedures:
  - CheckHook
  - ReadDb
  - ScalarValue
  - LockEmployee
  - TimeOffBalanceOn
  - TimeOffBalance
  - CompletedYears
  - AccrualHours
  - CountEarnedForPeriod
  - TimeOffEntryCheck
  - ExpectedRefusal
  - AppendEntry
  - PostOnePeriod
  - PostEarnedForPeriod
  - PostEarnedOnPurpose
  - PostDuePeriods
  - PostTimeTaken
  - CancelTimeOffEntry
  - ReplaceTimeOffEntry
platform_facts: [target-file, sql-insert-truncation, data-macro-rules, domain-function-transaction, row-lock-errors, recordset-append-crash, app-startup-autoexec]
steps: [01-open-and-confirm-tables, 02-rules-module, 03-posting-module, 04-checks]
warnings:
  - These procedures write real rows into the ledger, and the ledger is never edited or deleted once
    written, so a run that goes wrong cannot be undone by deleting what it wrote. A build against a
    database in real use is preceded by a backup copy of the file, and the developer is asked for one
    before anything is changed.
  - The function the Data Macro calls, TimeOffEntryCheck, and everything it calls must exist in the
    back end and in every front end. In a split database, two front ends holding different copies
    hold two people to different rules. Nothing keeps the copies in step, so a changed function has
    to be imported everywhere.
  - The checks, the lock and the replacement all depend on one rule about where a read comes from. A
    read that must see work a transaction has just done is made on the Database that transaction
    was begun on, never through a domain function. A build that breaks this compiles, runs, and
    returns a wrong answer without raising anything. See the transaction guard in error-handling.md.
  - CheckHook is empty in the delivered database and only a check copy fills it in. Delivering it
    filled in would run test code on every posting.
related:
  - "time-off-ledger-outcome-first — the same ledger, produced by an open route in place of these skeletons. The promise is identical and each template says so. It carries the checks this one is validated against."
  - "error-logging-scaffold — builds the shared logger these procedures call when an unexpected error occurs. Build it first if the database has none."
---

# Time Off Ledger — VBA Scaffold

**Status last determined:** 2026-10-05.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not
have to read it to use the template.

## Intent

Realize the logic that `time-off-ledger-schema` defers to code: the checks a new entry must pass
(Business Rules 4 to 10), the balance (Rule 2), hours earned by length of service (Rule 3), and the
six ways entries get posted. This scaffold supplies the **procedure skeletons**: signatures,
recordset plumbing, control flow, the transaction and lock structure, and the error-handling frame.
It does not write the domain logic. Each procedure marks where that goes, sourced from the table
template's numbered Business Rules. House style is deferred to the standards layer.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here.
- **`[STANDARDS]`** — house style, deferred to the standards layer.
- **`[BUSINESS LOGIC]`** — the domain rule you fill in, sourced from the table template's Business
  Rules.

**What this scaffold does not build.** The tables (`time-off-ledger-schema`), the Data Macro that
keeps a posted entry from changing and calls `TimeOffEntryCheck` (Business Rule 1 and the checks;
`time-off-ledger-outcome-first` states the routes and `templates/_materialization.md` has the
shape), and the open-time macro for posting. This is the VBA those pieces call.

## Prerequisites

| Object | Role |
|---|---|
| `time-off-ledger-schema` tables | The five tables the code runs against |
| A Data Macro on `tblTimeOffEntry` calling `TimeOffEntryCheck` | Where the checks run, on every route into the table |
| A central error logger | `error-handling.md`; `error-logging-scaffold` builds one |

### Ask before building

**Where this is a database already in use, ask for a backup copy before changing anything.** Two
questions, in this order:

1. *"Is this a database you already use, or a new one you're trying this out on?"*
2. Where it is one they already use: *"Make a copy of the file before I start. Say when it's done and
   the build continues; say there's no copy and it stops."*

**Stop and build nothing if they say there is no copy.** A new database needs no copy.

**One choice belongs to the developer and is never inferred:** when earned hours are posted. It is
item 7 under *Information and conditions you need to supply* in `time-off-ledger-outcome-first`.
Ask it as that template says.

### Where the code goes: it depends on how the database is built

**A single-file database.** Everything goes in that one file. There is nothing more to decide.

**A split database** (one file holding the tables, the back end, and a copy of a second file for
each person, the front end). Where the code goes follows from two facts:

- The Data Macro on `tblTimeOffEntry` calls `TimeOffEntryCheck`, and a Data Macro can call only a
  function that exists in the file where the change is made. So the rules code (the procedures up to
  and including `TimeOffEntryCheck`) must be in the back end.
- The posting procedures are called from whatever each person uses to post, so they go in each front
  end. They call the rules code, so each front end holds the rules code as well.

So the rules code is in the back end and in every front end, and the posting code is in the front
ends. Nothing keeps the copies in step: where the rules code changes, import it everywhere, or two
people are held to different rules.

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure, both shapes, the transaction guard and the
  line-number policy come from `error-handling.md`. A function a Data Macro calls reports a refusal
  to the macro and does not open a message box of its own. The posting procedures use the shared
  logger, never a message box, because nothing in a posting may interrupt a person with a dialog.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house query
  standard.
- **Naming** — procedure, variable and parameter names follow `naming-conventions.md`.
- **Design principles** — how these procedures divide is already settled here; a build that merges or
  splits them keeps every rule above.

**To the AI assistant: once the build is reported finished, and only then, mention each entry under
`related` in the front matter** (`_template-schema.md` §7.1), one line per entry, what it is and why.
This is not part of the build, never a gate, and never read before this point.

## Extra Options

*Named optional extensions, none of them filled in for an engagement.*

- **Role-based accrual, going below zero, a stored balance, date-range entries, holidays, carry-over
  and expiry**, all named in the schema's own *Extra Options*. This scaffold builds none of them.
  Each changes which procedures read or write what, and the outcome-first checks describe the base
  build only. Re-check every read against the transaction guard if one is taken.
- **A schedule that posts by itself.** A Windows scheduled task opens the database, calls
  `PostDuePeriods`, and closes it. This scaffold builds what the task runs; the task is the
  developer's.

## Parked / future considerations (not in this design)

- **A procedure that finds the open-time macro and adds the posting to it** lives in
  `app-startup-scaffold` and `standards/startup-conventions.md`, not here.
