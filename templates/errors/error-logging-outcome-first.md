---
template: error-logging-outcome-first
title: Error Logging — outcome-first method
domain: errors
type: outcome-first
version: 0.2.0
status: draft
implements: error-logging-schema
standards_layer:
  - error-handling
  - naming-conventions
  - query-style
  - design-principles
house_assumptions:
  - "Error log — the house audit columns (CreatedDate/CreatedBy, filled by a Before Change Data Macro calling a
    VBA function) are deliberately not applied to it. The two facts they would carry
    — when a row was written, and by whom — are instead written by the logger itself, as ordinary
    columns, so the record does not depend on a second mechanism being present and working at the
    exact moment something already is not. This mirrors the table template's own declared assumption;
    a build that attaches the house audit-column mechanism to this table has reintroduced the
    fragility its schema's Business Rule 2 forbids."
warnings:
  - Nothing about the log may be able to refuse a write. No required column the logger does not
    supply itself, no engine-evaluated default, no data macro, no relationship to another table. A log
    that rejects the row records nothing, and the error it was called about is lost with it.
  - In a split database, a log table in the back end cannot be written when the back end is the thing
    that failed — which is the error you most want kept. Where that risk matters, choose the
    table-falling-back-to-file route under *Information and conditions you need to supply*.
related:
  - "app-startup-outcome-first — worth considering if you haven't built it yet. The
    back-end-unreachable case it handles is one of the more valuable situations to have this log
    actually catch."
  - "standards/error-handling.md — worth a second look now that you have a logger. It offers two
    ways to report an error, and building this template is what makes the preferred one, calling a
    shared logger, available to you."
---

# Error Logging — outcome-first method

**Who reads this.** Everything from *Intent* down to *Standards Layer* is written for the developer
whose database this is. The section after that is addressed to the AI assistant building it, and
says so where it starts.

---

## Intent

**To produce, in an Access database, the results described in *What you end up with* below.** That
section is the specification: what the database does once this is built, in words the developer can
check for themselves, followed by the checks that confirm the result arrived and the behaviours that
must hold however the work was divided up. Those three sections together are the whole of what this
template promises.

**Nothing else is here.** No procedures, no module names, no code. This template states the result
and leaves the route to whoever builds it. No single mechanism is named, because the platform does
not force one: ordinary VBA, called from the end of an error handler, does everything this template
asks for. The other version of this template — the rules-based method, `error-logging-scaffold` —
produces the same result from a working module you import. Either one can be built against your own
database, and they can be built one after the other, against separate copies, to compare.

**This template realizes logic the table template leaves open.** `error-logging-schema` builds the
one table an application writes to when something goes wrong, and states, in its own Business Rules,
what has to be true of anything that writes to it — but leaves *how* to "the paired scaffold or this
template." This is that half, stated as an outcome rather than as code. It assumes that table already
exists, or is built as part of this same run.

---

## What you end up with

**Every unhandled error that reaches the logger produces exactly one record, and the log can never
refuse to accept it.** No required column depends on anything the logger doesn't supply itself, no
value depends on the database engine working one out, no rule ties the record to another table. A
failure serious enough to need logging is the worst possible moment for the log itself to have a
condition of its own that could go unmet.

**The record reflects the error exactly as it stood at the moment of failure — not a value changed by
anything the logger did afterward to capture or report it.** Reading and holding the error's number
and description happens before any of the logger's own work that could otherwise overwrite them, so
what lands in the log is always the error that was actually raised, never one produced by the act of
investigating it.

**A description too long for the log to hold is shortened, never rejected.** A shortened description
still names the problem; a refused write names nothing.

**The record always says who was running the code. Where that can't be determined, it says
"Unknown," never leaving the column blank and never blocking the write on it.**

**Where the log tracks which line failed, it names the exact line when the procedure carries line
numbers, and says plainly that no line is known when it doesn't — never a guess, and never an empty
column.** A procedure with no line numbers is a legitimate case, not a missing value.

**Records are never edited or deleted by the application afterward.** The log is a running history of
what went wrong; clearing old entries is a deliberate housekeeping act a person chooses to run, not
something the application does on its own.

**Where you chose to record to a table with a text file as fallback, a table the application cannot
reach is invisible to the person at the keyboard.** The entry still lands — in the file instead — and
nobody is told the table failed unless you chose to show full technical detail. What must never
happen is the reverse: a failure that reaches neither the table nor the file, with nothing said about
it.

**Where a reference number is shown to the person at the keyboard, it always names the row that was
actually just written.** It is never a stale number, and never invented when nothing was recorded.

**Whatever the person at the keyboard is told, it is never told that something was recorded when it
was not.** If nothing could be saved anywhere, they are told that plainly and asked to remember what
they were doing, rather than being reassured incorrectly.

**Nothing about recording an error stops the person using the application, beyond whatever you chose
they should be told.** Logging is bookkeeping that runs because something already went wrong; it does
not add a second interruption of its own.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

Perform each of these checks against a copy of your database. **Validate by asking: does this happen
or not?**

1. **An ordinary error is recorded correctly.**
   - Force an error in a procedure that calls the logger
   - Confirm the log holds that error's number, its description, the module and procedure it
     happened in, when it happened, and who was running the code
2. **An oversized description is shortened, not rejected.**
   - Force an error whose description is longer than the log's description column can hold
   - Confirm the log holds a shortened version of it and the write succeeded
3. **A procedure with no line numbers logs that plainly, not blank.**
   - Force an error in a procedure that carries no line numbers
   - Confirm the log's line column shows the known "no line" value for that procedure, not an empty
     column
4. **A procedure with line numbers logs the line that actually failed.**
   - Force an error on a specific numbered line
   - Confirm the log names that line, not the procedure's first line or a default
5. **Errors are never overwritten or lost.**
   - Force two different errors, one after the other
   - Confirm the log holds a separate row for each, with each one's own detail intact
6. **Choosing the table records to the table.**
   - Where you chose to record to a table, force an error
   - Confirm a row appears in the table
7. **Choosing the fallback still records when the table cannot be reached.**
   - Where you chose the table with a file fallback, make the table unreachable (rename it, or, in a
     split database, remove the link) and force an error
   - Confirm the entry appears in the file instead
   - Confirm the person at the keyboard sees only what you chose them to see — not a report that the
     table failed, unless you chose the full-detail option
8. **The reference number shown matches the row that was written.**
   - Where you chose to show a reference number, force an error and note the number shown
   - Confirm that number identifies the row just written, not a different one
9. **Running the build again does not duplicate anything.**
   - Run whatever created the log table and the logging code a second time
   - Confirm there is still exactly one log table, and no second copy of the logging code
10. **The logger cannot itself cause a new failure, even when every recording route fails at once.**
    - Make both the table and the file destination unreachable at the same time (rename the table and
      point the file path somewhere that cannot be written to)
    - Force an error
    - Confirm the application does not crash and no new, unhandled error appears
    - Confirm the person at the keyboard is told the problem could not be recorded, and asked to note
      what they were doing

### The same behavior every time, not the same structure

For an AI-assisted template, two builds need not produce the same code. They need to produce a
database that behaves the same way. The following must be true of every build:

- Every unhandled error in the application reaches the same recording logic — one shared destination
  for the fact of a failure, not a separate copy of the recording code sitting in every handler.
- The value recorded for the error's number and description is the value that was true at the moment
  the error was raised, regardless of what the logger does afterward to identify the module, the
  procedure, or the line — nothing the logger does to investigate an error can be allowed to change
  what that error was.
- Nothing the recording logic does can itself raise an error that escapes to the handler that called
  it. A failure to reach one recording destination steps down to another, or reports that nothing
  could be saved; it never propagates a new error out of code that was already handling one.
- **[your standards]** Where the recording path deviates from the house error-handling pattern for the
  reason above — a handler cannot call a handler of its own without risking exactly the failure it
  exists to prevent — that deviation is confined to the recording logic itself. Every other procedure
  in the build still follows the standards layer's own pattern and calls into the shared logger at the
  point that pattern names.
- Whatever creates the log table and the logging code is safe to run more than once. A second run
  reports what already exists rather than duplicating it.
- Nothing here interrupts the person using the application beyond what they chose under *Information
  and conditions you need to supply*. A shortfall in what could be recorded is never itself shown as
  an error dialog outside that choice.

Lines marked **[your standards]** come from your standards layer rather than from this template, and
move with that layer if your shop replaces it. Everything else in `standards/` applies here as it
does to every template.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences, say so while the
design is being worked out, before anything is built. Where you don't choose, the build will choose
based on the rules built into it. The template's promise holds either way.

- How the logic divides into procedures, what they are called, and how many there are. They will be
  functionally equivalent, not necessarily structurally the same.
- Whether the build reports in message boxes, as returned text, or both.
- How the code is laid out and commented, within whatever your standards already require.
- The order in which you are asked for the things only you can supply.

**If you want something not already in the template, you can have it.** Our promise is specified in
*What you end up with*, together with *The same behavior every time, not the same structure*. The
checks under *How you validate the template's output* are how you confirm you got it in any given
build. Asking for something different is an extension of the basic build, and you are welcome to
experiment — the result is then your responsibility, and the checks above may no longer describe what
you have.

### What the template does not do

- **A screen for browsing or filtering the log.** The log is a table. It can be queried. You can add
  your own form for it if you wish.
- **Removing old entries.** The log grows for as long as the database is used. You can add a feature
  to prune it; it is not in the base template.
- **Forwarding entries anywhere else** — a central server, an email summary — is yours to add.
- **Reducing repeated identical errors to one row with a count.** Every occurrence is its own row.
- **Recording anything other than an unhandled error.** This template does not audit changes to your
  data — see the audit-logging templates in this library for that.

---

## Information and conditions you need to supply

You need to provide information about five things; no guesses are made on your behalf.

1. **Where errors should be recorded** — a table in the database, a text file, or the table with the
   file as a fallback when the table cannot be reached. This changes what you actually get: a table is
   queryable; a file survives a database that will not open at all; the third loses nothing but costs
   a few more lines.
2. **Where that lives.** For a table: the back end of a split database (one shared log) or each front
   end (always writable, but scattered). For a file: the front end's own folder, the Documents folder,
   or the back end's folder. A single-file database has one place either way.
3. **What the person at the keyboard sees when something goes wrong** — a short plain message with a
   reference number where one exists, the full technical detail, or nothing at all.
4. **Whether you have a backup**, if this is a database you already use. If you tell the template
   there is no backup, the build stops rather than continuing.
5. **Permission to change your database**, asked immediately before anything is built.

**The accdb must be in a Trusted Location.** Code in an Access file outside one does not run at all,
and Access does not always say so — the failure can show up as the database reporting it cannot find a
procedure that is, in fact, right there. Trust is a setting on each machine, not a property of the
file.

---

## Standards Layer

**Your standards layer decides how this is built. This template decides only what it has to do.**

- **`error-handling`** — the pattern a generated procedure follows to report a failure, and the
  procedure the standards layer names as the one every handler calls. The recording logic itself is
  the documented exception to that pattern — see *The same behavior every time* — because it cannot
  call a handler of its own without risking the failure it exists to prevent.
- **`naming-conventions`** — what the log table, any new module, and their columns are called, within
  whatever your standards already require.
- **`query-style`** — how the generated code writes and holds its SQL, and specifically how it appends
  a row without risk of a description's own contents breaking the write.
- **`design-principles`** — how the work divides into procedures. *Free to choose alternatives* leaves
  that division open on purpose, and this layer is what it is open to.

A shop adopting this library replaces `standards/` with its own. Nothing in this file changes when
they do, which is the point of keeping the two apart.

---

## To the AI assistant building this

**Two sections are the specification, and only those two: *What you end up with* and *The same
behavior every time, not the same structure*, both under that heading above.** Build a system that
satisfies every promise there and passes every entry under *How you validate the template's output*.
Then run every one of those checks yourself, on a copy, and record what each one did — in
`build-record.md`, written before you report the build finished, never afterward and never only when
asked for it. The checks bind you twice: the build has to pass them, and you have to run them. How you
build is yours to decide, within *Free to choose alternatives*. Which checks you run is not — all of
them, every build. Every other section in this file — *Intent*, *What the template does not do* — is
context for reading those two. None of it binds on its own, and nothing that binds is stated only
there.

- **This template names no mechanism, because the platform leaves more than one workable route.**
  Ordinary VBA behind a shared logging procedure, called from the end of every handler, satisfies
  every promise above. Do not import a mechanism from the audit-logging templates in this library on
  the assumption that a promise this firm must mean one route: it doesn't, here.
- **You may read `error-logging-scaffold.md`, the rules-based method that produces this same result,
  for one worked decomposition — nothing more.** It shows procedure names, a control flow, and where
  the guarding logic sits. None of that is binding here. Copying its shape wholesale is a legitimate
  build; so is a different one that still satisfies every check.
- **Read `error-logging-schema.md`, the table template this realizes, for the table, its Business
  Rules, and the reasoning behind them** — in particular Business Rule 2 (nothing may refuse a write)
  and Business Rule 3 (the logger supplies every column itself; the database engine never evaluates a
  default for this table). This file restates their outcome; that file is where the field names and
  types live.
- **Reading `Err.Number` and `Err.Description` before any code that could clear them is not a route
  choice — it is the only way the promise about recording the error "as it actually stood" can hold.**
  Any `On Error` statement clears `Err`, and so does `Resume`, `Exit Sub`, or `Exit Function`. Whatever
  decomposition you choose, the values have to be captured — into parameters, into locals, however you
  divide the work — before anything in the recording path could clear them. This is stated as a
  behaviour under *The same behavior every time*, not as code to copy; build it however you choose, as
  long as the behaviour holds.
- **Read every file in `standards/` and apply it.** Error handling, naming, query style, and how the
  work divides into procedures all come from there and never from this file.
- **Ask for the five things under *Information and conditions you need to supply*,** one at a time,
  through the interactive selection control where the answer is a choice. Two of them are gates: a
  database in real use with no backup stops the build, and permission to change the database is asked
  immediately before anything is built.
- **After the fifth thing and before you present the design, offer the `Explore options` step**
  (`_template-schema.md` §12.5) over the list under *Free to choose alternatives*, and nothing outside
  it. A pick made there is recorded in the build record, holds for the rest of the run, and is
  restated in the design you present.
- **Never infer an answer that belongs to the developer** — not from what the database looks like, not
  from reasoning that makes an answer seem obvious. Where a check exists to answer a question, run the
  check at the point the sequence calls for it rather than working the answer out yourself.
- **Surface both house assumptions and every warning in the front matter** and get the developer's
  answer on each before building.
- **The build record reports against *How you validate the template's output*, one entry per check,
  each saying what was done and what was observed.** Passed and not passed are the only outcomes,
  including where the first method to run a check hits an obstacle — see `_template-schema.md` §12.2
  for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every entry opens with, and what
  to do before settling for a soft result.
- **While the build runs, do not narrate it.** Say once that it has started and what it will produce;
  say anything the developer must act on, as a question; say when it is finished, what was built, and
  where the build record is. Everything else — every procedure written, every check that passed — goes
  to the build record.
- **Once the build is reported finished, and only then, mention each entry under `related` in the
  front matter** (`_template-schema.md` §7.1) — one line per entry, what it is and why. This is not
  part of the build, never a gate, and never read before this point.

## Extra Options

*Named optional extensions, none of them filled in for an engagement.*

- **A context note the caller passes in** — an optional extra piece of information recording what the
  code was doing when it failed ("saving invoice 4471"), stored alongside the rest of the record.
- **Rate limiting** — where one failure repeats in a loop, record the first occurrence and a count
  rather than one row per repetition.
- **A "report this" action** on the message shown to the person at the keyboard, copying the reference
  number for a support request.
- **Forwarding** — a routine that copies new entries to a central location, where one team supports
  several installations.
