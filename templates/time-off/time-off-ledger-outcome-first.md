---
template: time-off-ledger-outcome-first
title: Time Off Ledger — outcome-first method
domain: time-off
type: outcome-first
version: 0.4.2
status: review
implements: time-off-ledger-schema
standards_layer:
  - audit-columns
  - naming-conventions
  - error-handling
  - query-style
  - design-principles
  - startup-conventions
house_assumptions:
  - "Balance — counts every entry, including Taken entries dated in the future, so time booked ahead is already set aside and cannot be booked twice. A shop that wants the balance as of today only changes this; the schema does not ask for it, so this template does not build it."
platform_facts: [target-file, sql-insert-truncation, data-macro-rules, domain-function-transaction, row-lock-errors, recordset-append-crash, app-startup-autoexec]
steps: [01-checks, 02-tables-and-data-macros, 03-code, 04-run-the-checks]
warnings:
  - Choosing the Data Macro route for Business Rule 1 or for the posting checks means this template attaches a Data Macro to a live table. A build against a database in real use is preceded by a backup copy of the file, and the developer is asked for one before anything is changed.
  - tblTimeOffEntry will already carry a Data Macro if the audit columns were built. The build reads it first and merges the new logic into it. It never loads a new one over it, because that replaces the stamping silently.
  - A Data Macro that calls a function needs that function in every front end as well as the back end. In a split database, if one front end holds an older copy than another, two people can be held to different rules. The build says where the functions went and the developer must re-import them everywhere whenever they change.
  - Choosing the posting-procedure route for the posting checks (Business Rules 4 to 10) means an entry inserted any other way, directly into the table or through an import, is not checked. This is a disclosed trade-off of that route, not a defect; the Data Macro route does not have this gap.
  - Two people can post at the same moment. For Business Rules 4, 5, 9 and 10 the build must turn that into one of them succeeding and the other being refused, never both succeeding. Checking first and then writing narrows the window and does not close it.
  - Replacing a wrong entry writes two entries, the corrected one and the cancellation of the wrong one, and the build must make them all or nothing. The check on the second entry has to see the first, which has not been saved yet, so the checks must read through the same transaction the two entries are written in. A build that checks on a separate connection will refuse every replacement, or, worse, accept some it should refuse.
  - Where the developer chooses to have earned hours posted when the database opens, this template adds to the database's open-time macro. If the database already has one, the build reads it first and adds to it. It never replaces it, because that silently removes whatever the database already did when it opened.
related:
  - "app-startup-outcome-first — worth building first, or alongside, if you haven't: it gives your application a place to confirm shared files and settings when it opens, which is where the functions this template relies on are best kept in step across front ends."
---

# Time Off Ledger — outcome-first method

**Status last determined:** 2026-10-05.

**Who reads this.** Everything from *Intent* down to *Standards Layer* is written for the developer
whose database this is. The section after that is addressed to the AI assistant building it, and
says so where it starts.

---

## Intent

**To produce, in an Access database, the results described in *What you end up with* below.** That
section is the specification. It says what the database does once this is built, in words the
developer can check for themselves, and it is followed by the checks that confirm the result
arrived and by the behaviours that must hold however the work was divided up. Those three sections
together are the whole of what this template promises.

**This template realizes ten Business Rules the table template states.**
`time-off-ledger-schema` builds the five tables and states, in its own Business Rules, what has to
be true once the database is in use: the ledger only grows, the balance is added up and never
stored, hours earned follow length of service, hours are earned once per period, no entry can take
the balance below zero, an employee who has stopped working earns and takes nothing afterwards, a
correction cancels exactly one entry, hours are signed, each day belongs to one category, and a
wrong entry is replaced by posting the corrected one first and then cancelling the wrong one. Each
needs something built on top of the tables to hold. This template states each as an outcome and
leaves the route to whoever builds it, within the choices named below. It assumes the schema's
tables already exist, or are built as part of this same run.

**How entries get posted.** Rules 3 to 10 are stated as outcomes, and something has to be the way a
person actually posts an entry. The build provides a way of doing each of six things: posting the
hours earned for one period, posting every period that is due, posting the hours for a period on
purpose, which is how a person gives hours that a Correction took away (Business Rule 4), recording
time taken, cancelling an entry, and replacing a wrong entry with a corrected one. Each answers with
a sentence saying what happened, and none shows a dialog of its own for a refusal. How each is built, and what it is
called, is the builder's to decide. No screen is built; a form of your own can call them.

**What the error-handling standard asks before the build.** The standards layer asks how errors
are reported in the code the build writes. Its preferred answer calls a shared error logger, which
`error-logging-outcome-first` builds. If you want that answer and your database has no logger yet,
build the logger first. Otherwise the standard's other answer, a message box, needs nothing.

**Nothing else is here.** No module names and no code beyond what is needed to reason through the
choices named below. The other version of this template, `time-off-ledger-scaffold`, produces the
same result from procedure skeletons you fill in, and is checked against the checks below.

---

## What you end up with

### Business Rule 1: the ledger only grows

**A posted entry in `tblTimeOffEntry` can never be changed and never be deleted: not through your
posting procedure, not through a direct edit to the table, not through an import.**

**Two routes exist to build this, and which one you pick changes what you actually get, so you are
asked, not defaulted:**

- **The Data Macro route (preferred).** A Data Macro is a small piece of logic Access attaches to a
  table itself, which runs whenever a row is added, changed or deleted, however the change is made.
  Attached to `tblTimeOffEntry`, it refuses every change to a posted row and every delete, on every
  route into the table.
- **The VBA route.** Simpler to build. Your own posting code never edits or deletes, and nothing in
  it can. An edit or delete made any other way, directly in the table or by a query, goes straight
  through and nothing catches it. That gap is a known, accepted trade-off of this route. Choose it
  only if you are comfortable living with it.

### Business Rule 2: the balance is added up, never stored

**An employee's balance in a category is always the sum of the hours on their entries in that
category, and it is always current.** Post an entry and the next time anyone looks, the balance
already reflects it, with no separate step to bring anything up to date. No table has a balance
column. Entries dated in the future count (see the house assumption above).

**No mechanism is named here.** A saved query and a VBA function both satisfy this promise equally.

### Business Rule 3: hours earned follow length of service

**For an accrual period, the hours earned are those of the schedule row for that category with the
greatest `MinYearsOfService` not above the employee's completed years of service on the first day of
the period.** An employee with two completed years and one with six, in a schedule with a band
starting at five years, are paid from different bands.

**Where no schedule row applies, nothing is guessed.** Nothing is posted for that employee and
period, and the gap is reported, as something a person notices. It is never posted as zero hours,
which would read as "this employee earns nothing" when the truth is "nobody has said what they
earn."

**No mechanism is named here.** This is ordinary lookup logic, finding the applicable row given a
date, and nothing about the platform forces one way of writing it.

### Business Rule 4: hours are earned once per period, and every period that is due gets posted

**An employee has at most one active Earned entry for a category and period.** An entry is active
unless a Correction cancels it. Posting the hours for the same employee, category and period a
second time, by posting every period that is due or by any other route into the table, leaves one
active entry, not two. That holds when the second attempt comes a day later, and it holds when two
posting attempts start at the same moment. The one exception is a replacement that is waiting to
cancel the entry it replaces (Business Rule 10).

**A period that was cancelled is not posted again by a run.** If a Correction cancelled the hours
for a period, and nothing replaced them, posting every period that is due treats that period as
already handled and leaves it alone. Taking hours away is not quietly undone by the next run. To
give the hours again, a person posts an Earned entry for that period on purpose.

**Posting every period that is due catches up whatever was missed.** A period is due when it has
started, is on or after the employee's `AccrualStartDate` (their `HireDate` where that is empty),
and starts no later than the employee's `InactiveDate` where there is one. If nobody posted hours
for three months, the next run posts all three, not only the latest. Periods before the
`AccrualStartDate` are never posted by a run, however long ago the employee was hired.

**What a run reports.** A run answers with how many periods it posted, how many were already there,
and how many it could not post, naming the first one it could not post and why. A period it could
not post is never silently skipped.

**Two runs reaching the same period at the same instant both finish without a failure.** One posts
the entry. The other finds that it is already there, or is held back until the first has finished
and then finds it, and carries on. Neither run reports an error for this to anyone, because nothing
went wrong.

### Business Rule 5: no entry can take the balance below zero

**An entry that lowers the employee's balance in its category, and would leave it below zero, is
refused, and the refusal tells the person who tried how many hours they have.** Time taken is the
usual case. A Correction that cancels hours earned is the other: if those hours have already been
used, cancelling them would leave the balance negative. An entry that leaves the balance at
exactly zero is accepted.

**A replacement is judged by its net effect.** For this rule an entry counts for its own hours,
except that a replacement (Business Rule 10) counts only for its hours less the hours of the entry
it replaces. Replacing 96 earned hours with 64 is therefore judged as lowering the balance by 32, and
the balance that counts is the one with the 64 in and the 96 out. If 80 hours have been taken, that
balance would be -16, so the replacement is refused as a whole, whichever of its two writes the
refusal comes at, and the result is the same: nothing is posted. The person
is told that 16 hours of time taken are not covered by the corrected hours, and the time taken
has to be dealt with first. Nothing is posted and nothing is lost track of.

**Two entries that would each be accepted alone, but not together, never both go in.** Suppose an
employee has eight hours and two people each post a full day at the same moment. One entry is
posted and the other is refused. The balance never goes negative because both checked first and
both then wrote.

**Which route builds the posting checks is your choice** (see *Information and conditions you need
to supply*, item 5), and it applies to the posting checks of Business Rules 4 to 10 together:

- **The Data Macro route (preferred).** The Data Macro on `tblTimeOffEntry` calls a VBA function
  that holds the logic and answers yes or no. It covers every route into the table: the posting
  procedure, a direct edit, an import.
- **The posting-procedure route.** The same checks run only in the procedure that posts an entry.
  An entry inserted any other way is not checked at all. This is a known, accepted trade-off of this
  route.

### Business Rule 6: an employee who has stopped working earns and takes nothing afterwards

**An entry dated after an employee's `InactiveDate` is refused.** An entry dated on the
`InactiveDate` or earlier is accepted. An employee whose `InactiveDate` is empty is still employed,
and nothing about their entries is refused for this reason.

### Business Rule 7: a correction cancels exactly one entry

**A Correction entry names the one entry it cancels, with hours equal and opposite to that entry's.
The cancelled entry is left exactly as it was. An entry can be cancelled once.** After the
correction, the balance is what it would have been had the cancelled entry never been posted. A
Correction that names no entry, names one that does not exist, names another employee's entry,
names an entry in a different category, names another Correction, or names one already cancelled,
is refused. **A Correction that cancels hours earned must carry a note saying why**, and is refused
without one, so that taking hours away from an employee is never anonymous.

### Business Rule 8: hours are signed, never zero, and in quarter-hour steps

**An Earned entry always has positive hours and a Taken entry always has negative hours.** A
Correction has the sign that offsets the entry it names. An entry with zero hours is refused, and so
is an entry whose hours are not a multiple of a quarter hour, so that sums are exact.

### Business Rule 9: one category per day

**An employee has at most one active Taken entry for any one date.** A Taken entry is active unless
a Correction cancels it. So vacation cannot be booked on a day that already has sick time, and sick
time cannot be booked twice for one day. A day that was booked and then cancelled can be booked
again. Two entries for the same employee and day posted at the same moment never both go in: one is
posted and the other is refused. The one exception is a replacement that is waiting to cancel the
entry it replaces (Business Rule 10).

### Business Rule 10: a wrong entry is replaced, not overwritten

**A wrong Earned or Taken entry is replaced by posting the corrected entry first, naming the entry
it replaces, and then cancelling the wrong one, and the two happen together or not at all.** If
anything stops the pair, because the cancellation is refused or an error occurs, the corrected
entry is not left behind: afterwards the ledger holds exactly what it held before, and the person
is told why.

**Examples the build has to handle.** The wrong accrual for a period is replaced by the right one.
A Taken entry on the wrong day is replaced by one on the right day, including when it used the
employee's whole balance, because the replacement counts only for the difference. A Taken entry
with the wrong hours on the right day is replaced by the right hours on the same day. In each, the
wrong entry and its replacement do not count against the rules about one entry per period or per
day for the moment between the two halves.

**What a replacement must be.** It has the same employee and category as the entry it replaces and
the same reason (Earned replaces Earned, Taken replaces Taken). An Earned replacement covers the
same accrual period. The entry replaced is still active, has not been replaced before, and is not a
Correction. A replacement that fails any of these is refused. A replacement posted on its own, by any
route into the table other than the replacement itself, is refused too: it would leave the wrong
entry and its replacement both counting. An entry can be replaced once, even when two replacements
of it are posted at the same moment. The cancellation that finishes the pair
carries a note saying why, as every Correction of hours earned does (Business Rule 7).

**Nothing gets in between the two halves.** While a replacement is waiting for its cancellation, no
other entry for that employee is posted in the gap.

### The same behavior every time, not the same structure

For an AI-assisted template, two builds need not produce the same code. They need to produce a
database that behaves the same way. The following must be true of every build:

- **Which route was chosen for Business Rule 1, and which for the posting checks, is stated plainly
  in the design and the build record**, not left for the developer to discover by testing. They are
  choosing a coverage tier and need to know which one they got.
- **The balance and the checks that read it are computed from the entries on disk at the moment they
  run.** Nothing reads a stored figure, because none exists.
- **Where a check and its write are separate steps, nothing a second person does between them can
  make both succeed.** One of two simultaneous entries is posted and the other is refused. Checking
  first and then writing does not satisfy this on its own. The checks that force the interleaving
  (checks 8, 10, 17, 31 and 32) are what tells a build that satisfies it from one that does not.
- **A replacement is all or nothing.** The corrected entry and the cancellation of the wrong one are
  written in one transaction. If the cancellation is refused, or anything else stops the pair, no
  trace of the corrected entry remains, and the person is told why.
- **A read that needs to see work its own transaction has just done is made on the same database
  object the transaction was begun on.** A domain function does not see that work. This includes the
  reads made by the posting checks while a replacement is being written: the check on the
  cancellation has to see the corrected entry written a moment before it, which is not saved yet.
- **The protection against two at once covers the whole replacement.** The employee's entries are
  held from the check on the corrected entry until the cancellation has been written or the pair has
  been undone, not released between the two halves.
- **Where a Data Macro calls a function, the function exists in the back end and in every front end,
  and the build record names where it went.**
- **The accrual lookup reads the schedule table**, never a figure typed into code.
- **A refusal tells the person what to do next:** the hours available, the date already taken, the
  entry already cancelled. It never reports only that something failed. Where a Data Macro refused
  the entry, the sentence the person sees is the macro's own sentence, carried through whatever
  code posted the entry, not a generic message in its place.
- **A failure nobody planned for is never mistaken for a refusal.** An unexpected error inside a
  function a Data Macro calls refuses the save. Access shows its own message for it, and it is not
  recorded unless the build makes it so; the build record says which. The posting procedures
  record their own unexpected errors with the shared logger where one exists.
- **[your standards]** The error-handling pattern applies to the build-time procedures and to any
  VBA function, not to a Data Macro's own `RaiseError` action. When a Data Macro refuses a save,
  Access shows its own dialog naming the problem; no `On Error` handler fires and nothing is logged
  unless something else is separately watching.
- Running the build again replaces what a previous build attached. It never adds a second copy
  alongside it.

Lines marked **[your standards]** come from your standards layer rather than from this template,
and move with that layer if your shop replaces it. Everything else in `standards/` applies here as
it does to every template.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences, say so while
the design is being worked out, before anything is built. Where you don't choose, the build will
choose based on the rules built into it. The template's promise holds either way.

- **How Business Rules 2 and 3 are exposed**: a saved query or a VBA function, for each
  independently.
- **How two simultaneous entries are kept from both going in** (Business Rules 4, 5, 9 and 10; the
  table's own unique indexes settle only cancelling and replacing an entry once): taking a write
  lock on the employee's row first, held
  from the check until the entry has been written, so that the second must wait; or writing and
  then checking inside the same transaction and undoing the write if it broke a rule. The first of
  these has passed the checks that force the collision when the checks run inside a function a Data
  Macro calls. The second has not been tried against them in that setting, so start from the first
  there. What must hold is stated above, and the checks that force the collision are how you know.
- **Whether the Data Macro for Business Rule 1 and the Data Macro for the posting checks are one
  document or are built together.** A table's Data Macros for one event live in one document, so
  this is settled by what is already attached.
- How VBA code, where any is used, divides into procedures, what it is called, and how many there
  are.
- The wording of everything the developer sees.
- Whether the build reports in message boxes, as returned text, or both.
- How the code is laid out and commented, within whatever your standards already require.
- The order in which you are asked for the things only you can supply.

**If you want something not already in the template, you can have it.** Our promise is specified in
*What you end up with*, together with *The same behavior every time, not the same structure*. The
checks under *How you validate the template's output* are how you confirm you got it in any given
build. Asking for something different is an extension of the basic build, and you are welcome to
experiment. The result is then your responsibility, and the checks above may no longer describe
what you have.

### What the template does not do

- **Catching a bad entry through a route the posting-procedure route doesn't cover.** Named plainly
  under *Business Rule 5* and in the warning above. This is the disclosed trade-off of that route,
  not something this template works around.
- **A trigger more elaborate than the three offered** under *Information and conditions you need to
  supply*, item 7. Reminders, per-employee schedules and an approval step are yours to add.
- **A screen for requesting time off or reviewing balances.** These are tables, functions and
  queries. You can add your own form or report.
- **Approval of requests.** Who may post a Taken entry, and whether a manager signs off first.
- **Anything about roles, going below zero, holidays or carry-over.** These are Extra Options in the
  schema.

---

## Information and conditions you need to supply

Eight things, and nothing here is guessed on your behalf:

1. **Whether you are trying this out on a new database or building it into one you already use.**
   The two behave differently from the first step onward.
2. **Which file holds the tables**, if your database is split.
3. **Whether you have a backup**, if this is a database you already use. If you tell the template
   there is no backup, the build stops rather than continuing.
4. **Which route you want for Business Rule 1: the Data Macro or VBA.** Read the trade-off under
   *Business Rule 1* before answering. This is a choice about what you actually get.
5. **Which route you want for the posting checks, Business Rules 4 to 10: the Data Macro or the
   posting procedure.** Read the trade-off under *Business Rule 5* before answering.
6. **Your accrual schedule: for each category, how many months make a period, and how many hours
   are earned per period at each length of service.** This template carries no schedule of its own.
   The build can propose a sample for you to change, and nothing is saved until you have confirmed
   it. Each band must earn more than zero hours.
7. **When earned hours should be posted.** Three choices, and they change what you get:
   - **You run it yourself.** Nothing posts until you start the posting of every period that is
     due. This is the simplest, and nothing happens that you did not start.
   - **When the database opens.** Each time a person opens the database, every period that is due
     is posted. It runs only when someone opens it, so a database nobody opens does not post. In a
     split database it runs from each person's front end, and the rule that a period is posted once
     is what stops two people posting it twice. Opening the database by an automated tool does not
     post anything, so that a tool's open does not change your data.
   - **On a Windows schedule.** A scheduled task opens the database, posts every period that is
     due, and closes it, with nobody present. It depends on a machine and an account that stay set
     up, and the task itself is yours to create; the template builds what the task runs.
8. **Permission to change your tables**, asked immediately before anything is built.

**The accdb must be in a Trusted Location.** Code in an Access file outside one does not run at
all, and Access does not always say so. The failure can show up as the database reporting it cannot
find a procedure that is, in fact, right there. Trust is a setting on each machine, not a property
of the file.

---

## Standards Layer

**Your standards layer decides how this is built. This template decides only what it has to do.**

- **`audit-columns`**: every table this template touches already carries the standard columns per
  the schema's own standards layer. This template adds nothing to that set. It does attach logic to
  the same event that stamps them, which is why the build merges rather than replaces.
- **`naming-conventions`**: what any function, module or setting is called, within whatever this
  layer already requires.
- **`error-handling`**: applies to the build-time procedures and to any VBA function a Data Macro
  calls, with one constraint that is not a preference: a function a Data Macro calls reports a
  refusal to the macro, which raises it, and does not open a message box of its own while someone is
  saving.
- **`query-style`**: how any generated code writes and holds its SQL, notably the balance sum and
  the overlap lookup.
- **`design-principles`**: how the work divides into procedures. *Free to choose alternatives*
  leaves that division open on purpose, and this layer is what it is open to.

A shop adopting this library replaces `standards/` with its own. Nothing in this file changes when
they do, which is the point of keeping the two apart.

---

## To the AI assistant building this

**Two sections are the specification, and only those two: *What you end up with* and *The same
behavior every time, not the same structure*, both under that heading above.** Build a system that
satisfies every promise there and passes every entry under *How you validate the template's
output*. Then run every one of those checks yourself, on a copy, and record what each one did, in
the build record, written before you report the build finished, never afterward and never only when
asked for it. The checks bind you twice: the build has to pass them, and you have to run them. How
you build is yours to decide, within *Free to choose alternatives*. Which checks you run is not:
all of them, every build. Every other section in this file is context for reading those two. None
of it binds on its own, and nothing that binds is stated only there.

The method and platform facts delivered with this template bind this build as fully as the sections
named above.

- **Business Rule 1 and the posting checks each offer a developer-facing choice between two named
  routes. Ask both, never infer either, never default either.** They are items 4 and 5 under
  *Information and conditions you need to supply*. A route chosen restricts the build to that
  route; if a route cannot pass its own checks, record the check as not passed and say what stopped
  it. Do not switch to the other route to make a check pass. The developer chose, and the other
  route has a different coverage.
- **Before you draft the design, fetch `time-off-ledger-schema`, the table template this realizes, with `get_template` (this run's route, `have_method=true`; its master is all you need), never by reading its file, for the fields, the full
  text of all ten Business Rules, and the reasoning behind them.** This file restates their
  outcomes. That file is where the field names, types and table shapes live.
- **Apply every standard delivered with this template.** Naming, audit columns, error handling, query
  style and how the work divides into procedures all come from there and never from this file.
- **Read the existing Data Macro on `tblTimeOffEntry` before writing anything to it** (fact
  `data-macro-rules`). The audit stamping lives in that document. Merge the new logic into it, and
  where the two do not obviously share a branch, ask before writing.
- **A function a Data Macro calls must exist in the back end and in every front end.** Say in the
  build record where each one went, and tell the developer that a changed function has to be
  re-imported everywhere, because nothing keeps the copies in step.
- **Where a transaction is used, begin it on a `Workspace` and take the `Database` every read and
  write goes through from that same `Workspace`, never from `CurrentDb`, as `error-handling.md`
  documents. Never read through a domain function a table the transaction writes** (fact
  `domain-function-transaction`). The balance sum is exactly such a read: make it a recordset on the
  transaction's `Database`.
- **For checks 8, 10, 17, 31 and 32, where the tool building this cannot produce two literally simultaneous
  processes, that is not evidence the check cannot run.** Open a second, independent connection to
  the same file and use it to insert the competing row between the first connection's own check and
  its own write. That forces the engine to show the real collision, rather than leaving you to infer
  one. See `_template-schema.md` §12.2. **The competing row has to land after the first writer has
  finished the read its decision rests on.** One that lands earlier is read by the first writer, so
  every build passes, and the check has shown nothing. Run it once on a copy with the protection
  against two at once switched off, and record what happened, as `_template-schema.md` §12.2
  requires; a copy carries test-only changes, and the record says what they were. **Give the
  posting code one empty procedure, called at each point a check needs a competing row** (after a
  check has passed, between the two halves of a replacement). The delivered database keeps it
  empty, and only a check copy fills it in. Without such a point there is no way to place the
  competing row at the moment these checks need.
- **Where the developer chose to have hours posted when the database opens**, follow
  `standards/startup-conventions.md`: the open-time macro's one action is the `Startup()` function,
  and the posting is one step inside it. Read any open-time macro the database already has before
  writing one, and add to it. Say in the build record, and to the developer when you hand the
  database over, that an open by an automated tool posts nothing and that a person's open was or
  was not tried. A scheduled run is not built: say what the task has to call.
- **A recovery from an expected refusal by the table must not reuse a variable that now holds a
  different statement.** Where the engine refuses a second cancellation or replacement and your code
  recovers by looking up the entry that is already there, give the lookup its own statement.
  Re-running the variable that held the failed insert runs that insert again.
- **A replacement is written in one transaction.** Begin it on a `Workspace` and take the `Database`
  both entries are written through from that same `Workspace`, as the transaction guard in
  `error-handling.md` says. The Data Macro checks run in a function Access calls on its own
  connection; that function must read through the transaction's database while a replacement is
  being written, or it cannot see the corrected entry and will judge the cancellation without it.
  Hand the function the transaction's database for the length of the replacement and take it away
  again, including when the replacement fails.
- **Never infer an answer that belongs to the developer**, not from what the database looks like,
  not from reasoning that makes an answer seem obvious. Where a check exists to answer a question,
  run the check at the point the sequence calls for it rather than working the answer out yourself.
- **Ask for the eight things under *Information and conditions you need to supply*,** one at a time,
  through the interactive selection control where the answer is a choice, and as a question phrased
  in plain language where it is a name or a number. Two of them are gates: a database in real use
  with no backup stops the build, and permission to change the tables is asked last, immediately
  before anything is changed and after the design has been approved. The accrual schedule in item 6
  is a value question: propose a sample if asked to, state it back in the design you present for
  approval, and accept whatever the developer confirms. Item 7 is a choice among three, and the
  developer may ask for something else, which is theirs to have as an extension.
- **After the seventh thing and before you present the design, offer the `Explore options` step**
  (`_template-schema.md` §12.5) over the list under *Free to choose alternatives*, and nothing
  outside it. The two routes for Business Rule 1 and the posting checks are not on that list. They
  were already asked directly.
- **The error-handling standard asks its own question before you write code.** Ask it, name the
  preferred answer, and never decide it. Where the answer is the shared logger and the database has
  none, the developer builds `error-logging-outcome-first` first; do not write a call to a logger
  that is not there.
- **Surface every `house_assumptions` entry and every `warnings` entry in the front matter** and get
  the developer's answer on each before building. Surfacing one is not asking about it: put each
  through the selection control and wait.
- **Checks 2 and 23 each carry a condition stated in
  advance.** Where the VBA route was chosen for Business Rule 1, check 2's expected result is stated
  in the check itself, and recording that expected, disclosed outcome is what "passed" means for it.
  Where the developer chose to run the posting themselves, the trigger part of check 23 does not
  apply, and the entry says so.

## Extra Options

*Named optional extensions, none of them filled in for an engagement.*

- **Role-based accrual, role history, going below zero, a stored balance, hooking into an existing
  employee table, date-range entries, holidays, and carry-over and expiry**, all named in the
  schema's own *Extra Options*. This template does not build any of them, and the checks above
  describe the base build only.