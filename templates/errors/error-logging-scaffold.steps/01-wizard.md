---
step: 01-wizard
title: "The wizard: the questions asked before anything is built"
platform_facts: []
route: both
when: before you draft the design
---

**Who reads this:** the AI assistant, which asks the developer these questions before it drafts the design.

## Wizard

Six questions, preceded by the entry question (`templates/_template-schema.md` §10.6) that asks
whether you want to answer them at all. Each one changes what the generated code does.

**If the entry question is answered `Just build it`:** every step here has a preferred choice, so
all six are used and nothing further is asked. State them before acting on them.

**Before step 1**, surface the two build-wide warnings from this template's front-matter: VBA does
not run outside a Trusted Location, and `LogError` is written so that no error inside it can ever
escape into the handler that called it. The other warnings belong to individual steps and appear
there.

### Step 1 — Do you want error handling and logging?

**Ask:** Do you want error handling and logging?

| Option | Short description |
|---|---|
| `Yes, include error handling and logging` | Carry on to the rest of the questions. |
| `No, do not include logging and error handling` | The code is written with no error handling and nothing recorded. |

**Preferred:** `Yes, include error handling and logging` — the standards these templates follow
expect a handler on every procedure.

**Skip when:** never. This is the first question after the entry question.

<details>
<summary>Tell me more about error handling and logging</summary>

**With a handler**, a failure is caught, recorded, and the procedure leaves through a single exit
that releases whatever it had open. You get a record naming the error, the module, the procedure,
and — if you want it — the line.

**Without one**, Access shows its own error dialog carrying a number the person using the
application can do nothing with, and the procedure stops on the line that failed. Anything it had
open stays open, and any work in progress is left as it was.

Choosing "No" is a real answer, not a discouraged one. It is the right call for a one-off utility
or a demonstration, where a log would be clutter. It is the wrong call for anything somebody else
relies on.

**Choosing "No" ends this wizard only.** Any other wizard in this template is asked independently,
and nothing about this answer carries into them.

Nothing here is permanent. Run the wizard again later and answer differently.

</details>

### Step 2 — How should the handler report the error?

**Ask:** When something goes wrong, how should it be reported?

| Option | Short description |
|---|---|
| `Write the names into the code` | Each module and procedure carries its own name, written at the same time as the code. |
| `Show a message box, keep no record` | Tells the person something went wrong and records nothing. |

**Preferred:** `Write the names into the code` — the first of the two ways to report an error,
in the order the standards these templates follow rank them.

**Skip when:** Step 1 was answered "No".

<details>
<summary>Tell me more about how the handler reports an error</summary>

These are the two options `standards/error-handling.md` ranks, where they are called **named
constants** and the plain **message box**. The first calls `LogError`. The second does not record
anything at all.

**Writing the names into the code.** A generator writes `MODULE_NAME` and `PROC_NAME` at the same
moment it writes the procedure, so they cannot be out of step with it — there is no point at which
the wrong name could be produced. It needs nothing installed, nothing enabled, and nothing
configured on the machine it runs on, so it behaves identically for everyone. If somebody renames a
procedure later, `PROC_NAME` is one line above the thing they renamed.

**The message box.** Once the person closes the box there is no trace the error happened — no
record to search, nothing to count, no way to find out afterward how often it occurs. Reasonable
for a one-off utility; not reasonable for anything someone else depends on.

</details>

### Step 3 — Where should errors be recorded?

**Ask:** Where should errors be recorded?

| Option | Short description |
|---|---|
| `A table in this database` | Errors go to `tblErrorLog`, which you can open, sort, and filter. |
| `A text file` | Errors are appended as lines of text to an external file. |
| `The table, falling back to the text file` | Uses the table, and writes to the file when the table cannot be reached. |

**Preferred:** `The table, falling back to the text file` — a table alone cannot record the one
error you would most want kept (the back end being unreachable), and the standards layer settles
the *shape* of a handler, not where its logger writes.

**Skip when:** Step 1 was answered "No", or Step 2 chose the message box — neither records
anything.

<details>
<summary>Tell me more about where errors are recorded</summary>

**A table is queryable.** "Every error in this module last week" is a question you can answer in a
few seconds, and you can count, sort, and filter the history.

**A table cannot record the failure where the database itself is the problem.** In a split
database, if the back end is unreachable — the share is down, the network dropped, the file is
locked — then the write to a log table in the back end fails too. That is precisely the error you
would most want kept, and it is the one a table-only log is guaranteed to miss.

**A text file is always writable** and survives a database that will not open at all. What you give
up is querying: you open it and read it, and that is all.

**The third option costs a few lines and covers both** — the table when it can be reached, the file
when it cannot. If you are unsure, this is the one that loses nothing.

</details>

### Step 4 — Where should that live?

**Ask (where Step 3 kept a table):** Which file holds the log table?

| Option | Short description |
|---|---|
| `The back end` | One shared log, holding everybody's errors together. |
| `Each front end` | Every person has their own log, inside their own copy of the application. |

**Preferred:** `The back end` — this template's own, following the general rule that tables live in
the file that holds the tables.

**Ask (where Step 3 kept a file):** Where should the log file go?

| Option | Short description |
|---|---|
| `The front end's folder` | Beside the copy of the application that person is running. |
| `Your Documents folder` | The standard Windows per-person location. |
| `The back end's folder` | Beside the shared data file, so everyone's files land together. |

**Preferred:** `The front end's folder` — the standards these templates follow treat a local log as
a per-user working folder, created per front end on each machine.

**Skip when:** Step 3 was skipped. Where Step 3 chose the fallback option, **this step is asked
twice** — once for the table's home, once for the file's folder.

<details>
<summary>Tell me more about where the log lives</summary>

**Table — the back end.** One place to look, and you see everybody's errors side by side, which is
how you notice that three people hit the same thing this morning. The catch is the one named in the
previous step: a log in the back end cannot be written when the back end is what failed. Choosing
this also adds one build step: the table is created once in the back end, then linked into each
front end — and into every front end, not just the first one.

**Table — each front end.** Always writable, because it is in the file that person is already
running. The cost is that there is no whole picture: to see what has been happening you have to
visit each machine and collect the copies.

On a single-file database this question does not arise — there is one file.

**File — the front end's folder.** Keeps the log beside the copy that produced it, which is exactly
what you want when you are helping one person with one problem. Where each person has their own
front end, this is per-person automatically. Where several people run the same copy from a share,
it is not, and Windows may refuse the write anyway.

**File — Documents.** Always writable by the person running the application, and it survives a
database that will not open at all. Note that a Documents folder redirected into OneDrive is not
where the plain path expects it; the code falls back to the temporary folder if it cannot find the
one it was told to use, so you get a log either way — just not always where you first looked.

**File — the back end's folder.** Everyone's files land together, which is convenient right up
until the shared folder is the thing that is unreachable. The code reads that folder off the link
the front end already has, so it only works where the log table is a linked table.

</details>

### Step 5 — Should the log name the failing line?

**Ask:** Should the log say which line failed?

| Option | Short description |
|---|---|
| `Yes, name the line` | Procedures get numbered lines, so the log can point at the one that failed. |
| `No, the procedure is enough` | The log names where the failure happened, but not the place inside it. |

**Preferred:** `Yes, name the line` — the standards these templates follow, on line numbering.

**Skip when:** Step 1 was answered "No", or Step 2 chose the message box.

<details>
<summary>Tell me more about line numbers</summary>

`Erl` is the VBA function that returns the number of the last numbered line executed before the
failure. In a numbered procedure it tells you where; in an unnumbered one it returns `0`, and the
log records which procedure failed but not the place inside it. On a short procedure that is often
enough. On a long one it is the difference between reading the code and searching it.

Numbering by hand is tedious and goes stale. If you have a tool that applies line numbering
automatically, it can normalize the numbering on import and after each edit.

**Stale numbers are worse than no numbers**, because they point confidently at the wrong line. If
you are not going to re-normalize after editing, choose No — an honest gap beats a confident lie.

Line numbering is itself a house choice, and `standards/error-handling.md` says so. A practice that
rejects line numbers replaces that file; `Erl` returns 0, the logger records the module and
procedure, and nothing else changes.

**Either answer works.** With line numbers, the log names the line that failed; without them, it
names the procedure only. The logger is built for both.

Line numbers can be added to your code by hand or by a tool you run over the code in the VBA editor.
If I put the code into your database directly, you can add line numbers yourself, or use a tool to
add them later.

</details>

### Step 6 — What does the person at the keyboard see?

**Ask:** When something goes wrong, what does the person using the application see?

| Option | Short description |
|---|---|
| `A short message, with a reference number when there is one` | Plain wording, nothing technical, quoting the log entry's number where the error reached the table. |
| `The full technical detail` | Error number, description, module, procedure, and line. |
| `Nothing` | The error is recorded silently and the person sees no message. |

**Preferred:** `A short message, with a reference number when there is one` — this template's own.

**Skip when:** Step 1 was answered "No". Where Step 2 chose the message box, only the first two
options apply: nothing was recorded, so there is never a reference number to quote, and the message
box is the only report there is — "Nothing" would mean nothing happens at all.

<details>
<summary>Tell me more about what the person sees</summary>

**A short message** is what most people want. They cannot act on error 3021, and a wall of
technical text mainly invites them to try something that makes it worse.

**The reference number** rides along with that short message wherever there is one. It turns a
support call from "it broke" into "it broke, reference 4127" — and that number opens the exact row.
Where the error went to a text file rather than the table there is no number, and the message
quietly drops it rather than inventing one, so this is not a separate choice you have to make.

**Full technical detail** is right while you are the only user and wrong the moment anybody else
is.

**Nothing at all** means the failure is invisible to the person it happened to. They will assume
their work saved, and retype it only when they discover otherwise. Choose it only where something
else tells them — a status bar, a validation message, a screen that visibly does not advance.

One thing the generated code does regardless: **it never claims a record was kept when the write
failed.** If nothing could be recorded, the message says so and asks the person to note what they
were doing.

</details>
