---
step: 04-code
title: Write and import the code
platform_facts: [mcp-module-import, mcp-line-numbers, vba-import-xml-entities, domain-function-transaction, row-lock-errors, recordset-append-crash, app-startup-autoexec.autoexec-build]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Code

**The tables and Data Macros exist from steps 2 and 3: do not rebuild them here.** Write and import only the
application's own procedures, including every function a Data Macro from step 3 calls by name.

Write and import the procedures the build needs, divided as the design-principles standard says,
within *Free to choose alternatives*. Each function a Data Macro calls goes into the back end and
every front end. Where the developer chose to post when the database opens, the open-time work
follows the startup-conventions standard and the fact on the open-time macro below.

**From the master's instructions to the AI assistant, for this step.** They bind as the master's do.

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
