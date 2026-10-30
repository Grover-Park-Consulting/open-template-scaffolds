---
template: error-logging-scaffold
title: Error Logging — VBA Scaffold
domain: errors
type: vba-scaffold
version: 0.10.0
status: stable
status_fingerprint: 23ca0aa40702
implements: error-logging-schema
requires_tables:
  - tblErrorLog
platform_facts: [target-file, sql-insert-truncation, host-helper-reuse]
steps: [01-wizard, 02-procedures-1, 03-procedures-2, 04-procedures-3, 05-checks]
standards_layer:
  - error-handling
  - naming-conventions
  - query-style
  - design-principles
  - startup-conventions
target_module: modErrorLog
new_procedures:
  - LogError
  - WriteErrorToTable
  - WriteErrorToFile
  - GetErrorLogFilePath
  - ShowErrorToUser
  - CreateErrorLogTable
  - SetFieldDescription
  - LinkErrorLogTable
warnings:
  - LogError runs inside a handler that is already dealing with a failure, so it must never raise
    an error of its own. An error escaping the logger re-enters the handler that called it and
    loops, or stops the application outright. Every write path here is guarded, and each failure
    steps down to the next option rather than propagating.
  - Any On Error statement clears the Err object, and so do Resume, Exit Sub, and Exit Function.
    LogError therefore copies Err.Number and Err.Description into local variables on its first two
    lines, BEFORE its own On Error Resume Next guard. Putting the guard first — the natural
    instinct when hardening a logger — silently records error 0 with an empty description on
    every single call. Do not reorder those lines.
  - VBA does not run at all outside a Trusted Location — Access silently disables it and nothing
    happens, with no error to explain why. In a split application this applies to every file that
    runs code, on every machine, so a front end copied somewhere untrusted fails for that one
    person and nobody else.
  - A log table in the back end cannot be written when the back end is unreachable, which is the
    error you most want kept. Step 3's fallback option exists for that case and nothing else.
  - Where the log table lives in the back end, each front end needs a link to it. Without that
    link every call to LogError records nothing and reports no problem — the write is guarded so
    that it can never raise into the handler that called it, so a missing link looks exactly like
    a working logger. LinkErrorLogTable creates the link; run it once in every front end.
  - This scaffold adds a table, a module, and a link rather than changing objects the developer
    already has, so less can go wrong here than in most templates. It still writes into their file.
    A build against a database in real use is preceded by a backup copy, and the developer is asked
    for one before anything is changed.
related:
  - "app-startup-scaffold — worth considering if you haven't built it yet. The back-end-unreachable
    case it handles is one of the more valuable situations to have this logger actually catch."
  - "standards/error-handling.md — worth a second look now that you have a logger. It offers two
    ways to report an error, and building this template is what makes the preferred one, calling a
    shared logger, available to you."
---

# Error Logging — VBA Scaffold

**Status last determined:** 2026-09-26.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Two things: a short **wizard** that walks a developer through the six decisions error logging
actually involves, one question at a time — and the **working code** those answers produce, built
around `LogError`, the shared logger that `standards/error-handling.md` names as the call at the
end of every error handler.

Until now the reasoning behind those six decisions lived in prose that a reader met all at once,
before they had chosen anything. The wizard does not remove any of it. It moves each piece to the
step where it is actually actionable, and keeps it closed until the reader asks.

**It is named for the Access wizards you already know**, but nothing of it runs inside your
database: no form is built, nothing is installed to run it, and nothing is left behind afterward.
The AI assistant asks the six questions in conversation; this file is where the questions, the
options, and the explanations are written down. See `templates/_template-schema.md` §10 for the
format.

**The procedures below are complete, working code, not skeletons.** Each one carries the wizard's
alternatives as commented blocks, marked `[WIZARD Step n]`, with the preferred choice left uncommented — so
the module compiles and runs as-is, and a different answer is a matter of moving which block is
commented out.

**The other version of this template — the outcome-first method, `error-logging-outcome-first` —
produces the same result from a specification rather than working code.** Either one can be built
against your own database, and they can be built one after the other, against separate copies, to
compare. Because the two promise the same result, they share one list of validation checks; that
file carries it, and *Validating the build* below says how to run it against this route.

> **If an AI assistant is running this wizard for someone:** ask the six questions one at a time
> and wait for each answer. Never infer one — not from the shape of the database, not from
> reasoning that makes an answer look obvious. Do not collapse the six into one upfront summary,
> even where every fact in it would be correct: the sequence exists so the developer decides each
> point, not so they approve the end state. Restate each decision in full at the moment you ask it,
> so it can be answered without scrolling back. An answer that departs from the standards layer
> holds for the rest of that run.

## Prerequisites

| Object | Role |
|---|---|
| `error-logging-schema`'s `tblErrorLog` | The table `LogError` writes to. `CreateErrorLogTable` below builds it; the paired template describes it |
| A Trusted Location | VBA does not run outside one, and there is no error message when it doesn't |

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

A **split database** is the usual shape for an Access application with more than one user: one file
holds the tables (the **back end**, normally on a shared drive), and each person runs their own copy
of a second file holding the forms, reports, and code (the **front end**), whose tables are links
pointing at the back end.

`modErrorLog` goes in **every file that runs code** — the front end each person uses, and the back
end too if any code runs there. It is ordinary VBA called by ordinary procedures, so it has to be
wherever those procedures are. Unlike a data macro's helper function, nothing invisible depends on
it being in a second place; but the copies must stay identical, and nothing keeps them in step for
you.

On a **single-file database** — one .accdb holding everything, a perfectly good choice for one
person — there is one file and one copy, and this does not arise.

## Standards Layer

- **Error handling** — the handler pattern the generated code follows comes from
  `standards/error-handling.md`: the `errHandler:` and `Cleanup:` labels, the `LogError` call, and
  the `Resume Cleanup` / `Resume` chain. Wizard Step 2 chooses among the two reporting options
  that file ranks; Step 5 answers its line-numbering question. `LogError` itself is the one
  documented exception to the pattern, for the reason given at that procedure.
- **Naming conventions** — module, table, and field names follow OTS house style. A forked practice
  builds the same objects under its own.
- **Query style** — the one data write here is a recordset append rather than a concatenated
  `INSERT`, so the contents of an error description can never break it.
- **Design principles** — one job per procedure: read the error, write to a table, write to a file,
  work out a path, show a message, build a table. Each is separately replaceable, which is what
  makes the wizard's answers a matter of swapping one block rather than editing the logger.

**To the AI assistant: once the build is reported finished, and only then, mention each entry under
`related` in the front matter** (`_template-schema.md` §7.1) — one line per entry, what it is and
why. This is not part of the build, never a gate, and never read before this point.

## Extra Options

*Named optional extensions, none of them filled in for an engagement.*

- **A context note from the caller** — an optional extra argument on `LogError` recording what the
  code was doing ("saving invoice 4471"), paired with the matching column in the schema template's
  Extra Options.
- **Rate limiting** — where one failure repeats in a loop, write the first occurrence and a count
  rather than thousands of rows.
- **A "report this" button** on the message box that copies the reference number and the last few
  log entries to the clipboard for a support email.
- **Forwarding** — a routine that copies new rows to a central SQL Server table, where one team
  supports several installations.
- **Reading the file back** — a small viewer for the text-file option, so the fallback log is as
  easy to look at as the table.
