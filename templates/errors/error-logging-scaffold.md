---
template: error-logging-scaffold
title: Error Logging — Set-up Wizard and Logger
domain: errors
type: vba-scaffold
version: 0.5.0
status: draft
implements: error-log-schema
requires_tables:
  - tblErrorLog
standards_layer:
  - error-handling
  - naming-conventions
  - query-style
  - design-principles
target_module: modErrorLog
wizard: true
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
  - The "look the names up while the code runs" option (Step 2) reads the Visual Basic Editor,
    which requires "Trust access to the VBA project object model" in Trust Center - Macro
    Settings. That setting is OFF by default and is per machine, not per file, so code that works
    on the machine it was written on can fail on someone else's. Confirm it is on everywhere the
    application runs before choosing that option.
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
---

# Error Logging — Set-up Wizard and Logger

## Intent

Two things: a short **wizard** that walks a developer through the six decisions error logging
actually involves, one question at a time — and the **working code** those answers produce, built
around `LogError`, the shared logger that `standards/error-handling.md` names as the call at the
end of every error handler.

Until now the reasoning behind those six decisions lived in prose that a reader met all at once,
before they had chosen anything. The wizard does not remove any of it. It moves each piece to the
step where it is actually actionable, and keeps it closed until the reader asks.

**Nothing here is an Access wizard.** No form is built, nothing is installed in the database to run
it, and no artifact is left behind afterward. The assistant asks the six questions in conversation;
this file is where the questions, the options, and the explanations are written down. See
`templates/_template-schema.md` §10 for the format.

**The procedures below are complete, working code, not skeletons.** Each one carries the wizard's
alternatives as commented blocks, marked `[WIZARD Step n]`, with the preferred choice left uncommented — so
the module compiles and runs as-is, and a different answer is a matter of moving which block is
commented out.

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
| `error-log-schema`'s `tblErrorLog` | The table `LogError` writes to. `CreateErrorLogTable` below builds it; the paired template describes it |
| A Trusted Location | VBA does not run outside one, and there is no error message when it doesn't |
| **Trust access to the VBA project object model** | Only for Step 2's "look the names up while the code runs" option. Off by default; per machine |

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

## Wizard

Six questions, preceded by the entry question (`templates/_template-schema.md` §10.6) that asks
whether you want to answer them at all. Each one changes what the generated code does.

**If the entry question is answered `Just build it`:** every step here has a preferred choice, so
all six are used and nothing further is asked. State them before acting on them.

**Before step 1**, surface the two build-wide warnings from this template's front-matter: VBA does
not run outside a Trusted Location, and every file that runs code needs its own copy of this
module. The other warnings belong to individual steps and appear there.

### Step 1 — Do you want error handling and logging?

**Ask:** Do you want error handling and logging?

| Option | Short description |
|---|---|
| `Yes, include error handling and logging` | Carry on to the rest of the questions. |
| `No, do not include logging and error handling` | The code is written with no error handling and nothing recorded. |

**Preferred:** `Yes, include error handling and logging` — from `standards/error-handling.md`, which
describes a handler on every procedure.

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
| `Look the names up while the code runs` | The code finds out for itself which module and procedure failed, so there is nothing to keep current. |
| `Show a message box, keep no record` | Tells the person something went wrong and records nothing. |

**Preferred:** `Write the names into the code` — from
`standards/error-handling.md`, "Three ways to report an error", option 1.

**Skip when:** Step 1 was answered "No".

<details>
<summary>Tell me more about how the handler reports an error</summary>

These are the three options `standards/error-handling.md` ranks, where they are called **named
constants**, **VBE reflection**, and the plain **message box**. The first two do the same thing and
differ only in how the handler learns which module and procedure it is sitting in; both call
`LogError`. The third does not record anything at all.

**Writing the names into the code.** A generator writes `MODULE_NAME` and `PROC_NAME` at the same
moment it writes the procedure, so they cannot be out of step with it — there is no point at which
the wrong name could be produced. It needs nothing installed, nothing enabled, and nothing
configured on the machine it runs on, so it behaves identically for everyone. If somebody renames a
procedure later, `PROC_NAME` is one line above the thing they renamed.

**Looking the names up while the code runs.** Nothing to keep current when a procedure is renamed,
because nothing is stored. The cost is a dependency: it reads the Visual Basic Editor's own object
model, which requires **Trust access to the VBA project object model** (Trust Center → Macro
Settings). That setting is **off by default**, and it is a per-machine Access setting rather than a
property of your file — so this can work perfectly on the machine you wrote it on and fail on
somebody else's, for a reason nothing in the code hints at. Confirm it is on everywhere the
application will run before choosing this.

Both options put the **same thing** in the log: a bare module name such as `modInventory`. The
lookup form has been run and checked against this, so the log carries one format either way and you
can change your mind later without splitting your history.

**The message box.** Once the person closes the box there is no trace the error happened — no
record to search, nothing to count, no way to find out afterward how often it occurs. Reasonable
for a one-off utility; not reasonable for anything someone else depends on.

</details>

### Step 3 — Where should errors be recorded?

**Ask:** Where should errors be recorded?

| Option | Short description |
|---|---|
| `A table in this database` | Errors go to `tblErrorLog`, which you can open, sort, and filter. |
| `A text file` | Errors are appended as lines of text to a file. |
| `The table, falling back to the text file` | Uses the table, and writes to the file when the table cannot be reached. |

**Preferred:** `A table in this database` — this template's own; the standards layer settles the
*shape* of a handler, not where its logger writes.

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

**Preferred:** `The back end` — this template's own, following the general rule in
`error-log-schema` that tables go in the back end.

**Ask (where Step 3 kept a file):** Where should the log file go?

| Option | Short description |
|---|---|
| `The front end's folder` | Beside the copy of the application that person is running. |
| `Your Documents folder` | The standard Windows per-person location. |
| `The back end's folder` | Beside the shared data file, so everyone's files land together. |

**Preferred:** `The front end's folder` — from `standards/startup-conventions.md` §4, which classes a
local log as a per-user working folder, created per front end on each machine.

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

**Preferred:** `Yes, name the line` — from `standards/error-handling.md`, "Line numbering".

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

## Procedures

Put these at the top of `modErrorLog`, below `Option Explicit`:

```vba
Option Compare Database
Option Explicit

' [STANDARDS - error-handling.md] This module's own name, for handlers inside it.
Private Const MODULE_NAME           As String = "modErrorLog"

' [SCAFFOLD] The table and file the logger writes to, in ONE place each.
Private Const ERROR_LOG_TABLE       As String = "tblErrorLog"
Private Const ERROR_LOG_FILE        As String = "ErrorLog.txt"

' [SCHEMA] tblErrorLog.ErrorDescription is Text(255). A longer description is
'          truncated to fit, never allowed to fail the write (schema Business Rule 4).
Private Const ERROR_DESCRIPTION_MAX As Long = 255
```

### LogError — `Public Sub` (the logger every handler calls)

The procedure `standards/error-handling.md` names at the end of every error handler. It takes the
module, the procedure, and the line, reads the error itself, records it wherever Step 3 said, and
tells the person whatever Step 6 said.

**This procedure deliberately does not follow the house handler pattern**, and it is the only place
in the library that doesn't. `error-handling.md` allows `On Error Resume Next` inside `Cleanup:`
and nowhere else — but this procedure *is* what a handler calls, so it cannot call a handler
itself. An error raised here would re-enter the handler that called it. It is guarded end to end
instead, and each failure steps down to the next option rather than propagating. See
`error-handling.md`, "The logger itself is the one exception".

**The order of the first two statements is load-bearing.** `Err` is cleared by any `On Error`
statement — including the guard on the third line. Read the error first, guard second. Swapping
them logs error 0 with an empty description, on every call, silently.

```vba
Public Sub LogError(ByVal sModule As String, _
                    ByVal sProcedure As String, _
                    ByVal lLine As Long)
    ' [SCAFFOLD] Called from an errHandler block that is already dealing with a failure.
    '            Nothing in here may raise an error of its own.
    Dim lErrNumber      As Long
    Dim sErrDescription As String
    Dim lLogID          As Long
    Dim bRecorded       As Boolean

    ' [SCAFFOLD] Read Err FIRST. The On Error statement below clears it, and so would
    '            any other On Error statement, Resume, Exit Sub, or Exit Function. These
    '            two lines must stay above the guard.
    lErrNumber = Err.Number
    sErrDescription = Err.Description

    ' [STANDARDS - error-handling.md] The documented exception: the logger is guarded
    '            rather than handled, because a handler here would call itself.
    On Error Resume Next

    ' [SCHEMA Business Rule 4] Truncate to fit rather than fail on length.
    If Len(sErrDescription) > ERROR_DESCRIPTION_MAX Then
        sErrDescription = Left$(sErrDescription, ERROR_DESCRIPTION_MAX)
    End If

    ' [WIZARD Step 3 - where errors are recorded] >>> keep ONE of the three blocks <<<

    ' --- "A table in this database" (preferred) ---
    bRecorded = WriteErrorToTable(lErrNumber, sErrDescription, sModule, sProcedure, _
                                  lLine, lLogID)

    ' --- "A text file" ---
    ' bRecorded = WriteErrorToFile(lErrNumber, sErrDescription, sModule, sProcedure, lLine)

    ' --- "The table, falling back to the text file" ---
    ' bRecorded = WriteErrorToTable(lErrNumber, sErrDescription, sModule, sProcedure, _
    '                               lLine, lLogID)
    ' If Not bRecorded Then
    '     bRecorded = WriteErrorToFile(lErrNumber, sErrDescription, sModule, sProcedure, lLine)
    ' End If

    ShowErrorToUser lErrNumber, sErrDescription, sModule, sProcedure, lLine, lLogID, bRecorded
End Sub
```

### WriteErrorToTable — `Private Function` → `Boolean`

Appends one row to `tblErrorLog` and hands back the new `ErrorLogID` so Step 6 can show it as a
reference number. **Returns `False` rather than raising** when the table cannot be reached — which
is what lets the fallback option in `LogError` step down to the file.

It appends through a recordset rather than an `INSERT` statement on purpose. An error description
containing an apostrophe would break a concatenated `INSERT`, and this is the one procedure in the
application that must not fail on the contents of its own input.

```vba
Private Function WriteErrorToTable(ByVal lErrNumber As Long, _
                                   ByVal sErrDescription As String, _
                                   ByVal sModule As String, _
                                   ByVal sProcedure As String, _
                                   ByVal lLine As Long, _
                                   ByRef lLogID As Long) As Boolean
    ' [SCAFFOLD] Guarded throughout; reports success as a Boolean instead of raising.
    Dim db    As DAO.Database
    Dim rs    As DAO.Recordset
    Dim sUser As String

    On Error Resume Next
    lLogID = 0

    ' [SCHEMA Business Rule 5] Never empty - the column is Required, and an empty string
    '            would be refused by the very write we are trying to guarantee.
    sUser = Environ$("USERNAME")
    If Len(sUser) = 0 Then sUser = "Unknown"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(ERROR_LOG_TABLE, dbOpenDynaset, dbAppendOnly)

    ' [SCAFFOLD] The table may be missing, or the back end unreachable. Either way this
    '            is a False, not an error - the caller has somewhere else to try.
    If Err.Number <> 0 Then GoTo Cleanup
    If rs Is Nothing Then GoTo Cleanup
    Err.Clear

    ' [STANDARDS - query-style.md] A recordset append, not a concatenated INSERT: an
    '            apostrophe in Err.Description must not be able to break the write.
    rs.AddNew
    rs!ErrorNumber = lErrNumber
    rs!ErrorDescription = sErrDescription
    rs!ModuleName = sModule
    rs!ProcedureName = sProcedure
    rs!ErrorLineNumber = lLine
    rs!ErrorOccurredOn = Now()
    rs!ErrorUser = sUser
    rs.Update

    If Err.Number = 0 Then
        ' [SCAFFOLD] The row is in. Read the reference number afterwards - if that part
        '            fails, the record still stands and only the number is missing.
        WriteErrorToTable = True
        rs.Bookmark = rs.LastModified
        lLogID = Nz(rs!ErrorLogID, 0)
    End If

Cleanup:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function
```

### WriteErrorToFile — `Private Function` → `Boolean`

Appends one tab-separated line to the log file. Used on its own where Step 3 chose the file, and as
the fallback where Step 3 chose both. **The timestamp at the start of the line is the reference**
in this case — there is no `ErrorLogID` to quote.

```vba
Private Function WriteErrorToFile(ByVal lErrNumber As Long, _
                                  ByVal sErrDescription As String, _
                                  ByVal sModule As String, _
                                  ByVal sProcedure As String, _
                                  ByVal lLine As Long) As Boolean
    ' [SCAFFOLD] Guarded throughout; reports success as a Boolean instead of raising.
    Dim iFile    As Integer
    Dim sPath    As String
    Dim sUser    As String
    Dim sEntry   As String
    Dim bWritten As Boolean

    On Error Resume Next

    sPath = GetErrorLogFilePath()
    If Len(sPath) = 0 Then Exit Function

    sUser = Environ$("USERNAME")
    If Len(sUser) = 0 Then sUser = "Unknown"

    ' [SCAFFOLD] Tab-separated so the file opens straight into a spreadsheet if anyone
    '            needs to sort it. The leading timestamp is what a support call quotes.
    sEntry = Format$(Now(), "yyyy-mm-dd hh:nn:ss") & vbTab & _
             sUser & vbTab & _
             sModule & "." & sProcedure & vbTab & _
             "line " & lLine & vbTab & _
             lErrNumber & vbTab & _
             sErrDescription

    Err.Clear
    iFile = FreeFile
    Open sPath For Append As #iFile
    If Err.Number = 0 Then
        Print #iFile, sEntry
        bWritten = (Err.Number = 0)
        Close #iFile
    End If

    WriteErrorToFile = bWritten
End Function
```

### GetErrorLogFilePath — `Private Function` → `String`

Works out the full path of the log file from Step 4's answer, and falls back to the temporary
folder if that location is not there — so a misconfigured path costs you the folder you wanted, not
the record.

```vba
Private Function GetErrorLogFilePath() As String
    ' [SCAFFOLD] Guarded: a bad path returns the fallback, never an error.
    Dim sFolder As String

    On Error Resume Next

    ' [WIZARD Step 4 - where the log file goes] >>> keep ONE of the three blocks <<<

    ' --- "The front end's folder" (preferred; standards/startup-conventions.md section 4
    '     classes a local log as a per-user working folder, created per front end) ---
    sFolder = CurrentProject.Path

    ' --- "Your Documents folder" ---
    ' [SCAFFOLD] A Documents folder redirected into OneDrive is not on this path; the
    '            fallback below catches that and the log lands in the temp folder.
    ' sFolder = Environ$("USERPROFILE") & "\Documents"

    ' --- "The back end's folder", read off the link the front end already has. Works
    '     only where ERROR_LOG_TABLE is a LINKED table - a local one has no Connect. ---
    ' Dim sConnect As String
    ' sConnect = CurrentDb.TableDefs(ERROR_LOG_TABLE).Connect
    ' sFolder = Mid$(sConnect, InStr(sConnect, "DATABASE=") + 9)
    ' sFolder = Left$(sFolder, InStrRev(sFolder, "\") - 1)

    ' [SCAFFOLD] Last resort. The temp folder always exists and is always writable, so
    '            there is no path by which this procedure returns nothing usable.
    If Len(sFolder) = 0 Then sFolder = Environ$("TEMP")
    If Len(Dir$(sFolder, vbDirectory)) = 0 Then sFolder = Environ$("TEMP")

    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    GetErrorLogFilePath = sFolder & ERROR_LOG_FILE
End Function
```

### ShowErrorToUser — `Private Sub`

Step 6's answer, in one place. It is handed everything it might need to show and uses only what the
chosen wording calls for.

**It never claims a record was kept when the write failed.** `bRecorded` is what the write actually
returned, and the wording follows it.

```vba
Private Sub ShowErrorToUser(ByVal lErrNumber As Long, _
                            ByVal sErrDescription As String, _
                            ByVal sModule As String, _
                            ByVal sProcedure As String, _
                            ByVal lLine As Long, _
                            ByVal lLogID As Long, _
                            ByVal bRecorded As Boolean)
    ' [SCAFFOLD] Guarded: a failure to show the message must not replace the failure we
    '            were called about.
    Dim sMsg As String

    On Error Resume Next

    ' [WIZARD Step 6 - what the person at the keyboard sees] >>> keep ONE block <<<

    ' --- "A short message, with a reference number when there is one" (preferred) ---
    If bRecorded And lLogID > 0 Then
        sMsg = "Something went wrong. The problem has been recorded." & vbCrLf & vbCrLf & _
               "Reference number: " & lLogID
    ElseIf bRecorded Then
        ' [SCAFFOLD] Recorded to the file, so there is no number to quote. Say the true
        '            thing rather than invent one.
        sMsg = "Something went wrong. The problem has been recorded."
    Else
        sMsg = "Something went wrong, and it could not be recorded. " & _
               "Please make a note of what you were doing."
    End If

    ' --- "A short message that something went wrong" ---
    ' If bRecorded Then
    '     sMsg = "Something went wrong. The problem has been recorded."
    ' Else
    '     sMsg = "Something went wrong, and it could not be recorded. " & _
    '            "Please make a note of what you were doing."
    ' End If

    ' --- "The full technical detail" ---
    ' sMsg = "Error " & lErrNumber & ": " & sErrDescription & vbCrLf & _
    '        "In: " & sModule & "." & sProcedure & vbCrLf & _
    '        "Line: " & lLine
    ' If Not bRecorded Then
    '     sMsg = sMsg & vbCrLf & vbCrLf & _
    '            "This could not be recorded. Please make a note of what you were doing."
    ' End If

    ' --- "Nothing" ---
    ' Exit Sub

    MsgBox sMsg, vbExclamation, "Unexpected problem"
End Sub
```

### CreateErrorLogTable — `Public Function` → `String` (run once, before anything else)

Builds `tblErrorLog` through DAO, exactly as the paired schema template describes it. Idempotent:
an existing table is reported and left alone, so re-running is safe.

Run it once in whichever file Step 4 chose. Called as a bare statement it shows a message box;
called as `sResult = CreateErrorLogTable(True)` it returns the same text with no dialog — which is
what an automated caller needs, since nothing is there to click a box away. **Passing `True` is
what suppresses the dialog**, not assigning the return value: VBA gives a procedure no way to know
whether it was called as a function or as a statement.

Three things here are not optional, and all three come from `templates/_materialization.md`: field
`Description`s are set **after** the table is appended (setting one before raises error 3219),
AutoNumber is a `dbLong` field with `dbAutoIncrField`, and the whole thing must be run from a
**Trusted Location** or Access disables the code with no message at all.

**`AllowZeroLength = True` on every text column is deliberate.** A Required text field that
disallows zero-length strings refuses an empty value — and `Err.Description` is sometimes empty.
That is exactly the kind of refusal schema Business Rule 2 forbids.

```vba
Public Function CreateErrorLogTable(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Creates tblErrorLog (schema template entity). Idempotent: an existing
    '            table is reported and skipped, so this is safe to re-run.
    '            bSilent:=True suppresses the message box and returns the same text.
    Dim db      As DAO.Database
    Dim tdf     As DAO.TableDef
    Dim fld     As DAO.Field
    Dim idx     As DAO.Index
    Dim sReport As String
    Dim bFailed As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb

    sReport = ERROR_LOG_TABLE & " created."

    On Error Resume Next
    Set tdf = db.TableDefs(ERROR_LOG_TABLE)
    On Error GoTo errHandler
    If Not tdf Is Nothing Then
        sReport = ERROR_LOG_TABLE & " already exists and was left exactly as it is."
        GoTo Cleanup
    End If

    Set tdf = db.CreateTableDef(ERROR_LOG_TABLE)

    ' [SCAFFOLD + _materialization.md rule 3] AutoNumber is dbLong + dbAutoIncrField.
    Set fld = tdf.CreateField("ErrorLogID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ErrorNumber", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    ' [SCHEMA Business Rule 2] AllowZeroLength on every text column: Err.Description is
    '            sometimes empty, and a Required text field that refuses "" would refuse
    '            the row.
    Set fld = tdf.CreateField("ErrorDescription", dbText, ERROR_DESCRIPTION_MAX)
    fld.Required = True
    fld.AllowZeroLength = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ModuleName", dbText, 100)
    fld.Required = True
    fld.AllowZeroLength = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ProcedureName", dbText, 100)
    fld.Required = True
    fld.AllowZeroLength = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ErrorLineNumber", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ErrorOccurredOn", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ErrorUser", dbText, 100)
    fld.Required = True
    fld.AllowZeroLength = True
    tdf.Fields.Append fld

    ' [SCHEMA] No DefaultValue anywhere on this table, deliberately. The ACE engine, not
    '          VBA, evaluates a default, and it cannot resolve Now() or Environ() there
    '          (_materialization.md rule 5) - the logger supplies every value itself.
    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ErrorLogID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCHEMA] What went wrong recently, and everything that ever failed in one module.
    Set idx = tdf.CreateIndex("ErrorOccurredOn")
    Set fld = idx.CreateField("ErrorOccurredOn")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("ModuleName")
    Set fld = idx.CreateField("ModuleName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCAFFOLD + _materialization.md rule 2] Descriptions in a SECOND pass, after the
    '            append - setting one during field-build raises runtime error 3219.
    db.TableDefs.Refresh
    SetFieldDescription db, "ErrorLogID", _
        "Surrogate key. Also the reference number shown to the user."
    SetFieldDescription db, "ErrorNumber", "Err.Number as captured by LogError."
    SetFieldDescription db, "ErrorDescription", _
        "Err.Description, truncated to fit rather than allowed to fail the write."
    SetFieldDescription db, "ModuleName", "The module the failing line sits in."
    SetFieldDescription db, "ProcedureName", "The procedure the failing line sits in."
    SetFieldDescription db, "ErrorLineNumber", _
        "Erl - the last numbered line before the failure. 0 where the procedure is unnumbered."
    SetFieldDescription db, "ErrorOccurredOn", "When the logger recorded the error."
    SetFieldDescription db, "ErrorUser", "Who was running the code. 'Unknown' if unavailable."

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing
    CreateErrorLogTable = sReport
    ' [SCAFFOLD] One message, whatever happened - the success text or the error text,
    '            never both. Building the report first and showing it here prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] This procedure runs at set-up time, before the
    '            logger it is building exists, so it reports rather than calls LogError.
    bFailed = True
    sReport = "ERROR creating " & ERROR_LOG_TABLE & ": " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### SetFieldDescription — `Private Sub` (helper to the table build)

Puts a field's comment on as its Access `Description`, so the table documents itself in design
view. The property does not exist until it is created, which is why this is `CreateProperty` and
`Append` rather than a plain assignment.

```vba
Private Sub SetFieldDescription(db As DAO.Database, _
                                ByVal sField As String, _
                                ByVal sDescription As String)
    ' [SCAFFOLD + _materialization.md rule 2] Called only after db.TableDefs.Append.
    '            A Description set during field-build raises runtime error 3219.
    Dim prp As DAO.Property

    On Error Resume Next
    Set prp = db.TableDefs(ERROR_LOG_TABLE).Fields(sField) _
                .CreateProperty("Description", dbText, sDescription)
    db.TableDefs(ERROR_LOG_TABLE).Fields(sField).Properties.Append prp

    ' [SCAFFOLD] Guarded because a documentation comment is never worth failing a build
    '            for, and because appending a property that already exists raises 3367.
    Set prp = Nothing
End Sub
```

### LinkErrorLogTable — `Public Function` → `String` (run once in each front end)

The second half of a two-step build. `CreateErrorLogTable` makes the table in the back end; this
puts a link to it in the front end, so `LogError` has somewhere to write. **Needed only where Step 4
put the log table in the back end** — a log table that lives in each front end is local already, and
there is nothing to link.

Run it once in **every** front end, not only the one the application was built in. Idempotent: a
table that is already there is reported and left alone, so re-running is safe — and running it
against a front end that already has the link tells you where that link points.

**Without the link, nothing announces the problem.** `WriteErrorToTable` is guarded end to end so
that it can never raise into the handler that called it, which means a table it cannot reach makes
it return `False` silently: every error recorded nowhere, no message, nothing to find later. That is
why this procedure reads the linked table before reporting success rather than trusting the append —
a build step that announced "linked" without looking would be the same silence one level up.

The back-end path is optional. Left out, it is read off a link this front end already has — the same
`;DATABASE=<path>` rule as `CurrentBackEndPath` in `app-startup-scaffold.md`, written out here so
this module does not depend on that scaffold having been built too. A front end with no links yet
has nothing to read, and says so instead of guessing. Passing `True` as the second argument
suppresses the message box in the same way, and for the same reason, as `CreateErrorLogTable`'s.

```vba
Public Function LinkErrorLogTable(Optional ByVal sBackEndPath As String = "", _
                                  Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Links tblErrorLog from the back end into this front end. Needed only
    '            where Step 4 put the log table in the back end. Idempotent: a table
    '            that is already here is reported and left alone, so this is safe to
    '            re-run. bSilent:=True suppresses the message box, returning the text.
    Dim db       As DAO.Database
    Dim tdf      As DAO.TableDef
    Dim tdfOther As DAO.TableDef
    Dim sReport  As String
    Dim bFailed  As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [SCAFFOLD] Already here? Say which kind. A local table means Step 4 chose "each
    '            front end" and there is nothing to link; a link says where it points,
    '            so one aimed at the wrong back end is visible rather than assumed.
    On Error Resume Next
    Set tdf = db.TableDefs(ERROR_LOG_TABLE)
    On Error GoTo errHandler
    If Not tdf Is Nothing Then
        If Len(tdf.Connect) = 0 Then
            sReport = ERROR_LOG_TABLE & " is a local table in this file and was left alone."
        ElseIf Left$(tdf.Connect, 10) = ";DATABASE=" Then
            sReport = ERROR_LOG_TABLE & " is already linked to " & _
                      Mid$(tdf.Connect, 11) & " and was left alone."
        Else
            sReport = ERROR_LOG_TABLE & " is already linked (" & tdf.Connect & _
                      ") and was left alone."
        End If
        GoTo Cleanup
    End If

    ' [SCAFFOLD] No path given? Read it off a link this front end already has - the same
    '            ";DATABASE=" rule as CurrentBackEndPath in app-startup-scaffold.md.
    '            A front end with no links yet has nothing to read, and says so below.
    If Len(sBackEndPath) = 0 Then
        For Each tdfOther In db.TableDefs
            If Left$(tdfOther.Connect, 10) = ";DATABASE=" Then
                sBackEndPath = Mid$(tdfOther.Connect, 11)
                Exit For
            End If
        Next tdfOther
    End If

    If Len(sBackEndPath) = 0 Then
        bFailed = True
        sReport = "No back end could be found from this file's existing links. " & _
                  "Run this again with the back end's full path, for example: " & _
                  "LinkErrorLogTable(""C:\Data\App_BE.accdb"")"
        GoTo Cleanup
    End If

    Set tdf = db.CreateTableDef(ERROR_LOG_TABLE)
    tdf.Connect = ";DATABASE=" & sBackEndPath
    tdf.SourceTableName = ERROR_LOG_TABLE
    db.TableDefs.Append tdf
    db.TableDefs.Refresh

    ' [SCAFFOLD] Prove the link reads before calling it done. Reading a field count
    '            forces the link to resolve; a back end that has moved since fails here,
    '            with its own message, rather than at the first error nobody records.
    If db.TableDefs(ERROR_LOG_TABLE).Fields.Count = 0 Then
        bFailed = True
        sReport = ERROR_LOG_TABLE & " was linked to " & sBackEndPath & _
                  ", but the link does not read. Check the back end."
        GoTo Cleanup
    End If

    sReport = ERROR_LOG_TABLE & " linked to " & sBackEndPath

Cleanup:
    Set tdfOther = Nothing
    Set tdf = Nothing
    Set db = Nothing
    LinkErrorLogTable = sReport
    ' [SCAFFOLD] One message, whatever happened - the success text or the error text,
    '            never both. Building the report first and showing it here prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] This procedure runs at set-up time, before the
    '            link the logger writes through exists, so it reports rather than calls
    '            LogError.
    bFailed = True
    sReport = "ERROR linking " & ERROR_LOG_TABLE & ": " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

## Standards Layer

- **Error handling** — the handler pattern the generated code follows comes from
  `standards/error-handling.md`: the `errHandler:` and `Cleanup:` labels, the `LogError` call, and
  the `Resume Cleanup` / `Resume` chain. Wizard Step 2 chooses among the three reporting options
  that file ranks; Step 5 answers its line-numbering question. `LogError` itself is the one
  documented exception to the pattern, for the reason given at that procedure.
- **Naming conventions** — module, table, and field names follow OTS house style. A forked practice
  regenerates the same objects under its own.
- **Query style** — the one data write here is a recordset append rather than a concatenated
  `INSERT`, so the contents of an error description can never break it.
- **Design principles** — one job per procedure: read the error, write to a table, write to a file,
  work out a path, show a message, build a table. Each is separately replaceable, which is what
  makes the wizard's answers a matter of swapping one block rather than editing the logger.

## Extra Options

*Empty in the base template. Filled per client engagement.*

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
