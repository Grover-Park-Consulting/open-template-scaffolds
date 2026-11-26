# Error Handling — OTS Default Standards Layer

**Who reads this:** the AI assistant, applying these rules to what it generates or a shop deciding
what to replace with their own.

If you are building something from a template, you don't need to read this file unless you are curious.
However, we have included clarifying comments to help you interpret what it says, just in case.

**Building from a template?** This file decides what happens when something goes wrong while
the code is running — what the user is told, and what gets written down.

You do not have to do anything with this information to use the templates. The rules apply to the code you receive
whether or not you read these files.

Using the standards as they are is the normal choice. If you came here because a template asked whether
you want to use these rules as they are or make them your own, you don't answer the question here.

*After you've read about the choice here, go back and answer it where it was asked.*

[`README.md`](README.md) lists all seven files if you want to see the others first.

> **This is the OTS default standards layer.** When you fork the library, replace this file with
> your own house error-handling pattern. You can also add to it here; **do so carefully to avoid
> breaking the template.** It covers any VBA generated alongside a
> schema or scaffold (templates that declare `standards_layer: [error-handling]`). The *structure*
> is the standard; the logger it calls is replaceable — a forked practice substitutes its own.

From here on, this file contains instructions for the AI.

**To the AI generating code:**

## One pattern, adopted whole

This file describes a complete error-handling pattern, and a forked practice substitutes a
complete one of its own. Either way it is a **single unit** — the labels, the reporting call, and
the Resume chain belong together and are used together. Choose one for a body of code and use all
of it.

Do not assemble a handler at the point of use from whichever parts seem convenient, and do not
introduce an exit the pattern doesn't have — `Resume Next`, a bare `Exit Function`, or simply
falling off the end — into a handler built on the Resume chain. Absorbing one expected error is the
recovery idiom below, with its own rules.

The failure this prevents is a quiet one: a handler that looks like it follows the standard while
behaving in a way the standard never describes, so nobody reading it thinks to check.

## Choose every handler from one table

Before you write a procedure, answer two questions about it. Together they give its handler. Put
the answer in the design spec, and name the shape in the handler's first comment:
`' [STANDARDS - error-handling.md] propagating shape`.

**1. Its role: who depends on how it ends?**

Ask in this order; the first yes is its role.

| Ask | If yes, its role is | For example |
|---|---|---|
| Is it what handlers call? | **The logger** | `LogError` |
| Does a Data Macro call it while a row is being saved? | **Called by a Data Macro** | `TimeOffEntryCheck`, `AuditUser` |
| Can its question be answered only by trying, under the three conditions in *When a failure is the answer*? | **Answer** | `BackEndIsReachable` |
| Does anything act on its having worked, or on the value it returns (a procedure, a query, a control's source), when its return value cannot say that it failed? | **Step** | `RecordScan` inside `ProcessScan`; `GetApplicablePayRate` in a query, whose Null already means "no rate" |
| Otherwise: nothing acts on its outcome, or its return value says whether it worked (`False`, a report the caller reads) | **Entry point** | an event procedure, `Startup`, `ProcessScan`, a setup procedure that returns its report |

**An entry point may also be called by code**, when it hands its outcome back as a value: it reports
the fault itself, returns failure, and the calling code acts on the value without logging it again.
That is how a setup procedure serves both a person at the keyboard and a caller with nobody there.
A procedure whose value cannot carry failure is never that: a Null or 0 that already means "none"
would turn an error into a believable answer.

A step called from a query or a control's source passes its error to Access, which shows `#Error` in
that row or control: visible, and not mistaken for a value.

**2. The kind of error: how was it raised?**

| Kind | Raised by | Recognised by |
|---|---|---|
| **Refusal** | The engine refusing a write the code made (a Data Macro's `RaiseError`, a validation rule, a required or empty field, a duplicate key, a missing or remaining related record, a value too large, a row locked by someone else), or our own code refusing, from the refusal range | `IsRefusal(Err.Number)` |
| **Expected** | One error this procedure means to treat as a normal outcome, named in its own code | that number, in that procedure only |
| **Fault** | Everything else, including a deliberate stop our own code raises from the stop range | neither of the above |

**3. The handler.** Find the row, then the column.

| Role | Fault | Refusal | Expected |
|---|---|---|---|
| **Entry point** | Reporting shape: roll back if it began a transaction, then report | Reporting shape's refusal branch: roll back, tell the person the sentence, do not log | Recovery idiom |
| **Step** | Propagating shape; no handler at all if there is nothing to clean up | Propagating shape, unchanged | Recovery idiom |
| **Answer** | Answering shape | Answering shape | — |
| **Called by a Data Macro** | Refusing shape: log quietly, return a sentence that refuses the save | Its own refusal: return the sentence | Recovery idiom |
| **The logger** | Guarded | — | — |

**Only an entry point sorts refusals from faults.** A step passes every error on as it came, so the
procedure that knows what the person was doing decides what they are told. Every shape is described
below, with its code.

## Labels

Always these exact spellings: `errHandler:` and `Cleanup:` (never `ErrorHandler`, `err_handler`).

## Two ways to report an error — pick one and use all of it

Option 1 calls the logger, `LogError`. Option 2 doesn't log at all.

**`LogError` also tells the user** — it is not the silent alternative to a message
box. It records the error *and* reports it to them, and **what they see is settled once, when the
logger is built**, rather than written into each handler: a short message with a reference number
they can quote back, the full technical detail, or — if that is deliberately chosen — nothing at
all. Deciding it in one place is why no option below puts a `MsgBox` in the handler itself.

**On a path where nobody may be interrupted, pass `booShow:=False`:** the error is recorded and
nothing is shown. Such a path is named by its template (scanning, a save a Data Macro is handling).
`LogError` returns the record's reference number, 0 when none was kept, so a quiet path can hand the
reference on in a sentence instead.

```vba
lngRef = LogError(ModuleName, ProcName, Erl, booShow:=False)
```

Option 2 has nowhere to put an error on a quiet path. A template with a quiet path either requires
option 1 or passes the error on to an entry point that may speak.

> **`LogError` is built by `templates/errors/error-logging-scaffold.md`**, along with the table it
> writes to. That template asks which of these two options you want — and where the record goes,
> what the user sees, and the rest — one question at a time. Choosing between
> them here, from this file, works just as well; the template exists so you don't have to.

### Option 1 — named constants *(preferred)*

```vba
' at the top of the module
Private Const ModuleName As String = "modInventory"

Public Sub RecountShelf(ByVal lShelfID As Long)
      Const ProcName As String = "RecountShelf"
100       On Error GoTo errHandler
          ' ... main logic ...

Cleanup:
210       Exit Sub

errHandler:
240       LogError ModuleName, ProcName, Erl
250       Resume Cleanup
260       Resume
End Sub
```

**Why this one is preferred: a template writes the names at the same moment it writes the
procedure**, so they are right by construction — there is no step at which a generator could get
them wrong. It needs nothing installed, nothing enabled, and nothing configured on the machine it
runs on, so it behaves the same way for every adopter. The constants are useful outside the handler
too: any `Debug.Print` or status message can reference them.

If someone later renames a procedure, `ProcName` is one line directly above the thing they renamed.

### Never identify the module or procedure by reading the VBA editor at run time

**Named constants (`ModuleName`/`ProcName`, above) are the only way this standard identifies where
an error happened.** A second technique exists in the wild — reading `Application.VBE.ActiveCodePane`
while the code runs, to work out which module and procedure raised the error, instead of writing
their names as constants. **It does not do what its name suggests, and generated code must never use
it, offer it, or describe it as an alternative equal to named constants.**

**Why it's wrong, not just less convenient.** `ActiveCodePane` returns the module that is *active or
was last active in the editor* — Microsoft's own description of the property, an IDE-focus property,
not a call-stack property. It reports whichever module a person last had open, not the module where
the error actually occurred. Confirmed live: an error raised in one form was logged against a
different form the developer had open moments earlier in the editor. **This failure is silent** — no
error, no blank field, nothing that looks wrong — which is worse than the technique's second problem:
it also depends on "Trust access to the VBA project object model" (Trust Center → Macro Settings),
off by default and set per machine, so code that happens to work where it was written can fail
differently on someone else's machine. Both are reasons this is prohibited; the silent-misattribution
problem is the one that matters even where Trust Center access happens to be on.

**Retrofit on encounter, without asking.** Where generated or existing code uses this pattern —
whether the AI assistant is building something new nearby, fixing an unrelated defect in the same
procedure, or simply reading through it — convert it to the named-constants form above as part of
that work, the same way a hardcoded literal name (`sFrm:="ModuleName"`) gets converted on sight. This
does not wait for the developer to ask, because the failure it produces is silent and gives them
nothing to notice and ask about.

### Option 2 — message box, no logging

```vba
errHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
```

Tells the person something went wrong and **records nothing** — once they close the box there is no
trace it happened. Reasonable for a one-off utility or a demonstration, where a log file would be
clutter. Not reasonable for anything someone else will rely on.

This handler does not report where the error happened, so its procedures carry **neither line
numbers nor `Erl`**. If a procedure
declares `ModuleName` and `ProcName` (option 1) but this handler never uses them, leave them in
and comment them: they are part of the fuller method above, kept as a forward-looking pre-set-up so
that adding a logger later does not mean editing every procedure.

### Substituting your own logger, and when it is really a handler

A routine of the host's is a **logger** only if it records the error and returns to the line after
the call on every path: no dialog, no re-raise, no value the caller acts on. A logger replaces the
call to `LogError` and nothing else; the procedure still captures `Erl`, calls it, then
`Resume Cleanup` and `Resume`. How `Erl` reaches the logger is the shop's to decide, and this file
doesn't prescribe it.

A routine that does anything else is an **error handler**, not a logger. It shows a dialog, or
re-raises the error, or returns a value, or ends the application, or has to be paired with a call on
entry or exit. The procedure's shape then belongs to the handler, and the next section applies.

### When the host already has an error handler

**To the AI generating code: this is a question for the developer, not a rule you apply.** It is one
question: accept the host's handler as it is, or use option 1 or option 2 above. Accepting means
every generated procedure follows the host's own procedure shape exactly as the host's own
procedures do: its labels, its line numbering, its call, and what follows the call. Never place the
host's call inside the shape this file gives. The two disagree about what happens after the call,
and a handler that re-raises never returns, so the lines after it never run.

1. **Read it first.** Read the handler and at least two of the host's own procedures that call it.
   Tell the developer what it does on each of these: returns or raises; records, shows a dialog, or
   both; needs a paired call or set-up; ends the application; writes to a table or file the build
   also writes.
2. **Offer accepting only when you can state every one of those.** If any is unclear, or the
   handler needs a paired call you cannot place correctly in every procedure, do not offer it. Say
   why, and offer option 1 and option 2.

   **A re-raising handler breaks any promise to return a value on failure.** Before you offer to
   accept it, list each procedure in the template that returns a sentence or a partial result on
   failure. In the design spec, show the developer one such procedure as a before-and-after, using the
   template's own case. The text below is the shape, not the words:

   > *"Posting a year of earned hours stops at month 5, because month 5 has no accrual schedule row.
   > With the template's own handler you would see: 'Posted 4 of 12 periods, then stopped at month 5:
   > no schedule row.' With your handler you see only the error and a line in your log file. Four rows
   > were posted, and nothing on screen tells you so. You would open the ledger to find out where it
   > stopped."*

   End it with one line on what to do: check the ledger after any failed posting, and read the log
   file.
3. **Ask, and name accepting as the preferred answer when it is offered.** Preferred does not
   mean decided (`templates/_template-schema.md` §10.7). Ask it even when the answer looks settled.
4. **Whichever is chosen, all five of these hold:**
   - A failure inside a transaction rolls the transaction back.
   - A caller whose next step depends on the outcome learns the call failed.
   - `Err` is read before anything clears it.
   - Generated code calls only a handler that is installed and compiles.
   - Anything that must happen on failure (a rollback, releasing a file or lock) comes before the
     handler call whenever that handler re-raises, because it never returns.
   - The roles in *Choose every handler from one table* still hold: the host's call replaces only
     the report in an entry point's fault branch. Refusals are still sorted with `IsRefusal` and
     told, steps still propagate, and code a Data Macro calls still returns a sentence.
5. **Prove it.** Force one failure inside a transaction, in a procedure called from another, and
   check the five conditions. Count the log entries and compare them with what the handler's
   behavior predicts: one per procedure the error passes through for a handler that records and
   re-raises, one in total for the standard shapes. The build record names which case applied and
   the entries seen.
6. **Whatever this section doesn't cover goes in the build record.** Do not add a rule for it.

### Which one to use, and who decides

**Option 1 is the preferred choice** when the host has no handler of its own. "Preferred" means the
one to put first when offering; it does **not** mean the one to use without asking (see
`templates/_template-schema.md` §10.7). Ask even when the answer looks settled.

One hard constraint bounds every answer: **generated code must compile on the machine it lands on,
so never emit a call to a handler or logger that isn't installed.** Where `LogError` is absent, say
so and offer the real choices: install the logger first, use option 2 now, or accept the host's
handler if it passes step 2 above. Do not quietly pick one and report it afterward; that is a
decision the developer never made.

## The logger itself is the one exception

`LogError` is what a handler calls, so it cannot call a handler of its own: an error raised inside
it would re-enter the handler that called it and loop, or stop the application outright. It is
therefore the single place in this library that is **guarded rather than handled** — `On Error
Resume Next` from top to bottom, with each write attempt reporting success or failure as a value
instead of raising.

Two rules come with that exception, and both are easy to get wrong:

- **Read `Err` before the guard, never after.** Any `On Error` statement clears the `Err` object,
  and so do `Resume`, `Exit Sub`, and `Exit Function`. A logger that sets up its guard first and
  then reads `Err.Number` records error 0 with an empty description, on every call, without ever
  failing visibly.
- **The exception does not travel.** `On Error Resume Next` stays confined to the logger and to
  `Cleanup:` blocks. Nothing else in generated code adopts it because the logger does.

## Full procedure skeleton

```vba
' at the top of the module
Private Const ModuleName As String = "modExample"

Public Sub ProcedureName(ByVal param1 As Type)
      Const ProcName As String = "ProcedureName"
100       On Error GoTo errHandler

      Dim db     As DAO.Database
      Dim strSQL As String

110       Set db = CurrentDb
120       strSQL = "SELECT ..."
130       ' ... main logic ...

Cleanup:
210       On Error Resume Next
220       Set db = Nothing
230       Exit Sub

errHandler:
240       LogError ModuleName, ProcName, Erl
250       Resume Cleanup
260       Resume
End Sub
```

**Trivial one-liners** (an event procedure that only calls an entry point, which reports for itself,
or a one-line helper that cannot realistically fail) take **no** `errHandler` and **no** line
numbers. An event procedure that calls a step is the entry point, and needs the reporting shape:

```vba
Private Sub cmdEdit_Click()
    AddEditRecord Me
End Sub
```

## Line numbering

- Add line numbers **only** where the procedure's error handler is designed to report where the
  error happened; that handler reads `Erl`. A procedure whose handler does not report the location,
  or that has no handler, carries neither.
- Increment by 10 from 100; restart at a round number for `Cleanup:` and `errHandler:`.
- **Renumber a procedure you edit, or strip its numbers entirely.** Numbers that no longer match the
  lines are worse than none — `Erl` reports a line where nothing failed.
- **Number the statements that can fail, never a declaration.** `Erl` returns the last numbered
  line that ran. A numbered `Dim` ahead of a failing unnumbered statement makes `Erl` report the
  `Dim`'s number.
- **Read `Erl` on the handler's first statement**, before any other numbered handler line runs.
  Once a numbered handler line has run, `Erl` reports that line, not the line that failed.
- **The `Cleanup:` and `errHandler:` lines may be left unnumbered.** `Erl` then reports the last
  numbered line that ran, which is the line that failed, so nothing has to be read first.
- **Line numbering is itself a house-specific choice.** This standard relies on `Erl`, which needs
  numbered lines, applied with a line-numbering tool on import. Other practices number manually, or
  reject line numbers entirely — in which case `Erl` returns 0 and the central handler simply logs
  without a line number. Templates and scaffolds **never hard-code line numbers**; they defer to
  whatever this file specifies, so a forked practice that doesn't number swaps this file and the
  scaffolds are unaffected.

## Resume chain

This is the **reporting shape** — capture `Erl`, report, `Resume Cleanup`, `Resume` — used by an
entry point. Where the entry point makes writes the engine may refuse, its handler sorts first:

```vba
errHandler:
240       lngLine = Erl                                      ' first: a numbered line below would change it
250       If booInTrans Then ws.Rollback: booInTrans = False
260       If IsRefusal(Err.Number) Then
270           MsgBox Err.Description, vbInformation          ' or wherever the template sends results
280       Else
290           LogError ModuleName, ProcName, lngLine
300       End If
310       Resume Cleanup
320       Resume
```

A refusal is told, never logged: nothing broke, the data was turned down. The sentence is the
refusal's own (a Data Macro's `RaiseError` description arrives unchanged in `Err.Description`) unless
the template words it. Where the template says results reach the person as returned text, the
sentence goes there instead of a message. `ws.Rollback` does not clear `Err` (observed 2026-11-18), so
`Err` is still the error when it is read after the rollback.

After the handler call: always `Resume Cleanup`, then `Resume`. The trailing `Resume` lets the
debugger step back to the error line during diagnosis.

## On Error Resume Next

Only inside the `Cleanup:` block, and inside the logger itself (see "The logger itself is the one
exception"). **Never** in main logic.

## Transaction guard

**The transaction and every write inside it must be on the same workspace.** A transaction is begun on a
`Workspace`; the writes are made through a `Database`. A write is inside the transaction when its
`Database` belongs to that `Workspace`. Take `db` from the workspace that began the transaction, as the
first two lines below do, and that is true by construction.

```vba
Dim ws       As DAO.Workspace
Dim db       As DAO.Database
Dim booInTrans As Boolean

Set ws = DBEngine.Workspaces(0)
Set db = ws.Databases(0)          ' same workspace as BeginTrans

ws.BeginTrans
booInTrans = True
db.Execute strSQL, dbFailOnError + dbSeeChanges
ws.CommitTrans
booInTrans = False
...
errHandler:
      If booInTrans Then ws.Rollback
      LogError ModuleName, ProcName, Erl
      Resume Cleanup
      Resume
```

**Why the source of `db` matters, and why reads inside the transaction use it:** which writes a
transaction covers, and what a domain function can see while one is open, are facts about the engine,
not house choices. They live in `templates/_materialization.md` (fact `domain-function-transaction`),
which every template that opens a transaction declares, so a shop replacing this file keeps them.
The house choice here is the rule above: take `db` from the workspace that began the transaction, and
read inside it with a recordset on that same `db`, never a domain function.

**One thing the compiler catches and one it does not.** `BeginTrans`, `CommitTrans` and `Rollback`
belong to `Workspace`; a `Database` has none of them, so calling them on the wrong object fails to
compile and you find it in seconds. Taking `db` from a different workspace than the one that began the
transaction compiles cleanly. Compile every build, and check this by reading.

**Reader: the AI assistant, and a shop replacing this file.**

## Errors from a called procedure must reach whoever depends on the outcome

**The standard handler shape — log, `Resume Cleanup`, return normally — is correct only for a
procedure that is itself the end of the line:** an event procedure with no caller relying on its
result, or a genuinely fire-and-forget call where nothing afterward is conditioned on whether it
worked. The moment a procedure is called as **one step inside something larger** — where the
caller's own next action depends on whether that step succeeded — the same pattern becomes a
defect. The callee logs the failure and returns exactly as it would on success; the caller has no
way to tell the two apart, and proceeds as if the step worked.

**The test is simple: does the caller do anything, next, that depends on this call having
succeeded?** If yes, the callee must propagate its error rather than swallow it. If nothing after
the call is conditioned on it — a status message, a log entry, cleanup that runs either way — the
standard pattern is fine exactly as written above.

**A transaction commit is the sharpest case, not a separate one.** A procedure called from inside
another procedure's transaction is a caller that depends on the outcome in the most literal sense:
if the nested call fails and the standard pattern swallows it, the procedure that began the
transaction has no way to learn that, and commits anyway. But the same hole opens any time A calls
B to do something A's own next step relies on — a validation check, a value B was supposed to
produce, a row B was supposed to insert — whether or not a transaction is involved.

**The fix is where the errHandler stops, not what it does first.** A procedure meant to be called
as a dependent step still cleans up its own local resources (close a recordset, release an object)
in its errHandler, but instead of swallowing the error, it hands it back. This is the
**propagating shape** — local cleanup only, then `Err.Raise`, no `Erl` capture and no report — the
shape the table gives a step:

```vba
errHandler:
          ' Adapted from AccessCodeLib (access-codelib.net), data/dao/DaoHandler.cls, BSD 3-Clause License.
240       lngErrNumber = Err.Number                           ' copy Err first: cleanup can clear it
250       strErrSource = Err.Source
260       strErrDescription = Err.Description
270       Set rs = Nothing                                    ' local cleanup only
280       Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
```

**Copy `Err` before the cleanup.** Cleanup can run code you cannot see: setting a class instance to
`Nothing` runs its `Class_Terminate`, and an `On Error` statement there, or in any procedure it calls,
clears `Err`. A re-raise that reads `Err` after the cleanup then raises error 0, which VBA reports as
error 5, "Invalid procedure call or argument", and the real error is lost. Observed 2026-11-18 in a
test database: a `Class_Terminate` with `On Error Resume Next`, and one calling a procedure with its
own handler, both cleared `Err`; `rs.Close` and `ws.Rollback` did not.

**Name this procedure in the source.** Each propagating procedure puts `ModuleName.ProcName->` in
front of `Err.Source`, so the error reaches the procedure that logs carrying the path it took
(`modOrder.PostOrder->modOrder.WriteLines->MyDatabase`). This tells the reader where the error began,
which `Erl` in the logging procedure cannot, since it reports only the line of the call. The logger
records the path only if it reads `Err.Source`.

The procedure whose next action depends on this one is the one that logs — and, if it began a
transaction, rolls back, per the transaction guard above (`If booInTrans Then ws.Rollback`). Logging
the same failure twice — once in the called procedure, once in the caller — is not a safeguard; it
is two records of one event, and whoever reads the log afterward cannot tell that from two separate
failures.

**A procedure with no local cleanup to do needs no `errHandler` of its own here at all** — an
unhandled error already propagates to its caller by VBA's own default behavior, so the block above
is only needed where there is something to release first.

### Recovering from one expected error and passing every other one on

Sometimes a procedure whose caller depends on its outcome must recover from one specific, expected
engine error (e.g. 3022, a duplicate key it means to treat as "already there") and let every other
error still reach the caller. The conditions above can't express this on their own: reading
`Err.Number` requires an `errHandler` to run at all, and passing an error on means `Err.Raise` from
that handler — which skips `Cleanup:` by design, so this idiom does its own local cleanup explicitly
rather than routing through it:

```vba
errHandler:
240       If Err.Number = 3022 Then
250           Set rs = Nothing
260           Resume Next
270       End If
280       lngErrNumber = Err.Number                           ' anything else: copy, clean up, propagate unlogged
290       strErrSource = Err.Source
300       strErrDescription = Err.Description
310       Set rs = Nothing
320       Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
```

Only the error number named as expected is ever recovered from inline; every other error takes the
propagate path above, unlogged, for the dependent caller to log and act on.

## Refusals: `IsRefusal`

When code writes a row and the engine turns it down, the refusal arrives as an ordinary error, by the
same route as a fault. One function tells them apart. Copy it, unchanged and `Private`, into each
module whose entry points write: the copy in this file is the only source, and a private copy cannot
collide with another module's.

```vba
' error-handling.md: the one list of refusals. Copy unchanged; change it here, not in a module.
Private Function IsRefusal(ByVal lngErr As Long) As Boolean
    Select Case lngErr
        Case 3939                     ' a Data Macro's RaiseError, whatever Number it gave
        Case 3022                     ' duplicate key
        Case 3200, 3201               ' related records remain; no related record
        Case 3314, 3315               ' required value missing; empty string not allowed
        Case 3316, 3317               ' table or field validation rule
        Case 3163                     ' value too large for the field
        Case 3188, 3218, 3260         ' row locked by someone else
        Case vbObjectError + 1000 To vbObjectError + 1999   ' refused by our own code
        Case Else: Exit Function
    End Select
    IsRefusal = True
End Function
```

A row locked by someone else is listed because the person can act on it (wait, then try again); it
is not a fault in the code. A Data Macro that fails rather than refuses raises nothing at all (see
*What never reaches a VBA handler*).

## Raising an error on purpose

Our own code raises an error for one of two reasons, and the number says which, so `IsRefusal` needs
no list per template:

- **A stop**: the code found something it cannot go on from, such as a setting with a value nobody
  defined. A fault. Number from `vbObjectError + 513` to `vbObjectError + 999`.
- **A refusal**: the code turns down what the person asked for, as a validation rule would. Number
  from `vbObjectError + 1000` to `vbObjectError + 1999`.

Declare each number once, at the top of the module, and raise it with the module and procedure as
the source and a sentence the person can act on as the description:

```vba
Private Const ErrScopeModeUnknown As Long = vbObjectError + 513   ' a stop
Private Const ErrEntryAfterInactive As Long = vbObjectError + 1000 ' a refusal
...
Err.Raise ErrScopeModeUnknown, ModuleName & "." & ProcName, _
          "AuditScopeMode is """ & AuditScopeMode & """. It has to be ""List"" or ""All""."
```

A number is unique within its module; the source says which module it came from, so two modules may
use the same number. Code that only needs to say no to its own caller returns the answer instead of
raising (a validation function returning its reason): raise only where the refusal has to cross a
procedure that would otherwise carry on.

## When a failure is the answer: the answering shape

Some questions can be answered only by trying: "can this file be opened as our back end?" A missing
file, a dead server and a file that is not a database all raise errors, and each means "no". A
procedure asking such a question sets the answer in its handler and reports nothing:

```vba
Private Function BackEndIsReachable(ByVal strPath As String) As Boolean
    Dim dbBackEnd As DAO.Database

    On Error GoTo errHandler
    If Len(Dir$(strPath)) = 0 Then GoTo Cleanup          ' tested first: no error needed
    Set dbBackEnd = DBEngine.OpenDatabase(strPath, False, True)   ' the one attempt
    BackEndIsReachable = True

Cleanup:
    On Error Resume Next
    If Not dbBackEnd Is Nothing Then dbBackEnd.Close
    Set dbBackEnd = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] answering shape
    BackEndIsReachable = False                            ' the answer, not a fault
    Resume Cleanup
    Resume
End Function
```

All three conditions hold, or the procedure is a step and propagates:

1. **Test first what can be tested without an error** (a file exists: `Dir`; a table, field or
   property exists: walk the collection) and try only what cannot. A question a test answers needs
   no handler.
2. **The procedure holds the attempt and nothing else**, so every error it can raise means "no".
3. **"No" is the safe answer for every caller.** A wrong "no" may cost a retry or a skipped step,
   never data. Where acting on a wrong "no" would delete, overwrite or lose something, or would show
   the person a false answer ("no pay rate", "not found"), the procedure is a step.

The caller decides what the person sees: it is the one that knows whether "no" means something is
wrong.

## Code a Data Macro calls: the refusing shape

A function a Data Macro calls runs inside someone's save. It never raises and never shows anything.
It returns `""` to let the save go on, or a sentence to refuse it, and the macro raises that sentence
with `RaiseError`. A fault inside it refuses the save too: it is logged quietly, and the sentence
carries the reference.

```vba
Public Function CheckEntry(ByVal lngEntryID As Long) As String
      Const ProcName As String = "CheckEntry"
      Dim lngRef As Long
100       On Error GoTo errHandler
          ' ... the check; a rule it enforces returns its own sentence:
          ' CheckEntry = "Not posted: the employee was inactive on that date."

Cleanup:
          On Error Resume Next
          ' release anything opened above
          Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] refusing shape
    lngRef = LogError(ModuleName, ProcName, Erl, booShow:=False)
    CheckEntry = "This change was not saved, because a check could not finish. " & _
                 "Quote reference " & lngRef & " if it keeps happening."
    Resume Cleanup
    Resume
End Function
```

```xml
<Action Name="SetLocalVar"><Argument Name="Name">strRefusal</Argument>
  <Argument Name="Value">CheckEntry([EntryID])</Argument></Action>
<ConditionalBlock><If><Condition>[strRefusal]&lt;&gt;""</Condition><Statements>
  <Action Name="RaiseError"><Argument Name="Number">1</Argument>
    <Argument Name="Description">=[strRefusal]</Argument></Action>
</Statements></If></ConditionalBlock>
```

**What the person sees.** On a form, Access shows the sentence in its own dialog; the record stays
on screen, unsaved, and they can try again or press Esc to undo. Through code, the writing procedure
receives error 3939 with the sentence as its description; `IsRefusal` sorts it as a refusal, so it is
told and not logged a second time. One event, one log entry, one sentence, whichever route made the
save.

**What it costs.** It fails closed. If the thing the function depends on breaks (an audit table, a
lookup), every save that runs it is refused until it is fixed. That is the point: a save the check
could not judge, or an edit the audit could not record, is never let through silently. A template
whose promise is "nothing changes unrecorded" depends on it.

Under option 2 the sentence carries `Err.Description` instead of a reference, since nothing is
recorded.

## What never reaches a VBA handler

- **A refused save made through a bound form.** Access shows its own dialog with the refusal's
  sentence, and no handler in our code runs. That is the default. A form that wants its own wording
  handles its `Error` event with `IsRefusal` and sets `Response = acDataErrContinue`.
- **A Data Macro that fails rather than refuses** (a bad expression, a missing table). It raises
  nothing; Access writes a row to `USysApplicationLog`, a table it creates on the first such failure.
  Every build that attaches a Data Macro checks, after its checks have run, that the table is absent
  or has no new rows.

## Proving the shapes

Every template's checklist carries these proofs, for the shapes its build uses, each run through a
removable seam: a hook the build adds so a check can make a step fail, and removes afterwards.

1. **A fault in a step**: make the deepest step of the longest chain fail. One log entry (one message
   under option 2), with the path in its source; a transaction rolled back, leaving no partial rows;
   the entry point reported it.
2. **A refusal**: make a write the template's tables refuse. The person is told the sentence; nothing
   is logged.
3. **An answer**, where there is one: give it the "no" case. No message, nothing logged, and the
   caller does what its "no" branch says.
4. **Code a Data Macro calls**, where there is some: make it fail. The save is refused with a
   sentence carrying the reference; one log entry.
5. **By reading**: every handler's first comment names its shape, and the shape matches the table;
   `On Error Resume Next` appears only in `Cleanup:` blocks and the logger.

## What a conforming build looks like

A build conforms to this file when every one of these is true of the code it produced. A practice
replacing this file replaces these conditions with its own. Where the developer accepted the host's
own handler, conditions 1 to 6 are replaced by the five conditions in "When the host already has an
error handler"; conditions 7 and 8 still apply.

1. Labels are spelled exactly `errHandler:` and `Cleanup:`.
2. Every procedure with an `errHandler:` block reaches it from `On Error GoTo errHandler` as its
   first executable line.
3. Every procedure's handler is the one *Choose every handler from one table* gives for its role,
   never a mix, and its first comment names the shape: **reporting** (entry point, with the refusal
   branch where it writes), **propagating** (step), **answering**, **refusing** (called by a Data
   Macro), or **guarded** (the logger). The recovery idiom may sit inside any of them, for one named
   error.
4. No procedure with an `errHandler:` block exits by a route that skips `Cleanup:`.
5. `On Error Resume Next` appears only inside a `Cleanup:` block and inside the logger.
6. A procedure whose error handler reports where the error happened carries line numbers and reads
   `Erl`. Every other procedure carries neither.
7. Every logger the generated code calls exists in the built database.
8. Where a transaction is used, the `Database` object every write inside it goes through was
   obtained from the same `Workspace` the transaction was begun on, and no domain function is read
   inside the transaction.
9. A procedure whose caller depends on its outcome propagates its errors to that caller, rather than
   logging and returning normally — only a procedure with no caller depending on the outcome, or the
   caller itself once it has the error, logs the failure (and rolls back, if it began a transaction).
10. Refusals are sorted only at entry points, with `IsRefusal` copied unchanged, and are told, never
    logged.
11. Code a Data Macro calls never raises and never shows anything; it returns `""` or a sentence.
12. Every error raised on purpose uses a declared constant from the stop range or the refusal range,
    with `ModuleName.ProcName` as its source.
13. Every quiet path the template names calls `LogError` with `booShow:=False`.
14. The checklist carries the proofs in *Proving the shapes* for the shapes the build uses.
