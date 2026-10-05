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
falling off the end — into a handler built on the Resume chain. "Continue past this one item" is a
different idiom with its own rules (see *On Error Resume Next* below), and it belongs in the code
that raised the error, never in the handler.

The failure this prevents is a quiet one: a handler that looks like it follows the standard while
behaving in a way the standard never describes, so nobody reading it thinks to check.

## Labels

Always these exact spellings: `errHandler:` and `Cleanup:` (never `ErrorHandler`, `err_handler`).

## Two ways to report an error — pick one and use all of it

Option 1 calls the logger, `LogError`. Option 2 doesn't log at all.

**`LogError` also tells the user** — it is not the silent alternative to a message
box. It records the error *and* reports it to them, and **what they see is settled once, when the
logger is built**, rather than written into each handler: a short message with a reference number
they can quote back, the full technical detail, or — if that is deliberately chosen — nothing at
all. Deciding it in one place is why no option below puts a `MsgBox` in the handler itself.

> **`LogError` is built by `templates/errors/error-logging-scaffold.md`**, along with the table it
> writes to. That template asks which of these two options you want — and where the record goes,
> what the user sees, and the rest — one question at a time. Choosing between
> them here, from this file, works just as well; the template exists so you don't have to.

### Option 1 — named constants *(preferred)*

```vba
' at the top of the module
Private Const MODULE_NAME As String = "modInventory"

Public Sub RecountShelf(ByVal lShelfID As Long)
      Const PROC_NAME As String = "RecountShelf"
100       On Error GoTo errHandler
          ' ... main logic ...

Cleanup:
210       Exit Sub

errHandler:
240       LogError MODULE_NAME, PROC_NAME, Erl
250       Resume Cleanup
260       Resume
End Sub
```

**Why this one is preferred: a template writes the names at the same moment it writes the
procedure**, so they are right by construction — there is no step at which a generator could get
them wrong. It needs nothing installed, nothing enabled, and nothing configured on the machine it
runs on, so it behaves the same way for every adopter. The constants are useful outside the handler
too: any `Debug.Print` or status message can reference them.

If someone later renames a procedure, `PROC_NAME` is one line directly above the thing they renamed.

### Never identify the module or procedure by reading the VBA editor at run time

**Named constants (`MODULE_NAME`/`PROC_NAME`, above) are the only way this standard identifies where
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
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
```

Tells the person something went wrong and **records nothing** — once they close the box there is no
trace it happened. Reasonable for a one-off utility or a demonstration, where a log file would be
clutter. Not reasonable for anything someone else will rely on.

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
3. **Ask, and name accepting as the preferred answer when it is offered.** Preferred does not
   mean decided (`templates/_template-schema.md` §10.7). Ask it even when the answer looks settled.
4. **Whichever is chosen, all five of these hold:**
   - A failure inside a transaction rolls the transaction back.
   - A caller whose next step depends on the outcome learns the call failed.
   - `Err` is read before anything clears it.
   - Generated code calls only a handler that is installed and compiles.
   - Anything that must happen on failure (a rollback, releasing a file or lock) comes before the
     handler call whenever that handler re-raises, because it never returns.
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
Private Const MODULE_NAME As String = "modExample"

Public Sub ProcedureName(ByVal param1 As Type)
      Const PROC_NAME As String = "ProcedureName"
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
240       LogError MODULE_NAME, PROC_NAME, Erl
250       Resume Cleanup
260       Resume
End Sub
```

**Trivial delegates** (a one-liner that just calls another procedure) take **no** `errHandler` and
**no** line numbers:

```vba
Private Sub cmdEdit_Click()
    AddEditRecord Me
End Sub
```

## Line numbering

- Add line numbers **only** in procedures that have an `errHandler:` block (so `Erl` returns a
  useful value). Procedures with no `errHandler` have **no** line numbers.
- Increment by 10 from 100; restart at a round number for `Cleanup:` and `errHandler:`.
- **Renumber a procedure you edit, or strip its numbers entirely.** Numbers that no longer match the
  lines are worse than none — `Erl` reports a line where nothing failed.
- **Number the statements that can fail, never a declaration.** `Erl` returns the last numbered
  line that ran. A numbered `Dim` ahead of a failing unnumbered statement makes `Erl` report the
  `Dim`'s number.
- **Read `Erl` on the handler's first statement**, before any other numbered handler line runs.
  Once a numbered handler line has run, `Erl` reports that line, not the line that failed.
- **Line numbering is itself a house-specific choice.** This standard relies on `Erl`, which needs
  numbered lines, applied with a line-numbering tool on import. Other practices number manually, or
  reject line numbers entirely — in which case `Erl` returns 0 and the central handler simply logs
  without a line number. Templates and scaffolds **never hard-code line numbers**; they defer to
  whatever this file specifies, so a forked practice that doesn't number swaps this file and the
  scaffolds are unaffected.

## Resume chain

This is the **reporting shape** — capture `Erl`, report, `Resume Cleanup`, `Resume` — used by a
procedure that is itself the end of the line (see "What a conforming build looks like" for the
other named shape, the propagating one).

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
Dim bInTrans As Boolean

Set ws = DBEngine.Workspaces(0)
Set db = ws.Databases(0)          ' same workspace as BeginTrans

ws.BeginTrans
bInTrans = True
db.Execute strSQL, dbFailOnError + dbSeeChanges
ws.CommitTrans
bInTrans = False
...
errHandler:
      If bInTrans Then ws.Rollback
      LogError MODULE_NAME, PROC_NAME, Erl
      Resume Cleanup
      Resume
```

**Why the source of `db` still matters.** What a transaction covers follows the workspace, not the
`Database` object. Observed, three repeats, SQL and recordset writes: with the transaction on the
default workspace (`DBEngine.Workspaces(0)`), a write through `CurrentDb` and a write through
`ws.Databases(0)` were both undone by `Rollback`, because `CurrentDb` belongs to the default
workspace. A transaction begun on a workspace made with `DBEngine.CreateWorkspace` did not cover
`CurrentDb` writes, and a write through a created workspace was not undone by a rollback on the default
one. In those cases the writes commit whatever happens next, `ws.Rollback` rolls back an empty
transaction, and nothing is raised or logged. The code looks correct, compiles, and runs, and the
damage is a wrong number rather than an error.

**Reads inside the transaction: domain functions do not see it.** `DLookup`, `DSum` and `DCount` run on
a connection of their own and never see uncommitted work. A recordset opened on `db` does see it.
Inside a transaction, read with a recordset on the same `db`. This includes reads made by a procedure
you call from inside the transaction, since that procedure may use a domain function of its own.
Otherwise the value that comes back is the last committed one and the code proceeds on it.

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
second of the two shapes named in "What a conforming build looks like":

```vba
errHandler:
240       Set rs = Nothing                                    ' local cleanup only
250       Err.Raise Err.Number, Err.Source, Err.Description   ' propagate - do not log, do not swallow
```

The procedure whose next action depends on this one is the one that logs — and, if it began a
transaction, rolls back, per the transaction guard above (`If bInTrans Then ws.Rollback`). Logging
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
280       Set rs = Nothing
290       Err.Raise Err.Number, Err.Source, Err.Description   ' anything else: propagate, unlogged
```

Only the error number named as expected is ever recovered from inline; every other error takes the
propagate path above, unlogged, for the dependent caller to log and act on.

## What a conforming build looks like

A build conforms to this file when every one of these is true of the code it produced. A practice
replacing this file replaces these conditions with its own. Where the developer accepted the host's
own handler, conditions 1 to 6 are replaced by the five conditions in "When the host already has an
error handler"; conditions 7 and 8 still apply.

1. Labels are spelled exactly `errHandler:` and `Cleanup:`.
2. Every procedure with an `errHandler:` block reaches it from `On Error GoTo errHandler` as its
   first executable line.
3. Every `errHandler:` block follows one of two named shapes, never a mix of the two: the
   **reporting shape** — capture `Erl`, report, `Resume Cleanup`, then `Resume` — for a procedure
   that is itself the end of the line; or the **propagating shape** — local cleanup only, then
   `Err.Raise Err.Number, Err.Source, Err.Description`, no `Erl` capture and no report — for a
   procedure whose caller depends on the outcome (see "Errors from a called procedure must reach
   whoever depends on the outcome").
4. No procedure with an `errHandler:` block exits by a route that skips `Cleanup:`.
5. `On Error Resume Next` appears only inside a `Cleanup:` block and inside the logger.
6. Every procedure with an `errHandler:` block is line-numbered; every procedure without one is not.
7. Every logger the generated code calls exists in the built database.
8. Where a transaction is used, the `Database` object every write inside it goes through was
   obtained from the same `Workspace` the transaction was begun on, and no domain function is read
   inside the transaction.
9. A procedure whose caller depends on its outcome propagates its errors to that caller, rather than
   logging and returning normally — only a procedure with no caller depending on the outcome, or the
   caller itself once it has the error, logs the failure (and rolls back, if it began a transaction).
