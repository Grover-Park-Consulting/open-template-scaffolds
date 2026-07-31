# Error Handling — OTS Default Standards Layer

> **OTS default; fork-and-replace.** The pattern for any VBA generated alongside a schema or
> scaffold (templates that declare `standards_layer: [error-handling]`). The *structure* is the
> standard; the logger it calls is replaceable — a forked practice substitutes its own.

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

## Three ways to report an error — pick one and use all of it

Options 1 and 2 both call the same logger, `LogError`, and differ only in how the handler learns
which module and procedure it sits in. Option 3 doesn't log at all.

> **`LogError` itself ships in a following change.** Until it lands, generated code uses option 3,
> which needs nothing installed. The call shape shown below is settled and won't change.

### Option 1 — named constants *(default)*

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

**Why this is the default: a template writes the names at the same moment it writes the
procedure**, so they are right by construction — there is no step at which a generator could get
them wrong. It needs nothing installed, nothing enabled, and nothing configured on the machine it
runs on, so it behaves the same way for every adopter. The constants are useful outside the handler
too: any `Debug.Print` or status message can reference them.

If someone later renames a procedure, `PROC_NAME` is one line directly above the thing they renamed.

### Option 2 — VBE reflection

```vba
errHandler:
240       LogError Application.VBE.ActiveCodePane.CodeModule, _
              Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), Erl
250       Resume Cleanup
260       Resume
```

Rather than storing the names, this asks the Visual Basic Editor which module and procedure the
failing line belongs to, so there is nothing to keep current.

**This option depends on a development environment that supports it.** Whether yours does is worth
establishing for yourself before you adopt it, and worth reading up on beyond what's here — this
library offers no guarantees on that point. If you'd rather not take that on, use option 1. Nothing
else about your code changes: both options call the same logger, and only the call site differs.

### Option 3 — message box, no logging

```vba
errHandler:
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
```

Tells the person something went wrong and **records nothing** — once they close the box there is no
trace it happened. Reasonable for a one-off utility or a demonstration, where a log file would be
clutter. Not reasonable for anything someone else will rely on.

### Substituting your own logger

If your practice already has a central error logger, swap it into the call site of whichever option
you picked and keep the rest. What this file fixes is the surrounding structure: capture `Erl`,
report, `Resume Cleanup`, `Resume`.

**To the AI generating code:** emit option 1 once `LogError` is available, and option 3 until then.
Generated code must compile on the adopter's machine as-is — never emit a call to a logger that
isn't in this library and that the developer hasn't supplied.

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
- **Do not manually renumber** after edits — MZ-Tools normalizes on import.
- **Line numbering is itself a house-specific choice.** This default relies on `Erl`, which needs
  numbered lines, applied with a tool such as MZ-Tools on import. Other practices number manually, or
  reject line numbers entirely — in which case `Erl` returns 0 and the central handler simply logs
  without a line number. Templates and scaffolds **never hard-code line numbers**; they defer to
  whatever this file specifies, so a forked practice that doesn't number swaps this file and the
  scaffolds are unaffected.

## Resume chain

After the handler call: always `Resume Cleanup`, then `Resume`. The trailing `Resume` lets the
debugger step back to the error line during diagnosis.

## On Error Resume Next

Only inside the `Cleanup:` block. **Never** in main logic.

## Transaction guard

```vba
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
