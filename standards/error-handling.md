# Error Handling — GPC Default Standards Layer

> **GPC default; fork-and-replace.** The house pattern for any VBA generated alongside a schema or
> scaffold (templates that declare `standards_layer: [error-handling]`). The *structure* is the
> standard; the specific central-handler function is house-specific — a forked practice substitutes
> its own central logger.

## Labels

Always these exact spellings: `errHandler:` and `Cleanup:` (never `ErrorHandler`, `err_handler`).

## The errHandler pattern — VBE reflection, never hardcode names

```vba
errHandler:
240       Call codearchive.GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
250       Resume Cleanup
260       Resume
```

The VBE-reflection form derives the module and procedure names at runtime, so they need not be
hardcoded and never go stale on a rename or copy. The hardcoded form (`sFrm:="ModuleName"`,
`sCtl:="ProcedureName"`) is **deprecated** — convert on sight.

**House-specific note:** `GlblErrMsg` is GPC's central handler, living in the CodeArchive library
(call it qualified — `codearchive.GlblErrMsg(...)` — except inside CodeArchive itself, where the
qualifier is omitted). A forked practice replaces this call with its own central error logger; the
surrounding structure (capture `Erl`, reflect the names via the VBE, log, `Resume Cleanup`, `Resume`)
is what the standard fixes.

**If you are not GPC — and almost every reader of this file is not — you do not have
`GlblErrMsg`, and it is not available anywhere.** It lives in GPC's private CodeArchive library
and appears here only as a real example of a central logger. **Do not search for it, and do not
ask your AI assistant to find or import it.** Use one of these instead:

- **Your shop's own central error logger**, if you have one — swap it into the call site.
- **The dependency-free default** — a plain message box, needs nothing installed:

```vba
errHandler:
240       MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
250       Resume Cleanup
260       Resume
```

**To the AI generating code:** unless this practice has supplied its own central logger, emit
the dependency-free block above. Never emit `codearchive.GlblErrMsg` for a non-GPC adopter, and
never search for, offer to obtain, or ask about CodeArchive. Generated code must compile on the
adopter's machine as-is.

## Full procedure skeleton

```vba
Public Sub ProcedureName(ByVal param1 As Type)
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
240       Call codearchive.GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
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
- **Line numbering is itself a house-specific choice.** The GPC pattern relies on `Erl`, which needs
  numbered lines, and GPC applies them with MZ-Tools on import. Other practices number manually, or
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
      Call codearchive.GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
      Resume Cleanup
      Resume
```
