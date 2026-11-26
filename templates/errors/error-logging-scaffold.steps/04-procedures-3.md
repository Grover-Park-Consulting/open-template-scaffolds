---
step: 04-procedures-3
title: "The procedures, part 3 of 3 (from LinkErrorLogTable)"
platform_facts: [mcp-module-import, mcp-line-numbers, host-helper-reuse, dao-table-build.rules, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

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
Public Function LinkErrorLogTable(Optional ByVal strBackEndPath As String = "", _
                                  Optional booSilent As Boolean = False) As String
    ' [SCAFFOLD] Links tblErrorLog from the back end into this front end. Needed only
    '            where Step 4 put the log table in the back end. Idempotent: a table
    '            that is already here is reported and left alone, so this is safe to
    '            re-run. booSilent:=True suppresses the message box, returning the text.
    Dim db       As DAO.Database
    Dim tdf      As DAO.TableDef
    Dim tdfOther As DAO.TableDef
    Dim strReport  As String
    Dim booFailed  As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [SCAFFOLD] Already here? Say which kind. A local table means Step 4 chose "each
    '            front end" and there is nothing to link; a link says where it points,
    '            so one aimed at the wrong back end is visible rather than assumed.
    ' A For Each that runs to the end leaves tdf Nothing: a missing table is an answer,
    ' not an error, so nothing here needs trapping.
    Set tdf = Nothing
    For Each tdf In db.TableDefs
        If tdf.Name = ErrorLogTable Then Exit For
    Next tdf
    If Not tdf Is Nothing Then
        If Len(tdf.Connect) = 0 Then
            strReport = ErrorLogTable & " is a local table in this file and was left alone."
        ElseIf Left$(tdf.Connect, 10) = ";DATABASE=" Then
            strReport = ErrorLogTable & " is already linked to " & _
                      Mid$(tdf.Connect, 11) & " and was left alone."
        Else
            strReport = ErrorLogTable & " is already linked (" & tdf.Connect & _
                      ") and was left alone."
        End If
        GoTo Cleanup
    End If

    ' [SCAFFOLD] No path given? Read it off a link this front end already has - the same
    '            ";DATABASE=" rule as CurrentBackEndPath in app-startup-scaffold.md.
    '            A front end with no links yet has nothing to read, and says so below.
    If Len(strBackEndPath) = 0 Then
        For Each tdfOther In db.TableDefs
            If Left$(tdfOther.Connect, 10) = ";DATABASE=" Then
                strBackEndPath = Mid$(tdfOther.Connect, 11)
                Exit For
            End If
        Next tdfOther
    End If

    If Len(strBackEndPath) = 0 Then
        booFailed = True
        strReport = "No back end could be found from this file's existing links. " & _
                  "Run this again with the back end's full path, for example: " & _
                  "LinkErrorLogTable(""C:\Data\App_BE.accdb"")"
        GoTo Cleanup
    End If

    Set tdf = db.CreateTableDef(ErrorLogTable)
    tdf.Connect = ";DATABASE=" & strBackEndPath
    tdf.SourceTableName = ErrorLogTable
    db.TableDefs.Append tdf
    db.TableDefs.Refresh

    ' [SCAFFOLD] Prove the link reads before calling it done. Reading a field count
    '            forces the link to resolve; a back end that has moved since fails here,
    '            with its own message, rather than at the first error nobody records.
    If db.TableDefs(ErrorLogTable).Fields.Count = 0 Then
        booFailed = True
        strReport = ErrorLogTable & " was linked to " & strBackEndPath & _
                  ", but the link does not read. Check the back end."
        GoTo Cleanup
    End If

    strReport = ErrorLogTable & " linked to " & strBackEndPath

Cleanup:
    Set tdfOther = Nothing
    Set tdf = Nothing
    Set db = Nothing
    LinkErrorLogTable = strReport
    ' [SCAFFOLD] One message, whatever happened - the success text or the error text,
    '            never both. Building the report first and showing it here prevents that.
    If Not booSilent Then MsgBox strReport, IIf(booFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] reporting shape (entry point, returns its report). It
    '            runs at set-up time, before the link the logger writes through exists, so it
    '            reports in its report rather than calling LogError.
    booFailed = True
    strReport = "ERROR linking " & ErrorLogTable & ": " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```
