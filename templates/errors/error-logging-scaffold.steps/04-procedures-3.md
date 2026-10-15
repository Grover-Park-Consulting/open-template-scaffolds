---
step: 04-procedures-3
title: "The procedures, part 3 of 3 (from LinkErrorLogTable)"
platform_facts: [mcp-module-import, mcp-line-numbers, host-helper-reuse, dao-table-build, vbe-line-continuation]
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
