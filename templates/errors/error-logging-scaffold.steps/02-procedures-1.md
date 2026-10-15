---
step: 02-procedures-1
title: "The procedures, part 1 of 3"
platform_facts: [open-existing-startup, mcp-module-import, mcp-line-numbers, host-helper-reuse, sql-insert-truncation, dao-table-build, vbe-line-continuation, sql-server-ddl]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

Open the database the developer supplied, as the facts below say. Change nothing before the developer has
given permission.

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

    ' --- "A table in this database" ---
    ' bRecorded = WriteErrorToTable(lErrNumber, sErrDescription, sModule, sProcedure, _
    '                               lLine, lLogID)

    ' --- "A text file" ---
    ' bRecorded = WriteErrorToFile(lErrNumber, sErrDescription, sModule, sProcedure, lLine)

    ' --- "The table, falling back to the text file" (preferred) ---
    bRecorded = WriteErrorToTable(lErrNumber, sErrDescription, sModule, sProcedure, _
                                  lLine, lLogID)
    If Not bRecorded Then
        bRecorded = WriteErrorToFile(lErrNumber, sErrDescription, sModule, sProcedure, lLine)
    End If

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
