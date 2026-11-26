---
step: 02-procedures-1
title: "The procedures, part 1 of 3"
platform_facts: [open-existing-startup, mcp-module-import, mcp-line-numbers, host-helper-reuse, sql-insert-truncation, vbe-line-continuation]
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

' [SCHEMA] tblErrorLog.ErrorSource is Text(255), truncated the same way.
Private Const ERROR_SOURCE_MAX      As Long = 255
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

**The order of the first three statements is load-bearing.** `Err` is cleared by any `On Error`
statement — including the guard on the fourth line. Read the error first, guard second. Swapping
them logs error 0 with an empty description, on every call, silently.

```vba
Public Sub LogError(ByVal strModule As String, _
                    ByVal strProcedure As String, _
                    ByVal lngLine As Long)
    ' [SCAFFOLD] Called from an errHandler block that is already dealing with a failure.
    '            Nothing in here may raise an error of its own.
    Dim lngErrNumber      As Long
    Dim strErrDescription As String
    Dim strErrSource      As String
    Dim lngLogID          As Long
    Dim booRecorded       As Boolean

    ' [SCAFFOLD] Read Err FIRST. The On Error statement below clears it, and so would
    '            any other On Error statement, Resume, Exit Sub, or Exit Function. These
    '            three lines must stay above the guard.
    lngErrNumber = Err.Number
    strErrDescription = Err.Description
    ' [STANDARDS - error-handling.md] For an error that came through propagating
    '            procedures, the path it took. Adapted from AccessCodeLib
    '            (access-codelib.net), base/modErrorHandler.bas, BSD 3-Clause License.
    strErrSource = Err.Source

    ' [STANDARDS - error-handling.md] The documented exception: the logger is guarded
    '            rather than handled, because a handler here would call itself.
    On Error Resume Next

    ' [SCHEMA Business Rule 4] Truncate to fit rather than fail on length.
    If Len(strErrDescription) > ERROR_DESCRIPTION_MAX Then
        strErrDescription = Left$(strErrDescription, ERROR_DESCRIPTION_MAX)
    End If
    If Len(strErrSource) > ERROR_SOURCE_MAX Then
        strErrSource = Left$(strErrSource, ERROR_SOURCE_MAX)
    End If

    ' [WIZARD Step 3 - where errors are recorded] >>> keep ONE of the three blocks <<<

    ' --- "A table in this database" ---
    ' booRecorded = WriteErrorToTable(lngErrNumber, strErrDescription, strErrSource, strModule, _
    '                               strProcedure, lngLine, lngLogID)

    ' --- "A text file" ---
    ' booRecorded = WriteErrorToFile(lngErrNumber, strErrDescription, strErrSource, strModule, _
    '                              strProcedure, lngLine)

    ' --- "The table, falling back to the text file" (preferred) ---
    booRecorded = WriteErrorToTable(lngErrNumber, strErrDescription, strErrSource, strModule, _
                                  strProcedure, lngLine, lngLogID)
    If Not booRecorded Then
        booRecorded = WriteErrorToFile(lngErrNumber, strErrDescription, strErrSource, strModule, _
                                     strProcedure, lngLine)
    End If

    ShowErrorToUser lngErrNumber, strErrDescription, strModule, strProcedure, lngLine, lngLogID, booRecorded
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
Private Function WriteErrorToTable(ByVal lngErrNumber As Long, _
                                   ByVal strErrDescription As String, _
                                   ByVal strErrSource As String, _
                                   ByVal strModule As String, _
                                   ByVal strProcedure As String, _
                                   ByVal lngLine As Long, _
                                   ByRef lngLogID As Long) As Boolean
    ' [SCAFFOLD] Guarded throughout; reports success as a Boolean instead of raising.
    Dim db    As DAO.Database
    Dim rst    As DAO.Recordset
    Dim strUser As String

    On Error Resume Next
    lngLogID = 0

    ' [SCHEMA Business Rule 5] Never empty - the column is Required, and an empty string
    '            would be refused by the very write we are trying to guarantee.
    strUser = Environ$("USERNAME")
    If Len(strUser) = 0 Then strUser = "Unknown"

    Set db = CurrentDb
    Set rst = db.OpenRecordset(ERROR_LOG_TABLE, dbOpenDynaset, dbAppendOnly)

    ' [SCAFFOLD] The table may be missing, or the back end unreachable. Either way this
    '            is a False, not an error - the caller has somewhere else to try.
    If Err.Number <> 0 Then GoTo Cleanup
    If rst Is Nothing Then GoTo Cleanup
    Err.Clear

    ' [STANDARDS - query-style.md] A recordset append, not a concatenated INSERT: an
    '            apostrophe in Err.Description must not be able to break the write.
    rst.AddNew
    rst!ErrorNumber = lngErrNumber
    rst!ErrorDescription = strErrDescription
    rst!ErrorSource = strErrSource
    rst!ModuleName = strModule
    rst!ProcedureName = strProcedure
    rst!ErrorLineNumber = lngLine
    rst!ErrorOccurredOn = Now()
    rst!ErrorUser = strUser
    rst.Update

    If Err.Number = 0 Then
        ' [SCAFFOLD] The row is in. Read the reference number afterwards - if that part
        '            fails, the record still stands and only the number is missing.
        WriteErrorToTable = True
        rst.Bookmark = rst.LastModified
        lngLogID = Nz(rst!ErrorLogID, 0)
    End If

Cleanup:
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Set db = Nothing
End Function
```

### WriteErrorToFile — `Private Function` → `Boolean`

Appends one tab-separated line to the log file. Used on its own where Step 3 chose the file, and as
the fallback where Step 3 chose both. **The timestamp at the start of the line is the reference**
in this case — there is no `ErrorLogID` to quote.

```vba
Private Function WriteErrorToFile(ByVal lngErrNumber As Long, _
                                  ByVal strErrDescription As String, _
                                  ByVal strErrSource As String, _
                                  ByVal strModule As String, _
                                  ByVal strProcedure As String, _
                                  ByVal lngLine As Long) As Boolean
    ' [SCAFFOLD] Guarded throughout; reports success as a Boolean instead of raising.
    Dim intFile    As Integer
    Dim strPath    As String
    Dim strUser    As String
    Dim strEntry   As String
    Dim booWritten As Boolean

    On Error Resume Next

    strPath = GetErrorLogFilePath()
    If Len(strPath) = 0 Then Exit Function

    strUser = Environ$("USERNAME")
    If Len(strUser) = 0 Then strUser = "Unknown"

    ' [SCAFFOLD] Tab-separated so the file opens straight into a spreadsheet if anyone
    '            needs to sort it. The leading timestamp is what a support call quotes.
    strEntry = Format$(Now(), "yyyy-mm-dd hh:nn:ss") & vbTab & _
             strUser & vbTab & _
             strModule & "." & strProcedure & vbTab & _
             "line " & lngLine & vbTab & _
             lngErrNumber & vbTab & _
             strErrDescription & vbTab & _
             strErrSource

    Err.Clear
    intFile = FreeFile
    Open strPath For Append As #intFile
    If Err.Number = 0 Then
        Print #intFile, strEntry
        booWritten = (Err.Number = 0)
        Close #intFile
    End If

    WriteErrorToFile = booWritten
End Function
```

### GetErrorLogFilePath — `Private Function` → `String`

Works out the full path of the log file from Step 4's answer, and falls back to the temporary
folder if that location is not there — so a misconfigured path costs you the folder you wanted, not
the record.

```vba
Private Function GetErrorLogFilePath() As String
    ' [SCAFFOLD] Guarded: a bad path returns the fallback, never an error.
    Dim strFolder As String

    On Error Resume Next

    ' [WIZARD Step 4 - where the log file goes] >>> keep ONE of the three blocks <<<

    ' --- "The front end's folder" (preferred; standards/startup-conventions.md section 4
    '     classes a local log as a per-user working folder, created per front end) ---
    strFolder = CurrentProject.Path

    ' --- "Your Documents folder" ---
    ' [SCAFFOLD] A Documents folder redirected into OneDrive is not on this path; the
    '            fallback below catches that and the log lands in the temp folder.
    ' strFolder = Environ$("USERPROFILE") & "\Documents"

    ' --- "The back end's folder", read off the link the front end already has. Works
    '     only where ERROR_LOG_TABLE is a LINKED table - a local one has no Connect. ---
    ' Dim sConnect As String
    ' sConnect = CurrentDb.TableDefs(ERROR_LOG_TABLE).Connect
    ' strFolder = Mid$(sConnect, InStr(sConnect, "DATABASE=") + 9)
    ' strFolder = Left$(strFolder, InStrRev(strFolder, "\") - 1)

    ' [SCAFFOLD] Last resort. The temp folder always exists and is always writable, so
    '            there is no path by which this procedure returns nothing usable.
    If Len(strFolder) = 0 Then strFolder = Environ$("TEMP")
    If Len(Dir$(strFolder, vbDirectory)) = 0 Then strFolder = Environ$("TEMP")

    If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    GetErrorLogFilePath = strFolder & ERROR_LOG_FILE
End Function
```
