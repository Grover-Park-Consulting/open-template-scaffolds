---
step: 02-procedures-1
title: "The procedures, part 1 of 3"
platform_facts: [open-existing-startup, mcp-module-import, mcp-line-numbers, host-helper-reuse, sql-insert-truncation, vbe-line-continuation, recordset-append-crash]
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
Private Const ModuleName           As String = "modErrorLog"

' [SCAFFOLD] The table and file the logger writes to, in ONE place each.
Private Const ErrorLogTable       As String = "tblErrorLog"
Private Const ErrorLogFile        As String = "ErrorLog.txt"

' [SCHEMA] tblErrorLog.ErrorDescription is Text(255). A longer description is
'          truncated to fit, never allowed to fail the write (schema Business Rule 4).
Private Const ErrorDescriptionMax As Long = 255

' [SCHEMA] tblErrorLog.ErrorSource is Text(255), truncated the same way.
Private Const ErrorSourceMax      As Long = 255
```

### LogError — `Public Function` → `Long` (the logger every handler calls)

The procedure `standards/error-handling.md` names at the end of every error handler. It takes the
module, the procedure, and the line, reads the error itself, records it wherever Step 3 said, and
tells the person whatever Step 6 said, unless the caller passes `booShow:=False` (a path where nobody
may be interrupted). It returns the record's reference number, 0 when none was kept, so a quiet path
can hand the reference on in a sentence. A handler that only reports calls it as a statement, as
before.

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
Public Function LogError(ByVal strModule As String, _
                         ByVal strProcedure As String, _
                         ByVal lngLine As Long, _
                         Optional ByVal booShow As Boolean = True) As Long
    ' [STANDARDS - error-handling.md] guarded shape: the logger.
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
    If Len(strErrDescription) > ErrorDescriptionMax Then
        strErrDescription = Left$(strErrDescription, ErrorDescriptionMax)
    End If
    If Len(strErrSource) > ErrorSourceMax Then
        strErrSource = Left$(strErrSource, ErrorSourceMax)
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

    ' [STANDARDS - error-handling.md] On a quiet path (booShow:=False) the error is recorded
    '            and nothing is shown.
    If booShow Then
        ShowErrorToUser lngErrNumber, strErrDescription, strModule, strProcedure, lngLine, lngLogID, booRecorded
    End If
    LogError = lngLogID
End Function
```

### WriteErrorToTable — `Private Function` → `Boolean`

Appends one row to `tblErrorLog` and hands back the new `ErrorLogID` so Step 6 can show it as a
reference number. **Returns `False` rather than raising** when the table cannot be reached — which
is what lets the fallback option in `LogError` step down to the file.

It writes with a parameterised `INSERT`. The values never enter the SQL text, so an error description
containing an apostrophe cannot break it, and this is the one procedure in the application that must
not fail on the contents of its own input. It is not a recordset append, which has crashed Access
inside a Data Macro's save, where the refusing shape logs from.

```vba
Private Function WriteErrorToTable(ByVal lngErrNumber As Long, _
                                   ByVal strErrDescription As String, _
                                   ByVal strErrSource As String, _
                                   ByVal strModule As String, _
                                   ByVal strProcedure As String, _
                                   ByVal lngLine As Long, _
                                   ByRef lngLogID As Long) As Boolean
    ' [STANDARDS - error-handling.md] guarded shape: part of the logger. Reports success as a
    '            Boolean instead of raising.
    Dim db             As DAO.Database
    Dim qdf            As DAO.QueryDef
    Dim strUser        As String
    Dim strParameters  As String
    Dim strDescription As String
    Dim strSource      As String

    On Error Resume Next
    lngLogID = 0

    ' [SCHEMA Business Rule 5] Never empty - the column is Required, and an empty string
    '            would be refused by the very write we are trying to guarantee.
    strUser = Environ$("USERNAME")
    If Len(strUser) = 0 Then strUser = "Unknown"

    ' [STANDARDS - query-style.md] A parameterised INSERT, never a recordset append: the values
    '            never enter the SQL text, so an apostrophe cannot break the write, and a recordset
    '            append inside a Data Macro's save has crashed Access (fact recordset-append-crash),
    '            which is where the refusing shape logs from.
    ' [FACT sql-insert-truncation] An empty string passed as a parameter arrives as Null, and
    '            ErrorDescription and ErrorSource are Required. An empty one is written as the
    '            literal '' instead, and only a non-empty one is passed as a parameter.
    strParameters = "PARAMETERS [pNumber] Long, [pModule] Text(255), [pProcedure] Text(255), " & _
                    "[pLine] Long, [pOn] DateTime, [pUser] Text(255)"
    If Len(strErrDescription) > 0 Then
        strParameters = strParameters & ", [pDescription] Text(255)"
        strDescription = "[pDescription]"
    Else
        strDescription = "''"
    End If
    If Len(strErrSource) > 0 Then
        strParameters = strParameters & ", [pSource] Text(255)"
        strSource = "[pSource]"
    Else
        strSource = "''"
    End If

    Set db = CurrentDb
    Set qdf = db.CreateQueryDef(vbNullString, strParameters & "; " & _
        "INSERT INTO " & ErrorLogTable & " (ErrorNumber, ErrorDescription, ErrorSource, " & _
        "ModuleName, ProcedureName, ErrorLineNumber, ErrorOccurredOn, ErrorUser) " & _
        "VALUES ([pNumber], " & strDescription & ", " & strSource & ", [pModule], " & _
        "[pProcedure], [pLine], [pOn], [pUser])")
    ' [SCAFFOLD] The table may be missing, or the back end unreachable. Either way this
    '            is a False, not an error - the caller has somewhere else to try.
    If Err.Number <> 0 Then GoTo Cleanup

    qdf.Parameters("pNumber").Value = lngErrNumber
    qdf.Parameters("pModule").Value = strModule
    qdf.Parameters("pProcedure").Value = strProcedure
    qdf.Parameters("pLine").Value = lngLine
    qdf.Parameters("pOn").Value = Now()
    qdf.Parameters("pUser").Value = strUser
    If Len(strErrDescription) > 0 Then qdf.Parameters("pDescription").Value = strErrDescription
    If Len(strErrSource) > 0 Then qdf.Parameters("pSource").Value = strErrSource
    qdf.Execute dbFailOnError
    If Err.Number <> 0 Then GoTo Cleanup

    ' [SCAFFOLD] The row is in. Read the reference number on the same Database object - if
    '            that part fails, the record still stands and only the number is missing.
    WriteErrorToTable = True
    lngLogID = Nz(db.OpenRecordset("SELECT @@IDENTITY")(0), 0)

Cleanup:
    Set qdf = Nothing
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
    ' [STANDARDS - error-handling.md] guarded shape: part of the logger. Reports success as a
    '            Boolean instead of raising.
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
    ' [STANDARDS - error-handling.md] guarded shape: part of the logger. A bad path returns
    '            the fallback, never an error.
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
    '     only where ErrorLogTable is a LINKED table - a local one has no Connect. ---
    ' Dim sConnect As String
    ' sConnect = CurrentDb.TableDefs(ErrorLogTable).Connect
    ' strFolder = Mid$(sConnect, InStr(sConnect, "DATABASE=") + 9)
    ' strFolder = Left$(strFolder, InStrRev(strFolder, "\") - 1)

    ' [SCAFFOLD] Last resort. The temp folder always exists and is always writable, so
    '            there is no path by which this procedure returns nothing usable.
    If Len(strFolder) = 0 Then strFolder = Environ$("TEMP")
    If Len(Dir$(strFolder, vbDirectory)) = 0 Then strFolder = Environ$("TEMP")

    If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    GetErrorLogFilePath = strFolder & ErrorLogFile
End Function
```
