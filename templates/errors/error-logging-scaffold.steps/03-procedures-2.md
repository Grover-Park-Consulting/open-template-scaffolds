---
step: 03-procedures-2
title: "The procedures, part 2 of 3 (from ShowErrorToUser)"
platform_facts: [mcp-module-import, mcp-line-numbers, host-helper-reuse, dao-table-build.rules, dao-table-build.sinks, dao-table-build.index-order, vbe-line-continuation, sql-server-ddl]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

### ShowErrorToUser — `Private Sub`

Step 6's answer, in one place. It is handed everything it might need to show and uses only what the
chosen wording calls for.

**It never claims a record was kept when the write failed.** `booRecorded` is what the write actually
returned, and the wording follows it.

```vba
Private Sub ShowErrorToUser(ByVal lngErrNumber As Long, _
                            ByVal strErrDescription As String, _
                            ByVal strModule As String, _
                            ByVal strProcedure As String, _
                            ByVal lngLine As Long, _
                            ByVal lngLogID As Long, _
                            ByVal booRecorded As Boolean)
    ' [SCAFFOLD] Guarded: a failure to show the message must not replace the failure we
    '            were called about.
    Dim strMsg As String

    On Error Resume Next

    ' [WIZARD Step 6 - what the person at the keyboard sees] >>> keep ONE block <<<

    ' --- "A short message, with a reference number when there is one" (preferred) ---
    If booRecorded And lngLogID > 0 Then
        strMsg = "Something went wrong. The problem has been recorded." & vbCrLf & vbCrLf & _
               "Reference number: " & lngLogID
    ElseIf booRecorded Then
        ' [SCAFFOLD] Recorded to the file, so there is no number to quote. Say the true
        '            thing rather than invent one.
        strMsg = "Something went wrong. The problem has been recorded."
    Else
        strMsg = "Something went wrong, and it could not be recorded. " & _
               "Please make a note of what you were doing."
    End If

    ' --- "A short message that something went wrong" ---
    ' If booRecorded Then
    '     strMsg = "Something went wrong. The problem has been recorded."
    ' Else
    '     strMsg = "Something went wrong, and it could not be recorded. " & _
    '            "Please make a note of what you were doing."
    ' End If

    ' --- "The full technical detail" ---
    ' strMsg = "Error " & lngErrNumber & ": " & strErrDescription & vbCrLf & _
    '        "In: " & strModule & "." & strProcedure & vbCrLf & _
    '        "Line: " & lngLine
    ' If Not booRecorded Then
    '     strMsg = strMsg & vbCrLf & vbCrLf & _
    '            "This could not be recorded. Please make a note of what you were doing."
    ' End If

    ' --- "Nothing" ---
    ' Exit Sub

    MsgBox strMsg, vbExclamation, "Unexpected problem"
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

**A `tblErrorLog` built from schema version 0.3.0 or earlier has no `ErrorSource` column.** Because
an existing table is left alone, add the column by hand first (Text(255), Required,
`AllowZeroLength = True`). Without it, every write to the table also lands in the text file, because
the missing column leaves an error that marks the write as failed.

```vba
Public Function CreateErrorLogTable(Optional booSilent As Boolean = False) As String
    ' [SCAFFOLD] Creates tblErrorLog (schema template entity). Idempotent: an existing
    '            table is reported and skipped, so this is safe to re-run.
    '            booSilent:=True suppresses the message box and returns the same text.
    Dim db      As DAO.Database
    Dim tdf     As DAO.TableDef
    Dim fld     As DAO.Field
    Dim idx     As DAO.Index
    Dim strReport As String
    Dim booFailed As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb

    strReport = ERROR_LOG_TABLE & " created."

    ' A For Each that runs to the end leaves tdf Nothing: a missing table is an answer,
    ' not an error, so nothing here needs trapping.
    Set tdf = Nothing
    For Each tdf In db.TableDefs
        If tdf.Name = ERROR_LOG_TABLE Then Exit For
    Next tdf
    If Not tdf Is Nothing Then
        strReport = ERROR_LOG_TABLE & " already exists and was left exactly as it is."
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

    Set fld = tdf.CreateField("ErrorSource", dbText, ERROR_SOURCE_MAX)
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
    SetFieldDescription db, "ErrorSource", _
        "Err.Source: the path through propagating procedures, or the project. Truncated to fit."
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
    CreateErrorLogTable = strReport
    ' [SCAFFOLD] One message, whatever happened - the success text or the error text,
    '            never both. Building the report first and showing it here prevents that.
    If Not booSilent Then MsgBox strReport, IIf(booFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] This procedure runs at set-up time, before the
    '            logger it is building exists, so it reports rather than calls LogError.
    booFailed = True
    strReport = "ERROR creating " & ERROR_LOG_TABLE & ": " & Err.Number & " - " & Err.Description
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
                                ByVal strField As String, _
                                ByVal strDescription As String)
    ' [SCAFFOLD + _materialization.md rule 2] Called only after db.TableDefs.Append.
    '            A Description set during field-build raises runtime error 3219.
    Dim fld As DAO.Field
    Dim prp As DAO.Property

    ' [SCAFFOLD] Appending a property that already exists raises 3367, so look for it first
    '            and set it where it is. Any other error is real and reaches the caller's
    '            handler: a field that cannot be described is a build that went wrong.
    Set fld = db.TableDefs(ERROR_LOG_TABLE).Fields(strField)
    For Each prp In fld.Properties
        If prp.Name = "Description" Then Exit For
    Next prp
    If prp Is Nothing Then
        fld.Properties.Append fld.CreateProperty("Description", dbText, strDescription)
    Else
        prp.Value = strDescription
    End If

    Set prp = Nothing
    Set fld = Nothing
End Sub
```
