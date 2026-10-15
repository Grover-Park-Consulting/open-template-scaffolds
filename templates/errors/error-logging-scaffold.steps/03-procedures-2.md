---
step: 03-procedures-2
title: "The procedures, part 2 of 3 (from ShowErrorToUser)"
platform_facts: [mcp-module-import, mcp-line-numbers, host-helper-reuse, dao-table-build, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

### ShowErrorToUser — `Private Sub`

Step 6's answer, in one place. It is handed everything it might need to show and uses only what the
chosen wording calls for.

**It never claims a record was kept when the write failed.** `bRecorded` is what the write actually
returned, and the wording follows it.

```vba
Private Sub ShowErrorToUser(ByVal lErrNumber As Long, _
                            ByVal sErrDescription As String, _
                            ByVal sModule As String, _
                            ByVal sProcedure As String, _
                            ByVal lLine As Long, _
                            ByVal lLogID As Long, _
                            ByVal bRecorded As Boolean)
    ' [SCAFFOLD] Guarded: a failure to show the message must not replace the failure we
    '            were called about.
    Dim sMsg As String

    On Error Resume Next

    ' [WIZARD Step 6 - what the person at the keyboard sees] >>> keep ONE block <<<

    ' --- "A short message, with a reference number when there is one" (preferred) ---
    If bRecorded And lLogID > 0 Then
        sMsg = "Something went wrong. The problem has been recorded." & vbCrLf & vbCrLf & _
               "Reference number: " & lLogID
    ElseIf bRecorded Then
        ' [SCAFFOLD] Recorded to the file, so there is no number to quote. Say the true
        '            thing rather than invent one.
        sMsg = "Something went wrong. The problem has been recorded."
    Else
        sMsg = "Something went wrong, and it could not be recorded. " & _
               "Please make a note of what you were doing."
    End If

    ' --- "A short message that something went wrong" ---
    ' If bRecorded Then
    '     sMsg = "Something went wrong. The problem has been recorded."
    ' Else
    '     sMsg = "Something went wrong, and it could not be recorded. " & _
    '            "Please make a note of what you were doing."
    ' End If

    ' --- "The full technical detail" ---
    ' sMsg = "Error " & lErrNumber & ": " & sErrDescription & vbCrLf & _
    '        "In: " & sModule & "." & sProcedure & vbCrLf & _
    '        "Line: " & lLine
    ' If Not bRecorded Then
    '     sMsg = sMsg & vbCrLf & vbCrLf & _
    '            "This could not be recorded. Please make a note of what you were doing."
    ' End If

    ' --- "Nothing" ---
    ' Exit Sub

    MsgBox sMsg, vbExclamation, "Unexpected problem"
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

```vba
Public Function CreateErrorLogTable(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Creates tblErrorLog (schema template entity). Idempotent: an existing
    '            table is reported and skipped, so this is safe to re-run.
    '            bSilent:=True suppresses the message box and returns the same text.
    Dim db      As DAO.Database
    Dim tdf     As DAO.TableDef
    Dim fld     As DAO.Field
    Dim idx     As DAO.Index
    Dim sReport As String
    Dim bFailed As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb

    sReport = ERROR_LOG_TABLE & " created."

    On Error Resume Next
    Set tdf = db.TableDefs(ERROR_LOG_TABLE)
    On Error GoTo errHandler
    If Not tdf Is Nothing Then
        sReport = ERROR_LOG_TABLE & " already exists and was left exactly as it is."
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
    CreateErrorLogTable = sReport
    ' [SCAFFOLD] One message, whatever happened - the success text or the error text,
    '            never both. Building the report first and showing it here prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] This procedure runs at set-up time, before the
    '            logger it is building exists, so it reports rather than calls LogError.
    bFailed = True
    sReport = "ERROR creating " & ERROR_LOG_TABLE & ": " & Err.Number & " - " & Err.Description
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
                                ByVal sField As String, _
                                ByVal sDescription As String)
    ' [SCAFFOLD + _materialization.md rule 2] Called only after db.TableDefs.Append.
    '            A Description set during field-build raises runtime error 3219.
    Dim prp As DAO.Property

    On Error Resume Next
    Set prp = db.TableDefs(ERROR_LOG_TABLE).Fields(sField) _
                .CreateProperty("Description", dbText, sDescription)
    db.TableDefs(ERROR_LOG_TABLE).Fields(sField).Properties.Append prp

    ' [SCAFFOLD] Guarded because a documentation comment is never worth failing a build
    '            for, and because appending a property that already exists raises 3367.
    Set prp = Nothing
End Sub
```
