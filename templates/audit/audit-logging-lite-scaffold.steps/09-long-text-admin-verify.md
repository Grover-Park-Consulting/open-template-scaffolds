---
step: 09-long-text-admin-verify
title: "The long-text, admin and verify modules"
platform_facts: [domain-function-transaction, mcp-module-import, mcp-line-numbers]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures (continued)

### AuditUser — `Public Function` → `String` (module `modAuditLongText` — back end AND front end)

**The single most important four lines in this template, and the easiest to leave out.** Every
generated stamping macro calls `AuditUser()` on **every** table — Long Text or not — to fill
`CreatedBy` and `ModifiedBy`. Those columns are `Required`, so a table whose macro cannot resolve
this function **rejects every insert**, with an error that names the function rather than the cause.

`modAuditLongText`'s name undersells it. A database with no Long Text field anywhere still needs
this module, in the back end **and in every front end**, because a macro fired by an edit through a
linked table looks for the function in *that person's* front end.

Returns `"Unknown"` rather than an empty string (schema Business Rule 9): `Environ$("USERNAME")`
comes back empty in some contexts — a scheduled task, a service account, a locked-down profile —
and because `CreatedBy` is `Required`, an empty result would block the insert outright. A row naming
an unknown user is a record; a refused write is not.

```vba
Public Function AuditUser() As String
    ' [STANDARDS — audit-columns.md] The identity every generated stamping macro calls, as
    '            =AuditUser(). Never returns an empty string: CreatedBy is Required, and an
    '            empty value would block the insert (schema Business Rule 9).
    '            CurrentUser() is a named alternative — see Extra Options — but it returns
    '            "Admin" for everyone unless workgroup security is in use.
    AuditUser = Environ$("USERNAME")
    If Len(AuditUser) = 0 Then AuditUser = "Unknown"
End Function
```

### BackupLongTextFieldsDM — `Public Function` (module `modAuditLongText` — back end AND front end)

The VBA half of the hybrid Long Text method, called *by the Before macros themselves*. Replaces
any earlier backup for the same table/field/record, reads the current Long Text value, and writes
it to `tblLongTextBackup` for the After macro to retrieve.

**Placement is the trap:** it must exist in the **back end** (so `LoadFromText` resolves the name
when the macros are created) **and in every front end** (a macro fired by a front-end edit
resolves the function in the front end's VBA project). Missing it in either place surfaces as a
data-macro execution error at save time.

```vba
Public Function BackupLongTextFieldsDM(strTableName As String, varPKValue As Variant, strFieldName As String)
    ' [SCAFFOLD] Stage one Long Text field's current value before an update or delete.
    '            Called by the generated BeforeChange / BeforeDelete Data Macros.
    ' [BUSINESS LOGIC — schema Business Rule 4] The key is a Variant, not a Long, for one
    '            reason: Null is the marker for "there is no prior row to copy". It is the only
    '            value that cannot be a primary key of any type, because a primary key is
    '            Required by definition. The guard this replaced tested the key value itself
    '            (If lngPKValue > 0), which silently skipped the backup for a row whose key is
    '            legitimately 0 — reachable on any Long Integer, Integer or Byte key that is
    '            not an AutoNumber.
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsOldValue As DAO.Recordset
    Dim strPKField As String
    Dim strOldValue As Variant

    On Error GoTo errHandler
    Set db = CurrentDb

    If strTableName = "" Then Exit Function
    If IsNull(varPKValue) Then Exit Function   ' an insert — there is no prior value to stage

    ' Replace any earlier backup for this table/field/record
    db.Execute "DELETE FROM tblLongTextBackup WHERE TableName='" & strTableName & _
        "' AND FieldName='" & strFieldName & "' AND PrimaryKey=" & _
        BackupKeyLiteral(varPKValue), dbFailOnError

    ' The audited table's PK field name comes from the config
    strPKField = DLookup("FieldName", "tblAuditLogConfig", _
        "TableName='" & strTableName & "' AND IsPrimaryKey=" & True)

    Set rsOldValue = db.OpenRecordset("SELECT " & strFieldName & " FROM " & strTableName & _
        " WHERE " & strPKField & "=" & SourceKeyLiteral(strTableName, strPKField, varPKValue))

    ' [SCAFFOLD] A key was supplied, so the row it names has to exist — it is the row being
    '            changed or deleted, and it is locked inside this transaction. Finding nothing
    '            means the literal above did not match, so no backup is staged. The delete
    '            above has already removed any earlier one, and the After macro creates its
    '            log row inside the LookUpRecord that reads this table — so the field is not
    '            logged at all, rather than logged wrongly. Raise it so the failure is not
    '            merely silent: the handler below stays quiet for the person editing, and a
    '            house logger sees it where the standards layer provides one.
    If rsOldValue.EOF Then
        rsOldValue.Close
        Err.Raise vbObjectError + 514, "BackupLongTextFieldsDM", _
            "No row found in " & strTableName & " for key " & CStr(varPKValue) & _
            ". The key could not be matched, so no Long Text backup was staged."
    End If

    strOldValue = rsOldValue.Fields(strFieldName).Value
    rsOldValue.Close

    Set rs = db.OpenRecordset("tblLongTextBackup", dbOpenDynaset)
    rs.AddNew
    rs!TableName = strTableName
    rs!PrimaryKey = varPKValue
    rs!FieldName = strFieldName
    rs!OldValue = strOldValue
    rs!DateChanged = Now()
    ' [STANDARDS / schema Business Rule 9] AuditUser() preferred choice, same as the macros. This
    '            row is the one the After macro reads back, so it is the easiest of the
    '            four sites to miss when applying the CurrentUser() Extra Option — and
    '            missing it puts two names on one edit.
    rs!ChangedBy = AuditUser()
    rs.Update
    rs.Close

Cleanup:
    On Error Resume Next
    Set rsOldValue = Nothing
    Set rs = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] deliberately quiet: this runs inside a data-macro save;
    '            a MsgBox here would interrupt every user's save. Log if your house pattern
    '            has a silent logger; never block.
    Resume Cleanup
    Resume
End Function

' [BUSINESS LOGIC — schema Business Rule 4] Two literals, because the two columns are not the
' same type, and both ASK rather than being told. This module ships to every front end on its
' own; the module holding the build settings is back end only, so a setting read here would not
' exist where this runs. Both tables are linked into the front end, so their column types are
' readable here and are the authority in any case.
Private Function BackupKeyLiteral(varPKValue As Variant) As String
    ' The staging table's own key column decides how the key is written into it.
    If CurrentDb.TableDefs("tblLongTextBackup").Fields("PrimaryKey").Type = dbText Then
        BackupKeyLiteral = "'" & CStr(varPKValue) & "'"
    Else
        BackupKeyLiteral = CStr(varPKValue)
    End If
End Function

Private Function SourceKeyLiteral(ByVal strTableName As String, ByVal strPKField As String, _
                                  varPKValue As Variant) As String
    ' [BUSINESS LOGIC — schema Business Rule 4] Asked per table, never per build. A build that
    '                 audits Replication ID keys still holds whole-number keyed tables, so one
    '                 global answer would wrap an ordinary key as {guid 123} and match nothing.
    '                 The audited table's own key field keeps its real type whatever the log's
    '                 key column holds, and Jet matches a Replication ID with the {guid {…}}
    '                 form and nothing else — ordinary quotes do not work.
    Dim sKey As String

    If CurrentDb.TableDefs(strTableName).Fields(strPKField).Type <> dbGUID Then
        SourceKeyLiteral = CStr(varPKValue)
        Exit Function
    End If

    ' A Data Macro hands the key over already printed, as "{C6C0FB4C-…}". DAO prints the same
    ' key as "{guid {C6C0FB4C-…}}" — the form Jet needs — so add the wrapper only when it is
    ' not already there, and this works whichever way the value arrived.
    sKey = CStr(varPKValue)
    If Left$(sKey, 5) = "{guid" Then
        SourceKeyLiteral = sKey
    Else
        SourceKeyLiteral = "{guid " & sKey & "}"
    End If
End Function
```

### BackupAndRemoveAllDataMacros — `Public Function` → `Boolean` (module `modAuditAdmin` — back end only)

The reset tool for regeneration (schema Business Rule 7): exports a table's current data macros to
a timestamped XML backup, then strips them by loading an empty macro document. Run it before
re-running `Four_GenerateAllAuditDataMacros` when the audit scope changes — the backups double as
your archive of prior macro states.

**It removes only the macros this generator created, and it never touches an Access system table.**
Two guards, and both matter because this is the tool a developer reaches for when something has
already gone wrong:

- **Two kinds of system table are skipped, for two different reasons.** Access's own — names
  starting `MSys` — are never touched, not even exported: this loop finds its work by asking which
  tables carry a data macro, and Access's own tables can answer yes. A table whose name starts
  `USys` is the *developer's* own hidden table, not Access's; it is theirs to manage, so it is left
  alone unless they opted it in themselves by putting it in `tblAuditLogConfig`. Both are name
  checks because the export has to be skipped too, and the ownership test below cannot run until
  something has been exported.
- **A macro set that is not this generator's is backed up and left in place** (`MacroBackupIsOurs`).
  Business-logic macros of the developer's own, and any who-and-when stamping the database had
  before this system arrived, are not this tool's to throw away. The export is still taken, so the
  developer has a record of what was found either way.

The trade this makes: a table whose generated macros are only the who-and-when stamping names no
audit table, so it reads as not-ours and is left alone. That is the safe direction — the tool
declines to strip rather than stripping something it should not have.

```vba
Public Function BackupAndRemoveAllDataMacros(Optional strBackupPath As String = "", _
                                             Optional bSilent As Boolean = False) As Boolean
    ' [SCAFFOLD] Back up every table's data macros, then remove the ones this generator wrote.
    '            Passing bSilent:=True suppresses both message boxes, so a caller with no one
    '            at the keyboard does not hang on a dialog. The Boolean return is the result.
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strTempFile As String
    Dim strBackupFile As String
    Dim strTableName As String
    Dim strSummary As String
    Dim intFileNum As Integer
    Dim intMacrosRemoved As Integer
    Dim intMacrosKept As Integer

    On Error GoTo errHandler
    Set db = CurrentDb
    intMacrosRemoved = 0
    intMacrosKept = 0

    If strBackupPath = "" Then
        strBackupPath = CurrentProject.Path & "\DataMacroBackups\"
    End If
    If Dir(strBackupPath, vbDirectory) = "" Then
        MkDir strBackupPath
    End If

    ' An empty macro document: loading it replaces (removes) a table's data macros
    strTempFile = Environ("TEMP") & "\BlankDataMacro.xml"
    intFileNum = FreeFile
    Open strTempFile For Output As intFileNum
    Print #intFileNum, "<?xml version=""1.0"" encoding=""UTF-16""?>"
    Print #intFileNum, "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/04/application"">"
    Print #intFileNum, "</DataMacros>"
    Close #intFileNum

    ' Tables with data macros: MSysObjects.LvExtra is non-null for them
    strSQL = "SELECT [Name] FROM MSysObjects " & _
             "WHERE Not IsNull(LvExtra) AND Type = 1 " & _
             "ORDER BY [Name]"
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)

    Do While Not rst.EOF
        strTableName = rst!Name

        ' [SCAFFOLD] Two kinds of system table, skipped for two different reasons. Access's own
        '            (MSys) are never ours to touch, not even to export — the query above finds
        '            tables by asking which ones carry a data macro, and Access's own can answer
        '            yes. A USys table is the DEVELOPER's own hidden table: theirs to manage, so
        '            it is left alone unless they opted it in themselves by putting it in
        '            tblAuditLogConfig. Both guards are on the name because the export has to be
        '            skipped too, and the ownership test below cannot run without one.
        If Left$(strTableName, 4) = "MSys" Then
            Debug.Print "Skipped (Access's own system table): " & strTableName
        ElseIf Left$(strTableName, 4) = "USys" _
            And DCount("*", "tblAuditLogConfig", "TableName='" & strTableName & "'") = 0 Then
            Debug.Print "Skipped (your own system table, not opted in): " & strTableName
        Else
            Debug.Print "Processing: " & strTableName

            strBackupFile = strBackupPath & strTableName & "_DataMacro_" & _
                            Format(Now(), "yyyymmdd_hhnnss") & ".xml"
            Application.SaveAsText acTableDataMacro, strTableName, strBackupFile

            ' [SCAFFOLD] Remove only what this generator wrote. The export is taken either
            '            way, so the developer has a record of what was on the table; only
            '            the removal is conditional. A table carrying someone else's business
            '            logic, or who-and-when stamping older than this system, keeps it.
            If MacroBackupIsOurs(strBackupFile) Then
                Application.LoadFromText acTableDataMacro, strTableName, strTempFile
                intMacrosRemoved = intMacrosRemoved + 1
            Else
                intMacrosKept = intMacrosKept + 1
                Debug.Print "  - left in place, not this generator's macros: " & strTableName
            End If
        End If

        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
    Kill strTempFile

    strSummary = "Backed up and removed data macros from " & intMacrosRemoved & " table(s)."
    If intMacrosKept > 0 Then
        strSummary = strSummary & vbCrLf & vbCrLf & intMacrosKept & " table(s) were backed up " & _
            "but left as they are, because their Data Macros were not created by this system. " & _
            "Removing those is a separate decision, and yours to make."
    End If
    strSummary = strSummary & vbCrLf & vbCrLf & "Backups saved to: " & strBackupPath

    If Not bSilent Then MsgBox strSummary, vbInformation, "Data Macros Removed"

    BackupAndRemoveAllDataMacros = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    '            One exit for every error: an error here stops the whole run rather than
    '            carrying on to the next table, because a failure part-way through means the
    '            database is in a half-reset state and the developer needs to know that now.
    '            The two skips inside the loop are decisions, not errors, and never come here.
    If Not bSilent Then MsgBox Err.Number & " Error: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    If Dir(strTempFile) <> "" Then Kill strTempFile
    BackupAndRemoveAllDataMacros = False
    Resume Cleanup
    Resume
End Function
```

### DumpTableMacros — `Public Function` → `String` (module `modAuditVerify` — back end only)

**Check the artifact, not the report.** Every procedure above returns a status line saying what it
did. That line is the generator's own account of itself — if the generator is wrong, the line is
wrong in exactly the same way, and everything still looks fine. This function and `ListMacroEvents`
below read the macros back **off the table**, and show what is actually attached. That is the only
answer capable of contradicting the generator.

`? DumpTableMacros("tblSupportTicket")` in the Immediate window returns the full macro XML — use it
when you need to see the actions *inside* a macro rather than just which macros exist.

It is read-only: it exports the macro set to a temporary file, reads it, and deletes the file.
Nothing is changed. `SaveAsText` writes UTF-16, which is why the file is opened with the `-1`
(TristateTrue) argument — reading it as ANSI returns unusable text.

```vba
Option Compare Database
Option Explicit

Public Function DumpTableMacros(ByVal sTable As String) As String
    ' [SCAFFOLD] Read a table's attached Data Macro set back out, so a build can be
    '            verified against the table itself rather than the generator's own report.
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim txt As Object
    Dim sPath As String
    Dim sOut As String

    On Error GoTo errHandler

    sPath = Environ$("TEMP") & "\" & sTable & "_verify.xml"
    If Dir(sPath) <> "" Then Kill sPath

    Application.SaveAsText acTableDataMacro, sTable, sPath

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.OpenTextFile(sPath, 1, False, -1)   ' -1 = TristateTrue, i.e. UTF-16
    sOut = txt.ReadAll
    txt.Close

    fso.DeleteFile sPath
    DumpTableMacros = sOut

Cleanup:
    On Error Resume Next
    Set txt = Nothing
    Set fso = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    DumpTableMacros = "ERROR: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### ListMacroEvents — `Public Function` → `String` (module `modAuditVerify` — back end only)

The compact check, and the one to run first. It lists just the macro events attached to a table, in
order, with the size of the set:

```text
? ListMacroEvents("tblSupportTicket")
tblSupportTicket: AfterInsert | AfterUpdate | AfterDelete | BeforeChange | BeforeDelete | (24063 chars)
```

Read it against Business Rule 2 — four macros on an ordinary table carrying the house audit columns,
five when it also has an audited Long Text field. **A missing event means the generator did not do
what its status line said it did**, and that is exactly the discrepancy worth catching before anyone
trusts the audit trail.

```vba
Public Function ListMacroEvents(ByVal sTable As String) As String
    ' [SCAFFOLD] Just the Event= names, in order — a one-line structural check that a table
    '            carries the macros the generator said it created.
    Dim sXml As String
    Dim lPos As Long
    Dim lEnd As Long
    Dim sOut As String

    On Error GoTo errHandler

    sXml = DumpTableMacros(sTable)
    If Left(sXml, 6) = "ERROR:" Then
        ListMacroEvents = sXml
        Exit Function
    End If

    lPos = InStr(1, sXml, "Event=", vbTextCompare)
    Do While lPos > 0
        lEnd = InStr(lPos + 7, sXml, Chr(34))
        sOut = sOut & Mid(sXml, lPos + 7, lEnd - lPos - 7) & " | "
        lPos = InStr(lEnd, sXml, "Event=", vbTextCompare)
    Loop

    ListMacroEvents = sTable & ": " & sOut & "(" & Len(sXml) & " chars)"

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    ListMacroEvents = "ERROR: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```
