---
step: 04-audit-tables-and-config
title: "The three audit tables and the config table"
platform_facts: [dao-table-build, sql-server-ddl, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures (continued)

### Two_CreateAuditTables — `Public Function` → `String` (setup step 2)

Creates the three system tables via DAO, idempotently — an existing table is reported and
skipped, so it is safe to re-run. Field-by-field DAO `CreateField` (never `CREATE TABLE`
DDL — see `templates/_materialization.md`). Each table is built with the indexes the schema
template declares for it, not the primary key alone.

**The skip is whole-table.** A table that already exists is left exactly as it is, indexes
included — re-running this procedure will not add a missing index to a table built by an earlier
version of this scaffold. If you have such a build, add the secondary and unique indexes by
hand, or start fresh in a copy.

```vba
Public Function Two_CreateAuditTables(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Creates tblAuditLog, tblLongTextBackup, tblAuditLogConfig (schema template
    '            entities). Idempotent: each block is skipped if its table already exists.
    '            Passing bSilent:=True suppresses the message box and returns the same text,
    '            so a caller with no one at the keyboard does not hang on a dialog.
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim sReport As String
    Dim bFailed As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [SCAFFOLD] The outcome text is set here, not at the end, because an "already exists"
    '            branch jumps straight to Cleanup. The errHandler replaces it on failure.
    sReport = "Audit tables are in place. Any that already existed were left exactly as " & _
        "they are, indexes included."

    ' ========== tblAuditLog ==========
    On Error Resume Next
    Set tdf = db.TableDefs("tblAuditLog")
    If Not tdf Is Nothing Then
        Debug.Print "tblAuditLog already exists"
        GoTo CreateLongTextBackup
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblAuditLog")

    Set fld = tdf.CreateField("AuditLogID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    ' [BUSINESS LOGIC - schema Business Rule 4] Long Integer unless the developer chose to
    '                 audit tables keyed by a Replication ID. Text holds the key's printed
    '                 form; AUDIT_KEY_TEXT_SIZE carries the width and why.
    If AUDIT_GUID_KEYS_AUDITED Then
        Set fld = tdf.CreateField("PrimaryKey", dbText, AUDIT_KEY_TEXT_SIZE)
    Else
        Set fld = tdf.CreateField("PrimaryKey", dbLong)
    End If
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OperationType", dbText, 25)
    fld.Required = True
    tdf.Fields.Append fld

    ' [_materialization.md rule 6] OldValue and NewValue are sinks: every audited field in
    '            every audited table writes into them, whatever its own type and width. dbMemo
    '            is the widest text ACE has. AllowZeroLength is True and stays True: a source
    '            that permits an empty string sends one, and AuditValueExpression's & "" yields
    '            one from a Null whatever the source permits.
    Set fld = tdf.CreateField("OldValue", dbMemo)
    fld.AllowZeroLength = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("NewValue", dbMemo)
    fld.AllowZeroLength = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DateChanged", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ChangedBy", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("AuditLogID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCHEMA] Secondary indexes the schema template declares: trail queries by table and
    '          date, per-record history by table and key value.
    Set idx = tdf.CreateIndex("TableNameDateChanged")
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("DateChanged")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("TableNamePrimaryKey")
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("PrimaryKey")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblAuditLog created"

CreateLongTextBackup:
    ' ========== tblLongTextBackup ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblLongTextBackup")
    If Not tdf Is Nothing Then
        Debug.Print "tblLongTextBackup already exists"
        GoTo CreateConfig
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblLongTextBackup")

    Set fld = tdf.CreateField("LongTextBackupID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    ' [BUSINESS LOGIC - schema Business Rule 4] Long Integer unless the developer chose to
    '                 audit tables keyed by a Replication ID. Text holds the key's printed
    '                 form; AUDIT_KEY_TEXT_SIZE carries the width and why.
    If AUDIT_GUID_KEYS_AUDITED Then
        Set fld = tdf.CreateField("PrimaryKey", dbText, AUDIT_KEY_TEXT_SIZE)
    Else
        Set fld = tdf.CreateField("PrimaryKey", dbLong)
    End If
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    ' [_materialization.md rule 6] OldValue is a sink, for the same reason and with the same
    '            settings as the log table's — see the comment there.
    Set fld = tdf.CreateField("OldValue", dbMemo)
    fld.AllowZeroLength = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DateChanged", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ChangedBy", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("LongTextBackupID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCHEMA] Unique per table/record/field — BackupLongTextFieldsDM replaces any earlier
    '          backup for the same field of the same row, and this enforces that one-row rule.
    Set idx = tdf.CreateIndex("TableNamePrimaryKeyFieldName")
    idx.Unique = True
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("PrimaryKey")
    idx.Fields.Append fld
    Set fld = idx.CreateField("FieldName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblLongTextBackup created"

CreateConfig:
    ' ========== tblAuditLogConfig ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblAuditLogConfig")
    If Not tdf Is Nothing Then
        Debug.Print "tblAuditLogConfig already exists"
        GoTo Cleanup
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblAuditLogConfig")

    Set fld = tdf.CreateField("AuditLogConfigID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldPosition", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DataType", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("IsPrimaryKey", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"     ' [SCAFFOLD] set before Append — never as DDL DEFAULT
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("IsAuditable", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "True"      ' [SCAFFOLD] audit scope is decided in this flag, as data
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("AuditLogConfigID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    ' [SCHEMA] Unique per table/field — one config row per scanned field, so a re-scan or a
    '          hand edit cannot leave two rows disagreeing about the same field.
    Set idx = tdf.CreateIndex("TableNameFieldName")
    idx.Unique = True
    Set fld = idx.CreateField("TableName")
    idx.Fields.Append fld
    Set fld = idx.CreateField("FieldName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblAuditLogConfig created"

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Two_CreateAuditTables = sReport
    ' [SCAFFOLD] One message, whatever happened — the success text or the error text, never
    '            both. Building the report first and showing it here is what prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    bFailed = True
    sReport = "ERROR creating tables: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### Three_PopulateConfigTable — `Public Function` → `String` (setup step 3)

Scans the schema into `tblAuditLogConfig`: **every field of every candidate table**, with its
ordinal position, DAO type code, a flag on the table's PK field, and `IsAuditable`. Nothing is
silently dropped — exclusions are *seeded* as `IsAuditable = False` rows: the three system
tables, the audited table's own primary-key field (its value is already on every log row, in
`tblAuditLog.PrimaryKey`), plus fields that would just add noise — this repo's house audit columns
(`CreatedDate`/`CreatedBy`/`ModifiedDate`/`ModifiedBy`/`AccessTS`, per `standards/audit-columns.md`)
and a few other always-changing system columns (`SSMA_TimeStamp`, `ValidFrom`, `ValidTo`). **After
running, open the config table and review the flags** — that review, in data, is where the audit
net is drawn (schema Business Rule 5).

**Two field types are seeded off and stay off**, whatever that review says — `IsUnauditableFieldType`
decides, and `Four_GenerateAllAuditDataMacros` asks it again rather than trusting the switch. An
**Attachment** field cannot be read or written by a Data Macro at all, so a macro referencing one
does not work. A **calculated** field is never edited by anyone: its value is derived from other
fields in the same row, and those fields are audited themselves, so a row in the log for it would
record a change nobody made. Seeding them off is not enough on its own — the config table exists to
be edited, and a switch that can be turned back on will be.

**What you will see, so it doesn't look wrong:** most of the rows belong to the three system
tables — `tblAuditLog`, `tblLongTextBackup`, `tblAuditLogConfig` — and on a small database that
can be well over half of them. They are all switched OFF and must stay OFF; auditing the audit
trail would loop. Sort by `TableName` and review only the rows for tables you recognise as your
own. The system rows are shown rather than hidden on purpose — nothing the scan did is invisible
to you.

Takes one optional Yes/No setting that decides the starting point for everything else:

- **Path A (try-it-out build):** run `Three_PopulateConfigTable` with nothing after it. Every
  field starts switched ON, and you switch OFF the few you don't want tracked.
- **Path B (a database you already use):** run `Three_PopulateConfigTable False`. Every field
  starts switched OFF, and you switch ON — table by table — only what you actually want a
  history of. This is the safer starting point on tables this system wasn't designed around,
  where "track everything" could sweep in more than you meant.

```vba
Public Function Three_PopulateConfigTable(Optional bDefaultAuditable As Boolean = True, _
                                        Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Rebuild the audit configuration from the live schema. Scope decisions
    '            live in the IsAuditable flags afterward, not in this code.
    '            Passing bSilent:=True suppresses the message box and returns the same text,
    '            so a caller with no one at the keyboard does not hang on a dialog.
    '            bDefaultAuditable sets the starting point for ordinary fields only:
    '            True  (preferred; demo build) — everything starts switched ON, you switch OFF
    '                  what you don't want tracked.
    '            False (Path B — call as Three_PopulateConfigTable(False)) — everything
    '                  starts switched OFF, you switch ON what you do want tracked.
    '            The three system tables and the noisy always-changing fields below are
    '            always switched OFF, no matter which way this is called.
    Dim db As DAO.Database
    Dim tdef As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim pkField As DAO.Field
    Dim sSql As String
    Dim isPK As Boolean
    Dim isAuditable As Boolean
    Dim pkFieldName As String
    Dim lRowCount As Long
    Dim lTablesInScope As Long
    Dim sLinkedNamed As String
    Dim sReport As String
    Dim bFailed As Boolean

    On Error GoTo errHandler
    Set db = CurrentDb
    lRowCount = 0
    lTablesInScope = 0
    sLinkedNamed = ""

    ' Clear existing config
    db.Execute "DELETE * FROM tblAuditLogConfig", dbFailOnError

    For Each tdef In db.TableDefs
        ' [BUSINESS LOGIC — scan boundary] Which tables are candidates at all. The test lives
        ' in IsAuditCandidateTable, which One_CheckAuditReadiness calls as well, so the two cannot
        ' disagree about what is in scope. It reads AUDIT_SCOPE_MODE; change that, not this.
        ' A table the developer named that turns out to be linked is collected here and named
        ' in the report below, so a table asked for by name is never dropped in silence.
        If (tdef.Attributes And (dbAttachedTable Or dbAttachedODBC)) <> 0 Then
            If IsNamedInScopeList(tdef.Name) Then
                sLinkedNamed = sLinkedNamed & "  " & tdef.Name & vbCrLf
            End If
        End If

        If IsAuditCandidateTable(tdef) Then
            lTablesInScope = lTablesInScope + 1

            ' Get the primary key field name for this table
            pkFieldName = ""
            For Each idx In tdef.Indexes
                If idx.Primary Then
                    For Each pkField In idx.Fields
                        pkFieldName = pkField.Name
                        Exit For
                    Next pkField
                    Exit For
                End If
            Next idx

            For Each fld In tdef.Fields
                isPK = (fld.Name = pkFieldName)

                ' [SCAFFOLD] Seed IsAuditable: False for the system tables themselves
                '            (schema Business Rule 5 — never audit the audit trail) and for
                '            noisy always-changing fields; the starting point set by
                '            bDefaultAuditable for everything else. Review and flip flags in
                '            tblAuditLogConfig after the scan.
                ' [STANDARDS — audit-columns.md] The four house audit columns come from the
                '            constants at the top of this module — one place to change if your
                '            standards/audit-columns.md names them differently. They are a
                '            VBA-side mirror of that file, not a live read of it.
                '            AccessTS is named here as a literal because it is a SQL Server
                '            rowversion that only appears on linked tables, so it is never one
                '            of the columns this generator creates or stamps.
                '            SSMA_TimeStamp/ValidFrom/ValidTo are not house audit columns; they
                '            are left here because a table carrying them already has its own
                '            change-tracking mechanism (e.g. SQL Server temporal system-versioning)
                '            that this scan would otherwise log as noisy, always-changing values.
                Select Case True
                    Case tdef.Name = "tblAuditLog", _
                         tdef.Name = "tblLongTextBackup", _
                         tdef.Name = "tblAuditLogConfig"
                        isAuditable = False
                    Case fld.Name = AUDIT_CREATED_DATE, fld.Name = AUDIT_CREATED_BY, _
                         fld.Name = AUDIT_MODIFIED_DATE, fld.Name = AUDIT_MODIFIED_BY, _
                         fld.Name = "AccessTS", fld.Name = "SSMA_TimeStamp", _
                         fld.Name = "ValidFrom", fld.Name = "ValidTo"
                        isAuditable = False
                    ' [SCAFFOLD] The table's own primary key. Its value is already on every
                    '            log row in tblAuditLog.PrimaryKey, which is what identifies
                    '            the record; a field row here would store the same value a
                    '            second time (schema Business Rule 5). A developer can flip
                    '            it on afterward if they want that.
                    Case isPK
                        isAuditable = False
                    ' [BUSINESS LOGIC] Attachment and calculated fields, which are never
                    '            audited on any build — see IsUnauditableFieldType for why
                    '            each one is impossible or pointless rather than merely
                    '            unwanted. Seeded off here so the developer can see them in
                    '            the config table and know they were considered; enforced
                    '            again in Four_GenerateAllAuditDataMacros, because this
                    '            table is meant to be edited and a switch that can be turned
                    '            back on will be.
                    Case IsUnauditableFieldType(fld)
                        isAuditable = False
                    Case Else
                        isAuditable = bDefaultAuditable
                End Select

                ' [STANDARDS — query-style.md] inline INSERT kept from the working source
                sSql = "INSERT INTO tblAuditLogConfig " & _
                    "(TableName, FieldName, FieldPosition, DataType, IsPrimaryKey, IsAuditable) " & _
                    "VALUES ('" & tdef.Name & "', '" & fld.Name & "', " & fld.OrdinalPosition & _
                    ", " & fld.Type & ", " & isPK & ", " & isAuditable & ")"
                db.Execute sSql, dbFailOnError
                lRowCount = lRowCount + 1
            Next fld
        End If
    Next tdef

    ' [SCAFFOLD] Nothing in scope is a stop, not a result. The config table ends up empty
    '            either way; what differs is that "0 field row(s) written", phrased like a
    '            successful run and shown with the same icon, reads as success — and the two
    '            steps after this one then report success as well, on nothing at all.
    If lTablesInScope = 0 Then
        bFailed = True
        sReport = "Stopped. No table in this file is in scope, so nothing was written to " & _
            "tblAuditLogConfig and nothing has been built." & vbCrLf & vbCrLf & _
            "The scope setting at the top of this module is """ & AUDIT_SCOPE_MODE & """:" & vbCrLf & _
            "  Standard - tables named tbl... or tlkp..." & vbCrLf & _
            "  All      - every table in this file." & vbCrLf & _
            "  List     - only the tables named in AUDIT_SCOPE_LIST." & vbCrLf & vbCrLf & _
            "Set it to match the tables in this file and run this again. System tables, " & _
            "temporary tables and linked tables are left out whichever one is set."
        GoTo Cleanup
    End If

    sReport = "Table list built: " & lRowCount & " field row(s) written to tblAuditLogConfig." & _
        vbCrLf & "Open tblAuditLogConfig and check the IsAuditable switches before you run " & _
        "the next step."

    ' [SCAFFOLD] A table the developer named by hand and did not get. Said here rather than
    '            left to be noticed, because the whole point of naming tables one at a time
    '            is that the developer decided which ones matter.
    If Len(sLinkedNamed) > 0 Then
        sReport = sReport & vbCrLf & vbCrLf & _
            "These tables are named in AUDIT_SCOPE_LIST but are linked to another file, so " & _
            "they were left out. Tracking attaches to the table itself and has to be built " & _
            "in the file where the table really lives:" & vbCrLf & sLinkedNamed
    End If

Cleanup:
    Set pkField = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set tdef = Nothing
    Set db = Nothing
    Three_PopulateConfigTable = sReport
    ' [SCAFFOLD] One message, whatever happened — the success text or the error text, never
    '            both. Building the report first and showing it here is what prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    bFailed = True
    sReport = "ERROR populating config: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```
