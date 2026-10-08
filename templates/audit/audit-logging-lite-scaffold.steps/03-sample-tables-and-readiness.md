---
step: 03-sample-tables-and-readiness
title: "The sample tables and the readiness check"
platform_facts: [table-in-use]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures

### Zero_CreateSampleTables — `Public Function` → `String` (Path A only — setup step 0)

**Skip this procedure entirely if you're doing Path B** (adding audit tracking to a database
you already use) — it only exists to build the made-up tables for trying the system out.

Creates the two made-up tables (`tblClient`, `tblSupportTicket`) and a short pick-list
(`tlkpTicketPriority`) described in the paired schema template, with the four starter pick-list
rows (Low, Normal, High, Urgent) and the two links between the tables. Same idempotent style as
`Two_CreateAuditTables` — an existing table is reported and skipped, so it's safe to re-run.

**All three carry the house audit columns** (`standards/audit-columns.md`), because a demo that
leaves them off doesn't demonstrate the thing most likely to bite on real tables: stamping and
change-auditing both wanting the Before Change macro. With them present, the Path A build shows the
two being generated together, which is what a real table needs. It also means the four seed rows
have to supply `CreatedDate` and `CreatedBy` themselves — they're inserted at step 0, before any
stamping macro exists, and those columns are `Required`. See the comment on the `INSERT` statements
below, and `templates/_materialization.md` rule 5 for why the values must be resolved in VBA first.

It returns the same wording it shows on screen. Run it from the Immediate window as
`Zero_CreateSampleTables` and you get the message box; call it as
`sResult = Zero_CreateSampleTables(True)` and you get the text back with **no** message box —
which is what an automated caller needs, because nothing is there to click a dialog away.

```vba
Public Function Zero_CreateSampleTables(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Creates tblClient, tlkpTicketPriority, tblSupportTicket (schema template
    '            entities) and seeds tlkpTicketPriority. Path A (try-it-out build) only —
    '            skip this procedure for Path B (an existing accdb's own tables). Idempotent:
    '            each block is skipped if its table already exists.
    '            Passing bSilent:=True suppresses the message box and returns the same text,
    '            so a caller with no one at the keyboard does not hang on a dialog.
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim rel As DAO.Relation
    Dim sReport As String
    Dim bFailed As Boolean
    Dim sAuditUser As String
    Dim sAuditNow As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' ========== tblClient ==========
    On Error Resume Next
    Set tdf = db.TableDefs("tblClient")
    If Not tdf Is Nothing Then
        Debug.Print "tblClient already exists"
        GoTo CreateTicketPriority
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblClient")

    Set fld = tdf.CreateField("ClientID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ClientName", dbText, 100)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("EmailAddress", dbText, 100)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("CellPhone", dbText, 25)
    tdf.Fields.Append fld

    AddAuditColumns tdf

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ClientID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("ClientName")
    idx.Unique = True
    Set fld = idx.CreateField("ClientName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblClient created"

CreateTicketPriority:
    ' ========== tlkpTicketPriority ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tlkpTicketPriority")
    If Not tdf Is Nothing Then
        Debug.Print "tlkpTicketPriority already exists"
        GoTo CreateSupportTicket
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tlkpTicketPriority")

    Set fld = tdf.CreateField("TicketPriorityID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketPriorityName", dbText, 30)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("SortOrder", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    AddAuditColumns tdf

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("TicketPriorityID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("TicketPriorityName")
    idx.Unique = True
    Set fld = idx.CreateField("TicketPriorityName")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    db.TableDefs.Refresh

    ' [SCAFFOLD + _materialization.md rule 5] Starter pick-list rows (schema:
    '            Low/Normal/High/Urgent). These are inserted at step 0, BEFORE any
    '            stamping macro exists — and CreatedDate/CreatedBy are Required, so the
    '            values have to be supplied here. They cannot be supplied as functions:
    '            the ACE engine, not VBA, evaluates the text of a db.Execute INSERT, and
    '            Environ() and AuditUser() are unknown to it ("Undefined function").
    '            Resolve each value in VBA first, then concatenate it in as a literal.
    sAuditUser = Environ$("USERNAME")
    If Len(sAuditUser) = 0 Then sAuditUser = "Unknown"
    sAuditNow = "#" & Format(Now(), "yyyy-mm-dd hh:nn:ss") & "#"

    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('Low', 10, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('Normal', 20, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('High', 30, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder, " & _
        AUDIT_CREATED_DATE & ", " & AUDIT_CREATED_BY & ") VALUES ('Urgent', 40, " & _
        sAuditNow & ", '" & sAuditUser & "')", dbFailOnError

    Debug.Print "tlkpTicketPriority created and seeded"

CreateSupportTicket:
    ' ========== tblSupportTicket ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblSupportTicket")
    If Not tdf Is Nothing Then
        Debug.Print "tblSupportTicket already exists"
        GoTo CreateRelationships
    End If
    On Error GoTo errHandler

    Set tdf = db.CreateTableDef("tblSupportTicket")

    Set fld = tdf.CreateField("SupportTicketID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ClientID", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketPriorityID", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketSubject", dbText, 255)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketDetail", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketOpenedDate", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TicketClosedDate", dbDate)
    tdf.Fields.Append fld

    AddAuditColumns tdf

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("SupportTicketID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("ClientID")
    Set fld = idx.CreateField("ClientID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Set idx = tdf.CreateIndex("TicketPriorityID")
    Set fld = idx.CreateField("TicketPriorityID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblSupportTicket created"

CreateRelationships:
    ' ========== Relationships (schema: enforced, no cascade) ==========
    On Error Resume Next
    db.Relations.Delete "tblClient_tblSupportTicket"
    db.Relations.Delete "tlkpTicketPriority_tblSupportTicket"
    On Error GoTo errHandler

    Set rel = db.CreateRelation("tblClient_tblSupportTicket", "tblClient", "tblSupportTicket", 0)
    Set fld = rel.CreateField("ClientID")
    fld.ForeignName = "ClientID"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("tlkpTicketPriority_tblSupportTicket", "tlkpTicketPriority", "tblSupportTicket", 0)
    Set fld = rel.CreateField("TicketPriorityID")
    fld.ForeignName = "TicketPriorityID"
    rel.Fields.Append fld
    db.Relations.Append rel

    Debug.Print "Relationships created"

    sReport = "Sample tables created. You're on Path A (try-it-out build) — nothing in your " & _
        "own database was touched."

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set rel = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Zero_CreateSampleTables = sReport
    ' [SCAFFOLD] One message, whatever happened — the success text or the error text, never
    '            both. Building the report first and showing it here is what prevents that.
    If Not bSilent Then MsgBox sReport, IIf(bFailed, vbCritical, vbInformation)
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    bFailed = True
    sReport = "ERROR creating sample tables: " & Err.Number & " - " & Err.Description
    Resume Cleanup
    Resume
End Function
```

### AddAuditColumns — `Private Sub` (Path A only — helper to step 0)

Appends the house audit set to a table being built, always last in column order (per
`standards/naming-conventions.md` §6.4). Called once per sample table by
`Zero_CreateSampleTables`, before the table is appended to the database.

`AccessTS` from the standards file is deliberately **not** created: it's a SQL Server rowversion
column, and a local Access table has no equivalent. None of the four is Long Text — a Data Macro
cannot write a Long Text field at all, which is why the audit set is Short Text and Date/Time by
design.

**On Path B this procedure is not used.** Your own tables already have whatever audit columns your
standards give them; the generator reads what's there rather than adding anything.

```vba
Private Sub AddAuditColumns(tdf As DAO.TableDef)
    ' [STANDARDS — audit-columns.md] The house audit set, always last in column order.
    '            Filled at run time by the Before Change stamping macro that
    '            BuildBeforeChangeMacro emits together with the audit macros.
    '            Column NAMES come from the constants at the top of this module — change
    '            them there, not here, if your house uses different names.
    Dim fld As DAO.Field

    Set fld = tdf.CreateField(AUDIT_CREATED_DATE, dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField(AUDIT_CREATED_BY, dbText, 100)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField(AUDIT_MODIFIED_DATE, dbDate)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField(AUDIT_MODIFIED_BY, dbText, 100)
    tdf.Fields.Append fld

    Set fld = Nothing
End Sub
```

### One_CheckAuditReadiness — `Public Function` → `String` (setup step 1 — the safety check, required for Path B)

A read-only check you can run any time, at no risk — it doesn't change anything. It looks at each
table you might track and tells you whether this system will actually work on it: **every table
needs one, single number field as its primary key, set to auto-number** (schema Business Rule 4).
Most tables you design yourself already look like this. Older or borrowed tables sometimes
don't — a table with no primary key set, one that uses two or more fields together as its key,
or one that uses a text code instead of a number, will not work with this system as-is.

**One kind of key is the developer's choice rather than a flat exclusion: a Replication ID.**
Where this check finds one, it says so and describes what auditing those tables would cost, so
the answer can be given before anything is built. Where it finds none, the question never
arises. `AUDIT_GUID_KEYS_AUDITED` carries that answer.

Run this **first**, before `Two_CreateAuditTables`. It reads table definitions and changes
nothing, so the developer learns which tables cannot be audited while there is still nothing to
undo — and the Replication ID question has to be answered before `Two_CreateAuditTables` builds
the log's key column. This step is **required for Path B**, since a database you didn't design
the audit system around is far more likely to have a table shaped this way. On Path A run it
after `Zero_CreateSampleTables`, which creates the tables it looks at; there it has nothing to
find, because those tables are already known to be shaped correctly.

If a table isn't ready, you have two choices: fix that table's primary key, or leave it out —
open `tblAuditLogConfig` and switch `IsAuditable` to No for every row belonging to that table.

**Returns the same report it shows in the message box, as text.** Called as a bare statement
(`One_CheckAuditReadiness`) it behaves exactly as before — it pops the `MsgBox` for a person sitting
at the keyboard. Called as `sResult = One_CheckAuditReadiness(True)`, the same text comes back as a
`String` and no dialog is shown — for a script, a test harness, or an AI assistant facilitating a
build to read directly, rather than needing to add its own throwaway diagnostic to see the
verdict. **The `True` is what suppresses the dialog**, not the assignment: VBA cannot tell whether
a procedure was called as a function or as a statement.

```vba
Public Function One_CheckAuditReadiness(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Read-only pre-flight check. Looks at the real table definitions (not the
    '            config table) so a multi-field primary key is never missed, and so this can
    '            run before anything at all is created. Run FIRST, before
    '            Two_CreateAuditTables — required on Path B, where tables were not designed
    '            around this system. On Path A run it after Zero_CreateSampleTables, which
    '            creates the tables it looks at; there it has nothing to find, because those
    '            tables were built to the shape this system needs.
    '            Returns the report as a String. Pass bSilent:=True to suppress the MsgBox so
    '            an automated caller is never left waiting on a dialog nobody can dismiss.
    Dim db As DAO.Database
    Dim tdef As DAO.TableDef
    Dim idx As DAO.Index
    Dim lPkFieldCount As Long
    Dim lPkFieldType As Long
    Dim lProblemCount As Long
    Dim lGuidKeyCount As Long
    Dim lTablesInScope As Long
    Dim sMsg As String
    Dim sReport As String
    Dim sOpen As String

    On Error GoTo errHandler
    Set db = CurrentDb
    lProblemCount = 0
    lGuidKeyCount = 0
    lTablesInScope = 0
    sMsg = ""

    ' [SCAFFOLD] The database has to be closed for the whole run, not only at the start, so
    '            every step that needs it asks again rather than trusting an earlier answer.
    '            Between the numbered steps the developer is expected to go and edit the
    '            config table, and opening something to look at it is the ordinary thing to
    '            do next.
    sOpen = ListOpenObjects()
    If Len(sOpen) > 0 Then
        sReport = "Stopped. Something in this database is still open:" & vbCrLf & vbCrLf & _
            sOpen & vbCrLf & _
            "Close it and run this again. Everything has to stay closed until the whole " & _
            "run is finished, not just when it starts."
        If Not bSilent Then MsgBox sReport, vbExclamation, "Close everything first"
        One_CheckAuditReadiness = sReport
        GoTo Cleanup
    End If

    For Each tdef In db.TableDefs
        ' [BUSINESS LOGIC — scan boundary] The same test Three_PopulateConfigTable uses, from the
        ' one place it lives. The three system tables are excluded here as well: they are
        ' scanned into the config table but never get macros (schema Business Rule 5).
        If IsAuditCandidateTable(tdef) _
            And tdef.Name <> "tblAuditLog" _
            And tdef.Name <> "tblLongTextBackup" _
            And tdef.Name <> "tblAuditLogConfig" Then

            lTablesInScope = lTablesInScope + 1
            lPkFieldCount = 0
            lPkFieldType = -1
            For Each idx In tdef.Indexes
                If idx.Primary Then
                    lPkFieldCount = idx.Fields.Count
                    If lPkFieldCount = 1 Then
                        lPkFieldType = tdef.Fields(idx.Fields(0).Name).Type
                    End If
                End If
            Next idx

            If lPkFieldCount = 0 Then
                lProblemCount = lProblemCount + 1
                sMsg = sMsg & tdef.Name & " — no primary key is set" & vbCrLf
                Debug.Print tdef.Name & ": NOT READY — no primary key"
            ElseIf lPkFieldCount > 1 Then
                lProblemCount = lProblemCount + 1
                sMsg = sMsg & tdef.Name & " — primary key uses more than one field" & vbCrLf
                Debug.Print tdef.Name & ": NOT READY — primary key has " & lPkFieldCount & " fields"
            ElseIf Not IsAuditableKeyType(lPkFieldType) Then
                lProblemCount = lProblemCount + 1
                If lPkFieldType = dbGUID Then
                    ' [BUSINESS LOGIC — schema Business Rule 4] Not a flat exclusion: this one is
                    ' the developer's to decide, and they can only decide it if they are told the
                    ' table exists. Counted separately so the report can raise the question.
                    lGuidKeyCount = lGuidKeyCount + 1
                    sMsg = sMsg & tdef.Name & " — primary key is a Replication ID, and this build " & _
                        "was not set up to audit those" & vbCrLf
                Else
                    sMsg = sMsg & tdef.Name & " — primary key is not an AutoNumber, Long Integer, " & _
                        "Integer or Byte field" & vbCrLf
                End If
                Debug.Print tdef.Name & ": NOT READY — primary key type is " & lPkFieldType
            Else
                If lPkFieldType = dbGUID Then lGuidKeyCount = lGuidKeyCount + 1
                Debug.Print tdef.Name & ": ready"
            End If
        End If
    Next tdef

    ' [SCAFFOLD] Nothing in scope means nothing was checked, and "every table checked is
    '            ready" is true of an empty set and useless to the developer. This is the
    '            worse of the two places to report success on nothing, because it is the
    '            last thing asked before the tables are changed.
    If lTablesInScope = 0 Then
        sReport = "Stopped. No table in this file is in scope, so there was nothing to " & _
            "check and there is nothing to build." & vbCrLf & vbCrLf & _
            "The scope setting at the top of modAddDataMacros does not match the tables in " & _
            "this file - it is set to """ & AUDIT_SCOPE_MODE & """. On the try-it-out path, " & _
            "run Zero_CreateSampleTables first: it creates the tables this looks at. " & _
            "Nothing has been changed."
        If Not bSilent Then MsgBox sReport, vbExclamation, "Nothing is in scope"
        One_CheckAuditReadiness = sReport
        GoTo Cleanup
    End If

    If lProblemCount = 0 Then
        sReport = "Every table checked is ready. Safe to run Two_CreateAuditTables."
    Else
        sReport = lProblemCount & " table(s) are NOT ready yet:" & vbCrLf & vbCrLf & sMsg & vbCrLf & _
            "A table can be audited when it has one primary key field that is an AutoNumber, " & _
            "Long Integer, Integer or Byte — or a Replication ID, where this build was set up " & _
            "to audit those. Either change that table's primary key, or leave it out — set " & _
            "IsAuditable to No for all of that table's rows in tblAuditLogConfig — before you " & _
            "run Four_GenerateAllAuditDataMacros."
    End If

    ' [BUSINESS LOGIC — schema Business Rule 4] The question is raised only where such a key
    ' was actually found, and it has to be raised before Two_CreateAuditTables runs, because
    ' that is what builds the log's key column. Reported whether or not the setting is on:
    ' when it is off, a found key is exactly what the developer needs to be asked about.
    If lGuidKeyCount > 0 Then
        sReport = sReport & vbCrLf & vbCrLf & _
            lGuidKeyCount & " table(s) have a Replication ID as their primary key. " & _
            IIf(AUDIT_GUID_KEYS_AUDITED, _
                "This build is set up to audit them, so the log's key column holds text. " & _
                "Every key in the log is then text, the ordinary numbers included, so " & _
                "sorting the log by key sorts as text rather than by number.", _
                "This build is not set up to audit them, so they are listed above as not " & _
                "auditable and will be left alone. Auditing them is a choice you can make: " & _
                "it builds the log's key column to hold text instead, which lets these " & _
                "tables be audited like any other, at the cost of the log sorting by key " & _
                "as text rather than by number.")
    End If

    If Not bSilent Then MsgBox sReport, IIf(lProblemCount = 0, vbInformation, vbExclamation)
    One_CheckAuditReadiness = sReport

Cleanup:
    Set idx = Nothing
    Set tdef = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    One_CheckAuditReadiness = "Error checking audit readiness: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox One_CheckAuditReadiness, vbCritical
    Resume Cleanup
    Resume
End Function
```

### IsAuditCandidateTable — `Private Function` → `Boolean`

**The one place that decides whether a table is in scope at all.** `Three_PopulateConfigTable` and
`One_CheckAuditReadiness` both call it, so the scan and the safety check cannot disagree about which
tables they are talking about — a disagreement that shows up as a table reported ready and then never
scanned, or the reverse.

**It carries out Step 4's answer; it does not make it.** `AUDIT_SCOPE_MODE` holds that answer, and
this function reads it. Everything finer-grained is a flag in the config table rather than code.

**Four exclusions are tested before the mode is read, so no answer can let one through.** Access's
own system tables and temporary tables are never audited; the developer's own hidden tables are
theirs, and are in scope only where they named one themselves; and a linked table cannot carry a
Data Macro at all, so including one would fail at the last step rather than the first. The linked
test is first of the four because it is the one a developer can ask for by name.

**It takes the table itself, not the table's name.** The linked test needs the table definition, and
a function that receives the whole thing cannot be called in a way that skips the check.

```vba
Private Function IsAuditCandidateTable(tdef As DAO.TableDef) As Boolean
    ' [BUSINESS LOGIC — scan boundary] Reads AUDIT_SCOPE_MODE, set in the module-level
    '            declarations from Step 4's answer. The four exclusions below hold under
    '            every mode and are tested first, so no setting and no table list can let
    '            one of them in.
    Dim sName As String

    sName = tdef.Name
    IsAuditCandidateTable = False

    ' A linked table lives in another file. Access cannot attach a Data Macro to one, so it
    ' is never a candidate — whatever the mode says, and even when named in the list.
    If (tdef.Attributes And (dbAttachedTable Or dbAttachedODBC)) <> 0 Then Exit Function

    ' Access's own system tables are never touched.
    If Left$(sName, 4) = "MSys" Then Exit Function

    ' A USys table is the DEVELOPER's own hidden table, theirs to manage. In scope only where
    ' they opted it in themselves by naming it.
    If Left$(sName, 4) = "USys" Then
        IsAuditCandidateTable = IsNamedInScopeList(sName)
        Exit Function
    End If

    ' Temporary and working tables.
    If Left$(sName, 3) = "tmp" Then Exit Function

    Select Case AUDIT_SCOPE_MODE
        Case "All"
            IsAuditCandidateTable = True
        Case "List"
            IsAuditCandidateTable = IsNamedInScopeList(sName)
        Case "Standard"
            IsAuditCandidateTable = (Left$(sName, 3) = "tbl" Or Left$(sName, 4) = "tlkp")
        Case Else
            ' [SCAFFOLD] A mode nobody defined is a typo in the constant, and the safe
            '            reading of a typo is not "audit nothing" — that is the failure
            '            this whole setting exists to prevent. Stop on the first table
            '            instead, where both callers' error blocks will report it.
            Err.Raise vbObjectError + 513, "IsAuditCandidateTable", _
                "AUDIT_SCOPE_MODE is set to """ & AUDIT_SCOPE_MODE & """. It has to be " & _
                """Standard"", ""All"" or ""List""."
    End Select
End Function
```

### IsNamedInScopeList — `Private Function` → `Boolean`

**Is this table one the developer named?** Reads `AUDIT_SCOPE_LIST`, which is empty unless Step 4
was answered with a list of table names. Two callers, both in `IsAuditCandidateTable`: the `List`
mode itself, and the `USys` opt-in, which works the same way under every mode.

Names are separated by semicolons and compared without regard to capitals, as Access compares table
names. Each name is trimmed, so a list typed with spaces after the semicolons still matches, and a
table name that itself contains spaces is unaffected.

```vba
Private Function IsNamedInScopeList(sTableName As String) As Boolean
    ' [SCAFFOLD] Exact match against one entry, never a partial one: "tblOrder" must not
    '            match "tblOrderLine".
    Dim vNames As Variant
    Dim i As Long

    IsNamedInScopeList = False
    If Len(AUDIT_SCOPE_LIST) = 0 Then Exit Function

    vNames = Split(AUDIT_SCOPE_LIST, ";")
    For i = LBound(vNames) To UBound(vNames)
        If StrComp(Trim$(vNames(i)), sTableName, vbTextCompare) = 0 Then
            IsNamedInScopeList = True
            Exit Function
        End If
    Next i
End Function
```

### IsAuditableKeyType — `Private Function` → `Boolean`

**The one place that decides whether a table's primary key can be stored in the log.** The log
keeps each audited row's key, so the key has to be a type that column can hold (schema Business
Rule 4).

**Three types are always accepted and one is the developer's choice.** AutoNumber, Long Integer,
Integer and Byte all fit the whole-number key column and are never in question. A **Replication
ID** fits only where the developer chose to audit such tables, which is what
`AUDIT_GUID_KEYS_AUDITED` records — that answer is also what decides whether the key column was
built to hold text in the first place, so the two cannot disagree. Everything else — text keys,
keys made of more than one field, and numbers too large for a Long Integer — is out.

```vba
Private Function IsAuditableKeyType(ByVal lFieldType As Long) As Boolean
    ' [BUSINESS LOGIC — schema Business Rule 4] The key types the log's key column can hold.
    '                 dbGUID is the only one the developer decides; the rest are fixed.
    Select Case lFieldType
        Case dbLong, dbInteger, dbByte
            IsAuditableKeyType = True
        Case dbGUID
            IsAuditableKeyType = AUDIT_GUID_KEYS_AUDITED
        Case Else
            IsAuditableKeyType = False
    End Select
End Function
```

### IsUnauditableFieldType — `Private Function` → `Boolean`

**The one place that decides whether a field can be audited at all.** `Three_PopulateConfigTable`
calls it to seed the switch off, and `Four_GenerateAllAuditDataMacros` calls it again to enforce
that whatever the switch now says — so the scan and the generator cannot disagree.

**Two types, for two different reasons.** An **Attachment** field cannot be read or written by a
Data Macro; a macro that references one does not work, so switching it on breaks the table's whole
macro set rather than adding a field to the log. A **calculated** field cannot be edited by anybody:
its value is derived from other fields in the same row, and those fields are audited themselves, so
a log row for it would record a change nobody made and attribute it to whoever changed something
else.

**Why the second test is a guarded property read.** A calculated field carries an `Expression`
property holding the formula; an ordinary field answers the same lookup with an empty one, so the
test is what the lookup gives back and never whether it can be made at all. Reading a property that
isn't there raises an error rather than returning empty, so the read is wrapped, and anything that
goes wrong leaves the answer `False` — a field this function
cannot classify is treated as ordinary, which is the behaviour every build before this one had.

```vba
Private Function IsUnauditableFieldType(fld As DAO.Field) As Boolean
    ' [BUSINESS LOGIC] Fields that are never audited, whatever tblAuditLogConfig says:
    '            Attachment — a Data Macro cannot read or write the type at all, so a macro
    '                         referencing one does not work.
    '            Calculated — nobody edits it; its value comes from other fields in the same
    '                         row, which are audited themselves, so a log row here would
    '                         record a change nobody made.
    '            No errHandler and therefore no line numbers: any error propagates to the
    '            caller, which is what should happen to a schema read this small.
    Dim sExpression As String

    If fld.Type = dbAttachment Then
        IsUnauditableFieldType = True
        Exit Function
    End If

    ' A calculated field carries a non-empty Expression property. Reading a property that is
    ' not there raises, so the probe is guarded and an unreadable field counts as ordinary.
    sExpression = ""
    On Error Resume Next
    sExpression = fld.Properties("Expression").Value & ""
    On Error GoTo 0

    IsUnauditableFieldType = (Len(sExpression) > 0)
End Function
```

### ListOpenObjects — `Public Function` → `String`

**Names everything open in this database, and returns an empty string when nothing is.** The
generator attaches macros by opening each table in design view, and Access refuses that while
anything is using the table, so the run has to start and stay in a closed database. This turns the
prerequisite into something the code can check rather than something it has to assume.

**A bound form is the usual cause, and it never names the table it uses.** That does not matter
here: the form is open, so the form is what gets reported, and the developer is told which object to
close rather than being asked to work out which forms reach which tables.

**It only sees this database.** A second copy of it, or a front end holding links from its own
session, is invisible from here and stays the developer's to close.

```vba
Public Function ListOpenObjects() As String
    ' [SCAFFOLD] Recognition only. Returns "" when nothing in this database is open, and
    '            otherwise one indented line per open object for a caller to print.
    '            Access's own tables are skipped: AllTables lists them (11 in an empty
    '            database), they are never this system's business, and an open one is not
    '            something a developer would be asked to close.
    '            No errHandler and therefore no line numbers: an object that cannot be asked
    '            is reported as not open, and the caller stops on whatever else it finds.
    Dim obj As Object
    Dim sList As String

    sList = ""

    On Error Resume Next

    For Each obj In CurrentProject.AllForms
        If obj.IsLoaded Then sList = sList & "  Form: " & obj.Name & vbCrLf
    Next obj

    For Each obj In CurrentProject.AllReports
        If obj.IsLoaded Then sList = sList & "  Report: " & obj.Name & vbCrLf
    Next obj

    For Each obj In CurrentData.AllTables
        If Left$(obj.Name, 4) <> "MSys" Then
            If SysCmd(acSysCmdGetObjectState, acTable, obj.Name) <> 0 Then
                sList = sList & "  Table: " & obj.Name & vbCrLf
            End If
        End If
    Next obj

    For Each obj In CurrentData.AllQueries
        If SysCmd(acSysCmdGetObjectState, acQuery, obj.Name) <> 0 Then
            sList = sList & "  Query: " & obj.Name & vbCrLf
        End If
    Next obj

    On Error GoTo 0

    ListOpenObjects = sList
End Function
```
