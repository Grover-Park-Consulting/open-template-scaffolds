---
template: audit-logging-lite-scaffold
title: Access Audit Logging (Lite) — Data Macro Generator VBA Scaffold
domain: audit
type: vba-scaffold
version: 0.1.0
status: draft
implements: audit-logging-lite-schema
requires_tables:
  - tblAuditLog
  - tblLongTextBackup
  - tblAuditLogConfig
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
  - design-principles
target_module: modAddDataMacros
new_procedures:
  - Zero_CreateSampleTables
  - One_CreateAuditTables
  - Two_PopulateConfigTable
  - CheckAuditReadiness
  - Three_GenerateAllAuditDataMacros
  - CreateAllDataMacros
  - BuildAfterInsertMacro
  - BuildAfterUpdateMacro
  - BuildAfterDeleteMacro
  - BuildBeforeChangeMacro
  - BuildBeforeDeleteMacro
  - GetComparisonExpression
  - BackupLongTextFieldsDM
  - BackupAndRemoveAllDataMacros
build_paths:
  - "Path A — Demo build: try the system out on three made-up tables the generator creates for
    you (Zero_CreateSampleTables). Nothing real is touched."
  - "Path B — Add it to a real database you already have: skip Zero_CreateSampleTables and point
    the generator at your own existing tables instead. Back up the file first (see warnings)."
warnings:
  - Data Macros cannot audit Long Text (Memo) fields on their own. Before building, list every
    Long Text field in the tables to be audited and confirm the list with the developer — any
    table carrying one takes the hybrid VBA path (BeforeChange/BeforeDelete backing values up
    through BackupLongTextFieldsDM); a table without one needs only the three After macros.
  - This module must run in the same accdb as the audited tables — the back end of a split
    design. BackupLongTextFieldsDM must additionally exist in every front end, because a data
    macro fired by a front-end edit resolves the function there.
  - DAO cannot create Data Macros. The only build path is writing UTF-16 XML to a file and
    loading it with Application.LoadFromText acTableDataMacro — exactly what this module does.
  - Every audited table is expected to have a single-column AutoNumber primary key. If any table
    to be audited has a different key design (composite, text, no PK), stop and tell the
    developer this template will not work for that table out of the box — they are free to adapt
    it, but the adaptation is theirs. CheckAuditReadiness checks for this automatically.
  - Path B (an existing accdb with real tables and real data) is much less forgiving than the
    demo. Make a copy of the .accdb file before running any of these steps against it — Data
    Macros get attached directly to your live tables, and this is not a step to redo casually.
  - If a table already carries its own Data Macros before this generator runs on it (for example,
    the standards/audit-columns.md Before Change stamping macro), they are silently replaced —
    Application.LoadFromText replaces a table's entire macro set, it does not merge. The generator
    backs up any table's existing macros automatically before replacing them, but nothing restores
    that stamping logic afterward — that is the developer's call.
  - If any importer un-escapes HTML/XML entities in VBA source on the way in (this repo's own
    Access Explorer MCP code-import tools do), a literal &lt;&gt; in GetComparisonExpression
    becomes a raw <> and breaks the generated macro XML (error 3870). The function builds the
    entity from Chr(38) at runtime for exactly this reason — don't revert it to a literal.
---

# Access Audit Logging (Lite) — Data Macro Generator VBA Scaffold

## Intent

The working half of the Lite audit system: the VBA that **creates the three system tables**,
**scans the schema into the config table**, and **generates + attaches the Data Macros** the
paired table template (`audit-logging-lite-schema`) describes. Unlike most scaffolds in this
library, the procedures here are **complete, working code**, not skeletons — the same design
(tables, macro set, Long Text hybrid path) stands behind a live production Access application
whose audit trail validates it end to end. Audit **scope** is decided in data: the scan writes
every candidate field to the config table and you flip `IsAuditable` flags — the `[BUSINESS
LOGIC]` markers land on that review step and on the one code filter (which table prefix to
scan); `[STANDARDS]` markers cover the usual deferred house style.

### Two ways to use this — pick one before you start

**Path A — Try it out first.** This creates three made-up tables (a client list, a support
ticket list, and a short pick-list of ticket priorities) so you can watch the audit trail work
before you touch anything real. Nothing in your own database is affected. Good for a first
look, a demo, or learning what this system does.

**Path B — Add this to a database you already use.** This turns on auditing for your own,
already-existing tables. It does **not** create the three made-up tables — it works directly
against the tables you already have. Because it changes a real database, **make a backup copy
of the .accdb file before you start**, the same way you'd back up before any change you can't
easily undo.

Both paths use the same three numbered steps below — only the middle step is set up slightly
differently, and Path B adds one extra safety check.

> **If an AI assistant is running these steps for someone:** the developer picks the path — don't
> infer it from what the database looks like, even if the answer seems obvious (e.g. "it already
> has real data, so it must be Path B"). Ask, then wait for their answer. Likewise, don't work out
> `CheckAuditReadiness`'s answer yourself by reading the tables directly — run the actual
> procedure at the point these steps call for it, show the developer what it says, and stop there.
> Each numbered step (and the review pause after it) is a separate decision point: present one,
> get the developer's go-ahead for that step specifically, then move to the next. Don't collapse
> the whole sequence into a single upfront report — that skips the review this template is built
> around, even if every fact in the report turns out correct.

```vba
' ---------- Path A — try it out first (nothing real is touched) ----------
Zero_CreateSampleTables          ' 0. create the two made-up tables and a short pick-list
One_CreateAuditTables            ' 1. create the 3 tables the audit trail itself lives in
Two_PopulateConfigTable          ' 2. make a list of every field in every table that could be
                                  '    audited, switched ON to start
'    ... open the list (tblAuditLogConfig) and switch OFF anything you don't want tracked ...
Three_GenerateAllAuditDataMacros ' 3. turn on tracking for everything still switched ON

' ---------- Path B — add this to a database you already use ----------
' >>> back up the .accdb file first — this step changes real, live tables <<<
One_CreateAuditTables            ' 1. create the 3 tables the audit trail itself lives in
Two_PopulateConfigTable False    ' 2. make a list of every field in every table that could be
                                  '    audited, switched OFF to start
CheckAuditReadiness              '    a safety check: tells you if any of your tables can't be
                                  '    tracked as-is (see Business Rule 4 below), before anything
                                  '    is changed
'    ... open the list (tblAuditLogConfig) and switch ON tracking, table by table, for whatever
'        you actually want a history of ...
Three_GenerateAllAuditDataMacros ' 3. turn on tracking for everything switched ON
```

`CheckAuditReadiness` and `Three_GenerateAllAuditDataMacros` are called above as bare statements
(the ordinary, interactive way — each pops its own `MsgBox`). Both also take an optional
`bSilent` argument: call them as `sResult = CheckAuditReadiness(True)` /
`sResult = Three_GenerateAllAuditDataMacros(True)` to read the same report back as a `String`
with no dialog at all — the way a script, test harness, or an AI assistant facilitating the build
should read the result, rather than adding a throwaway diagnostic just to see what happened.
**Passing `True` is what suppresses the dialog**, not assigning the return value: VBA gives a
procedure no way to detect whether it was called as a function or as a statement.

Then link the three system tables into the front end and import the Long Text helper module
there (see `BackupLongTextFieldsDM`).

**Module homes** (three modules, one job each):

| Module | Procedures | Lives in |
|---|---|---|
| `modAddDataMacros` | `Zero_CreateSampleTables`, the numbered steps, `CheckAuditReadiness`, the five `Build*` XML builders + `GetComparisonExpression` | Back end only |
| `modAuditLongText` | `BackupLongTextFieldsDM` | **Back end AND every front end** |
| `modAuditAdmin` | `BackupAndRemoveAllDataMacros` | Back end only |

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — the working structure provided here.
- **`[STANDARDS]`** — house style, deferred to the standards layer (`error-handling.md`,
  `query-style.md`, `naming-conventions.md`). The error blocks below use the dependency-free
  `MsgBox` default; substitute your house logger per `error-handling.md`.
- **`[BUSINESS LOGIC]`** — the audit-scope decisions you must make: the scan-boundary prefix
  test in `Two_PopulateConfigTable`, and the `IsAuditable` flag review in `tblAuditLogConfig`
  after the scan.

## Prerequisites

| Object | Role |
|---|---|
| `audit-logging-lite-schema` system tables | `tblAuditLog` / `tblLongTextBackup` / `tblAuditLogConfig` — created by `One_CreateAuditTables`, described in the paired template |
| The audited tables | Each with a single-column numeric PK (schema Business Rule 4) |
| A Trusted Location | The generator and the macros' VBA calls run only with code enabled |
| `Microsoft Scripting Runtime` (late-bound) | `FileSystemObject` writes the UTF-16 macro XML; `CreateObject` is used, no reference needed |

## Procedures

### Zero_CreateSampleTables — `Public Sub` (Path A only — setup step 0)

**Skip this procedure entirely if you're doing Path B** (adding audit tracking to a database
you already use) — it only exists to build the made-up tables for trying the system out.

Creates the two made-up tables (`tblClient`, `tblSupportTicket`) and a short pick-list
(`tlkpTicketPriority`) described in the paired schema template, with the four starter pick-list
rows (Low, Normal, High, Urgent) and the two links between the tables. Same idempotent style as
`One_CreateAuditTables` — an existing table is reported and skipped, so it's safe to re-run.

```vba
Public Sub Zero_CreateSampleTables()
    ' [SCAFFOLD] Creates tblClient, tlkpTicketPriority, tblSupportTicket (schema template
    '            entities) and seeds tlkpTicketPriority. Path A (try-it-out build) only —
    '            skip this Sub for Path B (an existing accdb's own tables). Idempotent: each
    '            block is skipped if its table already exists.
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim rel As DAO.Relation

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

    ' Starter pick-list rows (schema: Low/Normal/High/Urgent)
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder) VALUES ('Low', 10)", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder) VALUES ('Normal', 20)", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder) VALUES ('High', 30)", dbFailOnError
    db.Execute "INSERT INTO tlkpTicketPriority (TicketPriorityName, SortOrder) VALUES ('Urgent', 40)", dbFailOnError

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

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set rel = Nothing
    Set tdf = Nothing
    Set db = Nothing
    MsgBox "Sample tables created. You're on Path A (try-it-out build) — nothing in your " & _
        "own database was touched.", vbInformation
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    MsgBox "Error creating sample tables: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub
```

### One_CreateAuditTables — `Public Sub` (setup step 1)

Creates the three system tables via DAO, idempotently — an existing table is reported and
skipped, so the Sub is safe to re-run. Field-by-field DAO `CreateField` (never `CREATE TABLE`
DDL — see `templates/_materialization.md`).

```vba
Public Sub One_CreateAuditTables()
    ' [SCAFFOLD] Creates tblAuditLog, tblLongTextBackup, tblAuditLogConfig (schema template
    '            entities). Idempotent: each block is skipped if its table already exists.
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    On Error GoTo errHandler
    Set db = CurrentDb

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

    Set fld = tdf.CreateField("PrimaryKey", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OperationType", dbText, 25)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OldValue", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("NewValue", dbMemo)
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

    Set fld = tdf.CreateField("BackupID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("PrimaryKey", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OldValue", dbMemo)
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
    Set fld = idx.CreateField("BackupID")
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

    Set fld = tdf.CreateField("ConfigID", dbLong)
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
    Set fld = idx.CreateField("ConfigID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblAuditLogConfig created"

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing
    MsgBox "Audit tables created successfully!", vbInformation
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] dependency-free default; substitute your house logger.
    MsgBox "Error creating tables: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub
```

### Two_PopulateConfigTable — `Public Sub` (setup step 2)

Scans the schema into `tblAuditLogConfig`: **every field of every candidate table**, with its
ordinal position, DAO type code, a flag on the table's PK field, and `IsAuditable`. Nothing is
silently dropped — exclusions are *seeded* as `IsAuditable = False` rows: the three system
tables, plus fields that would just add noise — this repo's house audit columns
(`CreatedDate`/`CreatedBy`/`ModifiedDate`/`ModifiedBy`/`AccessTS`, per `standards/audit-columns.md`)
and a few other always-changing system columns (`SSMA_TimeStamp`, `ValidFrom`, `ValidTo`). **After
running, open the config table and review the flags** — that review, in data, is where the audit
net is drawn (schema Business Rule 5).

Takes one optional Yes/No setting that decides the starting point for everything else:

- **Path A (try-it-out build):** run `Two_PopulateConfigTable` with nothing after it. Every
  field starts switched ON, and you switch OFF the few you don't want tracked.
- **Path B (a database you already use):** run `Two_PopulateConfigTable False`. Every field
  starts switched OFF, and you switch ON — table by table — only what you actually want a
  history of. This is the safer starting point on tables this system wasn't designed around,
  where "track everything" could sweep in more than you meant.

```vba
Public Sub Two_PopulateConfigTable(Optional bDefaultAuditable As Boolean = True)
    ' [SCAFFOLD] Rebuild the audit configuration from the live schema. Scope decisions
    '            live in the IsAuditable flags afterward, not in this code.
    '            bDefaultAuditable sets the starting point for ordinary fields only:
    '            True  (default; Path A) — everything starts switched ON, you switch OFF
    '                  what you don't want tracked.
    '            False (Path B — call as Two_PopulateConfigTable(False)) — everything
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

    On Error GoTo errHandler
    Set db = CurrentDb

    ' Clear existing config
    db.Execute "DELETE * FROM tblAuditLogConfig", dbFailOnError

    For Each tdef In db.TableDefs
        ' [BUSINESS LOGIC — scan boundary] Which tables are candidates at all. This default
        ' scans tables prefixed tbl or tlkp (your data and lookup tables under
        ' naming-conventions.md) but never tmp (temporary/working tables); system (MSys)
        ' tables never match either prefix. On another naming policy, change this one test —
        ' every finer-grained decision is a flag in the config table, not code. Keep
        ' CheckAuditReadiness's copy of this test in sync if you change it here.
        ' >>> adjust the prefix test to your naming convention <<<
        If (Left(tdef.Name, 3) = "tbl" Or Left(tdef.Name, 4) = "tlkp") _
            And Left(tdef.Name, 3) <> "tmp" Then

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
                ' [STANDARDS — audit-columns.md] CreatedDate/CreatedBy/ModifiedDate/ModifiedBy/
                '            AccessTS are this repo's house audit columns as of the standards
                '            layer in use when this list was written. If your standards/
                '            audit-columns.md names different columns, update this list to
                '            match — it is a VBA-side mirror of that file, not a live read of it.
                '            SSMA_TimeStamp/ValidFrom/ValidTo are not house audit columns; they
                '            are left here because a table carrying them already has its own
                '            change-tracking mechanism (e.g. SQL Server temporal system-versioning)
                '            that this scan would otherwise log as noisy, always-changing values.
                Select Case True
                    Case tdef.Name = "tblAuditLog", _
                         tdef.Name = "tblLongTextBackup", _
                         tdef.Name = "tblAuditLogConfig"
                        isAuditable = False
                    Case fld.Name = "CreatedDate", fld.Name = "CreatedBy", _
                         fld.Name = "ModifiedDate", fld.Name = "ModifiedBy", _
                         fld.Name = "AccessTS", fld.Name = "SSMA_TimeStamp", _
                         fld.Name = "ValidFrom", fld.Name = "ValidTo"
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
            Next fld
        End If
    Next tdef

    MsgBox "Table list built. Open tblAuditLogConfig and check the IsAuditable switches " & _
        "before you run the next step.", vbInformation

Cleanup:
    Set pkField = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set tdef = Nothing
    Set db = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    MsgBox "Error populating config: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub
```

### CheckAuditReadiness — `Public Function` → `String` (safety check — run before step 3, required for Path B)

A read-only check you can run any time, at no risk — it doesn't change anything. It looks at each
table you might track and tells you whether this system will actually work on it: **every table
needs one, single number field as its primary key, set to auto-number** (schema Business Rule 4).
Most tables you design yourself already look like this. Older or borrowed tables sometimes
don't — a table with no primary key set, one that uses two or more fields together as its key,
or one that uses a text code instead of a number, will not work with this system as-is.

Run this after `Two_PopulateConfigTable` and before `Three_GenerateAllAuditDataMacros`. This
step is **required for Path B**, since a database you didn't design the audit system around is
far more likely to have a table shaped this way. It's optional on Path A, where the sample
tables are already known to be shaped correctly.

If a table isn't ready, you have two choices: fix that table's primary key, or leave it out —
open `tblAuditLogConfig` and switch `IsAuditable` to No for every row belonging to that table.

**Returns the same report it shows in the message box, as text.** Called as a bare statement
(`CheckAuditReadiness`) it behaves exactly as before — it pops the `MsgBox` for a person sitting
at the keyboard. Called as `sResult = CheckAuditReadiness(True)`, the same text comes back as a
`String` and no dialog is shown — for a script, a test harness, or an AI assistant facilitating a
build to read directly, rather than needing to add its own throwaway diagnostic to see the
verdict. **The `True` is what suppresses the dialog**, not the assignment: VBA cannot tell whether
a procedure was called as a function or as a statement.

```vba
Public Function CheckAuditReadiness(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Read-only pre-flight check. Looks at the real table definitions (not the
    '            config table) so a multi-field primary key is never missed. Run after
    '            Two_PopulateConfigTable and before Three_GenerateAllAuditDataMacros —
    '            required on Path B, where tables were not designed around this system.
    '            Returns the report as a String. Pass bSilent:=True to suppress the MsgBox so
    '            an automated caller is never left waiting on a dialog nobody can dismiss.
    Dim db As DAO.Database
    Dim tdef As DAO.TableDef
    Dim idx As DAO.Index
    Dim lPkFieldCount As Long
    Dim lPkFieldType As Long
    Dim lProblemCount As Long
    Dim sMsg As String
    Dim sReport As String

    On Error GoTo errHandler
    Set db = CurrentDb
    lProblemCount = 0
    sMsg = ""

    For Each tdef In db.TableDefs
        ' [BUSINESS LOGIC — scan boundary] Same tbl-or-tlkp-but-not-tmp test as
        ' Two_PopulateConfigTable; keep the two in sync if you change one.
        If (Left(tdef.Name, 3) = "tbl" Or Left(tdef.Name, 4) = "tlkp") _
            And Left(tdef.Name, 3) <> "tmp" _
            And tdef.Name <> "tblAuditLog" _
            And tdef.Name <> "tblLongTextBackup" _
            And tdef.Name <> "tblAuditLogConfig" Then

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
            ElseIf lPkFieldType <> dbLong Then
                lProblemCount = lProblemCount + 1
                sMsg = sMsg & tdef.Name & " — primary key is not an AutoNumber/Long Number field" & vbCrLf
                Debug.Print tdef.Name & ": NOT READY — primary key type is " & lPkFieldType
            Else
                Debug.Print tdef.Name & ": ready"
            End If
        End If
    Next tdef

    If lProblemCount = 0 Then
        sReport = "Every table checked is ready — each has one auto-number primary key. " & _
            "Safe to run Three_GenerateAllAuditDataMacros."
    Else
        sReport = lProblemCount & " table(s) are NOT ready yet:" & vbCrLf & vbCrLf & sMsg & vbCrLf & _
            "This system only works on tables with one auto-number (or plain number) " & _
            "primary key field. Either fix that table's primary key, or leave it out — set " & _
            "IsAuditable to No for all of that table's rows in tblAuditLogConfig — before you " & _
            "run Three_GenerateAllAuditDataMacros."
    End If

    If Not bSilent Then MsgBox sReport, IIf(lProblemCount = 0, vbInformation, vbExclamation)
    CheckAuditReadiness = sReport

Cleanup:
    Set idx = Nothing
    Set tdef = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    CheckAuditReadiness = "Error checking audit readiness: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox CheckAuditReadiness, vbCritical
    Resume Cleanup
End Function
```

### Three_GenerateAllAuditDataMacros — `Public Function` → `String` (setup step 3)

Reads the (reviewed) config, groups fields by table, and calls `CreateAllDataMacros` for each.
Re-runnable: reloading a table's macro XML replaces what was there (schema Business Rule 7).

**Returns a per-table report as text**, in addition to the summary `MsgBox` — one line per table
(`OK`, `SKIPPED`, or `ERROR: ...`), the same detail `CreateAllDataMacros` used to only send to
`Debug.Print`. Call it as `sResult = Three_GenerateAllAuditDataMacros(True)` to read that report
directly with no dialog — a script or an AI assistant facilitating the build no longer has to add
its own diagnostic wrapper to see which tables actually succeeded. `bSilent` is passed down into
`CreateAllDataMacros`, so a per-table error can't strand an automated caller either.

```vba
Public Function Three_GenerateAllAuditDataMacros(Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Generate and attach audit Data Macros for every configured table.
    '            Returns a per-table report so a caller — human or automated — can see exactly
    '            what happened to each table. Pass bSilent:=True to suppress every MsgBox,
    '            including the per-table ones raised inside CreateAllDataMacros.
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dictTables As Object          ' Scripting.Dictionary, late-bound
    Dim sTableName As String
    Dim sFieldName As String
    Dim lFieldDataType As Long
    Dim bFieldIsPK As Boolean
    Dim bFieldIsAuditable As Boolean
    Dim fieldList As Collection
    Dim fieldInfo As Variant
    Dim sTempPath As String
    Dim lTableCount As Long
    Dim vCurrentTable As Variant
    Dim sTableResult As String
    Dim sReport As String

    On Error GoTo errHandler
    Set db = CurrentDb
    Set dictTables = CreateObject("Scripting.Dictionary")
    sReport = ""

    ' Read configuration and group fields by table, in field order. All rows come along —
    ' the PK row is needed for plumbing even if its IsAuditable flag was flipped; the
    ' builders skip non-auditable fields when emitting audit actions.
    Set rs = db.OpenRecordset( _
        "SELECT TableName, FieldName, DataType, IsPrimaryKey, IsAuditable " & _
        "FROM tblAuditLogConfig ORDER BY TableName, FieldPosition", dbOpenSnapshot)

    Do While Not rs.EOF
        sTableName = Nz(rs!TableName, "")
        sFieldName = Nz(rs!FieldName, "")
        lFieldDataType = Nz(rs!DataType, 0)
        bFieldIsPK = Nz(rs!IsPrimaryKey, False)
        bFieldIsAuditable = Nz(rs!IsAuditable, False)

        ' [SCAFFOLD] Hard guard above the flags: the system tables never get macros
        '            (schema Business Rule 5), whatever their config rows say.
        If sTableName <> "tblAuditLog" _
            And sTableName <> "tblLongTextBackup" _
            And sTableName <> "tblAuditLogConfig" _
            And sTableName <> "" And sFieldName <> "" Then

            If Not dictTables.Exists(sTableName) Then
                Set fieldList = New Collection
                dictTables.Add sTableName, fieldList
            Else
                Set fieldList = dictTables(sTableName)
            End If
            ' Field info as array: (FieldName, DataType, IsPrimaryKey, IsAuditable)
            fieldInfo = Array(sFieldName, lFieldDataType, bFieldIsPK, bFieldIsAuditable)
            fieldList.Add fieldInfo
        End If
        rs.MoveNext
    Loop
    rs.Close

    sTempPath = Environ("TEMP") & "\"

    lTableCount = 0
    For Each vCurrentTable In dictTables.Keys
        sTableName = CStr(vCurrentTable)
        Set fieldList = dictTables(sTableName)
        sTableResult = CreateAllDataMacros(sTableName, fieldList, sTempPath, bSilent)
        sReport = sReport & sTableName & ": " & sTableResult & vbCrLf
        lTableCount = lTableCount + 1
    Next vCurrentTable

    If Not bSilent Then MsgBox "Generated audit data macros for " & lTableCount & " table(s). " & _
        "See the returned report for the per-table result.", vbInformation
    Three_GenerateAllAuditDataMacros = sReport

Cleanup:
    Set rs = Nothing
    Set db = Nothing
    Set dictTables = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Three_GenerateAllAuditDataMacros = sReport & "ERROR: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Function
```

### CreateAllDataMacros — `Private Function` → `String`

Builds one table's macro XML — the three After macros always, plus BeforeChange/BeforeDelete when
the field list contains a Long Text field (schema Business Rule 2) — writes it UTF-16, and loads
it with `LoadFromText acTableDataMacro`.

**Backs up a table's existing macros before replacing them.** `LoadFromText` replaces a table's
entire Data Macro set — it does not merge. On Path B a table may already carry its own macros
(most notably the `standards/audit-columns.md` Before Change stamping macro for `CreatedBy`/
`ModifiedBy`), and overwriting that silently would be a real, undocumented loss. Before loading
the new macros, this checks `MSysObjects` for an existing macro set on the table and, if one is
found, exports it to a timestamped backup file first — the same technique
`BackupAndRemoveAllDataMacros` already uses to detect a table's macros. This does **not** merge
the old logic into the new macros; it only makes sure nothing is destroyed without a copy and a
plain-language warning first. Restoring or re-implementing any lost stamping logic afterward is
the developer's call.

```vba
Private Function CreateAllDataMacros(sTableName As String, fieldList As Collection, sTempPath As String, Optional bSilent As Boolean = False) As String
    ' [SCAFFOLD] Generate the 3 (or 5, with Long Text) Data Macros for one table. Returns a
    '            one-line status ("OK ...", "SKIPPED ...", or "ERROR: ...") so the caller can
    '            report per-table results without relying on Debug.Print alone.
    Dim db As DAO.Database
    Dim rsCheck As DAO.Recordset
    Dim sXmlContent As String
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim txtFile As Object
    Dim sFilePath As String
    Dim sPrimaryKeyField As String
    Dim fieldInfo As Variant
    Dim bHasLongText As Boolean
    Dim lAuditableCount As Long
    Dim bHadExistingMacros As Boolean
    Dim sBackupFolder As String
    Dim sBackupFile As String
    Dim sBackupNote As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' Find the PK field, count auditable fields, and detect auditable Long Text
    ' (schema Business Rules 2, 4, 5)
    bHasLongText = False
    lAuditableCount = 0
    For Each fieldInfo In fieldList
        If fieldInfo(2) = True Then sPrimaryKeyField = fieldInfo(0)
        If fieldInfo(3) = True Then
            lAuditableCount = lAuditableCount + 1
            If fieldInfo(1) = dbMemo Then bHasLongText = True
        End If
    Next fieldInfo

    ' A table with nothing auditable gets no macros at all
    If lAuditableCount = 0 Then
        Debug.Print "  - Skipped (no auditable fields)"
        CreateAllDataMacros = "SKIPPED (no auditable fields)"
        Exit Function
    End If

    ' [SCAFFOLD] Safety net for Path B: LoadFromText replaces a table's WHOLE macro set. If
    '            this table already has one (e.g. the house audit-column stamping macro),
    '            back it up before it's overwritten — see the note above this code block.
    bHadExistingMacros = False
    Set rsCheck = db.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Name='" & sTableName & "' AND Type=1 AND Not IsNull(LvExtra)", _
        dbOpenSnapshot)
    bHadExistingMacros = Not rsCheck.EOF
    rsCheck.Close
    Set rsCheck = Nothing

    sBackupNote = ""
    If bHadExistingMacros Then
        sBackupFolder = CurrentProject.Path & "\DataMacroBackups\"
        If Dir(sBackupFolder, vbDirectory) = "" Then MkDir sBackupFolder
        sBackupFile = sBackupFolder & sTableName & "_PreAuditBackup_" & Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Application.SaveAsText acTableDataMacro, sTableName, sBackupFile
        sBackupNote = " (existing macros backed up to " & sBackupFile & " before replacing)"
        Debug.Print "  - " & sTableName & " already had Data Macros — backed up to " & sBackupFile
    End If

    ' One XML document carrying all of this table's macros
    sXmlContent = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    sXmlContent = sXmlContent & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application"">"

    sXmlContent = sXmlContent & BuildAfterInsertMacro(sTableName, fieldList, sPrimaryKeyField)
    sXmlContent = sXmlContent & BuildAfterUpdateMacro(sTableName, fieldList, sPrimaryKeyField)
    sXmlContent = sXmlContent & BuildAfterDeleteMacro(sTableName, fieldList, sPrimaryKeyField)

    If bHasLongText Then
        sXmlContent = sXmlContent & BuildBeforeChangeMacro(sTableName, fieldList, sPrimaryKeyField)
        sXmlContent = sXmlContent & BuildBeforeDeleteMacro(sTableName, fieldList, sPrimaryKeyField)
    End If

    sXmlContent = sXmlContent & "</DataMacros>"

    ' [SCAFFOLD] Write UTF-16 (CreateTextFile third argument True) — LoadFromText requires it —
    '            then load with the table held open in design view so the save sticks.
    sFilePath = sTempPath & sTableName & "_DataMacros.xml"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(sFilePath, True, True)
    txtFile.Write sXmlContent
    txtFile.Close
    Set txtFile = Nothing

    DoCmd.OpenTable sTableName, acViewDesign, acHidden
    Application.LoadFromText acTableDataMacro, sTableName, sFilePath
    DoCmd.Close acTable, sTableName, acSaveYes

    fso.DeleteFile sFilePath

    If bHasLongText Then
        Debug.Print "  - All 5 data macros created (3 After + BeforeChange/BeforeDelete)"
        CreateAllDataMacros = "OK — 5 data macros created (3 After + BeforeChange/BeforeDelete)" & sBackupNote
    Else
        Debug.Print "  - All 3 data macros created (AfterInsert, AfterUpdate, AfterDelete)"
        CreateAllDataMacros = "OK — 3 data macros created (AfterInsert, AfterUpdate, AfterDelete)" & sBackupNote
    End If

Cleanup:
    Set rsCheck = Nothing
    Set txtFile = Nothing
    Set fso = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    CreateAllDataMacros = "ERROR: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox "Error creating macros for " & sTableName & ": " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Function
```

### BuildAfterInsertMacro — `Private Function` → `String`

Emits the AfterInsert `<DataMacro>` fragment: one `tblAuditLog` row per configured field, marking
`OldValue` as `[NEW RECORD]`.

```vba
Private Function BuildAfterInsertMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterInsert: log every configured field of the new row.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String

    sXml = "<DataMacro Event=""AfterInsert""><Statements>"

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        sFieldName = fieldInfo(0)

        sXml = sXml & "<CreateRecord>"
        sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        sXml = sXml & "<Statements>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sXml = sXml & "</Action>"

        If sPrimaryKeyField <> "" Then
            sXml = sXml & "<Action Name=""SetField"">"
            sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sPrimaryKeyField & "]</Argument>"
            sXml = sXml & "</Action>"
        End If

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
        sXml = sXml & "</Action>"

        ' OperationType (schema Business Rule 6); OldValue stays Null on an insert
        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        sXml = sXml & "<Argument Name=""Value"">""Insert""</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
        sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sFieldName & "]</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
        sXml = sXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] CurrentUser() is the dependency-free default;
        ' the real-username Extra Option substitutes a public AuditUser() helper here.
        sXml = sXml & "<Action Name=""SetField"">"
        sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        sXml = sXml & "<Argument Name=""Value"">CurrentUser()</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "</Statements></CreateRecord>"
      End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterInsertMacro = sXml
End Function
```

### BuildAfterUpdateMacro — `Private Function` → `String`

Emits the AfterUpdate fragment. Ordinary fields: log only when the value actually changed
(`GetComparisonExpression`), reading the old value from `[Old]`. Long Text fields: always log
(schema Business Rule 6), retrieving the old value from `tblLongTextBackup` via `LookUpRecord` —
the BeforeChange macro put it there (schema Business Rule 3).

```vba
Private Function BuildAfterUpdateMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterUpdate: one conditional block per non-PK field; Long Text goes
    '            through the LookUpRecord retrieval path.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim lFldType As Long
    Dim bIsLongText As Boolean

    sXml = "<DataMacro Event=""AfterUpdate""><Statements>"

    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        lFldType = fieldInfo(1)
        bIsLongText = (lFldType = dbMemo)

        ' Auditable fields only (schema Business Rule 5); PK never changes, skip it
        If fieldInfo(3) = True And sFieldName <> sPrimaryKeyField Then
            sXml = sXml & "<ConditionalBlock><If>"
            sXml = sXml & "<Condition>" & GetComparisonExpression(sTableName, sFieldName, lFldType) & "</Condition>"
            sXml = sXml & "<Statements>"

            ' Long Text: fetch the backed-up old value (schema Business Rule 3)
            If bIsLongText Then
                sXml = sXml & "<LookUpRecord>"
                sXml = sXml & "<Data Alias=""BackupRec"">"
                sXml = sXml & "<Reference>tblLongTextBackup</Reference>"
                sXml = sXml & "<WhereCondition>"
                sXml = sXml & "[tblLongTextBackup].[TableName]=""" & sTableName & """ And "
                sXml = sXml & "[tblLongTextBackup].[PrimaryKey]=[" & sTableName & "].[" & sPrimaryKeyField & "] And "
                sXml = sXml & "[tblLongTextBackup].[FieldName]=""" & sFieldName & """"
                sXml = sXml & "</WhereCondition>"
                sXml = sXml & "</Data>"
                sXml = sXml & "<Statements>"
            End If

            sXml = sXml & "<CreateRecord>"
            If bIsLongText Then
                sXml = sXml & "<Data><Reference>tblAuditLog</Reference></Data>"
            Else
                sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
            End If
            sXml = sXml & "<Statements>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
            sXml = sXml & "</Action>"

            If sPrimaryKeyField <> "" Then
                sXml = sXml & "<Action Name=""SetField"">"
                If bIsLongText Then
                    sXml = sXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
                Else
                    sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
                End If
                sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sPrimaryKeyField & "]</Argument>"
                sXml = sXml & "</Action>"
            End If

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
            sXml = sXml & "</Action>"

            ' OperationType (schema Business Rule 6)
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">""Update""</Argument>"
            sXml = sXml & "</Action>"

            ' OldValue — from the backup for Long Text, from [Old] otherwise
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
                sXml = sXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
                sXml = sXml & "<Argument Name=""Value"">[Old].[" & sFieldName & "]</Argument>"
            End If
            sXml = sXml & "</Action>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">[" & sTableName & "].[" & sFieldName & "]</Argument>"
            sXml = sXml & "</Action>"

            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
            sXml = sXml & "</Action>"

            ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">CurrentUser()</Argument>"
            sXml = sXml & "</Action>"

            sXml = sXml & "</Statements></CreateRecord>"

            If bIsLongText Then
                sXml = sXml & "</Statements></LookUpRecord>"
            End If

            sXml = sXml & "</Statements></If></ConditionalBlock>"
        End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterUpdateMacro = sXml
End Function
```

### BuildAfterDeleteMacro — `Private Function` → `String`

Emits the AfterDelete fragment: one `tblAuditLog` row per field, `NewValue` marked `[DELETED]`,
old values read from `[Old]` — except Long Text, retrieved from the backup the BeforeDelete macro
staged.

```vba
Private Function BuildAfterDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] AfterDelete: log every configured field of the deleted row.
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim lFldType As Long
    Dim bIsLongText As Boolean

    sXml = "<DataMacro Event=""AfterDelete""><Statements>"

    For Each fieldInfo In fieldList
      If fieldInfo(3) = True Then    ' auditable fields only (schema Business Rule 5)
        sFieldName = fieldInfo(0)
        lFldType = fieldInfo(1)
        bIsLongText = (lFldType = dbMemo)

        If bIsLongText Then
            sXml = sXml & "<LookUpRecord>"
            sXml = sXml & "<Data Alias=""BackupRec"">"
            sXml = sXml & "<Reference>tblLongTextBackup</Reference>"
            sXml = sXml & "<WhereCondition>"
            sXml = sXml & "[tblLongTextBackup].[TableName]=""" & sTableName & """ And "
            sXml = sXml & "[tblLongTextBackup].[PrimaryKey]=[Old].[" & sPrimaryKeyField & "] And "
            sXml = sXml & "[tblLongTextBackup].[FieldName]=""" & sFieldName & """"
            sXml = sXml & "</WhereCondition>"
            sXml = sXml & "</Data>"
            sXml = sXml & "<Statements>"
        End If

        sXml = sXml & "<CreateRecord>"
        If bIsLongText Then
            sXml = sXml & "<Data><Reference>tblAuditLog</Reference></Data>"
        Else
            sXml = sXml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        End If
        sXml = sXml & "<Statements>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
        sXml = sXml & "</Action>"

        If sPrimaryKeyField <> "" Then
            sXml = sXml & "<Action Name=""SetField"">"
            If bIsLongText Then
                sXml = sXml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
            Else
                sXml = sXml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            End If
            sXml = sXml & "<Argument Name=""Value"">[Old].[" & sPrimaryKeyField & "]</Argument>"
            sXml = sXml & "</Action>"
        End If

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""" & sFieldName & """</Argument>"
        sXml = sXml & "</Action>"

        ' OperationType (schema Business Rule 6); NewValue stays Null on a delete
        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OperationType</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.OperationType</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">""Delete""</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
            sXml = sXml & "<Argument Name=""Value"">[Old].[" & sFieldName & "]</Argument>"
        End If
        sXml = sXml & "</Action>"

        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">Now()</Argument>"
        sXml = sXml & "</Action>"

        ' [STANDARDS / schema Business Rule 9] identity — see BuildAfterInsertMacro
        sXml = sXml & "<Action Name=""SetField"">"
        If bIsLongText Then
            sXml = sXml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
        Else
            sXml = sXml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        End If
        sXml = sXml & "<Argument Name=""Value"">CurrentUser()</Argument>"
        sXml = sXml & "</Action>"

        sXml = sXml & "</Statements></CreateRecord>"

        If bIsLongText Then
            sXml = sXml & "</Statements></LookUpRecord>"
        End If
      End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"
    BuildAfterDeleteMacro = sXml
End Function
```

### BuildBeforeChangeMacro — `Private Function` → `String`

Emits the BeforeChange fragment for a Long Text table. Distinguishes insert from update with
`IsNull([Old].[PK])` — on an insert there is nothing to back up; on an update it calls
`BackupLongTextFieldsDM` for each Long Text field. Returns an empty string when the table has no
Long Text fields.

```vba
Private Function BuildBeforeChangeMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] BeforeChange: stage Long Text old values ahead of an update
    '            (schema Business Rules 2-3).
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim bHasLongText As Boolean

    ' Emit nothing for a table without Long Text
    bHasLongText = False
    For Each fieldInfo In fieldList
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            bHasLongText = True
            Exit For
        End If
    Next fieldInfo
    If Not bHasLongText Then
        BuildBeforeChangeMacro = ""
        Exit Function
    End If

    sXml = "<DataMacro Event=""BeforeChange""><Statements>"

    ' Insert (PK is Null in [Old]) vs update
    sXml = sXml & "<ConditionalBlock><If>"
    sXml = sXml & "<Condition>IsNull([Old].[" & sPrimaryKeyField & "])</Condition>"
    sXml = sXml & "<Statements>"
    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">lngPKValue</Argument>"
    sXml = sXml & "<Argument Name=""Value"">0</Argument>"
    sXml = sXml & "</Action>"
    sXml = sXml & "</Statements></If>"

    sXml = sXml & "<Else><Statements>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">lngPKValue</Argument>"
    sXml = sXml & "<Argument Name=""Value"">=[" & sPrimaryKeyField & "]</Argument>"
    sXml = sXml & "</Action>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">strTableName</Argument>"
    sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
    sXml = sXml & "</Action>"

    ' One backup call per Long Text field — a data macro CAN call a public VBA
    ' function in the same accdb; that is the hinge of the whole hybrid method.
    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            sXml = sXml & "<Action Name=""SetLocalVar"">"
            sXml = sXml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
            sXml = sXml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & sFieldName & """)</Argument>"
            sXml = sXml & "</Action>"
        End If
    Next fieldInfo

    sXml = sXml & "</Statements></Else></ConditionalBlock>"
    sXml = sXml & "</Statements></DataMacro>"

    BuildBeforeChangeMacro = sXml
End Function
```

### BuildBeforeDeleteMacro — `Private Function` → `String`

Emits the BeforeDelete fragment for a Long Text table: captures the PK, then stages each Long
Text value through `BackupLongTextFieldsDM` so the AfterDelete macro can log it. Returns an empty
string when the table has no Long Text fields.

```vba
Private Function BuildBeforeDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
    ' [SCAFFOLD] BeforeDelete: stage Long Text values ahead of a delete
    '            (schema Business Rules 2-3).
    Dim sXml As String
    Dim fieldInfo As Variant
    Dim sFieldName As String
    Dim bHasLongText As Boolean

    bHasLongText = False
    For Each fieldInfo In fieldList
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            bHasLongText = True
            Exit For
        End If
    Next fieldInfo
    If Not bHasLongText Then
        BuildBeforeDeleteMacro = ""
        Exit Function
    End If

    sXml = "<DataMacro Event=""BeforeDelete""><Statements>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">lngPKValue</Argument>"
    sXml = sXml & "<Argument Name=""Value"">=[" & sPrimaryKeyField & "]</Argument>"
    sXml = sXml & "</Action>"

    sXml = sXml & "<Action Name=""SetLocalVar"">"
    sXml = sXml & "<Argument Name=""Name"">strTableName</Argument>"
    sXml = sXml & "<Argument Name=""Value"">""" & sTableName & """</Argument>"
    sXml = sXml & "</Action>"

    For Each fieldInfo In fieldList
        sFieldName = fieldInfo(0)
        If fieldInfo(1) = dbMemo And fieldInfo(3) = True Then
            sXml = sXml & "<Action Name=""SetLocalVar"">"
            sXml = sXml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
            sXml = sXml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & sFieldName & """)</Argument>"
            sXml = sXml & "</Action>"
        End If
    Next fieldInfo

    sXml = sXml & "</Statements></DataMacro>"

    BuildBeforeDeleteMacro = sXml
End Function
```

### GetComparisonExpression — `Private Function` → `String`

The change test the AfterUpdate macro embeds per field. Ordinary fields compare old and new with
`StrComp` over `Nz`-wrapped values; Long Text returns `True` — always log — because the macro
cannot read the old value to compare (schema Business Rule 6).

**The not-equal operator is built from `Chr(38)`, not typed as a literal `&lt;&gt;`.** A real
build hit error 3870 ("Microsoft Access cannot interpret the text you are pasting as a data
macro") on every table because the VBA source was imported through a tool that **un-escapes HTML
entities on the way in** — a literal `&lt;&gt;` in this function's source became a raw `<>` once
imported, which breaks the macro XML the moment it's loaded (the `<` is read as an opening tag).
Assembling the entity from character codes at runtime means no literal `&`-entity ever exists in
the source for an importer to touch — the string is only ever `&lt;&gt;` in memory, never in text
anyone's tool re-reads. See `templates/_materialization.md` for the general rule this follows.

```vba
Private Function GetComparisonExpression(sTableName As String, sFieldName As String, lFldType As Long) As String
    ' [SCAFFOLD] Per-type change test for the AfterUpdate conditional block.
    Dim sNotEqual As String
    Select Case lFldType
        Case dbMemo
            ' Long Text: always log (cannot compare the old value in-macro)
            GetComparisonExpression = "True"
        Case Else
            ' Built from Chr(38), not typed as a literal &lt;&gt; — see the note above this
            ' block. This is XML for "<>"; the expression lives inside <Condition>...</Condition>.
            sNotEqual = Chr(38) & "lt;" & Chr(38) & "gt;"
            GetComparisonExpression = "StrComp(NZ([" & sTableName & "].[" & sFieldName & "],""""),NZ([Old].[" & sFieldName & "],""""),0)" & sNotEqual & "0"
    End Select
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
Public Function BackupLongTextFieldsDM(strTableName As String, lngPKValue As Long, strFieldName As String)
    ' [SCAFFOLD] Stage one Long Text field's current value before an update or delete.
    '            Called by the generated BeforeChange / BeforeDelete Data Macros.
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsOldValue As DAO.Recordset
    Dim strPKField As String
    Dim strOldValue As Variant

    On Error GoTo errHandler
    Set db = CurrentDb

    If strTableName = "" Then Exit Function

    ' Replace any earlier backup for this table/field/record
    db.Execute "DELETE FROM tblLongTextBackup WHERE TableName='" & strTableName & _
        "' AND FieldName='" & strFieldName & "' AND PrimaryKey=" & lngPKValue, dbFailOnError

    ' The audited table's PK field name comes from the config
    strPKField = DLookup("FieldName", "tblAuditLogConfig", _
        "TableName='" & strTableName & "' AND IsPrimaryKey=" & True)

    If lngPKValue > 0 Then    ' updates and deletes only — a new record has no old value
        Set rsOldValue = db.OpenRecordset("SELECT " & strFieldName & " FROM " & strTableName & _
            " WHERE " & strPKField & "=" & lngPKValue)
        strOldValue = rsOldValue.Fields(strFieldName).Value
        rsOldValue.Close

        Set rs = db.OpenRecordset("tblLongTextBackup", dbOpenDynaset)
        rs.AddNew
        rs!TableName = strTableName
        rs!PrimaryKey = lngPKValue
        rs!FieldName = strFieldName
        rs!OldValue = strOldValue
        rs!DateChanged = Now()
        ' [STANDARDS / schema Business Rule 9] identity — CurrentUser() default; the
        ' real-username Extra Option substitutes AuditUser() here too.
        rs!ChangedBy = CurrentUser()
        rs.Update
        rs.Close
    End If

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
End Function
```

### BackupAndRemoveAllDataMacros — `Public Function` → `Boolean` (module `modAuditAdmin` — back end only)

The reset tool for regeneration (schema Business Rule 7): exports every table's current data
macros to timestamped XML backups, then strips them by loading an empty macro document. Run it
before re-running `Three_GenerateAllAuditDataMacros` when the audit scope changes — the backups
double as your archive of prior macro states.

```vba
Public Function BackupAndRemoveAllDataMacros(Optional strBackupPath As String = "") As Boolean
    ' [SCAFFOLD] Back up, then remove, the data macros on every table that has them.
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strTempFile As String
    Dim strBackupFile As String
    Dim intFileNum As Integer
    Dim intMacrosRemoved As Integer

    On Error GoTo errHandler
    Set db = CurrentDb
    intMacrosRemoved = 0

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
        Debug.Print "Processing: " & rst!Name

        strBackupFile = strBackupPath & rst!Name & "_DataMacro_" & _
                        Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Application.SaveAsText acTableDataMacro, rst!Name, strBackupFile

        Application.LoadFromText acTableDataMacro, rst!Name, strTempFile
        intMacrosRemoved = intMacrosRemoved + 1
        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
    Kill strTempFile

    MsgBox "Successfully backed up and removed data macros from " & intMacrosRemoved & " tables." & vbCrLf & _
           "Backups saved to: " & strBackupPath, vbInformation, "Data Macros Removed"

    BackupAndRemoveAllDataMacros = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] error 2950 = a table with no data macros to export;
    '            skip it and continue. Anything else: report and stop.
    If Err.Number <> 2950 Then
        MsgBox Err.Number & " Error: " & Err.Description, vbExclamation
        On Error Resume Next
        If Not rst Is Nothing Then rst.Close
        If Dir(strTempFile) <> "" Then Kill strTempFile
        BackupAndRemoveAllDataMacros = False
    Else
        Resume Cleanup
    End If
End Function
```

## Standards Layer

- **Error handling** — the blocks above ship the dependency-free `MsgBox` default; substitute
  your house pattern per `error-handling.md`. Two deliberate exceptions are annotated in place:
  `BackupLongTextFieldsDM` stays quiet (it runs inside every save), and
  `BackupAndRemoveAllDataMacros` skips error 2950 by design.
- **Query style** — the inline SQL kept here is from the proven source; rewrite per
  `query-style.md` if your house centralizes SQL differently.
- **Naming conventions** — the config scan's `tbl`/`tlkp` (excluding `tmp`) prefix filter is the
  naming convention made executable, and `CheckAuditReadiness` repeats the same filter so the two
  stay in step; adjust both with your prefix policy.
- **Design principles** — one job per procedure throughout: one sample-data setup (Path A only),
  three numbered entry points, one safety check, five single-macro builders, one comparison
  helper, one staging function, one admin reset.

## Extra Options

*Empty in the base template. Filled per client engagement.*

- **Real-username identity** — add a public `AuditUser()` function (`Environ("USERNAME")`) to
  both back end and front ends, and substitute it for `CurrentUser()` in the builders and in
  `BackupLongTextFieldsDM` (schema Business Rule 9 and its Extra Option). The production system
  this template is drawn from runs this upgrade — its trail stamps real Windows usernames.
- **Scheduled backup-table cleanup** — a maintenance routine clearing aged `tblLongTextBackup`
  rows (schema Business Rule 8).

## Parked / future considerations

- **Restore/undo tooling** — reconstructing a record from its `tblAuditLog` trail; the full
  (non-Lite) system's headline feature.
- **Composite/text primary keys** — `CheckAuditReadiness` detects these and tells you to fix or
  exclude the table; the staging plumbing and macro XML themselves still assume one numeric PK
  and don't support them (schema Business Rule 4).
