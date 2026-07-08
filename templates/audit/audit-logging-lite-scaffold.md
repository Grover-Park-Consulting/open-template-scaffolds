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
  - One_CreateAuditTables
  - Two_PopulateConfigTable
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
    it, but the adaptation is theirs.
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

Setup is three numbered steps, run once in the back end, in order:

```vba
One_CreateAuditTables            ' 1. create tblAuditLog / tblLongTextBackup / tblAuditLogConfig
Two_PopulateConfigTable          ' 2. scan the schema into config
'    ... then review tblAuditLogConfig and flip IsAuditable flags ...
Three_GenerateAllAuditDataMacros ' 3. build + attach 3 or 5 macros per audited table
```

Then link the three system tables into the front end and import the Long Text helper module
there (see `BackupLongTextFieldsDM`).

**Module homes** (three modules, one job each):

| Module | Procedures | Lives in |
|---|---|---|
| `modAddDataMacros` | the numbered steps + the five `Build*` XML builders + `GetComparisonExpression` | Back end only |
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
silently dropped — exclusions are *seeded* as `IsAuditable = False` rows (the three system
tables; noisy always-changing fields), and everything else defaults True. **After running, open
the config table and review the flags** — that review, in data, is where the audit net is drawn
(schema Business Rule 5).

```vba
Public Sub Two_PopulateConfigTable()
    ' [SCAFFOLD] Rebuild the audit configuration from the live schema. Scope decisions
    '            live in the IsAuditable flags afterward, not in this code.
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
        ' scans tables prefixed tbl (your data tables under naming-conventions.md); system
        ' (MSys) and temporary tables never match it. On another prefix policy, change this
        ' one test — every finer-grained decision is a flag in the config table, not code.
        ' >>> adjust the prefix test to your naming convention <<<
        If Left(tdef.Name, 3) = "tbl" Then

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
                '            noisy always-changing fields; True for everything else. Review
                '            and flip flags in tblAuditLogConfig after the scan.
                Select Case True
                    Case tdef.Name = "tblAuditLog", _
                         tdef.Name = "tblLongTextBackup", _
                         tdef.Name = "tblAuditLogConfig"
                        isAuditable = False
                    Case fld.Name = "AccessTS", fld.Name = "SSMA_TimeStamp", _
                         fld.Name = "ValidFrom", fld.Name = "ValidTo"
                        isAuditable = False
                    Case Else
                        isAuditable = True
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

    MsgBox "Configuration table populated. Review IsAuditable in tblAuditLogConfig " & _
        "before generating macros.", vbInformation

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

### Three_GenerateAllAuditDataMacros — `Public Sub` (setup step 3)

Reads the (reviewed) config, groups fields by table, and calls `CreateAllDataMacros` for each.
Re-runnable: reloading a table's macro XML replaces what was there (schema Business Rule 7).

```vba
Public Sub Three_GenerateAllAuditDataMacros()
    ' [SCAFFOLD] Generate and attach audit Data Macros for every configured table.
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

    On Error GoTo errHandler
    Set db = CurrentDb
    Set dictTables = CreateObject("Scripting.Dictionary")

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
        Debug.Print "Processing: " & sTableName & " (" & fieldList.Count & " fields)"
        Call CreateAllDataMacros(sTableName, fieldList, sTempPath)
        lTableCount = lTableCount + 1
    Next vCurrentTable

    MsgBox "Successfully generated audit data macros for " & lTableCount & " tables!", vbInformation

Cleanup:
    Set rs = Nothing
    Set db = Nothing
    Set dictTables = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub
```

### CreateAllDataMacros — `Private Sub`

Builds one table's macro XML — the three After macros always, plus BeforeChange/BeforeDelete when
the field list contains a Long Text field (schema Business Rule 2) — writes it UTF-16, and loads
it with `LoadFromText acTableDataMacro`.

```vba
Private Sub CreateAllDataMacros(sTableName As String, fieldList As Collection, sTempPath As String)
    ' [SCAFFOLD] Generate the 3 (or 5, with Long Text) Data Macros for one table.
    Dim sXmlContent As String
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim txtFile As Object
    Dim sFilePath As String
    Dim sPrimaryKeyField As String
    Dim fieldInfo As Variant
    Dim bHasLongText As Boolean
    Dim lAuditableCount As Long

    On Error GoTo errHandler

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
        Exit Sub
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
    Else
        Debug.Print "  - All 3 data macros created (AfterInsert, AfterUpdate, AfterDelete)"
    End If

Cleanup:
    Set txtFile = Nothing
    Set fso = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    MsgBox "Error creating macros for " & sTableName & ": " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub
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

```vba
Private Function GetComparisonExpression(sTableName As String, sFieldName As String, lFldType As Long) As String
    ' [SCAFFOLD] Per-type change test for the AfterUpdate conditional block.
    Select Case lFldType
        Case dbMemo
            ' Long Text: always log (cannot compare the old value in-macro)
            GetComparisonExpression = "True"
        Case Else
            ' Note &lt;&gt; — the expression lives inside XML, so <> must be escaped
            GetComparisonExpression = "StrComp(NZ([" & sTableName & "].[" & sFieldName & "],""""),NZ([Old].[" & sFieldName & "],""""),0)&lt;&gt;0"
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
- **Naming conventions** — the config scan's `tbl` prefix filter is the naming convention made
  executable; adjust it with your prefix policy.
- **Design principles** — one job per procedure throughout: three numbered entry points, five
  single-macro builders, one comparison helper, one staging function, one admin reset.

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
- **Composite/text primary keys** — the staging plumbing and macro XML assume one numeric PK
  (schema Business Rule 4).
