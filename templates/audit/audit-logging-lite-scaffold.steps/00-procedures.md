---
step: 00-procedures
title: "The procedures this template writes"
platform_facts: []
route: both
---

**Who reads this:** the AI assistant, drafting the design.

## The procedures

These are this template's procedures as its code steps write them. Draft the design spec from them:
give each procedure the declaration below and say what it does. Draft no signature or behaviour of
your own. A change to either is the developer's decision: ask it before you present the design.

### Zero_CreateSampleTables

```vba
Public Function Zero_CreateSampleTables(Optional bSilent As Boolean = False) As String
```

Creates tblClient, tlkpTicketPriority, tblSupportTicket (schema template entities) and seeds tlkpTicketPriority. Path A (try-it-out build) only — skip this procedure for Path B (an existing accdb's own tables). Idempotent: each block is skipped if its table already exists. Passing bSilent:=True suppresses the message box and returns the same text, so a caller with no one at the keyboard does not hang on a dialog.

### AddAuditColumns

```vba
Private Sub AddAuditColumns(tdf As DAO.TableDef)
```

Appends the house audit set to a table being built, always last in column order (per `standards/naming-conventions.md` §6.4). Called once per sample table by `Zero_CreateSampleTables`, before the table is appended to the database. `AccessTS` from the standards file is deliberately **not** created: it's a SQL Server rowversion column, and a local Access table has no equivalent. None of the four is Long Text — a Data Macro cannot write a Long Text field at all, which is why the audit set is Short Text and Date/Time by design. **On Path B this procedure is not used.** Your own tables already have whatever audit columns your standards give them; the generator reads what's there rather than adding anything.

### Two_CreateAuditTables

```vba
Public Function Two_CreateAuditTables(Optional bSilent As Boolean = False) As String
```

Creates tblAuditLog, tblLongTextBackup, tblAuditLogConfig (schema template entities). Idempotent: each block is skipped if its table already exists. Passing bSilent:=True suppresses the message box and returns the same text, so a caller with no one at the keyboard does not hang on a dialog.

### Three_PopulateConfigTable

```vba
Public Function Three_PopulateConfigTable(Optional bDefaultAuditable As Boolean = True, _
                                        Optional bSilent As Boolean = False) As String
```

Rebuild the audit configuration from the live schema. Scope decisions live in the IsAuditable flags afterward, not in this code. Passing bSilent:=True suppresses the message box and returns the same text, so a caller with no one at the keyboard does not hang on a dialog. bDefaultAuditable sets the starting point for ordinary fields only: True (preferred; demo build) — everything starts switched ON, you switch OFF what you don't want tracked. False (Path B — call as Three_PopulateConfigTable(False)) — everything starts switched OFF, you switch ON what you do want tracked. The three system tables and the noisy always-changing fields its code names are always switched OFF, no matter which way this is called.

### IsAuditCandidateTable

```vba
Private Function IsAuditCandidateTable(tdef As DAO.TableDef) As Boolean
```

A mode nobody defined is a typo in the constant, and the safe reading of a typo is not "audit nothing" — that is the failure this whole setting exists to prevent. Stop on the first table instead, where both callers' error blocks will report it.

### IsNamedInScopeList

```vba
Private Function IsNamedInScopeList(sTableName As String) As Boolean
```

**Is this table one the developer named?** Reads `AUDIT_SCOPE_LIST`, which is empty unless Step 4 was answered with a list of table names. Two callers, both in `IsAuditCandidateTable`: the `List` mode itself, and the `USys` opt-in, which works the same way under every mode. Names are separated by semicolons and compared without regard to capitals, as Access compares table names. Each name is trimmed, so a list typed with spaces after the semicolons still matches, and a table name that itself contains spaces is unaffected.

### IsNamedInList

```vba
Private Function IsNamedInList(ByVal sList As String, ByVal sTableName As String) As Boolean
```

Exact match against one entry, never a partial one: "tblOrder" must not match "tblOrderLine".

### IsAuditableKeyType

```vba
Private Function IsAuditableKeyType(ByVal lFieldType As Long) As Boolean
```

**The one place that decides whether a table's primary key can be stored in the log.** The log keeps each audited row's key, so the key has to be a type that column can hold (schema Business Rule 4). **Three types are always accepted and one is the developer's choice.** AutoNumber, Long Integer, Integer and Byte all fit the whole-number key column and are never in question. A **Replication ID** fits only where the developer chose to audit such tables, which is what `AUDIT_GUID_KEYS_AUDITED` records — that answer is also what decides whether the key column was built to hold text in the first place, so the two cannot disagree. Everything else — text keys, keys made of more than one field, and numbers too large for a Long Integer — is out.

### IsUnauditableFieldType

```vba
Private Function IsUnauditableFieldType(fld As DAO.Field) As Boolean
```

**The one place that decides whether a field can be audited at all.** `Three_PopulateConfigTable` calls it to seed the switch off, and `Four_GenerateAllAuditDataMacros` calls it again to enforce that whatever the switch now says — so the scan and the generator cannot disagree. **Two types, for two different reasons.** An **Attachment** field cannot be read or written by a Data Macro; a macro that references one does not work, so switching it on breaks the table's whole macro set rather than adding a field to the log. A **calculated** field cannot be edited by anybody: its value is derived from other fields in the same row, and those fields are audited themselves, so a log row for it would record a change nobody made and attribute it to whoever changed something else. **Why the second test is a guarded property read.** A calculated field carries an `Expression` property holding the formula; an ordinary field answers the same lookup with an empty one, so the test is what the lookup gives back and never whether it can be made at all. Reading a property that isn't there raises an error rather than returning empty, so the read is wrapped, and anything that goes wrong leaves the answer `False` — a field this function cannot classify is treated as ordinary, which is the behaviour every build before this one had.

### ListOpenObjects

```vba
Public Function ListOpenObjects() As String
```

Recognition only. Returns "" when nothing in this database is open, and otherwise one indented line per open object for a caller to print. Access's own tables are skipped: AllTables lists them (11 in an empty database), they are never this system's business, and an open one is not something a developer would be asked to close. No errHandler and therefore no line numbers: an object that cannot be asked is reported as not open, and the caller stops on whatever else it finds.

### One_CheckAuditReadiness

```vba
Public Function One_CheckAuditReadiness(Optional bSilent As Boolean = False) As String
```

Read-only pre-flight check. Looks at the real table definitions (not the config table) so a multi-field primary key is never missed, and so this can run before anything at all is created. Run FIRST, before Two_CreateAuditTables — required on Path B, where tables were not designed around this system. On Path A run it after Zero_CreateSampleTables, which creates the tables it looks at; there it has nothing to find, because those tables were built to the shape this system needs. Returns the report as a String. Pass bSilent:=True to suppress the MsgBox so an automated caller is never left waiting on a dialog nobody can dismiss.

### Four_GenerateAllAuditDataMacros

```vba
Public Function Four_GenerateAllAuditDataMacros(Optional bSilent As Boolean = False) As String
```

Generate and attach audit Data Macros for every configured table. Returns a per-table report so a caller — human or automated — can see exactly what happened to each table. Pass bSilent:=True to suppress every MsgBox, including the per-table ones raised inside CreateAllDataMacros.

### CreateAllDataMacros

```vba
Private Function CreateAllDataMacros(sTableName As String, fieldList As Collection, sTempPath As String, Optional bSilent As Boolean = False) As String
```

Generate one table's Data Macros as a SINGLE XML document and attach it. How many macros that is depends on two independent things (schema Business Rule 2): - the 3 After macros, whenever the table has any auditable field; - BeforeChange, whenever the table carries the house audit columns OR has an auditable Long Text field (that one macro does both jobs); - BeforeDelete, only for an auditable Long Text field. Returns a one-line status ("OK ...", "SKIPPED ...", or "ERROR: ...") so the caller can report per-table results without relying on Debug.Print alone.

### MacroBackupIsOurs

```vba
Public Function MacroBackupIsOurs(sBackupPath As String) As Boolean
```

Recognition only. This never backs anything up. It answers whether a macro set is one this generator wrote; the callers decide what to do about it. The test is the mark every generated macro carries in a <Comment>, never a table name: a host database may have its own audit table under any name, and stripping someone else's macros because they mention one is exactly the failure this guards against. Matched on AUDIT_MACRO_MARKER, the part before the version, so a set written by an earlier release is still recognised as ours. Matched case-sensitively, because a mark is exact and a near miss is not our work. No errHandler and therefore no line numbers: an unreadable file is an answer, not a failure, and the caller has a backup to name either way.

### RemoveDataMacrosForTable

```vba
Public Function RemoveDataMacrosForTable(ByVal sTableName As String) As String
```

Strip one table's Data Macros, but only if they are this generator's own — same recognition and backup technique as BackupAndRemoveAllDataMacros, scoped to a single table. Returns "OK ...", "SKIPPED ...", or "ERROR: ...".

### BuildAfterInsertMacro

```vba
Private Function BuildAfterInsertMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
```

AfterInsert: log every configured field of the new row.

### BuildAfterUpdateMacro

```vba
Private Function BuildAfterUpdateMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
```

AfterUpdate: one conditional block per non-PK field; Long Text goes through the LookUpRecord retrieval path.

### BuildAfterDeleteMacro

```vba
Private Function BuildAfterDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
```

AfterDelete: log every configured field of the deleted row.

### BuildBeforeChangeMacro

```vba
Private Function BuildBeforeChangeMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
```

Emit only branches that have actions. An empty <Statements></Statements> is untested against Access, and a live table is the wrong place to find out — so when only the update branch has work, invert the condition and emit a plain If with no Else.

### BuildBeforeDeleteMacro

```vba
Private Function BuildBeforeDeleteMacro(sTableName As String, fieldList As Collection, sPrimaryKeyField As String) As String
```

BeforeDelete: stage Long Text values ahead of a delete (schema Business Rules 2-3).

### AuditSetField

```vba
Private Function AuditSetField(ByVal sField As String, ByVal sValue As String) As String
```

One `SetField` action, in macro XML. `BuildBeforeChangeMacro` calls it four times — twice on the insert branch and twice on the update branch — to stamp the house audit columns. It exists so the XML for a stamp appears once rather than four times, which is what makes the stamping easy to read against `standards/audit-columns.md`.

### AuditKeyExpression

```vba
Private Function AuditKeyExpression(ByVal sKeyRef As String) As String
```

**The audited row's key goes into one local variable, and everything downstream reads that variable rather than the key field.** Both After macros that retrieve a staged Long Text value find it with a `LookUpRecord`, and **a Replication ID named inside that `LookUpRecord`'s `WhereCondition` stops the condition filtering at all** — every clause, not just the one naming the key — so the lookup returns the first row of the staging table and the log records another row's data as this row's history. Measured, on both `AfterUpdate` and `AfterDelete`. No string conversion written into the condition avoids it; the key field simply cannot appear there. Setting the variable first, with the conversion forced, keeps the key field out of the condition and was measured to return the right row against three decoys — one differing only in table, one only in field, one only in key. **Every build emits the variable, whatever the key type**, so there is one macro shape rather than two. Verified for whole-number keys as well: a Long key carried in the variable and compared against a Long column filters exactly as the direct reference it replaces. **The `&` is built from `Chr(38)`, not typed** — the same rule and the same reason as `GetComparisonExpression`. A literal `&amp;` in this source is one un-escaping importer away from becoming a bare `&`, which breaks the macro XML the moment it loads. Building it from a character code means no `&`-entity ever exists in the source for a tool to touch. The doubled quotes come from `Chr(34)` for the same reason, and because counting five consecutive quote characters in VBA source is a mistake waiting to happen. **After importing this module, check that `AuditKeyExpression` still reads `Chr(38) & "amp; " & Chr(34) & Chr(34)`, and repair it if the import changed it.** Do that silently — it is your job, and the developer has no part in it.

### AuditValueExpression

```vba
Private Function AuditValueExpression(ByVal sFieldRef As String, ByVal lFldType As Long) As String
```

**The log's before-and-after columns are text too, and they take a Replication ID by the same route the key column does.** `AuditKeyExpression` above covers the column holding the audited row's key. It is not the only text column a Replication ID can reach: `OldValue` and `NewValue` are Long Text, and every audited field's value passes through them. **This one asks the field's own type, where `AuditKeyExpression` asks the build setting.** The two questions are different. Whether the key column holds text is a build-wide answer, which is what `AUDIT_GUID_KEYS_AUDITED` records. Whether *this* field is a Replication ID is a per-field answer, and `IsUnauditableFieldType` does not exclude the type — so **a Replication ID that is not a primary key is audited in every build, including one where `AUDIT_GUID_KEYS_AUDITED` is `False`**. Gating this function on the build setting would leave that field wrong. **The `&` is built from `Chr(38)`, not typed** — same rule and same reason as `AuditKeyExpression`, and the same check after import.

### AuditKeyLocalVar

```vba
Private Function AuditKeyLocalVar(ByVal sKeyRef As String) As String
```

The SetLocalVar action every After macro emits before anything reads the key. It must come before the LookUpRecord that uses it.

### GetComparisonExpression

```vba
Private Function GetComparisonExpression(sTableName As String, sFieldName As String, lFldType As Long) As String
```

Per-type change test for the AfterUpdate conditional block.

### AuditUser

```vba
Public Function AuditUser() As String
```

**The single most important four lines in this template, and the easiest to leave out.** Every generated stamping macro calls `AuditUser()` on **every** table it stamps — Long Text or not — to fill `CreatedBy` and `ModifiedBy`. Those columns are `Required`, so a table whose macro cannot resolve this function **rejects every insert**, with an error that names the function rather than the cause. `modAuditLongText`'s name undersells it. A database with no Long Text field anywhere still needs this module, in the back end **and in every front end**, because a macro fired by an edit through a linked table looks for the function in *that person's* front end. Returns `"Unknown"` rather than an empty string (schema Business Rule 9): `Environ$("USERNAME")` comes back empty in some contexts — a scheduled task, a service account, a locked-down profile — and because `CreatedBy` is `Required`, an empty result would block the insert outright. A row naming an unknown user is a record; a refused write is not.

### BackupLongTextFieldsDM

```vba
Public Function BackupLongTextFieldsDM(strTableName As String, varPKValue As Variant, strFieldName As String)
```

Stage one Long Text field's current value before an update or delete. Called by the generated BeforeChange / BeforeDelete Data Macros.

### BackupAndRemoveAllDataMacros

```vba
Public Function BackupAndRemoveAllDataMacros(Optional strBackupPath As String = "", _
                                             Optional bSilent As Boolean = False) As Boolean
```

Back up every table's data macros, then remove the ones this generator wrote. Passing bSilent:=True suppresses both message boxes, so a caller with no one at the keyboard does not hang on a dialog. The Boolean return is the result.

### DumpTableMacros

```vba
Public Function DumpTableMacros(ByVal sTable As String) As String
```

Read a table's attached Data Macro set back out, so a build can be verified against the table itself rather than the generator's own report.

### ListMacroEvents

```vba
Public Function ListMacroEvents(ByVal sTable As String) As String
```

Just the Event= names, in order — a one-line structural check that a table carries the macros the generator said it created.
