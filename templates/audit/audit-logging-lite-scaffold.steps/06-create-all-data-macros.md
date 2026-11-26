---
step: 06-create-all-data-macros
title: "CreateAllDataMacros: writing and loading each table's macros"
platform_facts: [data-macro-rules, vba-import-xml-entities]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures (continued)

### CreateAllDataMacros — `Private Function` → `String`

Builds one table's macro XML into a single document, writes it UTF-16, and loads it with
`LoadFromText acTableDataMacro`. **How many macros a table gets varies** (schema Business Rule 2) —
the three After macros whenever anything is being audited, BeforeChange whenever the generator
stamps the table's house audit columns (`AuditStampList`) or it has an auditable Long Text field,
BeforeDelete only for Long Text.

**The insert-blocking trap this guards against.** Every table in `AuditStampList` gets its
**audit-column stamping** macro, including a table with no auditable fields. A table with nothing
worth auditing still has records created and changed on it, so it still needs the who-and-when stamp.
On a listed table nothing else fills `CreatedDate` and `CreatedBy`, they are `Required`, and no column
default can reach the Windows username, so a listed table left without that macro **would reject
every insert**. Nobody could add a record to it
through any interface. Switching auditing off for a
table is the ordinary Path B workflow, so it must never break writing to that table. The audit
actions are skipped; the stamping macro is still written. Only a table needing neither is skipped.

**Backs up a table's existing macros before replacing them.** `LoadFromText` replaces a table's
entire Data Macro set — it does not merge. That is why the house stamping and the change auditing
are generated together in one macro: written as two, whichever loaded second would replace the
first. What
remains is any **other** Data Macro a table carries — business logic written for reasons unrelated to
this system, which a Path B table may well have. Before loading, this checks `MSysObjects` for an
existing macro set and, if it finds one, exports it to a timestamped backup file first — the same
technique `BackupAndRemoveAllDataMacros` uses to detect a table's macros. It does **not** merge the
old logic into the new macros; it only makes sure nothing is destroyed without a copy and a
plain-language warning first. Re-implementing anything lost is the developer's call.

**A backup file is written only when a replacement actually happened.** The copy has to be taken
before `LoadFromText` — afterwards there is nothing left to copy — but it is written under a working
name and renamed to its real one only once the new macros have loaded. A run stopped part-way, by an
object left open in the database, deletes its working file instead: that table was not changed, so a
file claiming to hold what was replaced on it would be false. This matters because the folder is
where a developer goes to recover something, and a backup for a table that was never touched sends
them looking for a change that never happened.

**The file name says which kind of backup it is.** `<Table>_PreAuditBackup_<timestamp>.xml` holds
what the table had before this system was ever put on it. `<Table>_PreRegenBackup_<timestamp>.xml`
holds a set this generator wrote on an earlier run — restoring that one gives you this system back,
not the table's original macros, and regenerating is ordinary enough that the two would otherwise be
indistinguishable in the folder. `MacroBackupIsOurs` tells them apart by the mark every macro this
generator writes carries in a comment, so a set of ours is recognised whatever the host's tables are
called and whatever is switched on for the table, including a table that gets a stamping macro and
nothing else.

**`DataMacroBackups\` grows every time you regenerate, and nothing prunes it.** Once a table carries
macros, *every* subsequent run backs it up again — so a schema you regenerate ten times leaves ten
files per table beside the database. That is deliberate: throwing away the only copy of a macro set
to save a few kilobytes is the wrong trade. But it does mean the folder is yours to clear out, and
the newest file for a table is the one that matters.

```vba
Private Function CreateAllDataMacros(strTableName As String, fieldList As Collection, strTempPath As String, Optional booSilent As Boolean = False) As String
    ' [SCAFFOLD] Generate one table's Data Macros as a SINGLE XML document and attach it.
    '            How many macros that is depends on two independent things (schema
    '            Business Rule 2):
    '              - the 3 After macros, whenever the table has any auditable field;
    '              - BeforeChange, whenever the table carries the house audit columns OR
    '                has an auditable Long Text field (that one macro does both jobs);
    '              - BeforeDelete, only for an auditable Long Text field.
    '            Returns a one-line status ("OK ...", "SKIPPED ...", or "ERROR: ...") so the
    '            caller can report per-table results without relying on Debug.Print alone.
    Dim db As DAO.Database
    Dim rstCheck As DAO.Recordset
    Dim strXmlContent As String
    Dim strBeforeChange As String
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim txtFile As Object
    Dim strFilePath As String
    Dim strPrimaryKeyField As String
    Dim fieldInfo As Variant
    Dim booHasLongText As Boolean
    Dim booHasAuditColumns As Boolean
    Dim lngAuditableCount As Long
    Dim lngMacroCount As Long
    Dim strWhatWasBuilt As String
    Dim booHadExistingMacros As Boolean
    Dim strBackupFolder As String
    Dim strBackupTemp As String
    Dim strBackupFinal As String
    Dim strBackupNote As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' Find the PK field, count auditable fields, and detect auditable Long Text
    ' (schema Business Rules 2, 4, 5)
    booHasLongText = False
    booHasAuditColumns = False
    lngAuditableCount = 0
    For Each fieldInfo In fieldList
        If fieldInfo(2) = True Then strPrimaryKeyField = fieldInfo(0)
        If fieldInfo(3) = True Then
            lngAuditableCount = lngAuditableCount + 1
            If fieldInfo(1) = dbMemo Then booHasLongText = True
        End If
        ' [SCAFFOLD] Whether this table is stamped, for the report only — the same test
        '            BuildBeforeChangeMacro makes when it decides whether to emit stamping:
        '            listed in AuditStampList and carrying the tracking columns.
        '            Outside the IsAuditable test above, because tracking columns are always
        '            seeded off and would never be seen there.
        If IsNamedInList(AuditStampList, strTableName) Then
            If StrComp(fieldInfo(0), AuditCreatedDate, vbTextCompare) = 0 _
                Or StrComp(fieldInfo(0), AuditCreatedBy, vbTextCompare) = 0 _
                Or StrComp(fieldInfo(0), AuditModifiedDate, vbTextCompare) = 0 _
                Or StrComp(fieldInfo(0), AuditModifiedBy, vbTextCompare) = 0 Then booHasAuditColumns = True
        End If
    Next fieldInfo

    ' BeforeChange carries the audit-column stamping as well as Long Text staging, so it is
    ' built for every table; it comes back "" only when neither job applies to this one.
    strBeforeChange = BuildBeforeChangeMacro(strTableName, fieldList, strPrimaryKeyField)

    ' [SCAFFOLD] THE TRAP THIS GUARDS AGAINST. Every table in AuditStampList gets the
    '            audit-column stamping macro, including one with nothing auditable: records
    '            are still created and changed on it, so it still needs the who-and-when
    '            stamp. On a listed table nothing else fills CreatedDate and CreatedBy, they
    '            are Required, and no default can reach the username, so a listed table
    '            left without that macro would reject EVERY insert.
    '            Turning auditing off for a table is the normal Path B workflow, so it must
    '            never break writing to that table: the audit actions are skipped, the
    '            stamping macro is still emitted. Only a table that needs neither is skipped.
    '
    ' [SCAFFOLD] "Nothing to build" is not the same question as "nothing to do." A table can
    '            arrive here with every field just switched off and no audit columns, while
    '            still carrying the macros a PRIOR run attached when some field was on.
    '            Exiting unconditionally would leave that table auditing exactly as before,
    '            while reporting SKIPPED — which reads as "no change" when the true state is
    '            "still logging." So: before reporting SKIPPED, check whether this table
    '            already carries OUR OWN macros, and if it does, remove them and say so.
    If lngAuditableCount = 0 And Len(strBeforeChange) = 0 Then
        Dim strRemovalResult As String
        strRemovalResult = RemoveDataMacrosForTable(strTableName)
        If Left$(strRemovalResult, 2) = "OK" Then
            Debug.Print "  - Removed (nothing to audit or stamp; a prior run's macros were still attached)"
            CreateAllDataMacros = "REMOVED (nothing to audit or stamp; previous macros removed) - " & strRemovalResult
        Else
            Debug.Print "  - Skipped (nothing to audit, and no audit columns to stamp)"
            CreateAllDataMacros = "SKIPPED (nothing to audit, no audit columns to stamp)"
        End If
        Exit Function
    End If

    ' [SCAFFOLD] Safety net for Path B: LoadFromText replaces a table's WHOLE macro set. If
    '            this table already has one (e.g. the house audit-column stamping macro),
    '            back it up before it's overwritten — see the note above this code block.
    booHadExistingMacros = False
    Set rstCheck = db.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Name='" & strTableName & "' AND Type=1 AND Not IsNull(LvExtra)", _
        dbOpenSnapshot)
    booHadExistingMacros = Not rstCheck.EOF
    rstCheck.Close
    Set rstCheck = Nothing

    strBackupNote = ""
    strBackupTemp = ""
    strBackupFinal = ""
    If booHadExistingMacros Then
        strBackupFolder = CurrentProject.Path & "\DataMacroBackups\"
        If Dir(strBackupFolder, vbDirectory) = "" Then MkDir strBackupFolder
        ' [SCAFFOLD] Export under a working name. Taking the copy first is not optional —
        '            after LoadFromText there is nothing left to copy — but the file is only
        '            renamed to its real name once the new macros have actually loaded, and
        '            deleted if they have not. That is what makes a file in DataMacroBackups\
        '            mean a replacement really happened. Before this, a run stopped by an
        '            open object left behind a backup for a table it never touched.
        strBackupTemp = strBackupFolder & strTableName & "_backup-in-progress_" & _
            Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Application.SaveAsText acTableDataMacro, strTableName, strBackupTemp
        ' [SCAFFOLD] Whose macros are these? A set naming tblAuditLog is one this generator
        '            wrote on an earlier run, and calling that file a PRE-AUDIT backup would
        '            be false: restore it and you get this system back, not what the table
        '            had before any of it. Regenerating is ordinary, so this case is common.
        If MacroBackupIsOurs(strBackupTemp) Then
            strBackupFinal = strBackupFolder & strTableName & "_PreRegenBackup_" & _
                Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Else
            strBackupFinal = strBackupFolder & strTableName & "_PreAuditBackup_" & _
                Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        End If
    End If

    ' One XML document carrying all of this table's macros
    strXmlContent = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    strXmlContent = strXmlContent & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application"">"

    ' The three After macros are the audit trail itself — no auditable fields, none needed.
    lngMacroCount = 0
    If lngAuditableCount > 0 Then
        strXmlContent = strXmlContent & BuildAfterInsertMacro(strTableName, fieldList, strPrimaryKeyField)
        strXmlContent = strXmlContent & BuildAfterUpdateMacro(strTableName, fieldList, strPrimaryKeyField)
        strXmlContent = strXmlContent & BuildAfterDeleteMacro(strTableName, fieldList, strPrimaryKeyField)
        lngMacroCount = 3
    End If

    If Len(strBeforeChange) > 0 Then
        strXmlContent = strXmlContent & strBeforeChange
        lngMacroCount = lngMacroCount + 1
    End If

    ' BeforeDelete stages Long Text values for the AfterDelete log row, so it is only
    ' wanted when that log row is actually going to be written.
    If booHasLongText Then
        strXmlContent = strXmlContent & BuildBeforeDeleteMacro(strTableName, fieldList, strPrimaryKeyField)
        lngMacroCount = lngMacroCount + 1
    End If

    strXmlContent = strXmlContent & "</DataMacros>"

    ' [SCAFFOLD] Write UTF-16 (CreateTextFile third argument True) — LoadFromText requires it —
    '            then load with the table held open in design view so the save sticks.
    strFilePath = strTempPath & strTableName & "_DataMacros.xml"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(strFilePath, True, True)
    txtFile.Write strXmlContent
    txtFile.Close
    Set txtFile = Nothing

    ' [SCAFFOLD] This is the line that fails when something in the database is still open —
    '            Access cannot take a design lock on a table another object is using. The
    '            error is caught below, the table keeps the macros it already had, and the
    '            caller counts it as failed rather than built.
    DoCmd.OpenTable strTableName, acViewDesign, acHidden
    Application.LoadFromText acTableDataMacro, strTableName, strFilePath

    ' [SCAFFOLD] The load has happened, so the copy taken earlier is now a true record of what
    '            was replaced and earns its real name — promoted HERE, on the line after the
    '            load, and not further down. Everything below this point runs with the table
    '            already replaced, so a backup still carrying its working name past this line
    '            would be thrown away by Cleanup while the macros it recorded were already
    '            gone. Clearing strBackupTemp is also what tells Cleanup there is nothing left
    '            to throw away.
    If Len(strBackupTemp) > 0 Then
        Name strBackupTemp As strBackupFinal
        strBackupTemp = ""
        strBackupNote = " (existing macros backed up to " & strBackupFinal & " before replacing)"
        Debug.Print "  - " & strTableName & " already had Data Macros — backed up to " & strBackupFinal
    End If

    DoCmd.Close acTable, strTableName, acSaveYes

    fso.DeleteFile strFilePath

    ' [SCAFFOLD] Report what was actually built rather than a fixed count — the count varies
    '            by table now, and a reader comparing tables needs to see why.
    If lngAuditableCount = 0 Then
        strWhatWasBuilt = "BeforeChange only — audit-column stamping, no audit trail " & _
            "(this table has no auditable fields; inserts still work)"
    Else
        strWhatWasBuilt = "3 After"
        If Len(strBeforeChange) > 0 Then
            If booHasAuditColumns And booHasLongText Then
                strWhatWasBuilt = strWhatWasBuilt & " + BeforeChange (stamping and Long Text staging)"
            ElseIf booHasAuditColumns Then
                strWhatWasBuilt = strWhatWasBuilt & " + BeforeChange (stamping)"
            Else
                strWhatWasBuilt = strWhatWasBuilt & " + BeforeChange (Long Text staging only — " & _
                    "this table's tracking columns are not stamped here: none, or filled elsewhere)"
            End If
        Else
            strWhatWasBuilt = strWhatWasBuilt & " (this table's tracking columns are not stamped " & _
                "here: none, or filled elsewhere)"
        End If
        If booHasLongText Then strWhatWasBuilt = strWhatWasBuilt & " + BeforeDelete"
    End If

    Debug.Print "  - " & lngMacroCount & " data macro(s) created: " & strWhatWasBuilt
    CreateAllDataMacros = "OK — " & lngMacroCount & " macro(s): " & strWhatWasBuilt & strBackupNote

Cleanup:
    ' [SCAFFOLD] A working-name export still sitting here means the load never happened — the
    '            promotion above runs on the line straight after LoadFromText, so no failure
    '            past that point can leave one behind. Nothing was replaced, the file records
    '            nothing that was lost, and throwing it away is what keeps the rule that a file
    '            in DataMacroBackups\ means a replacement happened. On the success path
    '            strBackupTemp was cleared above and there is nothing to do.
    If Len(strBackupTemp) > 0 Then
        On Error Resume Next
        Kill strBackupTemp
    End If
    Set rstCheck = Nothing
    Set txtFile = Nothing
    Set fso = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] reporting shape: entry point that returns its report.
    '            A caller reads the report and does not report again. Option 1 also calls
    '            LogError here.
    CreateAllDataMacros = "ERROR: " & Err.Number & " - " & Err.Description
    If Not booSilent Then MsgBox "Error creating macros for " & strTableName & ": " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
    Resume
End Function
```
