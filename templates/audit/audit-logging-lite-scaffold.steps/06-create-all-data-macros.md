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
the three After macros whenever anything is being audited, BeforeChange whenever the table carries
the house audit columns or an auditable Long Text field, BeforeDelete only for Long Text.

**The insert-blocking trap this guards against.** Every table in scope gets its **audit-column
stamping** macro, including a table with no auditable fields. A table with nothing worth auditing
still has records created and changed on it, so it still needs the who-and-when stamp.
`CreatedDate` and `CreatedBy` are `Required`, and no column default can reach the Windows username,
so a table left without that macro **would reject every insert**. Nobody could add a record to it
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
Private Function CreateAllDataMacros(sTableName As String, fieldList As Collection, sTempPath As String, Optional bSilent As Boolean = False) As String
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
    Dim rsCheck As DAO.Recordset
    Dim sXmlContent As String
    Dim sBeforeChange As String
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim txtFile As Object
    Dim sFilePath As String
    Dim sPrimaryKeyField As String
    Dim fieldInfo As Variant
    Dim bHasLongText As Boolean
    Dim bHasAuditColumns As Boolean
    Dim lAuditableCount As Long
    Dim lMacroCount As Long
    Dim sWhatWasBuilt As String
    Dim bHadExistingMacros As Boolean
    Dim sBackupFolder As String
    Dim sBackupTemp As String
    Dim sBackupFinal As String
    Dim sBackupNote As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' Find the PK field, count auditable fields, and detect auditable Long Text
    ' (schema Business Rules 2, 4, 5)
    bHasLongText = False
    bHasAuditColumns = False
    lAuditableCount = 0
    For Each fieldInfo In fieldList
        If fieldInfo(2) = True Then sPrimaryKeyField = fieldInfo(0)
        If fieldInfo(3) = True Then
            lAuditableCount = lAuditableCount + 1
            If fieldInfo(1) = dbMemo Then bHasLongText = True
        End If
        ' [SCAFFOLD] Presence of the tracking columns, for the report only — the same test
        '            BuildBeforeChangeMacro makes when it decides whether to emit stamping.
        '            Outside the IsAuditable test above, because tracking columns are always
        '            seeded off and would never be seen there.
        If StrComp(fieldInfo(0), AUDIT_CREATED_DATE, vbTextCompare) = 0 _
            Or StrComp(fieldInfo(0), AUDIT_CREATED_BY, vbTextCompare) = 0 _
            Or StrComp(fieldInfo(0), AUDIT_MODIFIED_DATE, vbTextCompare) = 0 _
            Or StrComp(fieldInfo(0), AUDIT_MODIFIED_BY, vbTextCompare) = 0 Then bHasAuditColumns = True
    Next fieldInfo

    ' BeforeChange carries the audit-column stamping as well as Long Text staging, so it is
    ' built for every table; it comes back "" only when neither job applies to this one.
    sBeforeChange = BuildBeforeChangeMacro(sTableName, fieldList, sPrimaryKeyField)

    ' [SCAFFOLD] THE TRAP THIS GUARDS AGAINST. Every table in scope gets the audit-column
    '            stamping macro, including a table with nothing auditable: records are still
    '            created and changed on it, so it still needs the who-and-when stamp.
    '            CreatedDate and CreatedBy are Required, with no default able to reach the
    '            username, so a table left without that macro would reject EVERY insert.
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
    If lAuditableCount = 0 And Len(sBeforeChange) = 0 Then
        Dim sRemovalResult As String
        sRemovalResult = RemoveDataMacrosForTable(sTableName)
        If Left$(sRemovalResult, 2) = "OK" Then
            Debug.Print "  - Removed (nothing to audit or stamp; a prior run's macros were still attached)"
            CreateAllDataMacros = "REMOVED (nothing to audit or stamp; previous macros removed) - " & sRemovalResult
        Else
            Debug.Print "  - Skipped (nothing to audit, and no audit columns to stamp)"
            CreateAllDataMacros = "SKIPPED (nothing to audit, no audit columns to stamp)"
        End If
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
    sBackupTemp = ""
    sBackupFinal = ""
    If bHadExistingMacros Then
        sBackupFolder = CurrentProject.Path & "\DataMacroBackups\"
        If Dir(sBackupFolder, vbDirectory) = "" Then MkDir sBackupFolder
        ' [SCAFFOLD] Export under a working name. Taking the copy first is not optional —
        '            after LoadFromText there is nothing left to copy — but the file is only
        '            renamed to its real name once the new macros have actually loaded, and
        '            deleted if they have not. That is what makes a file in DataMacroBackups\
        '            mean a replacement really happened. Before this, a run stopped by an
        '            open object left behind a backup for a table it never touched.
        sBackupTemp = sBackupFolder & sTableName & "_backup-in-progress_" & _
            Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Application.SaveAsText acTableDataMacro, sTableName, sBackupTemp
        ' [SCAFFOLD] Whose macros are these? A set naming tblAuditLog is one this generator
        '            wrote on an earlier run, and calling that file a PRE-AUDIT backup would
        '            be false: restore it and you get this system back, not what the table
        '            had before any of it. Regenerating is ordinary, so this case is common.
        If MacroBackupIsOurs(sBackupTemp) Then
            sBackupFinal = sBackupFolder & sTableName & "_PreRegenBackup_" & _
                Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        Else
            sBackupFinal = sBackupFolder & sTableName & "_PreAuditBackup_" & _
                Format(Now(), "yyyymmdd_hhnnss") & ".xml"
        End If
    End If

    ' One XML document carrying all of this table's macros
    sXmlContent = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    sXmlContent = sXmlContent & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application"">"

    ' The three After macros are the audit trail itself — no auditable fields, none needed.
    lMacroCount = 0
    If lAuditableCount > 0 Then
        sXmlContent = sXmlContent & BuildAfterInsertMacro(sTableName, fieldList, sPrimaryKeyField)
        sXmlContent = sXmlContent & BuildAfterUpdateMacro(sTableName, fieldList, sPrimaryKeyField)
        sXmlContent = sXmlContent & BuildAfterDeleteMacro(sTableName, fieldList, sPrimaryKeyField)
        lMacroCount = 3
    End If

    If Len(sBeforeChange) > 0 Then
        sXmlContent = sXmlContent & sBeforeChange
        lMacroCount = lMacroCount + 1
    End If

    ' BeforeDelete stages Long Text values for the AfterDelete log row, so it is only
    ' wanted when that log row is actually going to be written.
    If bHasLongText Then
        sXmlContent = sXmlContent & BuildBeforeDeleteMacro(sTableName, fieldList, sPrimaryKeyField)
        lMacroCount = lMacroCount + 1
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

    ' [SCAFFOLD] This is the line that fails when something in the database is still open —
    '            Access cannot take a design lock on a table another object is using. The
    '            error is caught below, the table keeps the macros it already had, and the
    '            caller counts it as failed rather than built.
    DoCmd.OpenTable sTableName, acViewDesign, acHidden
    Application.LoadFromText acTableDataMacro, sTableName, sFilePath

    ' [SCAFFOLD] The load has happened, so the copy taken earlier is now a true record of what
    '            was replaced and earns its real name — promoted HERE, on the line after the
    '            load, and not further down. Everything below this point runs with the table
    '            already replaced, so a backup still carrying its working name past this line
    '            would be thrown away by Cleanup while the macros it recorded were already
    '            gone. Clearing sBackupTemp is also what tells Cleanup there is nothing left
    '            to throw away.
    If Len(sBackupTemp) > 0 Then
        Name sBackupTemp As sBackupFinal
        sBackupTemp = ""
        sBackupNote = " (existing macros backed up to " & sBackupFinal & " before replacing)"
        Debug.Print "  - " & sTableName & " already had Data Macros — backed up to " & sBackupFinal
    End If

    DoCmd.Close acTable, sTableName, acSaveYes

    fso.DeleteFile sFilePath

    ' [SCAFFOLD] Report what was actually built rather than a fixed count — the count varies
    '            by table now, and a reader comparing tables needs to see why.
    If lAuditableCount = 0 Then
        sWhatWasBuilt = "BeforeChange only — audit-column stamping, no audit trail " & _
            "(this table has no auditable fields; inserts still work)"
    Else
        sWhatWasBuilt = "3 After"
        If Len(sBeforeChange) > 0 Then
            If bHasAuditColumns And bHasLongText Then
                sWhatWasBuilt = sWhatWasBuilt & " + BeforeChange (stamping and Long Text staging)"
            ElseIf bHasAuditColumns Then
                sWhatWasBuilt = sWhatWasBuilt & " + BeforeChange (stamping)"
            Else
                sWhatWasBuilt = sWhatWasBuilt & " + BeforeChange (Long Text staging only — " & _
                    "no tracking columns found on this table, so nothing is stamped)"
            End If
        Else
            sWhatWasBuilt = sWhatWasBuilt & " (no tracking columns found on this table, " & _
                "so nothing is stamped)"
        End If
        If bHasLongText Then sWhatWasBuilt = sWhatWasBuilt & " + BeforeDelete"
    End If

    Debug.Print "  - " & lMacroCount & " data macro(s) created: " & sWhatWasBuilt
    CreateAllDataMacros = "OK — " & lMacroCount & " macro(s): " & sWhatWasBuilt & sBackupNote

Cleanup:
    ' [SCAFFOLD] A working-name export still sitting here means the load never happened — the
    '            promotion above runs on the line straight after LoadFromText, so no failure
    '            past that point can leave one behind. Nothing was replaced, the file records
    '            nothing that was lost, and throwing it away is what keeps the rule that a file
    '            in DataMacroBackups\ means a replacement happened. On the success path
    '            sBackupTemp was cleared above and there is nothing to do.
    If Len(sBackupTemp) > 0 Then
        On Error Resume Next
        Kill sBackupTemp
    End If
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
    Resume
End Function
```
