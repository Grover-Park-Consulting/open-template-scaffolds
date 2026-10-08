---
step: 05-generator
title: "The generator: its entry point and the macro backup"
platform_facts: [vba-import-xml-entities, table-in-use]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures (continued)

### Four_GenerateAllAuditDataMacros — `Public Function` → `String` (setup step 4)

Reads the (reviewed) config, groups fields by table, and calls `CreateAllDataMacros` for each.
Re-runnable: reloading a table's macro XML replaces what was there (schema Business Rule 7).

**Returns a per-table report as text**, in addition to the summary `MsgBox` — one line per table
(`OK`, `SKIPPED`, or `ERROR: ...`), the same detail `CreateAllDataMacros` sends to `Debug.Print`.
Call it as `sResult = Four_GenerateAllAuditDataMacros(True)` to read that report directly with no
dialog, so a script or an AI assistant facilitating the build can see which tables actually
succeeded without adding a diagnostic wrapper of its own. `bSilent` is passed down into
`CreateAllDataMacros`, so a per-table error can't strand an automated caller either.

**The summary counts outcomes, not attempts.** A table that fails is caught inside
`CreateAllDataMacros` and the loop carries on to the next one — which is the right behaviour, since
one blocked table should not cost the other five. The cost of it is that the number of tables
*processed* says nothing about how many were *built*, and a run where five of six tables were held
open would once report "Generated audit data macros for 6 table(s)" and mean it. The summary now
reads `6 table(s) processed: 1 built, 0 skipped, 5 failed`, the dialog carries a warning icon rather
than an information one, and on any failure it says what state those tables are left in and what to
do about it. **A run with nothing to do stops before the loop** rather than reporting `0 table(s)
processed` as though it had finished, and says whether the cause is switches left off or nothing in
scope at all. **Access raises its own error first** — the developer sees a lock failure on the table
before they ever reach this summary — so this is the message that has to agree with what they were
just told, not the one that breaks the news.

**It enforces the two unauditable field types rather than trusting the config table.** The switches
were seeded in step 3, and between step 3 and step 4 the developer is expected to go and edit them;
`IsUnauditableFieldType` is asked again here so an Attachment field switched back on cannot produce
a macro set that does not work.

```vba
Public Function Four_GenerateAllAuditDataMacros(Optional bSilent As Boolean = False) As String
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
    Dim fldLive As DAO.Field
    Dim sTempPath As String
    Dim lTableCount As Long
    Dim lBuiltCount As Long
    Dim lSkipCount As Long
    Dim lFailCount As Long
    Dim vCurrentTable As Variant
    Dim sTableResult As String
    Dim sSummary As String
    Dim sReport As String
    Dim sOpen As String

    On Error GoTo errHandler
    Set db = CurrentDb
    Set dictTables = CreateObject("Scripting.Dictionary")
    sReport = ""

    ' [SCAFFOLD] This is the step that actually fails on an open object, because it opens each
    '            table in design view to attach the macros. Stopping here costs one message;
    '            carrying on costs a partial run that has to be diagnosed from Access's own
    '            lock error. Asked again rather than relying on One_CheckAuditReadiness, which
    '            may have run some time ago and is optional on Path A.
    sOpen = ListOpenObjects()
    If Len(sOpen) > 0 Then
        sReport = "Stopped before anything was changed. Something in this database is " & _
            "still open:" & vbCrLf & vbCrLf & sOpen & vbCrLf & _
            "Close it and run this again. Everything has to stay closed until the whole " & _
            "run is finished, not just when it starts."
        If Not bSilent Then MsgBox sReport, vbExclamation, "Close everything first"
        Four_GenerateAllAuditDataMacros = sReport
        GoTo Cleanup
    End If

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
            ' [SCAFFOLD] Hard guard above the flags, on the field this time. tblAuditLogConfig
            '            exists to be edited, so a switch seeded off in step 3 may well be on
            '            by the time this runs — and for two field types that is not a choice
            '            the developer gets to make (see IsUnauditableFieldType). Ask the same
            '            function step 3 asked rather than trusting the row. A field that has
            '            since been deleted leaves fldLive Nothing and is simply not enforced;
            '            the builders skip a field they cannot resolve anyway.
            If bFieldIsAuditable Then
                Set fldLive = Nothing
                On Error Resume Next
                Set fldLive = db.TableDefs(sTableName).Fields(sFieldName)
                On Error GoTo errHandler
                If Not fldLive Is Nothing Then
                    If IsUnauditableFieldType(fldLive) Then bFieldIsAuditable = False
                End If
            End If

            ' Field info as array: (FieldName, DataType, IsPrimaryKey, IsAuditable)
            fieldInfo = Array(sFieldName, lFieldDataType, bFieldIsPK, bFieldIsAuditable)
            fieldList.Add fieldInfo
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' [SCAFFOLD] An empty config table is a stop. "0 table(s) processed: 0 built" is accurate
    '            and still reads as a run that finished, which is the reading that lets a
    '            build with nothing in it look like a build.
    If dictTables.Count = 0 Then
        sReport = "Stopped. There is nothing to build - no table in tblAuditLogConfig has " & _
            "any field switched on." & vbCrLf & vbCrLf & _
            "Run Three_PopulateConfigTable, then open tblAuditLogConfig and switch on the " & _
            "fields you want a history of. If that table is empty rather than switched " & _
            "off, no table in this file was in scope: check the scope setting at the top " & _
            "of this module. Nothing has been changed."
        If Not bSilent Then MsgBox sReport, vbExclamation, "Nothing to build"
        Four_GenerateAllAuditDataMacros = sReport
        GoTo Cleanup
    End If

    sTempPath = Environ("TEMP") & "\"

    lTableCount = 0
    lBuiltCount = 0
    lSkipCount = 0
    lFailCount = 0
    For Each vCurrentTable In dictTables.Keys
        sTableName = CStr(vCurrentTable)
        Set fieldList = dictTables(sTableName)
        sTableResult = CreateAllDataMacros(sTableName, fieldList, sTempPath, bSilent)
        sReport = sReport & sTableName & ": " & sTableResult & vbCrLf
        lTableCount = lTableCount + 1
        ' [SCAFFOLD] Count what actually happened, not how many tables were attempted. A
        '            per-table error is caught inside CreateAllDataMacros and the loop
        '            carries on, which is right — one blocked table should not cost the
        '            other five. But it means the number of tables processed says nothing
        '            about how many were built, and a summary reporting the first as though
        '            it were the second tells the developer a partial run succeeded.
        Select Case True
            Case Left$(sTableResult, 5) = "ERROR"
                lFailCount = lFailCount + 1
            Case Left$(sTableResult, 7) = "SKIPPED"
                lSkipCount = lSkipCount + 1
            Case Else
                lBuiltCount = lBuiltCount + 1
        End Select
    Next vCurrentTable

    sSummary = lTableCount & " table(s) processed: " & lBuiltCount & " built"
    If lSkipCount > 0 Then sSummary = sSummary & ", " & lSkipCount & " skipped"
    sSummary = sSummary & ", " & lFailCount & " failed."
    ' [SCAFFOLD] On a failure, say what state the failed tables are in and what to do — the
    '            developer is reading this in a dialog, not in the report, and the commonest
    '            cause by a distance is an object left open in the database.
    If lFailCount > 0 Then
        sSummary = sSummary & vbCrLf & vbCrLf & _
            "A table that failed almost always still has the Data Macros it had before, and " & _
            "where it had any, that set is saved in the DataMacroBackups folder beside this " & _
            "database. The usual cause is something left open: close every table, form, " & _
            "report and query in this database, then run this again. It is safe to run again."
    End If
    sReport = sSummary & vbCrLf & vbCrLf & sReport

    If Not bSilent Then MsgBox sSummary, IIf(lFailCount > 0, vbExclamation, vbInformation)
    Four_GenerateAllAuditDataMacros = sReport

Cleanup:
    Set fldLive = Nothing
    Set rs = Nothing
    Set db = Nothing
    Set dictTables = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Four_GenerateAllAuditDataMacros = sReport & "ERROR: " & Err.Number & " - " & Err.Description
    If Not bSilent Then MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
    Resume
End Function
```

### MacroBackupIsOurs — `Public Function` → `Boolean`

**Answers one question: did this generator write that macro set?** Given a macro set just exported
from a table, it looks for the mark every macro this generator emits carries in a comment.

**A table name cannot answer this, which is why the mark exists.** A database may have an audit
table of its own called anything at all, including the same name this template uses, and a macro
that mentions a table name says nothing about who wrote it. Deciding by name would eventually mean
stripping somebody else's macros. The mark is written by this generator and by nothing else, so a
false match is not something a developer can arrive at by accident.

**Two callers rely on the answer, which is why it is `Public` and not `Private`.**
`CreateAllDataMacros` uses it to name a backup file: a set of ours is filed as a `PreRegenBackup`,
anything else as a `PreAuditBackup`. `BackupAndRemoveAllDataMacros` uses it to decide what it is
allowed to strip, and lives in a different module — a `Private` function would not compile there.

**It reads the file as Unicode** — the `-1` argument — because `SaveAsText` writes UTF-16, and
reading it as plain text returns characters separated by nulls and finds nothing. **If the file
cannot be read the answer is `False`**, and `False` is the safe direction for both callers: it
produces the pre-audit name, which never claims a file holds less than it does, and it leaves a
macro set in place rather than stripping it.

```vba
Public Function MacroBackupIsOurs(sBackupPath As String) As Boolean
    ' [SCAFFOLD] Recognition only. This never backs anything up. It answers whether a macro
    '            set is one this generator wrote; the callers decide what to do about it.
    '            The test is the mark every generated macro carries in a <Comment>, never a
    '            table name: a host database may have its own audit table under any name,
    '            and stripping someone else's macros because they mention one is exactly
    '            the failure this guards against.
    '            Matched on AUDIT_MACRO_MARKER, the part before the version, so a set
    '            written by an earlier release is still recognised as ours. Matched
    '            case-sensitively, because a mark is exact and a near miss is not our work.
    '            No errHandler and therefore no line numbers: an unreadable file is an
    '            answer, not a failure, and the caller has a backup to name either way.
    Dim fso As Object                 ' Scripting.FileSystemObject, late-bound
    Dim tsBackup As Object
    Dim sContent As String

    sContent = ""
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 1 = ForReading, -1 = TristateTrue (Unicode) — SaveAsText writes UTF-16.
    Set tsBackup = fso.OpenTextFile(sBackupPath, 1, False, -1)
    sContent = tsBackup.ReadAll
    tsBackup.Close
    On Error GoTo 0

    Set tsBackup = Nothing
    Set fso = Nothing

    MacroBackupIsOurs = (InStr(1, sContent, AUDIT_MACRO_MARKER, vbBinaryCompare) > 0)
End Function
```

### RemoveDataMacrosForTable — `Public Function` → `String`

**The per-table sibling of `BackupAndRemoveAllDataMacros`.** `CreateAllDataMacros` calls this when a
table has nothing to build — no auditable fields and no audit columns to stamp — but already carries
macros this generator wrote on an earlier run: rather than leaving that table's audit trail running
under a stale configuration, this strips it and reports `REMOVED` instead of `SKIPPED`. A table that
has never carried this generator's macros is left alone and reported `SKIPPED`, unchanged.

**Same recognition, same backup, same empty-document technique as `BackupAndRemoveAllDataMacros` —
just scoped to one table instead of every table in the database.** It exists as its own callable
function, not only as an internal helper of `CreateAllDataMacros`, because it is also the right tool
for a developer who wants to strip one table's audit macros by hand without touching any other
table — the way `BackupAndRemoveAllDataMacros` is the tool for stripping all of them.

```vba
Public Function RemoveDataMacrosForTable(ByVal sTableName As String) As String
    ' [SCAFFOLD] Strip one table's Data Macros, but only if they are this generator's own —
    '            same recognition and backup technique as BackupAndRemoveAllDataMacros, scoped
    '            to a single table. Returns "OK ...", "SKIPPED ...", or "ERROR: ...".
    Dim db As DAO.Database
    Dim rsCheck As DAO.Recordset
    Dim bHasMacros As Boolean
    Dim sBackupFolder As String
    Dim sBackupPath As String
    Dim sEmptyPath As String
    Dim fso As Object
    Dim txtFile As Object

    On Error GoTo errHandler
    Set db = CurrentDb

    Set rsCheck = db.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Name='" & sTableName & "' AND Type=1 AND Not IsNull(LvExtra)", _
        dbOpenSnapshot)
    bHasMacros = Not rsCheck.EOF
    rsCheck.Close
    Set rsCheck = Nothing

    If Not bHasMacros Then
        RemoveDataMacrosForTable = "SKIPPED - " & sTableName & " carries no Data Macros"
        GoTo Cleanup
    End If

    sBackupFolder = CurrentProject.Path & "\DataMacroBackups\"
    If Dir(sBackupFolder, vbDirectory) = "" Then MkDir sBackupFolder
    sBackupPath = sBackupFolder & sTableName & "_PreRemoval_" & Format(Now(), "yyyymmdd_hhnnss") & ".xml"
    Application.SaveAsText acTableDataMacro, sTableName, sBackupPath

    ' [SCAFFOLD] Same guard as BackupAndRemoveAllDataMacros: a macro set that is not this
    '            generator's own is backed up (so the developer has a record) and left in
    '            place. Business logic of the developer's own is not this tool's to throw away.
    If Not MacroBackupIsOurs(sBackupPath) Then
        RemoveDataMacrosForTable = "SKIPPED - " & sTableName & _
            "'s Data Macros are not this generator's own work (backed up to " & sBackupPath & _
            " and left in place)"
        GoTo Cleanup
    End If

    sEmptyPath = Environ("TEMP") & "\" & sTableName & "_EmptyDataMacros.xml"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(sEmptyPath, True, True)
    txtFile.Write "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>" & _
        "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application""></DataMacros>"
    txtFile.Close
    Set txtFile = Nothing

    DoCmd.OpenTable sTableName, acViewDesign, acHidden
    Application.LoadFromText acTableDataMacro, sTableName, sEmptyPath
    DoCmd.Close acTable, sTableName, acSaveYes
    fso.DeleteFile sEmptyPath

    RemoveDataMacrosForTable = "OK - Data Macros removed from " & sTableName & " (backed up to " & sBackupPath & ")"

Cleanup:
    On Error Resume Next
    Set rsCheck = Nothing
    Set txtFile = Nothing
    Set fso = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    RemoveDataMacrosForTable = "ERROR: " & Err.Number & " - " & Err.Description
    Resume Cleanup
End Function
```
