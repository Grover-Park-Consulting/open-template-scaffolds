---
step: 03-procedures-2
title: "The procedures, part 2 of 2 (from GameLevelID)"
platform_facts: [mcp-module-import, mcp-line-numbers, recordset-append-crash, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

### GameLevelID — `Public Function` → `Long`

```vba
Public Function GameLevelID(ByVal lngGameID As Long) As Long
    Const ProcName As String = "GameLevelID"
    ' [SCAFFOLD] A game's play level, read through its home team - never stored on the game
    '            itself (schema Business Rule 5). Returns 0 when the game or its home team
    '            can't be found.
    Dim db   As DAO.Database
    Dim rst   As DAO.Recordset
    Dim strSQL As String

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #5] tblGame -> tblTeam (on HomeTeamID) -> PlayLevelID
    ' >>> join query per query-style.md <<<
    strSQL = vbNullString

    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rst.EOF Then GameLevelID = Nz(rst!PlayLevelID, 0)

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step. Its 0 already means "not found",
    '            so a failure must not return one. In a query or a control it shows #Error instead.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set rst = Nothing: Set db = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### OfficialAge — `Public Function` → `Variant`

```vba
Public Function OfficialAge(ByVal lngOfficialID As Long) As Variant
    Const ProcName As String = "OfficialAge"
    ' [SCAFFOLD] An official's age, computed from BirthDate as of today - never stored
    '            anywhere (schema Business Rule 8). Returns Null when BirthDate is unknown,
    '            which is why this is a Variant rather than an Integer.
    Dim db      As DAO.Database
    Dim rst      As DAO.Recordset
    Dim strSQL    As String
    Dim dtBirth As Variant

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #8] look up tblOfficial.BirthDate for lngOfficialID
    ' >>> lookup query per query-style.md <<<
    strSQL = vbNullString

    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rst.EOF Then
        dtBirth = rst!BirthDate
        If Not IsNull(dtBirth) Then
            ' Subtract a year when this year's birthday hasn't happened yet.
            OfficialAge = DateDiff("yyyy", dtBirth, Date) - _
                IIf(Format(Date, "mmdd") < Format(dtBirth, "mmdd"), 1, 0)
        End If
    End If

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step. Its Null already means "birth
    '            date unknown", so a failure must not return one. In a query or a control it shows
    '            #Error instead.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set rst = Nothing: Set db = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### EnsurePhotoFolder — `Public Function` → `Boolean`

```vba
Public Function EnsurePhotoFolder() As Boolean
    Const ProcName As String = "EnsurePhotoFolder"
    ' [SCAFFOLD] Confirm the shared folder named in tblAppSetting.OfficialPhotoFolder exists -
    '            never create one locally in its place (schema Business Rule 9; see
    '            templates/_materialization.md -> External file assets, "shared content
    '            folders: VERIFY, never create"). Returns False, with a message naming the
    '            folder, when it can't be reached.
    Dim strFolder As String

    On Error GoTo errHandler

    strFolder = GetAppSetting("OfficialPhotoFolder")

    If Len(strFolder) = 0 Then
        MsgBox "No photo folder is configured (OfficialPhotoFolder is blank).", vbExclamation
        GoTo Cleanup
    End If

    If Len(Dir$(strFolder, vbDirectory)) = 0 Then
        MsgBox "The shared photo folder cannot be reached:" & vbCrLf & strFolder, vbExclamation
        GoTo Cleanup
    End If

    EnsurePhotoFolder = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] reporting shape: entry point that hands its outcome back. It
    '            tells the person itself, returns False, and SetOfficialPhoto acts on the False
    '            without reporting again.
    LogError ModuleName, ProcName, Erl          ' option 2 puts its message box here instead
    Resume Cleanup
    Resume
End Function
```

### SetOfficialPhoto — `Public Function` → `Boolean`

```vba
Public Function SetOfficialPhoto(ByVal lngOfficialID As Long, ByVal strSourcePath As String) As Boolean
    Const ProcName As String = "SetOfficialPhoto"
    ' [SCAFFOLD] Copy a picked photo into the shared folder under a controlled name, and
    '            record that name on the official - never the source path, never the file
    '            itself (schema Business Rule 9). The controlled name carries the picked
    '            file's own extension, which can change between picks for the same official
    '            (a .jpg replaced by a .png) - so the new name is not always the old name.
    '            The record change is held in a transaction until the new file is in place:
    '            the file is copied beside the old one under a temporary name, swapped in,
    '            and only then is the record committed. A failure before the commit rolls
    '            the record back and leaves the old photo. After the commit, the official's
    '            previous file is deleted by its recorded full name, when that name differs
    '            from the new one; guessing the old extension would miss an extension change.
    Dim ws         As DAO.Workspace
    Dim db         As DAO.Database
    Dim rst         As DAO.Recordset
    Dim strSQL       As String
    Dim strFolder    As String
    Dim strExt       As String
    Dim strName      As String
    Dim strOldName   As String
    Dim lngDot       As Long
    Dim booInTrans As Boolean

    On Error GoTo errHandler
    ' [STANDARDS — error-handling.md, "Transaction guard"] db comes from the Workspace the
    '            transaction is begun on.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)

    If Not EnsurePhotoFolder() Then GoTo Cleanup

    strFolder = GetAppSetting("OfficialPhotoFolder")
    If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    ' [BUSINESS LOGIC #9] read this official's current PhotoFileName (full name, with extension)
    ' >>> lookup query per query-style.md <<<
    strSQL = vbNullString
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rst.EOF Then strOldName = Nz(rst!PhotoFileName, vbNullString)
    rst.Close

    lngDot = InStrRev(strSourcePath, ".")
    strExt = IIf(lngDot > 0, Mid$(strSourcePath, lngDot), vbNullString)
    strName = "Official_" & lngOfficialID & strExt

    ws.BeginTrans
    booInTrans = True

    ' [BUSINESS LOGIC #9] UPDATE tblOfficial SET PhotoFileName = strName WHERE OfficialID = lngOfficialID
    ' >>> update per query-style.md, on db <<<

    ' Copy beside the old file, never over it, then swap the copy in.
    FileCopy strSourcePath, strFolder & strName & ".tmp"
    If Len(Dir$(strFolder & strName)) > 0 Then Kill strFolder & strName
    Name strFolder & strName & ".tmp" As strFolder & strName

    ws.CommitTrans
    booInTrans = False

    ' Delete the superseded file, by its own recorded full name, now that nothing points at it.
    ' Only when a prior file is on record and its name differs from the new one - never delete
    ' on a guess, and never delete when nothing changed.
    If Len(strOldName) > 0 And strOldName <> strName Then
        If Len(Dir$(strFolder & strOldName)) > 0 Then Kill strFolder & strOldName
    End If

    SetOfficialPhoto = True

Cleanup:
    On Error Resume Next
    ' A temporary copy left by a failure is removed; after a success there is none.
    If Len(strName) > 0 Then
        If Len(Dir$(strFolder & strName & ".tmp")) > 0 Then Kill strFolder & strName & ".tmp"
    End If
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing: Set db = Nothing: Set ws = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] reporting shape: entry point; returns False when refused or
    '            failed. The transaction guard's rollback comes first: a failure before the commit
    '            leaves the record and the old photo as they were.
    If booInTrans Then ws.Rollback: booInTrans = False
    If IsRefusal(Err.Number) Then
        MsgBox Err.Description, vbInformation
    Else
        LogError ModuleName, ProcName, Erl          ' option 2 puts its message box here instead
    End If
    Resume Cleanup
    Resume
End Function
```
