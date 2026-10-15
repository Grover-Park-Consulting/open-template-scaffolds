---
step: 03-procedures-2
title: "The procedures, part 2 of 2 (from GameLevelID)"
platform_facts: [mcp-module-import, mcp-line-numbers, recordset-append-crash, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

### GameLevelID — `Public Function` → `Long`

```vba
Public Function GameLevelID(ByVal lGameID As Long) As Long
    ' [SCAFFOLD] A game's play level, read through its home team - never stored on the game
    '            itself (schema Business Rule 5). Returns 0 when the game or its home team
    '            can't be found.
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #5] tblGame -> tblTeam (on HomeTeamID) -> PlayLevelID
    ' >>> join query per query-style.md <<<
    sSql = vbNullString

    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then GameLevelID = Nz(rs!PlayLevelID, 0)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```

### OfficialAge — `Public Function` → `Variant`

```vba
Public Function OfficialAge(ByVal lOfficialID As Long) As Variant
    ' [SCAFFOLD] An official's age, computed from BirthDate as of today - never stored
    '            anywhere (schema Business Rule 8). Returns Null when BirthDate is unknown,
    '            which is why this is a Variant rather than an Integer.
    Dim db      As DAO.Database
    Dim rs      As DAO.Recordset
    Dim sSql    As String
    Dim dtBirth As Variant

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #8] look up tblOfficial.BirthDate for lOfficialID
    ' >>> lookup query per query-style.md <<<
    sSql = vbNullString

    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then
        dtBirth = rs!BirthDate
        If Not IsNull(dtBirth) Then
            ' Subtract a year when this year's birthday hasn't happened yet.
            OfficialAge = DateDiff("yyyy", dtBirth, Date) - _
                IIf(Format(Date, "mmdd") < Format(dtBirth, "mmdd"), 1, 0)
        End If
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```

### EnsurePhotoFolder — `Public Function` → `Boolean`

```vba
Public Function EnsurePhotoFolder() As Boolean
    ' [SCAFFOLD] Confirm the shared folder named in tblAppSetting.OfficialPhotoFolder exists -
    '            never create one locally in its place (schema Business Rule 9; see
    '            templates/_materialization.md -> External file assets, "shared content
    '            folders: VERIFY, never create"). Returns False, with a message naming the
    '            folder, when it can't be reached.
    Dim sFolder As String

    On Error GoTo errHandler

    sFolder = GetAppSetting("OfficialPhotoFolder")

    If Len(sFolder) = 0 Then
        MsgBox "No photo folder is configured (OfficialPhotoFolder is blank).", vbExclamation
        GoTo Cleanup
    End If

    If Len(Dir$(sFolder, vbDirectory)) = 0 Then
        MsgBox "The shared photo folder cannot be reached:" & vbCrLf & sFolder, vbExclamation
        GoTo Cleanup
    End If

    EnsurePhotoFolder = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```

### SetOfficialPhoto — `Public Function` → `Boolean`

```vba
Public Function SetOfficialPhoto(ByVal lOfficialID As Long, ByVal sSourcePath As String) As Boolean
    ' [SCAFFOLD] Copy a picked photo into the shared folder under a controlled name, and
    '            record that name on the official - never the source path, never the file
    '            itself (schema Business Rule 9). The controlled name carries the picked
    '            file's own extension, which can change between picks for the same official
    '            (a .jpg replaced by a .png) - so the new name is not always the old name.
    '            Before writing the new file, this reads the official's CURRENT
    '            PhotoFileName - the full stored name, extension included - and deletes
    '            exactly that file if it differs from the name about to be written. Deleting
    '            by the recorded full name, not by guessing the old extension matches the
    '            new one, is what keeps this accurate across an extension change.
    Dim db       As DAO.Database
    Dim rs       As DAO.Recordset
    Dim sSql     As String
    Dim sFolder  As String
    Dim sExt     As String
    Dim sName    As String
    Dim sOldName As String
    Dim lDot     As Long

    On Error GoTo errHandler
    Set db = CurrentDb

    If Not EnsurePhotoFolder() Then GoTo Cleanup

    sFolder = GetAppSetting("OfficialPhotoFolder")
    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"

    ' [BUSINESS LOGIC #9] read this official's current PhotoFileName (full name, with extension)
    ' >>> lookup query per query-style.md <<<
    sSql = vbNullString
    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then sOldName = Nz(rs!PhotoFileName, vbNullString)
    rs.Close

    lDot = InStrRev(sSourcePath, ".")
    sExt = IIf(lDot > 0, Mid$(sSourcePath, lDot), vbNullString)
    sName = "Official_" & lOfficialID & sExt

    ' Delete the superseded file, by its own recorded full name, before writing the new one.
    ' Only when a prior file is on record and its name differs from the one about to be
    ' written - never delete on a guess, and never delete when nothing changed.
    If Len(sOldName) > 0 And sOldName <> sName Then
        If Len(Dir$(sFolder & sOldName)) > 0 Then Kill sFolder & sOldName
    End If

    FileCopy sSourcePath, sFolder & sName

    ' [BUSINESS LOGIC #9] UPDATE tblOfficial SET PhotoFileName = sName WHERE OfficialID = lOfficialID
    ' >>> update per query-style.md <<<

    SetOfficialPhoto = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```
