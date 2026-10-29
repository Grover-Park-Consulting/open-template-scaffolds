---
step: 03-procedures-3
title: "The procedures, part 3 of 3 (from RelinkAllTables)"
platform_facts: [mcp-module-import, mcp-line-numbers, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

### RelinkAllTables — `Private Function` → `Boolean`

```vba
Private Function RelinkAllTables(ByVal sNewPath As String, _
                                 ByVal sOldPath As String) As Boolean
    ' [SCAFFOLD] Point every link that used to reach sOldPath at sNewPath instead.
    '            Two deliberate exclusions:
    '              - local tables (no Connect string) are not links and are left alone;
    '              - ODBC links (Connect starts "ODBC;") reach a different server entirely
    '                and must NOT be rewritten with a file path.
    '            Passing an empty sOldPath repoints every Access file link - the right
    '            behaviour on a front end with one back end, and the wrong one on a front
    '            end linked to two (see Parked).
    '            Returns False if nothing was relinked: an application whose links all
    '            failed, yet where no link matched, has a problem worth stopping for.
    Dim db        As DAO.Database
    Dim tdf       As DAO.TableDef
    Dim lRelinked As Long

    On Error GoTo errHandler

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If Len(PathFromConnect(tdf.Connect)) > 0 Then
            If Len(sOldPath) = 0 Or _
               InStr(1, tdf.Connect, sOldPath, vbTextCompare) > 0 Then
                ' Only the path is replaced. Anything else the link carries - a database
                ' password, for one - is kept, because rebuilding the string from scratch
                ' would drop it and the refresh below would then fail.
                tdf.Connect = ConnectWithNewPath(tdf.Connect, sNewPath)
                ' RefreshLink is what actually re-establishes the connection; setting
                ' Connect on its own only changes the stored text.
                tdf.RefreshLink
                lRelinked = lRelinked + 1
            End If
        End If
    Next tdf

    RelinkAllTables = (lRelinked > 0)

Cleanup:
    On Error Resume Next
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Function
```

### AskUserForBackEnd — `Private Function` → `String`

```vba
Private Function AskUserForBackEnd() As String
    ' [SCAFFOLD] Show a file picker and return the chosen path, or an empty string if the
    '            person cancels. Declared As Object (late binding) and using the numeric
    '            dialog constant so this module compiles without a reference to the Office
    '            object library - which is not present in every Access application.
    Dim objDialog As Object

    On Error GoTo errHandler

    Set objDialog = Application.FileDialog(mlngFilePicker)

    With objDialog
        ' [BUSINESS LOGIC] substitute your application's name.
        .Title = "Locate the data file for this application"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Access data file", "*.accdb; *.mdb"
        ' Show returns -1 when a file was chosen, 0 when the person cancelled.
        If .Show = -1 Then AskUserForBackEnd = .SelectedItems(1)
    End With

Cleanup:
    On Error Resume Next
    Set objDialog = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Function
```

### LocalSettingValue — `Public Function` → `String`

```vba
Public Function LocalSettingValue(ByVal sSettingName As String) As String
    ' [SCAFFOLD] Read one setting from the front end's own USysLocalSetting table. Returns
    '            an empty string when the setting has never been saved - which is the normal
    '            state of a newly built front end, not an error.
    '            This procedure and SaveLocalSetting are the only two that know WHERE the
    '            setting is kept; swap them both to move it to a text file or a constant.
    Dim db   As DAO.Database
    Dim rst  As DAO.Recordset
    Dim sSQL As String

    On Error GoTo errHandler

    Set db = CurrentDb

    ' [STANDARDS - query-style.md] a literal in a WHERE clause has its single quotes
    '            doubled, so a value containing an apostrophe cannot break the statement.
    sSQL = "SELECT SettingValue FROM USysLocalSetting " & _
           "WHERE SettingName = '" & Replace$(sSettingName, "'", "''") & "';"

    Set rst = db.OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rst.EOF Then LocalSettingValue = Nz(rst!SettingValue, vbNullString)

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Function
```

### SaveLocalSetting — `Public Sub`

```vba
Public Sub SaveLocalSetting(ByVal sSettingName As String, ByVal sSettingValue As String)
    ' [SCAFFOLD] Write one setting to USysLocalSetting, adding the row if it isn't there
    '            yet. Editable recordset rather than an UPDATE followed by an INSERT: one
    '            trip, and no window in which the row exists twice.
    Dim db   As DAO.Database
    Dim rst  As DAO.Recordset
    Dim sSQL As String

    On Error GoTo errHandler

    Set db = CurrentDb

    ' [STANDARDS - query-style.md] same quote-doubling rule as LocalSettingValue.
    sSQL = "SELECT SettingName, SettingValue FROM USysLocalSetting " & _
           "WHERE SettingName = '" & Replace$(sSettingName, "'", "''") & "';"

    Set rst = db.OpenRecordset(sSQL, dbOpenDynaset)

    If rst.EOF Then
        rst.AddNew
        rst!SettingName = sSettingName
    Else
        rst.Edit
    End If

    rst!SettingValue = sSettingValue
    rst.Update

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Sub
```
