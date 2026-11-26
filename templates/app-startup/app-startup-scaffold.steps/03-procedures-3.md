---
step: 03-procedures-3
title: "The procedures, part 3 of 3 (from RelinkAllTables)"
platform_facts: [mcp-module-import, mcp-line-numbers, vbe-line-continuation, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

### RelinkAllTables — `Private Function` → `Boolean`

```vba
Private Function RelinkAllTables(ByVal strNewPath As String, _
                                 ByVal strOldPath As String) As Boolean
    Const ProcName As String = "RelinkAllTables"
    ' [SCAFFOLD] Point every link that used to reach strOldPath at strNewPath instead.
    '            Two deliberate exclusions:
    '              - local tables (no Connect string) are not links and are left alone;
    '              - ODBC links (Connect starts "ODBC;") reach a different server entirely
    '                and must NOT be rewritten with a file path.
    '            Passing an empty strOldPath repoints every Access file link - the right
    '            behaviour on a front end with one back end, and the wrong one on a front
    '            end linked to two (see Parked).
    '            Returns False if nothing was relinked: an application whose links all
    '            failed, yet where no link matched, has a problem worth stopping for.
    Dim db        As DAO.Database
    Dim tdf       As DAO.TableDef
    Dim lngRelinked As Long

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If Len(PathFromConnect(tdf.Connect)) > 0 Then
            If Len(strOldPath) = 0 Or _
               InStr(1, tdf.Connect, strOldPath, vbTextCompare) > 0 Then
                ' Only the path is replaced. Anything else the link carries - a database
                ' password, for one - is kept, because rebuilding the string from scratch
                ' would drop it and the refresh below would then fail.
                tdf.Connect = ConnectWithNewPath(tdf.Connect, strNewPath)
                ' RefreshLink is what actually re-establishes the connection; setting
                ' Connect on its own only changes the stored text.
                tdf.RefreshLink
                lngRelinked = lngRelinked + 1
            End If
        End If
    Next tdf

    RelinkAllTables = (lngRelinked > 0)

Cleanup:
    On Error Resume Next
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step. Its False already means "no link
    '            matched", so a failure must not return one.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set tdf = Nothing: Set db = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### AskUserForBackEnd — `Private Function` → `String`

```vba
Private Function AskUserForBackEnd() As String
    Const ProcName As String = "AskUserForBackEnd"
    ' [SCAFFOLD] Show a file picker and return the chosen path, or an empty string if the
    '            person cancels. Declared As Object (late binding) and using the numeric
    '            dialog constant so this module compiles without a reference to the Office
    '            object library - which is not present in every Access application.
    Dim objDialog As Object

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler

    Set objDialog = Application.FileDialog(FilePickerDialog)

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
    ' [STANDARDS - error-handling.md] propagating shape: a step. Its "" already means "the person
    '            cancelled", so a failure must not return one.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set objDialog = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### LocalSettingValue — `Public Function` → `String`

```vba
Public Function LocalSettingValue(ByVal strSettingName As String) As String
    Const ProcName As String = "LocalSettingValue"
    ' [SCAFFOLD] Read one setting from the front end's own USysLocalSetting table. Returns
    '            an empty string when the setting has never been saved - which is the normal
    '            state of a newly built front end, not an error.
    '            This procedure and SaveLocalSetting are the only two that know WHERE the
    '            setting is kept; swap them both to move it to a text file or a constant.
    Dim db   As DAO.Database
    Dim qdf  As DAO.QueryDef
    Dim rst  As DAO.Recordset

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler

    Set db = CurrentDb

    ' [STANDARDS - query-style.md] the name is a parameter, never part of the SQL text, so
    '            a value containing an apostrophe cannot break the statement.
    Set qdf = db.CreateQueryDef(vbNullString, _
          "PARAMETERS [pSettingName] Text(255); " & _
          "SELECT SettingValue FROM USysLocalSetting WHERE SettingName = [pSettingName];")
    qdf.Parameters("pSettingName").Value = strSettingName

    Set rst = qdf.OpenRecordset(dbOpenSnapshot)
    If Not rst.EOF Then LocalSettingValue = Nz(rst!SettingValue, vbNullString)

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step. Its "" already means "never
    '            saved", so a failure must not return one.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set rst = Nothing: Set qdf = Nothing: Set db = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### SaveLocalSetting — `Public Sub`

```vba
Public Sub SaveLocalSetting(ByVal strSettingName As String, ByVal strSettingValue As String)
    Const ProcName As String = "SaveLocalSetting"
    ' [SCAFFOLD] Write one setting to USysLocalSetting: UPDATE the row, and INSERT it only
    '            when the UPDATE changed nothing. Parameterised statements, never a recordset
    '            write. The table sits in this person's own front end, so no one else can add
    '            the same row between the two statements.
    Dim db   As DAO.Database
    Dim qdf  As DAO.QueryDef

    On Error GoTo errHandler

    Set db = CurrentDb

    ' [FACT sql-insert-truncation] A SQL statement cuts a value longer than the field without
    '            an error, so refuse it first. SettingValue is Text(255).
    If Len(strSettingValue) > 255 Then
        Err.Raise ErrSettingTooLong, ModuleName & "." & ProcName, _
                  "The setting " & strSettingName & " was not saved: its value is longer than " & _
                  "255 characters."
    End If

    ' [STANDARDS - query-style.md] the values are parameters, as in LocalSettingValue.
    Set qdf = db.CreateQueryDef(vbNullString, _
          "PARAMETERS [pSettingName] Text(255), [pSettingValue] Text(255); " & _
          "UPDATE USysLocalSetting SET SettingValue = [pSettingValue] " & _
          "WHERE SettingName = [pSettingName];")
    qdf.Parameters("pSettingName").Value = strSettingName
    qdf.Parameters("pSettingValue").Value = strSettingValue
    qdf.Execute dbFailOnError

    If qdf.RecordsAffected = 0 Then
        Set qdf = db.CreateQueryDef(vbNullString, _
              "PARAMETERS [pSettingName] Text(255), [pSettingValue] Text(255); " & _
              "INSERT INTO USysLocalSetting (SettingName, SettingValue) " & _
              "VALUES ([pSettingName], [pSettingValue]);")
        qdf.Parameters("pSettingName").Value = strSettingName
        qdf.Parameters("pSettingValue").Value = strSettingValue
        qdf.Execute dbFailOnError
    End If

Cleanup:
    On Error Resume Next
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS - error-handling.md] reporting shape: entry point. Nothing acts on whether the
    '            setting was saved: the next start simply asks again.
    If IsRefusal(Err.Number) Then
        MsgBox Err.Description, vbInformation
    Else
        LogError ModuleName, ProcName, Erl          ' option 2 puts its message box here instead
    End If
    Resume Cleanup
    Resume
End Sub
```
