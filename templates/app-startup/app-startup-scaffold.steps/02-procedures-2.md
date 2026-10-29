---
step: 02-procedures-2
title: "The procedures, part 2 of 3 (from CurrentBackEndPath)"
platform_facts: [mcp-module-import, mcp-line-numbers, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.
### CurrentBackEndPath — `Private Function` → `String`

```vba
Private Function CurrentBackEndPath() As String
    ' [SCAFFOLD] Where do this front end's links currently point? Returns the path held by
    '            the first file link found, or an empty string when there are none - which
    '            is how a single-file database identifies itself, without asking anyone.
    '            Which links count, and how the path is dug out of one, is PathFromConnect's
    '            job - there is more than one shape and getting it wrong fails silently.
    Dim db  As DAO.Database
    Dim tdf As DAO.TableDef

    On Error GoTo errHandler

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        CurrentBackEndPath = PathFromConnect(tdf.Connect)
        If Len(CurrentBackEndPath) > 0 Then Exit For
    Next tdf

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

### PathFromConnect — `Private Function` → `String`

```vba
Private Function PathFromConnect(ByVal sConnect As String) As String
    ' [SCAFFOLD] Given what one link records, return the Access file it points at, or an
    '            empty string if it doesn't point at one. This is the single place that
    '            decides what counts as a link to this application's data file.
    '            Exactly TWO shapes name an Access file, and both are ten characters long:
    '              ";DATABASE=" - an ordinary link, path immediately after
    '              "MS Access;" - a link to a file that HAS A DATABASE PASSWORD, which
    '                             records the password first: MS Access;PWD=x;DATABASE=y
    '            Testing the START of the string, rather than searching anywhere in it, is
    '            what keeps a spreadsheet or text-file link out - those carry ";DATABASE="
    '            too, further along ("Excel 12.0 Xml;HDR=YES;...;DATABASE=..."), and
    '            repointing one at an .accdb would break it.
    Dim lStart As Long
    Dim lEnd   As Long
    Dim sTail  As String

    On Error GoTo errHandler

    Select Case Left$(sConnect, 10)
        Case ";DATABASE=", "MS Access;"
            ' one of ours - carry on
        Case Else
            GoTo Cleanup
    End Select

    lStart = InStr(1, sConnect, ";DATABASE=", vbTextCompare)
    If lStart = 0 Then GoTo Cleanup

    sTail = Mid$(sConnect, lStart + 10)

    ' Whatever follows the path is separated from it by a semicolon.
    lEnd = InStr(sTail, ";")
    If lEnd > 0 Then sTail = Left$(sTail, lEnd - 1)

    PathFromConnect = sTail

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Function
```

### BackEndExtras — `Private Function` → `String`

```vba
Private Function BackEndExtras() As String
    ' [SCAFFOLD] Everything OpenDatabase needs about the back end BESIDES its path. Today
    '            that means one thing: the database password, where the back end has one.
    '            Returned in the form OpenDatabase wants - ";PWD=<password>" - or an empty
    '            string, which OpenDatabase accepts happily.
    '            Read out of the links the front end already holds, so nobody types a
    '            password and none is stored anywhere new. Note what that means: a link to
    '            a password-protected file carries the password IN CLEAR TEXT, readable by
    '            anyone who can open the front end. That is Access's doing, not this
    '            scaffold's, but it is worth knowing before choosing to use one.
    Dim db     As DAO.Database
    Dim tdf    As DAO.TableDef
    Dim lStart As Long
    Dim lEnd   As Long
    Dim sTail  As String

    On Error GoTo errHandler

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If Left$(tdf.Connect, 10) = "MS Access;" Then
            lStart = InStr(1, tdf.Connect, ";PWD=", vbTextCompare)
            If lStart > 0 Then
                sTail = Mid$(tdf.Connect, lStart + 5)
                lEnd = InStr(sTail, ";")
                If lEnd > 0 Then sTail = Left$(sTail, lEnd - 1)
                BackEndExtras = ";PWD=" & sTail
                Exit For
            End If
        End If
    Next tdf

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

### ConnectWithNewPath — `Private Function` → `String`

```vba
Private Function ConnectWithNewPath(ByVal sConnect As String, _
                                    ByVal sNewPath As String) As String
    ' [SCAFFOLD] Swap the file path inside a link's Connect string and change nothing else.
    '            Building a fresh ";DATABASE=" & path string instead would be shorter and
    '            would silently throw away a database password, after which RefreshLink
    '            fails and the relink reports failure for a reason nobody can see.
    Dim lStart As Long
    Dim lEnd   As Long
    Dim sTail  As String

    On Error GoTo errHandler

    lStart = InStr(1, sConnect, ";DATABASE=", vbTextCompare)
    If lStart = 0 Then GoTo Cleanup

    sTail = Mid$(sConnect, lStart + 10)
    lEnd = InStr(sTail, ";")

    If lEnd > 0 Then
        ' Something follows the path - keep it, starting at its semicolon.
        ConnectWithNewPath = Left$(sConnect, lStart + 9) & sNewPath & Mid$(sTail, lEnd)
    Else
        ConnectWithNewPath = Left$(sConnect, lStart + 9) & sNewPath
    End If

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Function
```

### BackEndIsReachable — `Private Function` → `Boolean`

```vba
Private Function BackEndIsReachable(ByVal sPath As String) As Boolean
    ' [SCAFFOLD] Two questions in one: can this file be opened at all, and is it THIS
    '            application's back end? The second question is the guard against a picked
    '            file that opens perfectly and holds the wrong data.
    '            Returns False by default. Any error - missing file, unreachable server, a
    '            file that isn't a database, a table that isn't there - lands in the handler,
    '            resumes to Cleanup, and leaves that False in place. That is the whole
    '            mechanism; there is no On Error Resume Next in the main logic.
    Dim dbBackEnd As DAO.Database
    Dim tdf       As DAO.TableDef

    On Error GoTo errHandler

    If Len(Dir$(sPath)) = 0 Then GoTo Cleanup

    ' Opened read-only and non-exclusively, so this check never blocks anyone else, and
    ' closed again in Cleanup a few lines later. BackEndExtras supplies the database
    ' password where the back end has one - without it this open fails on a file that is
    ' perfectly reachable, and the person is told their own data file is the wrong file.
    Set dbBackEnd = DBEngine.OpenDatabase(sPath, False, True, BackEndExtras())

    ' [BUSINESS LOGIC] name a table this application's back end must contain. Asking for a
    '            table that isn't there raises an error, which is exactly the answer wanted.
    ' >>> Set tdf = dbBackEnd.TableDefs("<ATableOnlyThisBackEndHas>") <<<

    BackEndIsReachable = True

Cleanup:
    On Error Resume Next
    Set tdf = Nothing
    If Not dbBackEnd Is Nothing Then dbBackEnd.Close
    Set dbBackEnd = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup).
    '            Note this handler logs a failed probe. That is deliberate: an application
    '            that cannot reach its data is worth a log entry.
    Resume Cleanup
    Resume
End Function
```
