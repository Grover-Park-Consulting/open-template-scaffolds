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

### Startup

```vba
Public Function Startup() As Boolean
```

The only thing AutoExec runs, and the one place open-time work happens - including, once you switch it on, what your existing AutoExec did (step 0). Returns False when the application must not carry on; each step has already told the user why.

Error handling (`error-handling.md`): reporting shape.

### EnsureBackEndLink

```vba
Public Function EnsureBackEndLink() As Boolean
```

Make sure the linked tables reach a real back end, relinking if they don't. Four outcomes, in the order they are tried: a. no linked tables at all -> single-file database, nothing to do b. the current links work -> record where they point, carry on c. the remembered path works -> relink to it silently, carry on d. neither works -> tell the user, let them find the file

Error handling (`error-handling.md`): reporting shape.

### EnsureAppFolders

```vba
Public Function EnsureAppFolders() As Boolean
```

The folder slot named in startup-conventions.md sections 2 and 4, and the extension point for open-time setup: add each new folder here as the application grows. Idempotent by design - the SAME routine creates a folder on the first run and verifies it on every open after that.

Error handling (`error-handling.md`): reporting shape.

### CurrentBackEndPath

```vba
Private Function CurrentBackEndPath() As String
```

Where do this front end's links currently point? Returns the path held by the first file link found, or an empty string when there are none - which is how a single-file database identifies itself, without asking anyone. Which links count, and how the path is dug out of one, is PathFromConnect's job - there is more than one shape and getting it wrong fails silently.

Error handling (`error-handling.md`): propagating shape.

### PathFromConnect

```vba
Private Function PathFromConnect(ByVal strConnect As String) As String
```

Given what one link records, return the Access file it points at, or an empty string if it doesn't point at one. This is the single place that decides what counts as a link to this application's data file. Exactly TWO shapes name an Access file, and both are ten characters long: ";DATABASE=" - an ordinary link, path immediately after "MS Access;" - a link to a file that HAS A DATABASE PASSWORD, which records the password first: MS Access;PWD=x;DATABASE=y Testing the START of the string, rather than searching anywhere in it, is what keeps a spreadsheet or text-file link out - those carry ";DATABASE=" too, further along ("Excel 12.0 Xml;HDR=YES;...;DATABASE=..."), and repointing one at an .accdb would break it.

Error handling (`error-handling.md`): propagating shape.

### BackEndExtras

```vba
Private Function BackEndExtras() As String
```

Everything OpenDatabase needs about the back end BESIDES its path. Today that means one thing: the database password, where the back end has one. Returned in the form OpenDatabase wants - ";PWD=<password>" - or an empty string, which OpenDatabase accepts happily. Read out of the links the front end already holds, so nobody types a password and none is stored anywhere new. Note what that means: a link to a password-protected file carries the password IN CLEAR TEXT, readable by anyone who can open the front end. That is Access's doing, not this scaffold's, but it is worth knowing before choosing to use one.

Error handling (`error-handling.md`): propagating shape.

### ConnectWithNewPath

```vba
Private Function ConnectWithNewPath(ByVal strConnect As String, _
                                    ByVal strNewPath As String) As String
```

Swap the file path inside a link's Connect string and change nothing else. Building a fresh ";DATABASE=" & path string instead would be shorter and would silently throw away a database password, after which RefreshLink fails and the relink reports failure for a reason nobody can see.

Error handling (`error-handling.md`): propagating shape.

### BackEndIsReachable

```vba
Private Function BackEndIsReachable(ByVal strPath As String) As Boolean
```

Two questions in one: can this file be opened at all, and is it THIS application's back end? The second question is the guard against a picked file that opens perfectly and holds the wrong data. Returns False by default. Any error - missing file, unreachable server, a file that isn't a database, a table that isn't there - lands in the handler, resumes to Cleanup, and leaves that False in place. That is the whole mechanism; there is no On Error Resume Next in the main logic.

Error handling (`error-handling.md`): answering shape.

### RelinkAllTables

```vba
Private Function RelinkAllTables(ByVal strNewPath As String, _
                                 ByVal strOldPath As String) As Boolean
```

Point every link that used to reach strOldPath at strNewPath instead. Two deliberate exclusions: - local tables (no Connect string) are not links and are left alone; - ODBC links (Connect starts "ODBC;") reach a different server entirely and must NOT be rewritten with a file path. Passing an empty strOldPath repoints every Access file link - the right behaviour on a front end with one back end, and the wrong one on a front end linked to two (see Parked). Returns False if nothing was relinked: an application whose links all failed, yet where no link matched, has a problem worth stopping for.

Error handling (`error-handling.md`): propagating shape.

### AskUserForBackEnd

```vba
Private Function AskUserForBackEnd() As String
```

Show a file picker and return the chosen path, or an empty string if the person cancels. Declared As Object (late binding) and using the numeric dialog constant so this module compiles without a reference to the Office object library - which is not present in every Access application.

Error handling (`error-handling.md`): propagating shape.

### LocalSettingValue

```vba
Public Function LocalSettingValue(ByVal strSettingName As String) As String
```

Read one setting from the front end's own USysLocalSetting table. Returns an empty string when the setting has never been saved - which is the normal state of a newly built front end, not an error. This procedure and SaveLocalSetting are the only two that know WHERE the setting is kept; swap them both to move it to a text file or a constant.

Error handling (`error-handling.md`): propagating shape.

### SaveLocalSetting

```vba
Public Sub SaveLocalSetting(ByVal strSettingName As String, ByVal strSettingValue As String)
```

Write one setting to USysLocalSetting: UPDATE the row, and INSERT it only when the UPDATE changed nothing. Parameterised statements, never a recordset write. The table sits in this person's own front end, so no one else can add the same row between the two statements.

Error handling (`error-handling.md`): reporting shape.
