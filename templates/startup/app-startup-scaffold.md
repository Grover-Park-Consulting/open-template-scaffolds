---
template: app-startup-scaffold
title: Application Startup and Back-End Relinking — VBA Scaffold
domain: startup
type: vba-scaffold
version: 0.1.0
status: draft
requires_tables:
  - USysLocalSetting
standards_layer:
  - design-principles
  - error-handling
  - naming-conventions
  - query-style
  - startup-conventions
target_module: modAppStartup
new_procedures:
  - Startup
  - EnsureBackEndLink
  - EnsureAppFolders
  - CurrentBackEndPath
  - BackEndIsReachable
  - RelinkAllTables
  - AskUserForBackEnd
  - LocalSettingValue
  - SaveLocalSetting
seeds:
  - USysLocalSetting.BackEndPath
house_assumptions:
  - USysLocalSetting — a name/value settings table rather than one column per setting, mirroring the shared tblAppSetting so both read the same way; a practice preferring one column per setting changes the two accessor procedures and nothing else
---

# Application Startup and Back-End Relinking — VBA Scaffold

## Intent

Realize the open-time entry point that `standards/startup-conventions.md` describes — the
`AutoExec` → `Startup()` convention — and solve the problem that makes it urgent: **the data file
has moved and the application can no longer find it.**

**What a link is, and why it breaks.** In a split database the tables live in one file (the **back
end**) and each person runs their own copy of a second file holding the forms and code (the **front
end**). The front end doesn't contain the tables; it contains **links** to them. A link stores the
back end's **path, recorded at the moment the link was made** — nothing keeps it up to date. Move
the back end to a new server, rename the share, deploy the application to a second site, or hand the
front end to someone whose drive letters differ, and every link in that front end points at a file
that isn't there. The application opens normally and then fails on the first form that touches
data — which looks, to the person using it, like the application is broken.

This scaffold supplies the **procedure skeletons** for an application that handles that itself:
check the connection before anything else runs, relink when the back end has moved, remember the
answer, and say something a person can act on when it can't. It also carries `EnsureAppFolders()` —
the folder-ensure slot `startup-conventions.md` §2 names — because the two are ordered: the folder
paths are settings **in the back end**, so folders cannot be ensured until the connection works.

**On a single-file database this scaffold does nothing.** One .accdb holding everything is an
acceptable choice for one user; it has no linked tables, so the relink check finds nothing to check
and returns immediately. Nothing here needs to be removed for that case.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here (the sequence, the probes, the relink loop, the error
  structure).
- **`[STANDARDS]`** — house style, deferred (`error-handling.md`, `query-style.md`,
  `naming-conventions.md`, `startup-conventions.md`).
- **`[BUSINESS LOGIC]`** — what you fill in per application (the table that proves a chosen file is
  the right back end, your startup form, your folder settings, the application's name in messages).

## Prerequisites

| Object | Role |
|---|---|
| `USysLocalSetting` | A table **in the front end, not linked** — this scaffold's memory of where the back end is. Defined below. |
| An `AutoExec` macro | Its only action is `RunCode Startup()`. Built per `templates/_materialization.md` → *The AutoExec build gotcha*. |
| One table the back end must contain | Named by you; `BackEndIsReachable` looks for it to prove a chosen file really is this application's data file. |
| A startup form | The switchboard, menu, or home form `Startup()` opens once everything checks out. |
| A central error logger | `error-handling.md` (GPC default: `codearchive.GlblErrMsg`). **It must not write to the back end** — see *Standards Layer*. |

### `USysLocalSetting` — the front end's own memory

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `LocalSettingID` | AutoNumber | PK | Surrogate key |
| `SettingName` | Text(50) | Required | Unique. This scaffold uses one row: `BackEndPath` |
| `SettingValue` | Text(255) | Nullable | For `BackEndPath`, the full path of the back-end file |
| `SettingDescription` | Text(255) | Nullable | Why the setting exists, for whoever reads the table later |

Indexes: PK on `LocalSettingID`; unique on `SettingName`.

**Why this table is local and not shared.** A settings table in the back end — the pattern
`templates/sports/officiating-assignment-schema.md` uses for `tblAppSetting` — is right for settings
everybody must share, because changing one row changes it for everyone at once. It is exactly wrong
here: **a setting stored in the back end cannot tell you where the back end is.** When the back end
is unreachable, that is the one table you cannot read. So this one lives in the front end, is never
linked, and opens whether or not the back end can be found.

**Why the name starts with `USys`.** `standards/naming-conventions.md` gives system and
configuration tables the `USys` prefix, and `standards/audit-columns.md` exempts `USys` tables from
the audit columns — which matters here, because the house audit machinery lives in the back end this
table exists to locate. **What you will see:** Access treats an object whose name begins with `USys`
as a system object and **hides it from the navigation pane**. To look at its contents, turn on *Show
System Objects* in the navigation options. It is not missing.

**It fills itself in.** On a fresh build the table has no `BackEndPath` row and needs none. The first
time the application opens with working links, `EnsureBackEndLink` records where they point. Nobody
has to type a path to get started.

**Two other places the path could live**, ranked below the local table rather than omitted — each is
a legitimate choice with a different cost:

| Where | What it buys | What it costs |
|---|---|---|
| **A local table** (this scaffold) | Opens even when every link is broken, because it is local. Configuration lives in data, so changing it is an edit, not a code change. | The table travels inside each front-end copy, so each person's copy remembers its own answer. |
| **A text file beside the front end** | Someone can fix it in Notepad **without opening Access at all** — which is precisely the situation you are in when the application won't start. | File reading in the startup path, and a file that can be deleted, moved, or edited into nonsense with nothing to catch it. |
| **A constant in code** | Simplest to write and to read; nothing to create, nothing to seed. | Changing where the back end lives means editing code and giving everybody a new front end. |

Swapping choices changes `LocalSettingValue` and `SaveLocalSetting` and nothing else — every other
procedure asks those two for the path and doesn't know where it came from.

### Where this module goes in a split database

A **split database** is the normal shape for Access applications, especially those in multi-user
environments: one file holds the tables (the **back end**, usually on a shared network drive — never
on OneDrive, Dropbox, or any other file-syncing cloud folder, which corrupts a shared Access back
end), and each person runs their own copy of a second file holding the forms, reports, and code (the
**front end**), whose tables are *links* pointing at the back end. A single-file database — one
.accdb holding everything — is an acceptable choice for one user, and everything here works there
too; put everything in that one file and ignore the distinction.

| Object | Back end | Front end | Why |
|---|---|---|---|
| `modAppStartup` (this scaffold) | No | **Yes** | `AutoExec` runs when a *person* opens a file. A back end is opened by the database engine on behalf of a front end, not by a person, so an `AutoExec` placed there never runs at all. |
| `AutoExec` macro | No | **Yes** | Same reason. |
| `USysLocalSetting` | No | **Yes** | Its whole purpose is to be readable when the back end is not. |

**Every front end needs its own copy of all three**, and each one remembers its own path. That is
correct, not a flaw: two people can legitimately reach the same back end by different routes — one
by drive letter, one by network path — and each front end keeps the route that works for it.

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. The **`errHandler` block is
shown in full in `Startup` and referenced thereafter** — the VBE-reflection form
(`error-handling.md`) is identical in every procedure by design. Line numbers are deliberately
absent (house-specific; see `error-handling.md`).

**Every procedure here carries a handler**, including the two that expect to fail. `BackEndIsReachable`
is asked a question whose answer is often "no," and a missing file, a dead server, or a file that
isn't an Access database all raise errors. Rather than reach for `On Error Resume Next` — which
`error-handling.md` permits only inside `Cleanup:` — the probe is **its own procedure that returns
False by default**: any error jumps to the standard handler, which resumes to `Cleanup`, and the
function returns the False it started with. The question is answered by the handler doing exactly
what the standard says it does.

Module header:

```vba
Option Compare Database
Option Explicit

' msoFileDialogFilePicker, written as its number so this module needs no reference to
' the Office object library - one less thing to go wrong on someone else's machine.
Private Const mlngFilePicker As Long = 3

' The one setting this module keeps in USysLocalSetting.
Private Const mstrBackEndPathSetting As String = "BackEndPath"
```

### Startup — `Public Function` → `Boolean`

```vba
Public Function Startup() As Boolean
    ' [SCAFFOLD] The only thing AutoExec runs, and the one place open-time work happens.
    '            The order below is not stylistic - see the comment at step 2.
    '            Returns False when the application must not carry on; each step has
    '            already told the user why.
    ' [STANDARDS - startup-conventions.md] RunCode can only call a Function, never a Sub,
    '            which is why this is a Function. Everything it does must be safe to run
    '            on every open, because it runs on every open.

    On Error GoTo errHandler

    ' 1. The data connection comes first. Nothing else is meaningful without it.
    If Not EnsureBackEndLink() Then GoTo Cleanup

    ' 2. Folders second, and this order is forced: the folder paths are settings that live
    '    in the BACK END, so they cannot be read until step 1 has succeeded.
    If Not EnsureAppFolders() Then GoTo Cleanup

    ' 3. Only now is it safe to open a form that touches data.
    ' [BUSINESS LOGIC] your application's startup form - switchboard, menu, or home form.
    ' >>> DoCmd.OpenForm "<AppStartupForm>" <<<

    Startup = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] house-specific central logger, shown as a demo.
    '            GPC-private - you won't have it and shouldn't look for it; substitute your
    '            own logger or the dependency-free MsgBox block in error-handling.md.
    Call codearchive.GlblErrMsg(iLn:=Erl, _
        sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
        sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), bLog:=True)
    Resume Cleanup
    Resume
End Function
```

### EnsureBackEndLink — `Public Function` → `Boolean`

```vba
Public Function EnsureBackEndLink() As Boolean
    ' [SCAFFOLD] Make sure the linked tables reach a real back end, relinking if they don't.
    '            Four outcomes, in the order they are tried:
    '              a. no linked tables at all -> single-file database, nothing to do
    '              b. the current links work   -> record where they point, carry on
    '              c. the remembered path works -> relink to it silently, carry on
    '              d. neither works            -> tell the user, let them find the file
    Dim sCurrentPath As String
    Dim sStoredPath  As String
    Dim sNewPath     As String

    On Error GoTo errHandler

    ' (a) A single-file database has no file links. Not an error - there is nothing to check.
    sCurrentPath = CurrentBackEndPath()
    If Len(sCurrentPath) = 0 Then
        EnsureBackEndLink = True
        GoTo Cleanup
    End If

    ' (b) The links already work. Record where they point, so a later move has something
    '     to fall back on. This is what seeds the setting on a brand-new front end.
    If BackEndIsReachable(sCurrentPath) Then
        If LocalSettingValue(mstrBackEndPathSetting) <> sCurrentPath Then
            SaveLocalSetting mstrBackEndPathSetting, sCurrentPath
        End If
        EnsureBackEndLink = True
        GoTo Cleanup
    End If

    ' (c) The links are stale. Try the path we remember from last time.
    sStoredPath = LocalSettingValue(mstrBackEndPathSetting)
    If Len(sStoredPath) > 0 And sStoredPath <> sCurrentPath Then
        If BackEndIsReachable(sStoredPath) Then
            If RelinkAllTables(sStoredPath, sCurrentPath) Then
                EnsureBackEndLink = True
                GoTo Cleanup
            End If
        End If
    End If

    ' (d) Ask. Name the path that was tried - it is the single most useful fact for
    '     whoever has to sort this out, and it costs nothing to include.
    ' [BUSINESS LOGIC] substitute your application's name for "This application".
    MsgBox "This application cannot find its data file." & vbCrLf & vbCrLf & _
           "It looked here:" & vbCrLf & sCurrentPath & vbCrLf & vbCrLf & _
           "Click OK and choose the data file to continue.", _
           vbExclamation, "Data file not found"

    sNewPath = AskUserForBackEnd()
    If Len(sNewPath) = 0 Then
        MsgBox "No data file was chosen, so the application cannot continue." & vbCrLf & _
               "Ask whoever looks after this application where the data file has moved to.", _
               vbExclamation, "Data file not found"
        GoTo Cleanup
    End If

    ' The guard that matters. A file picker will happily hand back last year's backup, or
    ' a copy sitting on this one machine - and everything would then work, quietly, against
    ' the wrong data. Check it is the right file before connecting anything to it.
    If Not BackEndIsReachable(sNewPath) Then
        MsgBox "That file is not this application's data file:" & vbCrLf & sNewPath & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, "Wrong file"
        GoTo Cleanup
    End If

    If Not RelinkAllTables(sNewPath, sCurrentPath) Then
        MsgBox "The tables could not be reconnected to:" & vbCrLf & sNewPath, _
               vbExclamation, "Could not reconnect"
        GoTo Cleanup
    End If

    ' Remember it only now - after a relink that actually worked. A remembered bad path is
    ' worse than no remembered path, because it gets tried first on every future open.
    SaveLocalSetting mstrBackEndPathSetting, sNewPath
    EnsureBackEndLink = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Function
```

### EnsureAppFolders — `Public Function` → `Boolean`

```vba
Public Function EnsureAppFolders() As Boolean
    ' [SCAFFOLD] The folder slot named in startup-conventions.md sections 2 and 4, and the
    '            extension point for open-time setup: add each new folder here as the
    '            application grows. Idempotent by design - the SAME routine creates a
    '            folder on the first run and verifies it on every open after that.
    ' [STANDARDS - startup-conventions.md] the two kinds of folder are ensured DIFFERENTLY,
    '            and getting this backwards fails silently. See the two blocks below.
    Dim sSharedFolder As String
    Dim sLocalFolder  As String

    On Error GoTo errHandler

    ' --- Shared-content folders: VERIFY, never create. ---------------------------------
    ' These hold files that a row in a SHARED table points at - photos, scans, attachments.
    ' There must be exactly one such folder for everybody, named by an absolute path.
    ' Creating a local one instead SUCCEEDS, which is the problem: the person who added the
    ' file sees it, nobody else does, every front end did as it was told, and nothing raises
    ' an error.
    ' [BUSINESS LOGIC] read each shared folder's absolute path from the settings table in
    '            the back end - e.g. tblAppSetting.OfficialPhotoFolder.
    ' >>> sSharedFolder = <the shared folder setting> per query-style.md <<<
    sSharedFolder = vbNullString

    If Len(sSharedFolder) > 0 Then
        If Len(Dir$(sSharedFolder, vbDirectory)) = 0 Then
            MsgBox "The shared folder this application needs cannot be reached:" & vbCrLf & _
                   sSharedFolder & vbCrLf & vbCrLf & _
                   "Files stored there will not open until it is available.", _
                   vbExclamation, "Shared folder not reachable"
            GoTo Cleanup
        End If
    End If

    ' --- Per-user working folders: create if missing. ----------------------------------
    ' These hold files belonging to one person and one session - exports, temporary output,
    ' a local log. One per front end is correct, so a relative path is right for them.
    ' MkDir creates ONE level only: a folder two levels deep needs one MkDir per level.
    ' [BUSINESS LOGIC] one block per per-user folder the application uses.
    sLocalFolder = CurrentProject.Path & "\Exports"
    If Len(Dir$(sLocalFolder, vbDirectory)) = 0 Then
        MkDir sLocalFolder
    End If

    EnsureAppFolders = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] standard errHandler block (see Startup)
    Resume Cleanup
    Resume
End Function
```

### CurrentBackEndPath — `Private Function` → `String`

```vba
Private Function CurrentBackEndPath() As String
    ' [SCAFFOLD] Where do this front end's links currently point? Returns the path held by
    '            the first file link found, or an empty string when there are none - which
    '            is how a single-file database identifies itself, without asking anyone.
    '            A link to an Access file stores its path as ";DATABASE=<path>"; an ODBC
    '            link stores something else entirely and is deliberately ignored here.
    Dim db  As DAO.Database
    Dim tdf As DAO.TableDef

    On Error GoTo errHandler

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If Left$(tdf.Connect, 10) = ";DATABASE=" Then
            CurrentBackEndPath = Mid$(tdf.Connect, 11)
            Exit For
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
    ' closed again in Cleanup a few lines later.
    Set dbBackEnd = DBEngine.OpenDatabase(sPath, False, True)

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
    Dim sConnect  As String
    Dim lRelinked As Long

    On Error GoTo errHandler

    Set db = CurrentDb
    sConnect = ";DATABASE=" & sNewPath

    For Each tdf In db.TableDefs
        If Left$(tdf.Connect, 10) = ";DATABASE=" Then
            If Len(sOldPath) = 0 Or _
               InStr(1, tdf.Connect, sOldPath, vbTextCompare) > 0 Then
                tdf.Connect = sConnect
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

## Standards Layer

- **Startup conventions** — the `AutoExec` → `Startup()` entry point, `Startup()` being a `Public
  Function` because `RunCode` cannot call a `Sub`, the idempotence rule, the relink convention, and
  the two kinds of folder all come from `startup-conventions.md`. A practice that starts its
  applications differently swaps that file and rewrites `Startup()` accordingly; the other eight
  procedures are unaffected.
- **Error handling** — the `errHandler`/`Cleanup` structure, the central logger, and the line-number
  policy come from `error-handling.md`. The `GlblErrMsg` call shown is the GPC default
  (house-specific); a forked practice substitutes its own logger, or the dependency-free `MsgBox`
  block, and may number lines or not.

  **One requirement this scaffold adds, and it is not optional: the logger used here must not write
  to the back end.** Every procedure in this module can run at a moment when the back end is
  unreachable — that is the situation the module exists for. A logger that writes to a table in the
  back end would fail inside the handler that was reporting the original failure, and the person
  would see neither. A log **file** beside the front end, or a message box, both survive it.
- **Query style** — the two `SELECT` statements and their quote-doubling follow `query-style.md`.
- **Naming** — procedure, variable, and parameter names, the `USys` prefix on a configuration table,
  and the `[Entity]ID` primary key follow `naming-conventions.md`.
- **Design principles** — one job per procedure is why the path accessors, the probe, the picker,
  and the relink loop are separate: the storage choice can change without touching the relink logic,
  and the probe can be reused anywhere the back end's health matters.
- **Audit columns** — `USysLocalSetting` carries none. `audit-columns.md` exempts `USys`
  configuration tables, which is the right answer twice over here: the house audit machinery lives
  in the back end, and this table has to work when the back end does not.

## Extra Options

*Empty in the base template. Filled per engagement; the filled copy is saved to the developer's own
library, not committed here.*

- **Close the application instead of sitting idle.** As written, a failed relink opens no form and
  leaves Access open with nothing in it. `Application.Quit acQuitSaveNone` after the final message
  is the firmer alternative — it removes any chance of someone poking at a half-started application,
  and it also removes their chance to read the message twice.
- **An explainer form instead of a message box.** A form can hold more than a message box
  comfortably will: who to contact, the path that was tried, and a *Try again* button.
- **Treat a folder problem as a warning, not a stop.** `EnsureAppFolders` currently returns False on
  a missing shared folder, which stops the application. An application where the folder matters to
  one screen out of twenty may prefer to warn and carry on.
- **Relink only when needed, silently otherwise.** Already the behaviour — worth stating to whoever
  reads the code, because a relink that works produces no visible sign at all.
- **A "where is my data?" menu item.** The same `EnsureBackEndLink` call on a button, so a move can
  be handled deliberately rather than waiting for the next failure.

## Parked / future considerations (not in this design)

- **A version check.** `startup-conventions.md` §1 names relink checks and version checks together,
  but they are different problems: a version check compares something stored in the front end
  against something stored in the back end, so it can only run *after* this scaffold has succeeded.
  It belongs in `Startup()` as a fourth step whenever it is built.
- **Two back ends.** An application linked to more than one data file needs `RelinkAllTables` called
  once per back end, with the old path supplied each time so the two sets of links stay apart. The
  parameter is already there; the calling logic is not.
- **ODBC links.** Deliberately untouched. A SQL Server link fails for its own reasons — a server
  name, a driver, credentials — and repointing it at a file path would make things worse.
- **Automatic discovery.** Searching likely folders for the back end rather than asking. Tempting,
  and the reason it is parked is the same reason the picked file gets validated: a search that finds
  the wrong copy is worse than a question that gets the right answer.
