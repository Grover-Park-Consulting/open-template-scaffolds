---
step: 01-procedures-1
title: "The procedures, part 1 of 3 (to EnsureAppFolders)"
platform_facts: [open-existing-startup, mcp-module-import, mcp-line-numbers, app-startup-autoexec, vbe-line-continuation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

Open the database the developer supplied, as the facts below say. Change nothing before the developer has
given permission.

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. The **`errHandler` block is
shown in full in `Startup` and referenced thereafter** — the errHandler block's form
(`error-handling.md`) is identical in every procedure by design. Line numbers are deliberately
absent (house-specific; see `error-handling.md`).

**Each handler is the one `error-handling.md`'s table gives for the procedure's role**, and its first
comment names it. `Startup` is the entry point; `EnsureBackEndLink` and `EnsureAppFolders` are entry
points that hand their outcome back to it; `SaveLocalSetting` is an entry point nothing depends on;
`BackEndIsReachable` is an answer; every other procedure is a step. `BackEndIsReachable` is asked a
question whose answer is often "no", and a missing file, a dead server, or a file that isn't an
Access database all raise errors. It uses the **answering shape**: its handler sets False and resumes
to `Cleanup`, reporting nothing, and `EnsureBackEndLink`, which concludes there is no back end
anywhere, is the one that tells the person.

Module header:

```vba
Option Compare Database
Option Explicit

' [STANDARDS - error-handling.md] This module's own name, for its handlers. Copy IsRefusal
'            from error-handling.md into this module too, unchanged and Private.
Private Const ModuleName As String = "modAppStartup"

' [STANDARDS - error-handling.md] A refusal this module raises itself (refusal range).
Private Const ErrSettingTooLong As Long = vbObjectError + 1000

' msoFileDialogFilePicker, written as its number so this module needs no reference to
' the Office object library - one less thing to go wrong on someone else's machine.
Private Const FilePickerDialog As Long = 3

' The one setting this module keeps in USysLocalSetting.
Private Const BackEndPathSetting As String = "BackEndPath"
```

### Startup — `Public Function` → `Boolean`

```vba
Public Function Startup() As Boolean
    Const ProcName As String = "Startup"
    ' [SCAFFOLD] The only thing AutoExec runs, and the one place open-time work happens -
    '            including, once you switch it on, what your existing AutoExec did (step 0).
    '            The order below is not stylistic - see the comment at step 2.
    '            Returns False when the application must not carry on; each step has
    '            already told the user why.
    ' [STANDARDS - startup-conventions.md] RunCode can only call a Function, never a Sub,
    '            which is why this is a Function. Everything it does must be safe to run
    '            on every open, because it runs on every open.

    On Error GoTo errHandler

    ' 0. PROPOSED, commented out: what your existing AutoExec does today, translated to VBA
    '    from reading it. Your AutoExec has not been changed. Choose one:
    '    - Review these lines against your AutoExec, switch them on, and make
    '      RunCode Startup() your AutoExec's only action; or
    '    - Leave them off, and add RunCode Startup() at the top of your AutoExec, above its
    '      own actions.
    '    Doing both runs that work twice. Empty when there was no AutoExec, or it did no
    '    open-time work. AutoExec_NewCommandsToAdd holds the RunCode Startup() action.
    ' >>> the existing AutoExec's open-time work, each line commented out <<<

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
    ' [STANDARDS - error-handling.md] reporting shape: entry point; returns False when the
    '            application must not carry on. The person has been told why.
    ' Whatever reports here must not write to the back end, which may be what failed.
    If IsRefusal(Err.Number) Then
        MsgBox Err.Description, vbInformation
    Else
        LogError ModuleName, ProcName, Erl          ' option 2 puts its message box here instead
    End If
    Resume Cleanup
    Resume
End Function
```

### EnsureBackEndLink — `Public Function` → `Boolean`

```vba
Public Function EnsureBackEndLink() As Boolean
    Const ProcName As String = "EnsureBackEndLink"
    ' [SCAFFOLD] Make sure the linked tables reach a real back end, relinking if they don't.
    '            Four outcomes, in the order they are tried:
    '              a. no linked tables at all -> single-file database, nothing to do
    '              b. the current links work   -> record where they point, carry on
    '              c. the remembered path works -> relink to it silently, carry on
    '              d. neither works            -> tell the user, let them find the file
    Dim strCurrentPath As String
    Dim strStoredPath  As String
    Dim strNewPath     As String

    On Error GoTo errHandler

    ' (a) A single-file database has no file links. Not an error - there is nothing to check.
    strCurrentPath = CurrentBackEndPath()
    If Len(strCurrentPath) = 0 Then
        EnsureBackEndLink = True
        GoTo Cleanup
    End If

    ' (b) The links already work. Record where they point, so a later move has something
    '     to fall back on. This is what seeds the setting on a brand-new front end.
    If BackEndIsReachable(strCurrentPath) Then
        If LocalSettingValue(BackEndPathSetting) <> strCurrentPath Then
            SaveLocalSetting BackEndPathSetting, strCurrentPath
        End If
        EnsureBackEndLink = True
        GoTo Cleanup
    End If

    ' (c) The links are stale. Try the path we remember from last time.
    strStoredPath = LocalSettingValue(BackEndPathSetting)
    If Len(strStoredPath) > 0 And strStoredPath <> strCurrentPath Then
        If BackEndIsReachable(strStoredPath) Then
            If RelinkAllTables(strStoredPath, strCurrentPath) Then
                EnsureBackEndLink = True
                GoTo Cleanup
            End If
        End If
    End If

    ' (d) Ask. Name the path that was tried - it is the single most useful fact for
    '     whoever has to sort this out, and it costs nothing to include.
    ' [BUSINESS LOGIC] substitute your application's name for "This application".
    MsgBox "This application cannot find its data file." & vbCrLf & vbCrLf & _
           "It looked here:" & vbCrLf & strCurrentPath & vbCrLf & vbCrLf & _
           "Click OK and choose the data file to continue.", _
           vbExclamation, "Data file not found"

    strNewPath = AskUserForBackEnd()
    If Len(strNewPath) = 0 Then
        MsgBox "No data file was chosen, so the application cannot continue." & vbCrLf & _
               "Ask whoever looks after this application where the data file has moved to.", _
               vbExclamation, "Data file not found"
        GoTo Cleanup
    End If

    ' The guard that matters. A file picker will happily hand back last year's backup, or
    ' a copy sitting on this one machine - and everything would then work, quietly, against
    ' the wrong data. Check it is the right file before connecting anything to it.
    If Not BackEndIsReachable(strNewPath) Then
        MsgBox "That file is not this application's data file:" & vbCrLf & strNewPath & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, "Wrong file"
        GoTo Cleanup
    End If

    If Not RelinkAllTables(strNewPath, strCurrentPath) Then
        MsgBox "The tables could not be reconnected to:" & vbCrLf & strNewPath, _
               vbExclamation, "Could not reconnect"
        GoTo Cleanup
    End If

    ' Remember it only now - after a relink that actually worked. A remembered bad path is
    ' worse than no remembered path, because it gets tried first on every future open.
    SaveLocalSetting BackEndPathSetting, strNewPath
    EnsureBackEndLink = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] reporting shape: entry point that hands its outcome back. It
    '            tells the person itself and returns False; Startup acts on the False without
    '            reporting again.
    ' Whatever reports here must not write to the back end, which may be what failed.
    If IsRefusal(Err.Number) Then
        MsgBox Err.Description, vbInformation
    Else
        LogError ModuleName, ProcName, Erl          ' option 2 puts its message box here instead
    End If
    Resume Cleanup
    Resume
End Function
```

### EnsureAppFolders — `Public Function` → `Boolean`

```vba
Public Function EnsureAppFolders() As Boolean
    Const ProcName As String = "EnsureAppFolders"
    ' [SCAFFOLD] The folder slot named in startup-conventions.md sections 2 and 4, and the
    '            extension point for open-time setup: add each new folder here as the
    '            application grows. Idempotent by design - the SAME routine creates a
    '            folder on the first run and verifies it on every open after that.
    ' [STANDARDS - startup-conventions.md] the two kinds of folder are ensured DIFFERENTLY,
    '            and getting this backwards fails silently. See the two blocks below.
    Dim strSharedFolder As String
    Dim strLocalFolder  As String

    On Error GoTo errHandler

    ' --- Shared-content folders: VERIFY, never create. ---------------------------------
    ' These hold files that a row in a SHARED table points at - photos, scans, attachments.
    ' There must be exactly one such folder for everybody, named by an absolute path.
    ' Creating a local one instead SUCCEEDS, which is the problem: the person who added the
    ' file sees it, nobody else does, every front end did as it was told, and nothing raises
    ' an error.
    ' [BUSINESS LOGIC] read each shared folder's absolute path from the settings table in
    '            the back end - e.g. tblAppSetting.OfficialPhotoFolder.
    ' >>> strSharedFolder = <the shared folder setting> per query-style.md <<<
    strSharedFolder = vbNullString

    If Len(strSharedFolder) > 0 Then
        If Len(Dir$(strSharedFolder, vbDirectory)) = 0 Then
            MsgBox "The shared folder this application needs cannot be reached:" & vbCrLf & _
                   strSharedFolder & vbCrLf & vbCrLf & _
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
    strLocalFolder = CurrentProject.Path & "\Exports"
    If Len(Dir$(strLocalFolder, vbDirectory)) = 0 Then
        MkDir strLocalFolder
    End If

    EnsureAppFolders = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] reporting shape: entry point that hands its outcome back. It
    '            tells the person itself and returns False; Startup acts on the False without
    '            reporting again.
    ' Whatever reports here must not write to the back end, which may be what failed.
    LogError ModuleName, ProcName, Erl          ' option 2 puts its message box here instead
    Resume Cleanup
    Resume
End Function
```
