# Lessons database and review form: how the AI assistant builds them

**Who reads this:** the AI assistant, when the developer asks to set up or review their lessons. The developer needs nothing from this file.

Two files make the lessons store, both in the developer's lessons folder (`set_lessons_store` chose it):

- `lessons.accdb`, the back end. The sync script `mcp-server/Sync-LessonsDb.ps1` creates it, with its six lookup tables and nine relationships, from `lessons/_lookups.md`, `lessons/_mechanisms.md` and the lesson files. The first lessons call runs the sync; never build this file by hand. A lessons database from the first version is not migrated: the sync stops with a message, and the developer deletes the file.
- `LessonsReview.accdb`, the front end. It links to the back end and finds it by looking beside itself. The files in this folder are its source; they are exported text, never edited by hand.

Only an Access MCP server can finish the front end, as for any build.

## Before you start

- `lessons.accdb` exists in the folder. If it does not, run any lessons call first.
- Never build over an existing `LessonsReview.accdb`. `New-LessonsReview.ps1` refuses.
- The folder must be a trusted location, or the form's code will not run. Say so in one line.
- Release `lessons.accdb` first. Nothing else may have it open.

## Build, in this order

1. Run `New-LessonsReview.ps1 -FolderPath <lessons folder>` in the 32-bit PowerShell (`SysWOW64`) when Office is 32-bit, otherwise the 64-bit one. It creates `LessonsReview.accdb` with the nine linked tables and the eight named queries.
2. Import `basLessonStore.bas` and `basLessonTerms.bas` as modules.
3. Run `EnsureStoreLinks`. Call it by its plain name: a call written `Module.Procedure` is not found. It checks that the nine links point beside the file.
4. Run `EnsureTerms`. It creates the help table `tlkpLessonTerm` and fills it.
5. Import the forms from their text files, in this order: `sfrmLesson_Event`, `frmLessonTerm_Help`, `frmLesson_Review`. A text file for a form is UTF-16, which is how it was exported.
6. Run `ApplyAppSettings`. It sets the application title, makes `frmLesson_Review` the startup form, hides the navigation window, and writes the developer's seven startup switches (status bar, shortcut menus, full menus, built-in toolbars, toolbar changes, special keys, datasheet schema, all on) into the file instead of leaving Access to default them.
7. Compile, then close the file and release it. Tell the developer it is ready to open.

## Check, on a copy

- The application title is `Lessons Learned from Builds` and the startup form is `frmLesson_Review`.
- The review form has 14 help buttons (`cmdHelp_`), and each names a term that exists in `tlkpLessonTerm`.
- Opening the help form with a filter on one term shows exactly one record.
- Approve, Reject and Proposed each save, and stamp who and when.
- A library lesson's text is locked; only its approval changes.
- The filter box offers All plus the three approvals, and All shows every lesson.
