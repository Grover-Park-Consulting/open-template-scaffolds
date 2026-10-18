# Lesson review form: how the AI assistant builds it

**Who reads this:** the AI assistant, when the developer asks to review their lessons. The developer needs nothing from this file.

The review form is a front end. It links to the lessons database and is built in the folder that holds that database, so it finds the database by looking beside itself. Only an Access MCP server can build it, as for any build. The files in this folder are the form's source; they are exported text, never edited by hand.

## Before you start

- The lessons database `lessons.accdb` is in the developer's lessons folder (`set_lessons_store` chose it, and the first lessons call built it). Build in that folder. Never build over an existing `LessonsReview.accdb` without asking.
- The folder must be a trusted location, or the form's code will not run. Say so in one line.
- Release the lessons database first. Nothing else may have it open.

## Build, in this order

1. Create `LessonsReview.accdb` in the lessons folder.
2. Import `basLessonStore.bas` and `basLessonTerms.bas` as modules, from their text files.
3. Run `EnsureStoreLinks`. Call it by its plain name: a call written `Module.Procedure` is not found. It links `tblLesson`, `tblLessonEvent` and `tlkpMechanism`.
4. Create three queries:
   - `qryLesson_frm`: `SELECT tblLesson.* FROM tblLesson ORDER BY tblLesson.LessonSlug;`
   - `qryLessonEvent_sfrm`: `SELECT tblLessonEvent.* FROM tblLessonEvent ORDER BY tblLessonEvent.LessonEventID;`
   - `qryMechanism_cbo`: `SELECT tlkpMechanism.MechanismID, tlkpMechanism.MechanismName FROM tlkpMechanism ORDER BY tlkpMechanism.MechanismName;`
5. Run `EnsureTerms`. It creates the help table `tlkpLessonTerm` and fills it.
6. Import the forms from their text files, in this order: `sfrmLesson_Event`, `frmLessonTerm_Help`, `frmLesson_Review`. A text file for a form is UTF-16, which is how it was exported.
7. Run `ApplyAppSettings`. It sets the application title and makes `frmLesson_Review` the startup form.
8. Compile, then close the file and release it. Tell the developer it is ready to open.

## Check, on a copy

- The application title is `Lessons Learned from Builds` and the startup form is `frmLesson_Review`.
- The review form has 14 help buttons (`cmdHelp_`), and each names a term that exists in `tlkpLessonTerm`.
- Opening the help form with a filter on one term shows exactly one record.
- Approve, Reject and Proposed each save, and stamp who and when.
- A library lesson's text is locked; only its approval changes.
