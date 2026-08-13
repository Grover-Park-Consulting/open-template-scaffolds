---
template: _materialization
title: Open Template Scaffolds — Materialization (table-schema + form-spec)
domain: _meta
type: spec
version: 0.7.0
status: draft
---

# Materialization — building real artifacts from a template

**Who reads this:** the AI assistant turning an approved design into real tables, code, or forms.

**If you are the developer:** read it to see exactly what will be done in your database. It is written as instructions only to the AI assistant, not for anyone else.

This is a **format/process reference** (like `_template-schema.md`): it defines how a template's
approved design becomes a real, buildable artifact, and **proves each mapping by hand** so the
template library MCP server's generator (phase B3) is never built against an unproven mapping.
It is meta, not a template
(`type: spec`, `domain: _meta`). Two builds are covered:

- **Table-schema → tables** — Access local tables via a VBA DAO `Sub`, or SQL Server via `CREATE TABLE` DDL.
- **Form-spec → a form** — importable Access form text (`SaveAsText` / `LoadFromText`).

**A requirement before any of this runs: the database has to sit in a folder Access trusts.** Outside
one, Access silently disables VBA — nothing is created, and the failure reports that a procedure
cannot be found rather than mentioning trust at all. This is not a case to design around; it is a
condition of using these templates, and a developer who cannot arrange it cannot use them. Say so
plainly, once, before building. What it means for a split application, and why the symptom can appear
on one person's machine alone, is under *Five rules that make the generated Sub actually run* below.

---

## The build record

**Every build leaves a build record**: one file, written beside the artifact, saying how it was
built. It is not optional, and it is not a summary of the conversation. **This applies to every
template and every build** — whether a wizard ran or the developer answered everything at once,
whether the artifact is tables, code, or a form.

It is also what `_template-schema.md` §10.4 promises the developer in your own words, and it is the
reason the detail can stay out of the messages they read while they are still deciding things.

**Name it `build-record.md`**, and write it in the folder that holds the artifact — beside the
`.accdb`, not in the library. Where a build touches two files, such as a front end and a back end,
one record covers both.

Six parts, in this order:

1. **What was built, and where.** Every object created, and which file it went into. The developer
   should be able to open the database and find each one.
2. **What was checked before building.** The state you found: what already existed, which references
   were present, whether the folder was trusted. This is the part that is worth nothing on the day
   and a great deal three months later, when something has changed and nobody remembers what it
   used to be.
3. **The decisions taken.** Every question the developer answered and what they chose, and any
   decision they handed back to you. Where a wizard ran, that is its steps and their answers.
4. **What was verified afterwards, and how.** The tests actually run, with their real results —
   never "tested and working". Say which paths were exercised, and name the ones that were not.
5. **Anything that did not match what the template said.** Divergences, surprises, and anything you
   worked around. This is the section a template author needs and nobody else will write.
6. **What is left for the developer to do.** Every follow-up the build could not complete, including
   anything the standards layer calls for that the build route could not deliver.

**Write it before you say the build is finished**, not when you are asked for it. A record written
later is written from memory, and the details worth keeping are the first ones to go.

**It records what happened, not what was meant to happen — good or bad, and whatever it reflects on
you.** A step that failed and was retried belongs in it. A test that was skipped belongs in it, named
as skipped. A build record in which everything went to plan is either untrue or not worth keeping.

**Two kinds of entry go missing, and they are the two worth most.** Anything the developer had to do
to unblock the run — a message they had to clear, a file they had to close. From where you sit that is
a moment of waiting; from where they sit the build stopped and demanded something, and it is the part
they will remember. Say what stopped, what they had to do, and how many times — not only a technical
note about the cause. And anything that went wrong because of how you worked rather than because of
their database: a file written in the wrong encoding, code that would not compile, a tool used the
wrong way. Both belong in part 5, with everything else that did not go as the template said. **Neither
is a new part of the record** — nothing here needs a seventh.

---

## The runbook — every build the developer finishes themselves

**A build the developer has to run any part of leaves a runbook**: one file, beside the code, saying
what to run and in what order. It is not optional, and it is not the build record. The build record
says what was done; the runbook says what they do next.

**When it applies:** the file-handoff route always, and any build that leaves a procedure for the
developer to run. Where you built everything directly and nothing is left to run, there is nothing
to write.

**Why it exists.** You know the order because you generated the code. They have a folder of files. A
procedure named `Three_…` tells them it is third; it does not tell them what the first two are,
which one needs an answer written into a table before the next will do anything, or which ones are
optional. **Naming a procedure is not documenting it**, and a developer who cannot tell an optional
tool from a required step either runs everything or runs nothing.

**Name it `runbook.md`**, and write it in the folder that holds the code.

Five parts:

1. **Before you start.** Everything that has to be true first, each one paired with what happens if
   it isn't: the folder trusted, the file writable, every object in the database closed, a backup
   taken. A precondition with no stated consequence gets skipped.
2. **Importing the files.** Which files, and where each one goes — in a split design, which belong
   in the back end and which have to be in every front end as well.
3. **What to run, in order.** Every procedure they run, numbered, each with: what it does, how to
   run it, what they should see when it worked, and what to do when it didn't. **Where a step needs
   them to go and look at something** — review rows in a table, set switches, check a value — that
   is its own numbered step, not a remark attached to another one.
4. **What is optional.** Procedures that exist but are no part of a normal run, and what each is
   for.
5. **How to tell it worked.** What they can open or run to confirm the result in their own database,
   without asking you.

**Write it in the second person and name every procedure exactly as it appears in the code.** A
runbook that says "run the setup procedures" has told them nothing they had not already guessed.

**Say it exists when you hand the files over.** A file in a folder nobody was told about is a file
nobody reads.

---

## Table-schema build — Access (VBA-DAO) or SQL Server

After a `table-schema` design is approved and the developer asks you to build it, **ask which platform**
and generate the matching artifact. Both carry the tables, fields (with their **comments**), keys,
indexes, relationships, and lookup **seed rows**.

### First, know which file you are building into

A **split database** is the normal shape for Access applications, especially those in multi-user
environments: one file holds the tables (the **back end**, usually on a shared network drive — never
on OneDrive, Dropbox, or any other file-syncing cloud folder, which corrupts a shared Access back
end), and each person runs their own copy of a second file holding the forms, reports, and code (the
**front end**), whose tables are *links* pointing at the back end. A single-file database — one
.accdb holding everything — is an acceptable choice for one user; everything below works there too,
and the distinction simply collapses.

When you build for a split application, every artifact on this page has a home, and putting one in
the wrong file fails in ways that are hard to see:

| Artifact | Where it goes | What happens if it's in the wrong file |
|---|---|---|
| Tables, indexes, relationships (the DAO build Sub below) | **Back end** — run the Sub in the file that holds, or will hold, the tables | Tables built in a front end are local to that one person and invisible to everybody else. Relationships cannot be enforced between tables in two different files at all. |
| Data macros (the audit-stamping macro below) | **Back end** — they attach to the table, so they attach where the table is | You cannot attach one to a linked table; the attempt is against the link, not the table. |
| **VBA functions a data macro calls** (e.g. `AuditUser()`) | **Back end *and* every front end** | The macro fires where the edit happened and looks for the function *there*. Missing in a front end, the stamp fails — and because the audit columns are `Required`, that front end cannot insert a row at all. |
| Forms, reports, queries, ordinary modules | **Front end** | — |
| `AutoExec` / `Startup()` | **Front end** — it runs when a person opens the application | Nothing runs at all: a back end is opened by the engine, not by a person. |
| A local settings table holding **where the back end is** (e.g. `USysLocalSetting`) | **Front end**, as a real local table — **never** a link | Put it in the back end and it is unreadable at exactly the moment it is needed: the application cannot find the back end, and the answer is inside the file it cannot find. |

The copies of a shared module are kept in step **by hand**; nothing enforces it. Treat the back
end's copy as the original and re-import it to each front end after any change.

### SQL Server → DDL

`CREATE TABLE` statements with primary keys, foreign keys, indexes, the lookup tables, and `INSERT` seed
rows. Field comments become inline comments or extended properties.

### Access (ACE) → a VBA `Sub` using DAO

Access local tables are **not** built from a `CREATE TABLE` query — they're built in a VBA `Sub` with
DAO, because that is the only way to carry a field's **`Description`** (the integral comments) and to set
AutoNumber, indexes, and relationships fully. This pattern is **proven by running it** (see the four
rules below); generalize it to any `table-schema`:

```vba
Option Compare Database
Option Explicit

' Build the <domain> tables in the current database, using DAO.
'
' RUN THIS IN THE FILE THAT HOLDS THE TABLES — the back end, if the application
' is split. Run in a front end, the tables are created there instead: local to
' one person, invisible to everyone else.
'
' RUN THIS FROM A TRUSTED LOCATION. Outside one, Access silently disables VBA
' and nothing is created.
'
' The errHandler here is self-contained (a message box) so the Sub runs with no
' outside dependencies. A shop with a central error logger swaps the errHandler
' block for its own (per standards/error-handling.md).

Public Sub BuildCatalogTables()
    Dim db  As DAO.Database
    Dim tdf As DAO.TableDef
    Dim rel As DAO.Relation
    Dim fld As DAO.Field

    On Error GoTo errHandler
    Set db = CurrentDb

    ' Safety guard: never overwrite existing tables.
    If TableExists(db, "tlkpMediaType") Or TableExists(db, "tblPublication") Then
        MsgBox "One or more target tables already exist. Build aborted.", vbExclamation
        GoTo Cleanup
    End If

    ' ---- tlkpMediaType (lookup) ----
    Set tdf = db.CreateTableDef("tlkpMediaType")
    AddField tdf, "MediaTypeID", dbLong, 0, True, True        ' AutoNumber primary key
    AddField tdf, "MediaTypeName", dbText, 255, True, False
    AddField tdf, "SortOrder", dbLong, 0, False, False
    AddPrimaryKey tdf, "MediaTypeID"
    db.TableDefs.Append tdf
    ' Descriptions must be set AFTER the table is appended (else error 3219).
    SetDesc db, "tlkpMediaType", "MediaTypeID", "Surrogate key"
    SetDesc db, "tlkpMediaType", "MediaTypeName", "Media-type label"
    SetDesc db, "tlkpMediaType", "SortOrder", "Display order"

    ' ---- tblPublication (entity) ----
    Set tdf = db.CreateTableDef("tblPublication")
    AddField tdf, "PublicationID", dbLong, 0, True, True      ' AutoNumber primary key
    AddField tdf, "PublicationTitle", dbMemo, 0, True, False
    AddField tdf, "MediaTypeID", dbLong, 0, False, False      ' foreign key to tlkpMediaType
    AddField tdf, "MultiVolumeSet", dbBoolean, 0, True, False
    AddField tdf, "AddedOn", dbDate, 0, True, False, "Now()"  ' audit default = field property, not DDL
    AddPrimaryKey tdf, "PublicationID"
    AddIndex tdf, "MediaTypeID", False
    db.TableDefs.Append tdf
    SetDesc db, "tblPublication", "PublicationID", "Surrogate key"
    SetDesc db, "tblPublication", "PublicationTitle", "The full title, untruncated"
    SetDesc db, "tblPublication", "MediaTypeID", "Format / media type (FK to tlkpMediaType)"
    SetDesc db, "tblPublication", "MultiVolumeSet", "True = a multi-volume set"
    SetDesc db, "tblPublication", "AddedOn", "When the row was created (defaults to Now())"

    ' ---- relationship: tlkpMediaType (1) -> (many) tblPublication ----
    Set rel = db.CreateRelation("tlkpMediaTypetblPublication", "tlkpMediaType", "tblPublication")
    Set fld = rel.CreateField("MediaTypeID")
    fld.ForeignName = "MediaTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel

    ' ---- seed rows for the lookup ----
    db.Execute "INSERT INTO tlkpMediaType (MediaTypeName, SortOrder) VALUES ('Book', 1)", dbFailOnError
    db.Execute "INSERT INTO tlkpMediaType (MediaTypeName, SortOrder) VALUES ('Periodical', 2)", dbFailOnError
    db.Execute "INSERT INTO tlkpMediaType (MediaTypeName, SortOrder) VALUES ('Map', 3)", dbFailOnError

    Application.RefreshDatabaseWindow
    MsgBox "Catalog tables built successfully.", vbInformation

Cleanup:
    On Error Resume Next
    Set fld = Nothing: Set rel = Nothing: Set tdf = Nothing: Set db = Nothing
    Exit Sub

errHandler:
    MsgBox "Build failed: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Sub AddField(tdf As DAO.TableDef, ByVal sName As String, ByVal lngType As Long, _
                     ByVal lngSize As Long, ByVal booRequired As Boolean, ByVal booAutoNum As Boolean, _
                     Optional ByVal sDefault As String = "")
    Dim fld As DAO.Field
    If lngSize > 0 Then
        Set fld = tdf.CreateField(sName, lngType, lngSize)
    Else
        Set fld = tdf.CreateField(sName, lngType)
    End If
    If booAutoNum Then
        fld.Attributes = dbAutoIncrField        ' AutoNumber = Long + auto-increment
    ElseIf booRequired Then
        fld.Required = True
    End If
    ' A field default is a property set BEFORE append (e.g. sDefault = "Now()") --
    ' never a DDL DEFAULT clause; the DAO/ANSI-89 engine rejects DEFAULT in CREATE TABLE.
    If Len(sDefault) > 0 Then fld.DefaultValue = sDefault
    tdf.Fields.Append fld
End Sub

Private Sub SetDesc(db As DAO.Database, ByVal sTable As String, ByVal sField As String, ByVal sDesc As String)
    Dim prp As DAO.Property
    Set prp = db.TableDefs(sTable).Fields(sField).CreateProperty("Description", dbText, sDesc)
    db.TableDefs(sTable).Fields(sField).Properties.Append prp
End Sub

Private Sub AddPrimaryKey(tdf As DAO.TableDef, ByVal sField As String)
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    Set fld = idx.CreateField(sField)
    idx.Fields.Append fld
    tdf.Indexes.Append idx
End Sub

Private Sub AddIndex(tdf As DAO.TableDef, ByVal sField As String, ByVal booUnique As Boolean)
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Set idx = tdf.CreateIndex("idx" & sField)
    idx.Unique = booUnique
    Set fld = idx.CreateField(sField)
    idx.Fields.Append fld
    tdf.Indexes.Append idx
End Sub

Private Function TableExists(db As DAO.Database, ByVal sName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = db.TableDefs(sName)
    TableExists = (Not tdf Is Nothing)
End Function
```

**Five rules that make the generated Sub actually run** — each learned by running it against a real
database:

1. **Trusted Location.** Stated as a requirement at the top of this file, and this is where it bites:
   outside one, Access silently disables VBA and nothing is created (a description-less failure).
   **In a split application this applies to every file that runs code, on every machine** — each
   person's front end as well as the back end you built the tables in. Trusted Locations are a
   per-machine Access setting, not a property of the file, so a front end copied to a new machine and
   dropped somewhere untrusted has its VBA disabled there and nowhere else. The symptom is an
   application that works for everyone except one person, for no visible reason.
2. **Descriptions after append.** Setting a field's `Description` during field-build throws **error
   3219**; create the `Description` property in a **second pass, after `db.TableDefs.Append`**.
3. **AutoNumber** is a `dbLong` field with `Attributes = dbAutoIncrField` (and no explicit `Required`).
4. **Defaults are a field property, never DDL.** Set `fld.DefaultValue` (e.g. `"Now()"`) **before**
   appending the field. Do **not** emit a `CREATE TABLE … DEFAULT …` statement — the DAO/ANSI-89
   engine rejects the `DEFAULT` keyword outright (proven: even `DEFAULT 'literal'` fails with a syntax
   error). This is the exact trap that turns a generated ACE-DDL build into "Syntax error in CREATE
   TABLE statement."
5. **VBA-only functions never go in engine-evaluated SQL or defaults.** A function the ACE
   engine doesn't recognize — `Environ()`, and most VBA-runtime functions — fails with
   **"Undefined function '…' in expression"** the moment it's embedded in an `INSERT` passed
   to `db.Execute` or set as a field's `DefaultValue`; the engine's expression service, not
   VBA, evaluates those. Resolve the value in VBA first and concatenate the literal. To stamp
   `CreatedBy` on a seed row, for example: `Dim sUser As String: sUser = Environ$("USERNAME")`
   then `db.Execute "INSERT INTO … (…, CreatedBy) VALUES (…, '" & sUser & "')", dbFailOnError`.
   (`CurrentUser()` *is* engine-known, but returns `"Admin"` without workgroup security — the
   resolved Windows user name is preferred.)

**Both index orders are valid.** The sample above appends indexes to `tdf.Indexes` *before*
`db.TableDefs.Append`; `templates/errors/error-logging-scaffold.md` appends the table first and builds
its indexes after. Both are proven by running them. Neither corrects the other — **don't rewrite
working code to match whichever you saw first.** Only field `Description`s are order-bound, and only
for the reason rule 2 gives.

Naming, audit columns, and types follow the active standards; the `errHandler` is the standards-layer
one (the **dependency-free message-box default** unless `error-handling.md` specifies a central logger).

### Audit-field stamping — the Before Change data macro

The audit columns (`standards/audit-columns.md`) are **Required** but can't be filled by a default —
`CreatedBy` needs the current user, and the engine can't evaluate `Environ()` in a default (rule 5),
so a generated table with a `Required` `CreatedBy` and no macro **rejects every insert**. Access's
answer is a **Before Change data macro**, which **cannot be built with DAO** — write it as a UTF-16
XML file and load it. Proven by running it against a generated table:

```vba
' Run in the accdb that holds the tables (the BE for a split app).
Public Sub BuildAuditStampMacro(ByVal sTable As String, ByVal sPK As String)
    Dim xml As String, fso As Object, txt As Object, sPath As String
    On Error GoTo errHandler
    xml = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    xml = xml & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/11/application"">"
    xml = xml & "<DataMacro Event=""BeforeChange""><Statements><ConditionalBlock>"
    xml = xml & "<If><Condition>IsNull([Old].[" & sPK & "])</Condition><Statements>"
    xml = xml & SetFieldXml("CreatedDate", "Now()") & SetFieldXml("CreatedBy", "AuditUser()")
    xml = xml & "</Statements></If><Else><Statements>"
    xml = xml & SetFieldXml("ModifiedDate", "Now()") & SetFieldXml("ModifiedBy", "AuditUser()")
    xml = xml & "</Statements></Else></ConditionalBlock></Statements></DataMacro></DataMacros>"
    sPath = Environ$("TEMP") & "\" & sTable & "_BeforeChange.xml"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.CreateTextFile(sPath, True, True)   ' arg3 True = UTF-16 (required)
    txt.Write xml: txt.Close
    Application.LoadFromText acTableDataMacro, sTable, sPath
    fso.DeleteFile sPath
Cleanup:
    On Error Resume Next
    Set txt = Nothing: Set fso = Nothing: Exit Sub
errHandler:
    MsgBox "BuildAuditStampMacro failed on " & sTable & ": " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Function SetFieldXml(ByVal sField As String, ByVal sValue As String) As String
    SetFieldXml = "<Action Name=""SetField""><Argument Name=""Field"">" & sField & _
                  "</Argument><Argument Name=""Value"">" & sValue & "</Argument></Action>"
End Function
```

The helper it calls lives in a standard module in the accdb **where the edit happens** (a data macro
can call a public function there — the only way to reach the Windows user, since `Environ()` is out
of engine reach). **In a split application that means the back end *and* every front end**: the macro
is attached to the table in the back end, but when it fires because someone edited through a *link*,
it looks for `AuditUser()` in **that person's front end**. Missing there, the stamp fails — and since
`CreatedBy` is `Required`, that front end cannot insert a row at all. This is the placement mistake a
single-file test can never catch, because in one file there is only one place for the function to be.

```vba
Public Function AuditUser() As String
    ' The fallback is not decoration: Environ$ returns an empty string in some contexts
    ' (a scheduled task, a service account, a locked-down profile), and CreatedBy is
    ' Required — an empty string would block the write outright, which is the one thing a
    ' stamping helper must never do.
    AuditUser = Environ$("USERNAME")
    If Len(AuditUser) = 0 Then AuditUser = "Unknown"
End Function
```

**Rules learned by running it:**
1. **Before Change fires before Required validation** — so the macro satisfies a `Required` `CreatedBy`.
2. **`IsNull([Old].[<PK>])`** is the INSERT-vs-UPDATE discriminator inside Before Change.
3. **A table's whole Data Macro set lives in one document** — `SaveAsText` exports it that way and
   `LoadFromText` replaces it that way. The **2010/12** namespace carries all five events
   (AfterInsert, AfterUpdate, AfterDelete, BeforeChange, BeforeDelete) together, and that is the form
   to generate. If you are emitting Before and After events for the same table, they share one
   document — there is no other way to attach both.
4. **Data macros cannot set Long Text (Memo) fields** — keep audit fields Short Text / Date-Time.
5. `CurrentUser()` is engine-known but returns `"Admin"` without workgroup security; `AuditUser()` gets
   the real Windows user.

**Stamping on its own vs. stamping alongside change-auditing.** The `BuildAuditStampMacro` above is the
proven form for a table that needs **only** the audit-column stamping — it emits one Before Change
event and nothing else. The moment a table also needs After Insert/Update/Delete macros (for example
`templates/audit/audit-logging-lite-scaffold.md`, which logs every field change), you cannot load the
two separately: per rule 3 the second `LoadFromText` replaces the whole set, taking the first with it.
Generate both jobs **together**, into one 2010/12 document, with the stamping actions and the
change-auditing actions sharing the same `IsNull([Old].[<PK>])` branch. The audit-logging scaffold's
`BuildBeforeChangeMacro` is the worked example.

### VBA code import — an import path can unescape XML entities

Proven by a real failure: a `vba-scaffold` module whose Data Macro builder emits **escaped XML
entities** (`&lt;`, `&gt;`, `&amp;`) as literal string content compiled and ran fine as a
standalone import, then threw error **3870** ("Microsoft Access cannot interpret the text you are
pasting as a data macro") on every table when the *same* source was imported through **an Access
MCP server's code-import tools**. The import path **HTML-unescapes entities on the way in** — a
literal `&lt;&gt;` in the source becomes a raw `<>` once
stored, and that raw `<` is then read as an opening tag inside `<Condition>…</Condition>`,
malforming the macro XML `LoadFromText` is asked to load.

**The rule this proves:** any scaffold whose VBA assembles XML (or HTML) containing escaped
entities must build those entities from character codes at runtime — `Chr(38) & "lt;" & Chr(38) &
"gt;"` for `&lt;&gt;`, for example — never write the escape sequence as literal text in the
module's source. A `Chr(38)`-built entity exists only in memory as the string `&lt;&gt;`; it is
never present as literal text for an importer to re-interpret. See
`templates/audit/audit-logging-lite-scaffold.md`'s `GetComparisonExpression` for the worked fix.

**The build route is asked, never assumed.** Where an Access MCP server is connected, say you have
it and ask — `CLAUDE.md` → "After approval — building it" carries the question, and its preferred
answer is `Use it`. **A connected Access MCP server is not authorization to use one:** presence is
not an answer, the developer's is. Where none is connected there is nothing to ask about — generate
the script, hand it over, and say that is what you are doing. **The template library MCP server
that ships in `mcp-server/` is not an Access MCP server** — it reads this library's files and
cannot build anything.

*This passage avoids the word "default" on purpose. It means both the route we would point at first
and what happens when nobody chooses, and only the first is true here (`_template-schema.md` §10.7).*

The entity corruption above is a reason to **check the imported source** where a scaffold warns
about it, not a reason to steer away from that route. Handing the developer a script to import the
ordinary way does avoid the failure mode entirely — but a scaffold that emits escaped XML should
assemble entities from `Chr()` codes either way, as a second line of defence.

### Opening a host database that has a startup routine — it can block, and it is not a dropped connection

**This applies to every template, not only the one that found it.** Any build against a database the
developer already uses has to open that database first, and opening it is where this goes wrong.

**What happens.** Opening a database runs whatever it is set up to run on open — an `AutoExec`
macro, a startup form, or both. If any of that puts a form on screen and waits for an answer, a
sign-in prompt being the usual case, the open never completes: an Access MCP server is not a person
and has nobody to answer it. The call sits there and eventually comes back as a **remote-procedure-call
failure**. Nothing is wrong with the file, and nothing is wrong with the server.

**Why this needs saying.** The failure arrives with no mention of startup, so it reads as the tool
being broken. It has already been misread once, as a restriction on launching Access as a separate
process, in a build that then took the file-handoff route for the wrong reason. And `CLAUDE.md`
tells you a dropped Access MCP server is yours to reconnect rather than a question to hand over —
correct in general, and here it will put you in a retry loop against something that was never going
to succeed. **Before you treat a failed open as a dropped connection, rule this out.**

**You may not be able to check first**, which is the awkward part: reading the startup settings
means opening the database, and that is the thing that blocks. The system catalog is not a way
around it either — an external read of it fails for an unrelated reason, on permissions, and that
failure is *not* evidence about startup. Two failures in the same minute can have two causes.

**So ask the developer, who knows the answer instantly**, and ask it in their world:

> **Ask:** When you open this database yourself, does it show you anything before you can use it — a
> sign-in, a menu screen, a form of any kind?

**If the answer is yes, tell them to turn it off. Do not offer a choice about it.** A startup routine
earns its place when people are using the application. Nobody is using it right now — it is being
built or changed, and the part that greets a person at the door is standing between the work and the
database. Switching it off for the duration is the ordinary thing to do, not a concession.

Say it as a direction, with the reason and the reassurance in the same breath:

> **Say:** That sign-in runs every time the database opens, including when I open it, and it waits
> for an answer I have no way to give — so it needs to be off while we work. Nothing about your
> application changes: you switch it back on when we're done, and I'll remind you. Two things to
> switch off, and both go back the same way:
>
> - The macro named `AutoExec`, if there is one. Rename it — `AutoExec_off` will do — and it stops
>   running. Renaming it back starts it again.
> - The form the database opens by itself, set in the database's own options. Clear it there.
>
> To get in and make those two changes, hold the Shift key down while you open the file. That skips
> both of them for that one open.

**Why this one is in the developer's own hands.** Everywhere else here, the assistant does the work.
This is a deliberate exception. Startup is how an application meets the people who use it, and
changing that on somebody's behalf is a step beyond any template's remit. Only two people can be at
the keyboard, and the reasoning holds for both. If they built the application, they already know how
to get past its own startup, so nothing is being asked of them that they cannot do. If they did not
build it — if they are changing somebody else's application without going back to whoever wrote it —
then switching the startup off themselves is the moment they take ownership of it. That is worth
requiring rather than doing for them.

**Put it back, and say so.** When the build is finished, tell them to restore both, in the same plain
terms. **Record in the build record what was switched off, that the developer switched it off, and
that it was restored** — an application that greets nobody at the door three months from now is a
defect somebody will lose an afternoon to, and this is the only written trace of why.

**Handing over the code is the last resort, and it is a real loss.** It exists for the developer who
*cannot* switch startup off — a database they do not administer, a sign-in their organization
requires — or who will not. Take it then, and only then. It is not the polite option and it is not
the cautious one: an AI-assisted template build that ends in a folder of files the developer has to
run themselves has handed back the thing they came for. Say plainly why you are falling back to it,
and put that in the build record too.

*Not to be confused with "Application startup — `AutoExec`, `Startup()`, and external file assets"
below, which is about the startup routine you **build** for a generated application. This one is
about a startup routine that is **already there**, in somebody else's database, getting in the way
of opening it.*

### Importing a VBA module through an Access MCP server — the encoding, and one wrong document

**A module goes in as a file, and the file must be UTF-8 with no byte-order mark.**

⚠ **The import tool's own documentation says the input must be UTF-16 LE. It is wrong**, and the
failure is quiet enough to cost a build: a UTF-16 LE file imports without reporting an error and
produces a module of null-padded garbage whose first characters are `ÿþA` — the byte-order mark
stored as text. Re-encoding the identical content as UTF-8 without a BOM imports correctly.
*(Observed 2026-08-02 building the audit-logging scaffold: a 2,954-line corrupt module on the first
attempt, clean on the second.)*

**A later build wrote the same content as ANSI and it imported correctly too, which proves less than
it looks.** For content that is entirely plain ASCII, ANSI and UTF-8-without-a-BOM are the same
bytes, so both "working" is one result, not two. The durable rule is the one stated above — **not
UTF-16 LE** — and UTF-8 without a BOM is what to write.

**Where that distinction starts to matter is the first non-ASCII character**, and a comment is the
likeliest place for one: a dash typed as an em dash, a curly quote pasted in from prose. Whether such
a character survives the import is **untested**. Until somebody establishes it, keep the code blocks
a scaffold ships ASCII-only — a plain hyphen in a comment costs nothing, and a module that will not
compile costs a build.

**The first line of each file is the module name**, in the form the editor expects:

```text
Attribute VB_Name = "modAddDataMacros"
```

followed by that module's `Option` header, then its procedures.

**One `Option` header per module, not per procedure.** A `vba-scaffold`'s `## Procedures` section is
written to be *read* — one fenced block per procedure — and some blocks carry their own
`Option Compare Database` / `Option Explicit` lines while others do not. Concatenating them yields
duplicate headers in some modules and none in others, and the compile stops on the duplicate. Strip
every `Option` line out of the bodies and emit one header per module.

**Not every fenced `vba` block is code to import.** The same file may hold illustrative fragments —
a call sequence showing what the procedures do in order, say — fenced exactly like real code. There
is no marker distinguishing them; read each block before including it.

**A failed import can take the COM session with it.** After the corrupt import above, the next two
calls failed with "This operation requires an open database" and "refers to an object that is closed
or doesn't exist." Closing the session and calling any tool re-established it, with no work lost.
Treat it as a reconnect, not a rebuild.

### Running a procedure through an Access MCP server — bare name, and use eval for arguments

Three separate failures, each observed while driving a `vba-scaffold`'s staged procedures.

**1. Never module-qualify the name.** `Application.Run "modAddDataMacros.One_CreateAuditTables"`
fails with "cannot find the procedure." `Application.Run` reads a dotted name as
*project*.*procedure*, not *module*.*procedure*, so a module-qualified name is never found. Pass the
procedure name on its own — `One_CreateAuditTables` — which resolves as long as that name is unique
in the project.

**2. To pass an argument, evaluate an expression instead of running a procedure.** A run-style tool
that marshals arguments separately can fail to pass a VBA `Boolean` at all, and a failed marshal has
been seen to kill the COM session outright — leaving the database held and the next call dead. Use
the Access MCP server's **expression-evaluation** tool and write the call as ordinary VBA:

```vba
Three_GenerateAllAuditDataMacros(True)    ' works — evaluated as an expression
```

This matters more than it looks. Staged procedures take an optional `bSilent` argument precisely so
an automated caller can suppress the message box and read the returned text instead; a caller that
cannot pass `True` gets the dialog, and with nobody at the keyboard it waits forever. **A scaffold
whose procedures accept arguments should say which tool to call them with**, or the first automated
run hangs.

**3. Expression-evaluation is not a general substitute for running a procedure.** Point 2 is narrow:
reach for it *when you need to pass an argument*. A bare zero-argument call through the same
evaluation tool has failed with **"Subscript out of range"** where the ordinary run tool handled the
identical call without complaint. Use the run tool by default; the evaluation tool is the exception,
not the upgrade.

### Code imported through an Access MCP server arrives without line numbers

Line numbers are added to VBA code by hand, or by a tool run over it in the editor. Neither happens
when an Access MCP server writes code straight into a database, so **every procedure built that way arrives
unnumbered**, and `Erl` in an error handler returns **0** instead of the line that failed. Nothing
breaks; the handler still fires and still reports the error number and description. What is lost is
the one thing line numbers buy: knowing *where* it failed.

Whether that matters is a house decision, not a rule this library sets — `standards/error-handling.md`
owns it. Say it out loud when handing over a module built through an Access MCP server, because a shop that relies on `Erl`
for diagnostics will otherwise lose it silently — and tell them the numbers can be added afterwards,
by hand or with a tool.

### After a build through an Access MCP server, confirm the file was actually released

A successful close reported by the Access MCP server is **not** proof the file is free. An `MSACCESS` process can
survive that close and keep the .accdb exclusively locked, leaving a `.laccdb` file beside it.
Before telling the developer to open the file and look at what was built — the natural next step
after any build — check for the leftover `.laccdb` and test an exclusive open; if the file is still
held, the surviving process has to be ended manually. This is intermittent: it happened on one build
and not on the next one against the same file, so test for it rather than assuming it either way.

**Check the file you actually built into** — for a split application that is the back end, not the
front end the developer usually opens. A held back end is worse than a held front end: it blocks
*everybody*, and the person who reports it is rarely the person whose machine is holding it.

### Application startup — AutoExec, Startup(), and external file assets

Per `standards/startup-conventions.md`, a generated Access **application** opens through one entry
point: an **`AutoExec` macro** whose only action is `RunCode Startup()`, and a `Public Function
Startup()` that runs open-time initialization. Build both whenever you materialize an app (not for a
bare table-schema or a single form in isolation). Typical `Startup()` entries include ensuring the
app's working folders (`EnsureAppFolders()`) and opening the app's startup form —
e.g. `DoCmd.OpenForm "AppStartupForm"`.

**Both belong in the front end**, along with everything else a person interacts with. `AutoExec` runs
when someone *opens* a file; a back end is opened by the database engine on behalf of a front end, not
by a person, so an `AutoExec` placed there never runs.

**Build the front end's own settings table at the same time.** A split application's `Startup()` also
confirms it can reach its data and reconnects the links if the back end has moved
(`standards/startup-conventions.md` §5) — and it remembers where the back end is in a **local table
in the front end**, `USysLocalSetting` by default. Create that table with the same DAO build Sub used
for any other table, but run it **in the front end**, and leave it unlinked. Four fields
(`LocalSettingID` AutoNumber PK, `SettingName` Text(50) unique, `SettingValue` Text(255),
`SettingDescription` Text(255)); no seed row is needed, because the first open that succeeds records
the path itself. `USys` is the naming-conventions prefix for a configuration table, and it has a
visible effect worth expecting: **Access hides objects whose names begin with `USys` from the
navigation pane** until *Show System Objects* is turned on. The runnable code that uses it is
`templates/startup/app-startup-scaffold.md`.

**External file assets — create the folder *and* copy the file in.** When a template stores a
**two-part reference** to an external asset — a file **name** in a table plus a folder from a
settings row, e.g. `tblOfficial.PhotoFileName` + `tblAppSetting.OfficialPhotoFolder` — the build must
do two things the reference alone does not imply:

1. **Ensure the folder** — create it at instantiation and re-ensure it idempotently at startup, via
   `EnsureAppFolders()` (`startup-conventions.md`). The reference is worthless if the folder never
   gets made.

   **In a split application, ask first whether the folder is shared or per-user**, because the two
   are ensured differently (`startup-conventions.md` §4). A folder holding files that a **shared
   table points at** — photos, scans, attachments — must be **one absolute network path** for
   everybody, and `EnsureAppFolders()` **verifies** it rather than creating a local one. Get this
   wrong and every front end obediently makes its own folder: the person who added a photo sees it,
   nobody else does, and **nothing raises an error**. A folder holding one person's temporary output
   is genuinely per-user and a relative path is correct for it.
2. **Copy the chosen file in** — the selection UI (a file-dialog picker) must **copy the picked file
   into the managed folder** under a controlled name, then store *that* name. Capturing the picked
   file's original name only points the record at a file outside the app, which blanks on retrieval.
   A collision-free convention such as `Official_<OfficialID>.<ext>` keeps one asset per record and
   lets a re-select overwrite cleanly. (This means the record must be saved first, so its key exists.)

**The AutoExec build gotcha** (sits beside the data-macro and DDL-`DEFAULT` gotchas above). DAO
**cannot** create a macro. `Application.LoadFromText acMacro` on modern Access **rejects both** the
modern `SaveAsText` XML **and** the `Version =20` classic text format — either throws runtime **2128
"errors while importing."** The accepted form is **classic macro text with `Version =196611`**:

```text
Version =196611
PublishOption =1
ColumnsShown =0
Begin
    Action ="RunCode"
    Argument ="Startup()"
End
Begin
    Action ="StopMacro"
End
```

Once you have that text, an Access MCP server's code-setting tool round-trips it directly (proven
with object type `macro`) — no FSO / UTF-16 file dance (unlike the data macros above).

---

*The rest of this document covers **form-spec** materialization.*

---

## Layout fidelity — a limitation of the approach

The generated layout is a **functional default**, not a reproduction of a real 2D form design —
multiple columns, multiple controls per line, landscape balance. A structural spec deliberately avoids
specifying detailed layouts, partly because of the variability of controls, design patterns and user
preferences, and partly in recognition of the fact that the adopter will implement their own preferred
layout in any case. **This is a limit of capturing forms as structural specifications, not a limitation
of the AI** — implementation-specific information can't be contained in the spec to begin with. When it
materializes a form, the AI **says so up front**: *"This is a functional default layout; arranging it
into your columns and lines is your styling pass."*

---

## Mapping rules

1. **Form shell** — `record_source` → `RecordSource`; `title` → `Caption`; default view.
2. **Regions → sections** — Form Header / Detail / Form Footer → the Access form's sections.
3. **Control row → control block** — name + type → `Begin <Type>`; `Bound to` → `ControlSource`; a
   lookup combo → `RowSource` (+ hidden bound column); Boolean → `CheckBox`; a subform → `SourceObject`
   + `LinkMasterFields` / `LinkChildFields`.
4. **Matching label — every user-facing control gets one.** Each data control (textbox, combo,
   checkbox, image, subform) is emitted with its own `Begin Label`, captioned from the field's friendly
   name (the inventory's intent, not the raw field name). The `form-spec` inventory lists only the data
   controls — the labels are **derived here, one per control, and must never be skipped.** A check box
   may carry its own caption and hidden controls need none, but emit the label anyway by default:
   deleting a spare beats hunting for a missing one.
5. **Default-layout rule** — single-field controls (textbox, combo, checkbox) stack vertically in
   inventory order (label-left / control-right, fixed row pitch). **Subform controls are never
   interleaved into that stack — they are placed to the *right* of it, or *below* it when width doesn't
   allow.** That is the *one* layout judgment the default makes; beyond it, no 2D optimization is
   attempted (see Layout fidelity).
6. **Code-behind** — event handlers wired to the named framework helpers, each with the standard
   `errHandler` block (`error-handling.md`; line numbering per the house policy). Helpers are **called,
   not defined**.

---

## Hand-validation — publication form (focused, importable fragment)

A representative slice of `templates/library/publication-form.md` materialized to Access text: the form
shell, a Detail section with a label + textbox, a lookup combo, and a subform placed to the right of the
single-field stack (rule 4), plus one framework-wired event handler.

```text
Begin Form
    RecordSource ="qryPublication_frm"
    Caption ="Library Catalog — Publication"
    Begin
        Begin Label
            Name ="lblPublicationTitle"
            Caption ="Publication Title"
            Left =120  Top =120  Width =2160  Height =480
        End
        Begin TextBox
            Name ="txtPublicationTitle"
            ControlSource ="PublicationTitle"
            Left =2400 Top =120 Width =6000 Height =960     ' Memo -> taller, multi-line
        End
        Begin ComboBox
            Name ="cboMediaTypeID"
            ControlSource ="MediaTypeID"
            RowSource ="SELECT MediaTypeID, MediaTypeName FROM tlkpMediaType ORDER BY SortOrder"
            ColumnCount =2
            ColumnWidths ="0;2880"
            Left =2400 Top =1200 Width =6000 Height =480
        End
        Begin SubForm                                       ' [rule 4] subform to the RIGHT of the stack
            Name ="sfrmPublication_Creator"
            SourceObject ="sfrmPublication_Creator"
            LinkMasterFields ="PublicationID"
            LinkChildFields ="PublicationID"
            Left =10680 Top =120 Width =6000 Height =1440
        End
    End
End
```

```vba
' CodeBehindForm
Private Sub cboSelectPublication_AfterUpdate()
100   On Error GoTo errHandler
110   With Me
120       .RecordSource = RefreshSQLWhere(.RecordSource, _
              IIf(.cboSelectPublication = 0, "1=1", "PublicationID = " & .cboSelectPublication), _
              "", "PublicationSortTitle")
130   End With
Cleanup:
140   Exit Sub
errHandler:
150   MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
160   Resume Cleanup
170   Resume
End Sub
```

*(The generator also emits the boilerplate `SaveAsText` headers — `Version` / `Checksum` / `NameMap` /
`PrtMip`. This fragment shows the meaningful mapping — control blocks, default positions, the
subform-to-the-right rule, and one standards-layer handler — which is what the hand-validation
confirms. A byte-perfect importable file is the generator's job in B3.)*

---

## Alternative path — build live through an Access MCP server

The same mapping drives an Access MCP server's form-creation and control-creation tools: instead of
emitting text for import, the generator creates the form and its controls directly, applying the same
default-layout rule and wiring the same code-behind. The `form-spec` markdown remains the source of
truth; both paths are generated targets.

**Create the label controls explicitly.** A control-creation tool makes only the control you name — it does
**not** auto-create an attached label — so each data control needs a second call for
its `lbl…` label (Mapping rule 4). Skipping this is why a live-built form comes up with no captions on
any control; every control the inventory lists must get its matching label.
