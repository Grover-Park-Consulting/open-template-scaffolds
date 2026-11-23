---
step: 02-before-the-modules
title: "Open the database, ask the two questions the code depends on, and write the module declarations"
platform_facts: [open-existing-startup, mcp-module-import, mcp-line-numbers, vbe-line-continuation, vba-import-xml-entities]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Before the code

### Before you write the modules

**To the AI assistant.** `standards/error-handling.md` requires you to ask the developer how errors
are reported, and to ask it even when the answer looks settled. **That question is not one of this
template's nine wizard steps and is not meant to be** — the standards file owns it. Ask it through
the selection control before you write any of the four modules, and name it as a question this build
added rather than folding it into the template's count. The answer changes the handler in every
procedure you emit, so asking afterwards means writing them twice. The blocks below ship option 3
inline; anything else means replacing the handler in all four modules, not some of them.

**Look before you ask.** The question is the same on every build; the answers are not, and two of
them exist only if you go and see. Open the file wizard Step 2 named and find out what is there
before you compose the question — asking first offers three answers when there were four.

**What the answers are here.** The three in `error-handling.md` — and on Path B usually a fourth:
**the error handler this database already has.** A database that has been in use generally has one,
and it is usually the right answer, because the code you are adding should fail the way the rest of
the database already fails. Offer it by name if you find one. If it records the failure without
telling anyone, `error-handling.md` says what to do about that.

**When there is no logger at all** — none built here, none of the developer's own — options 1 and 2
are not available, because generated code has to compile on the machine it lands on. Offer the two
real choices: install a logger first, or use the message box now. Never quietly pick the message box
and report it afterwards.

**One procedure keeps its own handler whatever is chosen.** `BackupLongTextFieldsDM` runs inside
every save, called by the data macro itself. A handler that shows a message there interrupts someone
who was only editing a record, over a failure in bookkeeping they never asked for — and a handler
that blocks costs them the edit outright. It stays quiet: log silently if the chosen pattern can, and
never block. Anything else you add to that path follows the same rule.

**Whose name is recorded — ask this the same way, and look first.** The template's answer is
`AuditUser()`, which returns the Windows account name (schema Business Rule 9, and
`standards/audit-columns.md` names it). `CurrentUser()` is the named alternative in Extra Options.
**On Path B there is usually a third, and it is often the best one: the identity function this
database already has.** A database built by a team frequently has one, and it may return the
signed-in person's actual name rather than their Windows account. Find whatever fills the host's
existing tracking columns, offer it by name, and say what it returns that the template's answer does
not.

**If the host's function is chosen, use it in all six places, not four.** Extra Options lists four
sites and those four are the audit *log*. `BuildBeforeChangeMacro` calls the identity function twice
more, for the stamped `CreatedBy` and `ModifiedBy` on the *record*. Change the log's four, leave the
record's two, and you get two different names on one edit — the failure Extra Options already
describes. Choosing the host's function departs from `audit-columns.md`, which names `AuditUser()`:
it holds for this build and is not written back to `standards/`.

**The front-end placement requirement belongs to whichever function is chosen, not to
`AuditUser()`.** The generated macros call it on every audited table, and a macro firing because
somebody edited through a link resolves the function in **that person's front end** — absent, and
that front end cannot insert a row into a table this system stamps. Everything this template says about putting `modAuditLongText` in
every front end is that requirement wearing `AuditUser()`'s name. Choose the host's function and the
requirement does not go away: it moves to whichever module holds it, which is a module this template
does not own and will not place for you. Say which module that is, and tell the developer it now has
to be in every front end.

**And one thing stops being true when that happens.** `modAuditLongText` is described here as
required on every build, which holds because `AuditUser()` lives in it. Replace `AuditUser()` with
the host's function and it holds only `BackupLongTextFieldsDM` — needed where there is a Long Text
field, and nowhere else.

## Module-level declarations

Put these at the top of `modAddDataMacros`, below `Option Explicit`.

```vba
Option Compare Database
Option Explicit

' [BUSINESS LOGIC — scan boundary] Which tables this build audits: the answer to Step 4 of
' the wizard, and the only audit-scope decision that is written into code rather than set as
' data afterward. Set it before the module is imported; nothing at run time changes it, so a
' re-run of the scan always draws the same boundary as the first run did.
'   "List" — only the tables named in AUDIT_SCOPE_LIST, separated by semicolons: the list
'            the developer confirmed or gave. Chosen by what each table holds, not its name.
'   "All"  — every table in this file.
' Three kinds of table are out of scope under every mode, and IsAuditCandidateTable tests for
' them before it reads this setting: Access's own (MSys), the developer's own hidden tables
' (USys) unless named in the list, and linked tables, which cannot carry a Data Macro at all.
' >>> set these to the developer's answer before you import this module <<<
Private Const AUDIT_SCOPE_MODE As String = "List"
Private Const AUDIT_SCOPE_LIST As String = ""

' [BUSINESS LOGIC — schema Business Rule 4] Whether tables keyed by a Replication ID are
' audited. False is the normal case and what every build has produced until now: the log's
' key column is a Long Integer, and a table keyed by a Replication ID is reported as not
' auditable and left alone.
' Set it True ONLY when BOTH conditions hold — One_CheckAuditReadiness found at least one
' such table, AND the developer chose to audit those tables. It retypes
' tblAuditLog.PrimaryKey and tblLongTextBackup.PrimaryKey to Text(64), which makes every key
' in the log text, the ordinary whole numbers included, so the log then sorts by key as text
' rather than by number. A build that produces text key columns without both conditions
' holding is wrong.
' >>> set this to the developer's answer before you import this module <<<
Private Const AUDIT_GUID_KEYS_AUDITED As Boolean = False

' [SCAFFOLD] Width of the log's key column when the setting above is True. The widest value
' that can land there is 45 characters — DAO renders a Replication ID as
' "{guid {C6C0FB4C-...}}" where a Data Macro renders the same key as the bare 38-character
' "{C6C0FB4C-...}". Measured: both composite indexes over this column build at every width up
' to 255 and accept full-width data, so this is headroom rather than a limit being avoided.
Private Const AUDIT_KEY_TEXT_SIZE As Long = 64

' [SCAFFOLD] The local variable each generated After macro uses to carry the audited row's key.
' It exists because a Replication ID named directly inside a LookUpRecord WhereCondition stops
' that condition filtering at all — see the warning in the front matter. Every build emits it,
' whatever the key type, so there is one macro shape rather than two.
Private Const AUDIT_KEY_LOCAL_VAR As String = "varKeyText"

' [STANDARDS — audit-columns.md] The house audit column names, in ONE place.
' Three procedures below use them: Three_PopulateConfigTable (to seed them
' not-auditable), BuildBeforeChangeMacro (to stamp them, on the tables in
' AUDIT_STAMP_LIST), and Zero_CreateSampleTables (to create them on the sample tables).
' A shop that forks this library and renames its audit columns changes these
' four lines and nothing else. AccessTS is deliberately absent — it is a SQL
' Server rowversion and does not apply to a local Access table.
Private Const AUDIT_CREATED_DATE  As String = "CreatedDate"
Private Const AUDIT_CREATED_BY    As String = "CreatedBy"
Private Const AUDIT_MODIFIED_DATE As String = "ModifiedDate"
Private Const AUDIT_MODIFIED_BY   As String = "ModifiedBy"

' [STANDARDS — audit-columns.md] The tables whose four columns above this generator fills,
' separated by semicolons. The generator adds to what a table already does; it never takes
' over what already fills those columns. So a table is listed only where its existing Data
' Macro stamped them (the replacement must carry that on), where nothing fills them and the
' developer chose that the generator should, or where this build added them. A table whose
' columns a default or the database's own code fills is left out. Empty = stamp nothing.
' >>> set this from what you found, before you import this module <<<
Private Const AUDIT_STAMP_LIST As String = ""

' [SCAFFOLD] The mark that says this generator wrote a macro set. Every macro it emits
' carries the full form in a <Comment>, and MacroBackupIsOurs looks for the stable form
' before anything is removed. Public because BackupAndRemoveAllDataMacros, which decides
' what may be stripped, lives in another module.
' Two constants, not one: emit the version so an old build can be identified later, but
' match on the part before it, so releasing a new version never stops this system
' recognising work it did itself. Never edit either string and never reuse them
' elsewhere: the way back out depends on this mark being unique to macros we wrote.
Public Const AUDIT_MACRO_MARKER      As String = "OTS-AUDIT-DATAMACRO-GENERATED"
Public Const AUDIT_MACRO_MARKER_FULL As String = "OTS-AUDIT-DATAMACRO-GENERATED v0.9.3"
```

**To the AI assistant — before you import the module.** `AUDIT_SCOPE_MODE` carries Step 4's answer
and `AUDIT_SCOPE_LIST` carries the table names where that answer was a list. Set them here, in the
module text, before the import. There is no run-time argument for either and no dialog that asks
later: a module imported with the shipped values has an empty list, and the scan stops, saying no
table is in scope.

**Never fill either in from what the tables are called.** Step 4 is the only source for these two
values: the list the developer confirmed or gave, never one they have not seen.

**Write the list exactly as the developer confirmed or gave it**, one table name per entry, separated by
semicolons: `"tblClient;tblSupportTicket"`. Spaces around a semicolon are ignored, so a name may
itself contain spaces. A name that is not a table in this file simply never matches, and the scan
then reports fewer tables than the developer expects — read the names back to them before you set
this.

**To the AI assistant — Path B, before you import the module.** These four values name the
who-and-when columns, and `AUDIT_STAMP_LIST` names the tables where the generator fills them. A host
database that already carries tracking columns may name them something else —
`AddedOn`/`AddedBy`/`ModifiedOn`/`ModifiedBy` is one form you will meet, and there are others.

**The generator adds to what a table already does; it never takes over what already fills those
columns.** So first find out, table by table, how the host fills them today. There are three ways
and you cannot tell them apart without looking:

- **A data macro on the table.** The generator replaces a table's entire macro set, so the host's
  stamping would be lost. List the table in `AUDIT_STAMP_LIST`: the generated macro carries the
  stamping on, and the original is written to `DataMacroBackups\` first. Read the old macro before
  you replace it and tell the developer if it did anything the generated one will not.
- **A column default, or VBA** — a form's Before Update, a save routine, an import. Search the
  database's own modules and form code for the four column names. Leave the table out of the list:
  the generator touches neither defaults nor VBA, so they keep filling the columns as before.
  Stamping on top of them would overwrite every value they write, including any they set on
  purpose, such as an import that carries a record's original creator.
- **Nothing fills them.** Ask, through the selection control, whether the generator should.

**Say in the design spec, table by table, which of these you found and whether the table is
listed.** It is a change to how the developer's tables behave, so it is never left to the code.

**When the host's names differ from the four above, ask. Never settle it yourself.** Put it through
the selection control as its own question, and name it as a question this database added rather than
folding it into the template's own count. The two answers, and both belong in the question:

- **Use the names this database already has.** No columns are added. Set the four values above to
  the host's names. `Three_PopulateConfigTable` then seeds those columns not-auditable, and
  `BuildBeforeChangeMacro` stamps them on the tables in `AUDIT_STAMP_LIST`.
- **Use the library's names.** The host's four columns are replaced by the library's four. Leave the
  values above as shipped; they seed and stamp the library's columns. **Say what this costs before
  they choose it, because nothing here does it for them:** the four columns have to be added, the
  values already in the host's columns carried across, the old four dropped, and every form, query,
  report and line of the database's own code that names them corrected. **What it costs depends on
  what you found above** — a data macro is replaced anyway and a default goes with the column it sits
  on, but **their own VBA stops working the moment the old columns are dropped, because the columns
  it names are gone.** Check whether the tables hold data and say so when you ask: on an empty table
  this is cheap; on a table with history it is not.

**When a table carries no tracking columns, the build adds none:** attaching Data Macros does not
change a table's fields (`standards/audit-columns.md`). Offer adding them as a question, naming the
tables. Two things belong in it: `CreatedDate` and `CreatedBy` are required, so a table that already
holds rows needs the column added, filled, and only then marked required; and every row already there
will carry the date the columns were added rather than the date the record was really created,
because that history was never kept. Without them a table still gets its audit trail and simply has
nothing to stamp. A table that gets them goes in `AUDIT_STAMP_LIST`, since nothing else fills them.

Left as shipped on a host that names them differently, with nobody asked, they match no column: the
host's tracking columns get seeded as ordinary tracked fields, `BuildBeforeChangeMacro` emits no
stamping for any table, and because loading a macro set replaces the whole set, a table that arrived
with stamping loses it. A table whose stamping macro is left out of `AUDIT_STAMP_LIST` loses it the
same way. Nothing errors.
