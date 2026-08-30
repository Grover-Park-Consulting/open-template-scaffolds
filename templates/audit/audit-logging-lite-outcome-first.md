---
template: audit-logging-lite-outcome-first
title: Access Audit Logging (Lite) — outcome-first method
domain: audit
type: outcome-first
version: 0.1.0
status: draft
implements: audit-logging-lite-schema
standards_layer:
  - audit-columns
  - design-principles
  - error-handling
  - naming-conventions
  - query-style
house_assumptions:
  - Audited tables — chosen by the developer, table by table, rather than inferred from anything
    about them. A practice that audits everything by default, or that decides by a naming
    convention, changes that.
  - Windows account name — the name recorded against a change. A practice whose database already
    knows the signed-in person by their real name changes that.
warnings:
  - This template attaches Data Macros to live tables. A build against a database in real use is
    preceded by a backup copy of the file, and the developer is asked for one before anything is
    changed.
---

# Access Audit Logging (Lite) — outcome-first method

**Who reads this.** Everything from *Intent* down to *Standards Layer* is written for the developer
whose database this is. The two sections after that are addressed to the AI assistant building it,
and each says so where it starts.

---

## Intent

**To produce, in an Access database, the results described in *What you end up with* below.** That
section is the specification. It says what the database does once this is built, in words a person
using the database can check for themselves, and it is followed by the checks that confirm the
result arrived and by the behaviours that must hold however the work was divided up. Those three
sections together are the whole of what this template promises.

**Nothing else is here.** No procedures, no module names, no code. This template states the result
and leaves the route to whoever builds it. The other version of this template — the rules-based
method, `audit-logging-lite-scaffold` — produces the same result from working code you import.
Either one can be run against your own database, and they can be run one after the other on
separate copies to compare.

---

## What you end up with

Every change to the tables you choose to audit is recorded, by the database itself, regardless of who makes it and
how they make it. Somebody editing through a form, somebody typing straight into a table,
somebody running a query that updates a thousand rows at once — all of it lands in the same log.
The recording mechanism is attached to the tables rather than to any screen or any code that reaches
them. **No one can switch auditing off for their own session. It happens automatically, every time.**

**The log answers two questions about each change: what did this field hold before, what does it hold now.**
One row per field that actually changed. It does so by naming which table, which record, which field. It does so
whether the record was created, changed or deleted: the value before, the value after, when, and who. A record deleted
outright is logged field by field, so what it held is still there afterwards.

**Nothing is recorded for a field that did not change.** Editing one field of a forty-field record
produces one row in the log, not forty, because the 39 unchanged fields aren't recorded. A person opening a record, looking at it, and closing it produces nothing in the audit log at all.

**Long text is handled separately.** Access Data Macros cannot retrieve the previous contents of a
long-text field. Therefore, the template captures the old value before the change occurs and retrieves
it for the log after the change. That way, the log holds values from long text fields the same as any other field.
Without that special handling, the one field where people write the longest text strings would be the one field whose
history you could not see.

**Your tables may already carry four fields to identify who created each record and when, and who last changed it and when.**
These are the four stamping columns your standards layer names. Those columns, when they exist, are filled by the same automatic Data Macro behaviour, on every table in scope, including a table where you have switched audit logging off.

**You decide what gets audited, one field at a time. Your decisions are recorded in a table.**
It is data, not code. Changing your mind later means editing rows and running the build again, not rewriting anything.

**There is a way back out.** Everything this template puts on your tables during a build can be taken off again. The template saves what was on them beforehand to a file before it does anything so you can put it back yourself if you want to.

**You can try the whole thing on made-up tables before it goes anywhere near your own.** Simply tell the template which way to build.

### How you validate the template's output

You can validate that the template produced the results we promised. Perform each of these checks. Successfully completing these checks indicates the template ran as expected. Each validation check can be expressed as a yes/no question. **Validate by asking: Does this happen or not?**

Perform these checks on a copy of the database after creating audit logging on it.
Where your database is split into two files, you need both of them.

1. **A change in one field is recorded as one row.**
   - Open one of the audited tables
   - Change one field of one record
   - Save the record
   - Open the log to confirm
   - The log has one row for
     - that table
     - that record
     - that field
     - marked as a change
     - showing what the field held before and what it holds now
     - stamped with the time and your name
2. **Only what changed is recorded.**
   - Change two fields of one record at once
   - Open the log to confirm
     - it has exactly two rows for that change
   - Move to a different record
   - Change nothing
   - Move off that record
   - Open the log to confirm
     - it has no new rows
3. **A new record is recorded.**
   - Add a record to a table
   - Open the log to confirm
     - there is a row for each field that was filled in
     - each row is marked as a creation
     - the "before" column is empty
     - each row is stamped with the time and your name
4. **A deleted record leaves its contents behind.**
   - Delete a record
   - Open the log to confirm
      - there is a row for each filled in field the deleted record held
      - marked as a deletion
      - showing what each field contained
   - The record is gone; however, what it held is still readable in the audit log.
5. **Long text survives being deleted.**
   - Repeat check 4 (delete a record) on a record with a long-text field which had something in it
   - Open the log to confirm
     - the log holds the full previous contents of the long-text field, not a blank and not a truncation of it
6. **Long text survives being changed.**
   - Edit a long-text field to something different and save
   - Open the log to confirm
     - the log shows the whole of what was in that long-text field before the edit
     - the log shows the whole of what is in that field after the edit
7. **Stamping columns fill themselves.**
   - If your tables also carry their own stamping columns, confirm they were not changed.
   - Add a record without touching those four columns.
     - Confirm that they fill in.
   - Change any field of that record
     - the "last changed" pair updates
     - the "created" pair does not
8. **A table for which you switched off audit logging is not logged, but stamping still works.**
   - Pick a table
   - Switch every one of its fields off in the configuration table
   - Build again
   - Change one or more fields in a record in that table
   - Open the log to confirm
     - nothing appears in the log
   - Confirm in the table itself
     - its *changed* columns fill in
   - **Add a new record to that table**
   - Confirm in the table itself
     - its *created* and *changed* columns fill in
9. **Turn logging on again.**
   - In that table switch a field back on
   - Build again
   - Change that field
     - Open the log to confirm
       - the change is recorded
   - Nothing you did earlier has to be undone first.
10. **Editing through a form is recorded the same way as editing in a table.**
    - Open a form bound to an audited table
    - Edit a record through it
    - Open the log to confirm
      - the log rows are the same as in check 1
    - If your database is split, do this with a form in the front end.
11. **A bulk update is recorded.**
    - Run an update query that changes several records at once
    - Open the log to confirm
      - every changed field of every changed record is in the log
12. **The way back out works.**
    - Use a copy of the database you can throw away
    - Remove the audit logging Data Macros from the tables
    - Confirm
      - the tables still accept records
      - a file holding what was there before has been written
      - a table which had Data Macros of your own — unrelated to this template's Data Macros
        - if the build was stopped, those Data Macros are intact
        - if the build ran, those Data Macros are saved to a file you can read
13. **A table that cannot be audited is reported, not attempted.**
    Some tables are not auditable. This validation check applies only if you have a table like this.
    Such tables include
    - those whose primary key is text
    - those whose primary keys consist of more than one field (composite keys)
    - those which have no primary key
    - The build names unauditable tables and tells you why, before anything is changed.
14. **When a build is blocked, it stops rather than leaving your tables half-finished.**
    An open table or form in the database to which you are adding audit logging blocks the template
    - Open the database
    - Leave a table or form open
    - Start the build
    - Confirm that the build
      - stops
      - names what is open
      - changes nothing

### The same behavior every time, not the same structure

For an AI assisted template, two runs need not produce the same code. They need to produce a database
that behaves the same way: the same things are recorded, in the same situations, for the people using it.

The audit logging work may be divided into different procedures, given different names, or commented differently.

**Comparing one run against another is not the test.**

The following **behaviors** must be true of every run. When you see these results, you know the run was successful.
You can confirm these behaviors with the validation checks listed above.

- One log row per field that actually changed. A field whose value is the same before and after
  produces nothing, and an edit that changes nothing produces nothing.
- A value that was in a field in a table before a change or a deletion is in the log afterwards, whatever its type,
  including the long-text types Access cannot hand to Data Macros directly.
- Every log row says what event produced it: create, change or delete.
- Every row carries the time and the name of the person responsible for that edit.
- **[your standards]** When present, the four created-and-changed stamping columns are filled by the database
  itself, on every table in scope, including tables with all auditing switched off. A table in scope
  that ends up without that behaviour would reject every insert, because those columns are required
  and nothing else can supply a person's name.
- **[your standards]** The name recorded on a log row and the name stamped on the record are the same
  name, produced the same way. One edit never writes two different people into two places.
- Whatever supplies that name must be reachable from every file a person edits through, not only from
  the file holding the tables. Where a database is split, that means it exists in both the back end
  and in every copy of the front end.
- What is audited is held as data in a configuration table the developer can read and edit, one row per field.
  The build never identifies audited tables and fields by reading a table's or field's name.
- Field types the database cannot meaningfully audit are never audited. This can't be overridden.
  Those field types are
  - attachment fields
  - calculated fields
- Building again replaces what a previous build attached; it never adds a second copy alongside it.
  Running the build twice leaves the same result as running it once.
- Data Macros a table already carried, which you wrote for reasons of your
  own, are saved to a file before they are replaced. The tables affected are named in the build report.
- The system identifies the Data Macros it adds to tables by incorporating a mark. The system can only remove
  Data Macros carrying its mark.
- Access's own hidden tables are never exported or changed. Hidden tables belonging to the
  developer are left alone unless the developer opted those tables into auditing.
- The whole run only happens when every object in the database is closed. Each step that needs exclusive control
  confirms it and stops rather than half-finishing.
- **[your standards]** Nothing in the audit logging path interrupts a person while editing data. Audit logging is
  bookkeeping that runs when somebody saves a change. A message box there would stop a person who was only
  changing a record because of something they did not ask for and cannot act on.
- Before running a build against a database in real use, best practice is to make a backup copy of the file.
  The template asks for one before changing anything.
- Every audited table has a single-field number primary key (usually an AutoNumber). The template reports tables
  that do not qualify for audit before the build rather than failing during it.

Lines marked **[your standards]** come from your standards layer rather than from this template, and
move with that layer if your shop replaces it. Everything else in `standards/` applies here as it
does to every template.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences about any of them,
say so while the design is being worked out, before anything is built. Where you don't choose, the build will choose based
on the rules built into it. The template's promise holds either way.

- How VBA code is divided into procedures, what they are called, and how many there are. They will be functionally equivalent,
  although not necessarily structurally the same.
- Names for the log table, the settings table and the staging table, and the names of their columns, within whatever your naming
  rules already require.
- How the build identifies which of your tables to offer to audit in the first place.
  **A build must not decide on its own what to change or remove by reading the table and field names.**
  You can choose from
  - a list of tables which you confirm
  - a table naming convention you specify
  - every table in the file
- Whether the settings table starts with auditing switched on or switched off for everything.
- The wording of everything the developer sees.
- Whether the build reports in message boxes, as returned text, or both.
- How the previous contents of a long-text field are held between the moment before a change and the
  moment after, and how long they are kept afterwards.
- Where the saved copies of replaced Data Macros are written, and whether anything ever tidies them up.
- How the code is laid out and commented, within whatever your standards already require.
- The order in which you are asked for the things only you can supply.

**If you want something not already in the template, you can have it.**
Our promise is specified in *What you end up with*, together with the list under *The same behavior every time, not the same structure*. The
fourteen checks are how you confirm you got it in any given build.

Everything in this section leaves that promise intact.

Asking for something different is a different act, an extension of the basic build.

You are welcome, even encouraged to experiment, now or later. If you do so, the result is then your responsibility rather than this template's. The validation checks above may no longer describe what you have.

### Facts about the platform

- **Data Macros.** Access can attach automatic behaviour to a table itself using Data Macros. This behaviour occurs regardless of who makes
  the change to data and how they make it. It is the only mechanism in Access that a person editing a table cannot bypass.
  Audit logging takes advantage of this by creating and attaching Data Macros to tables you select.
- **The ordinary VBA programming interface cannot programmatically create a Data Macro.** They can
  only be written programmatically as an XML document and loaded into the table. They can be read back out the same way.
- That XML has to be written in the two-byte character encoding, **UTF-16**. Written any other way it
  loads without complaint but does nothing.
- Loading XML for a table's Data Macros **replaces the Data Macros already attached to that table**. It never merges. Two separate
  loads means the second wipes out the first. Everything a table needs has to be built into one XML document.
- The table has to be held open in design view while the load happens, or the change does not stick.
- Reading the attachment back out also produces UTF-16, so anything reading that file has to be told
  to expect it.
- **Access rewrites part of the XML document on the way out.** The version identifier it returns is not
  the one that went in, so an exported file never matches its input exactly. Nothing depends on it,
  but anyone comparing the two will see a difference that is not a fault. Behavior, not structure, determines success.
- **A comment placed inside a Data Macro survives being loaded and read back, character for character.**
  It holds whether the comment sits before a plain action or before a conditional block.
  Therefore, the template takes advantage of comments to identify its own work because it is the only reliable way to tell
  whether a given attachment is one this build wrote.
- Which tables carry Data Macros can be discovered from Access's own System table. That System table
  also lists Access's own tables, which carry attachments of their own. The template can read from the System table, but must never attempt to change anything in it.
- **Access cannot hand the previous contents of a long-text field to its own Data Macro.** This is a limitation in Data Macros.
  The data types affected are the long ones — the memo type and anything built on it. This shapes the design for long-text fields: the old value in a long-text field has to be captured a moment before the change, held somewhere, and retrieved for the audit log after the change.
- **A hyperlink field is a long-text field.** Access reports it as one and it needs the same handling,
  although a developer sees "Hyperlink" in the table designer. A list that refers to a hyperlink field as a long text field names
  something the developer cannot easily find. The two are told apart by a flag on the field rather than by its type. Hyperlink fields are audited the same way ordinary long-text fields are audited.
- **Attachment fields and calculated fields cannot usefully be audited.** A calculated field cannot be
  edited directly. Only its components are ever changed. Therefore, calculated fields themselves are not auditable.
- **Data Macros can call functions (but not subs) in the file where they run.** They look for functions called that way in
  **the same file the person is editing from**, not the file holding the tables. In a split database that is each person's own front end. If the functions are not found in the same file, the change fails outright and nobody can add a record at all.
- Attaching Data Macros to a table requires a design lock on that table, and Access refuses one while anything is using it.
  An open form bound to the table is one cause. **If we allowed a failure to bypass that table, the run would
  carry on to the next, so what you would get is some tables done and others not, with no apparent sign that
  had happened.**
- Access can be asked whether a given form or report is open, and whether a given table or query is
  open, without opening anything. Asking about a name that does not exist returns "not open" rather
  than failing.
- **A file marked read-only by Windows opens in Access anyway**, with no warning, and fails only when
  something tries to write.
- Certain characters have special meaning in the XML document and have to be written differently
  inside it. A value containing one, written literally, makes the document unreadable and the load
  stops.
- **Line number must be within the body of a Function and in front of certain statements. Otherwise the VBA will not compile.**
  Access rejects line numbers placed between the start of a multi-branch decision and its first branch, though the same numbering is accepted everywhere else.
- Importing a code file saved in the two-byte encoding can appear to succeed while producing
  nonsense. A module that imports and then reports no procedures has not really imported.
- **Code in a file that is not in a trusted location does not run.** Access may not say so; it may
  report that it cannot find the procedure, which can send people looking at the code. Trust is a
  setting on each machine.
- A database that several people share must not sit in a folder that syncs to the cloud, e.g. OneDrive, DropBox and similar cloud folders.
  Those folders copy a whole file once it stops changing. That does not support several people writing to one database
  at the same time. The file can be corrupted rather than merged.
- Access's own session user name and the Windows account name are different things and are obtained
  differently. A database with no workgroup security reports the same placeholder for every person
  under the first of the two.

### What the template does not do

Some of these features are left out by design; we chose not to include them.

Some of these are things the template cannot do — a limitation rather than a decision.

Either way nothing here does them. The template build won't tell you what it isn't going to do, only what it will do.

- **Putting a change back.** The log records what a field held. You can add your own feature to restore logged changes, but that's not in the base template.
- **Removing old log rows.** The log grows for as long as the database is used. You can add a feature to prune it although it is
  not in the base template.
- **A screen for reading the log.** The audit log is a table. It can be queried. You can add your own form for it if you wish.
- **Auditing a table whose primary key is text, or made up of more than one field, or absent.** The build reports such tables and leaves them out of it. You can adapt the design to a different kind of key. If you choose to do so, it becomes your adaptation, not subject to the promise of the template.
- **Auditing attachment fields or calculated fields.** You cannot override the setting for these fields. Auditing them would be
  an advanced feature worth considering, should you feel ambitious.
- **Monitoring your tables for design changes since the last build.** If you add a field to a table, that field is not audited
  until the build is run again to include it in the audit. Nothing watches for this and nothing reminds you.
- **Recording a change made while the auditing was switched off for that field.** A gap in the log looks exactly
  like a period when nothing was edited.
- **Protecting the log from being edited.** The log is an ordinary table in the same database. Anyone
  who can open the database can change or delete its rows. Whether that matters depends on who has
  access and to what degree you secure your data.
- **Recording when someone merely looked at or read something.** Only changes are recorded, not views.
- **Keeping the copies of this system's code in step across a split database.** If the code in one copy
  of the front end changes, you must update other copies of the front end to match. A front end
  running an older copy of a build fails only when differences in the audited tables change how data is tracked.
- **Telling you that a front end is missing that part.** The failure appears as an insert that will
  not work, at the moment somebody tries.
- **Auditing tables that live in another file the front end also links to.** The build runs where the
  tables are.
- **Anything about a database server.** This is Access's own mechanism, on Access's own tables. Audit logging for
  a database server is often easier or more powerful, but it is not available within Access.
- **Restoring Data Macros of your own that this template replaced.** Existing Data Macros are saved to a file and named in
  the report. Putting them back is yours to do if you wish, after the template build is complete.
- **Recognising Data Macros written by a version of this system that predates the mark it now writes.**
  Adding a mark to Data Macros was added to a later version of the template. Data Macros added in an earlier version
  are treated as somebody else's work and left in place.
- **Preventing two people building this into the same database at once.** The whole run assumes one
  person with everything else closed. That's not under the control of the template. You are responsible for how you work.

---

## Information and conditions you need to supply

You need to provide information about seven things; we make no assumptions or guesses.

1. **Whether you are trying this out on a new database or building it into a database you already use.** The two behave
   differently from the first step onward. We can't reliably determine facts about your database to answer this.
   A database full of real data may be a copy you are experimenting on or it may be a production database.
   **We don't know and cannot guess.**
2. **Which file holds the tables.** When your database is split, that is the file holding the
   data, which is generally referred to as the back end, not the one holding the forms, code and queries, referred to as the front end.
3. **Whether you have a backup copy** if this is a database you already use. This template changes live
   tables. If you tell the template there is no backup, the build stops rather than continuing. We won't risk your data.
4. **Which tables to audit.** You confirm the list before we make changes. We do not work out what to audit
   from what your tables are called. A shop's tables may be named anything at all and any guess we make would fail silently.
5. **Which fields within the audited tables to audit.** The selection of fields to audit is switched on or off one at a time.
   You do this by manually editing the settings table rather than by answering a question. Nothing else does it
   for you. If you skip this step in a build, it audits everything or nothing.
6. **Confirmation that the long-text fields found are the ones you expected.** Hyperlink fields also appear
   here, because Access holds them as long text; that is not an error. You confirm that the build has correctly identified those fields in your database.
7. **Permission to change your tables.** We ask you immediately before changing anything.

**The accdb must be in a Trusted Location.** Accdb files have to sit in a folder Access trusts. That's called a Trusted Location. Trust is a setting on each machine, not a property of the accdb files Access uses.

Code in an Access file anywhere else does not run at all. Access normally informs you about Trusted Locations when it opens an accdb file, although that is not the same as what happens in the template. The build reports that it cannot find the procedure it needs. Either way, you must put the accdb file in a trusted location for the build to succeed.

**Nothing here works without a trusted location, and nothing here can arrange it for you.**

---

## Standards Layer

**Your standards layer decides everything about how this is built. This template decides only what
it has to do.** The rules-based method carries working code with the house-specific parts marked for
substitution, so its standards layer overrides particular lines. This one carries no code at all,
so there is nothing here for a standards rule to override: whatever `standards/` says is simply what
gets written.

- **`audit-columns`** — the names of the four columns recording who created a record and when, and
  who last changed it and when. This template never names them. The lines marked **[your standards]**
  under *The same behavior every time, not the same structure* are promises that exist only where
  this layer asks for those columns.
- **`naming-conventions`** — what the log table, the settings table, the staging table and their
  columns are called. Those names are listed as free choices under *Free to choose alternatives*,
  which means free within whatever this layer already requires.
- **`error-handling`** — how a generated procedure reports a failure. One constraint on the answer is
  not a preference and is stated as a promise above: nothing in the recording path interrupts the
  person editing. A message box raised while somebody saves a record stops them for something they
  did not ask for and cannot act on.
- **`query-style`** — how the generated code writes and holds its SQL.
- **`design-principles`** — how the work divides into procedures. *Free to choose alternatives*
  leaves that division open on purpose, and this layer is what it is open to.

A shop adopting this library replaces `standards/` with its own. Nothing in this file changes when
they do, which is the point of keeping the two apart.

---

## Standards Gate

**To the AI assistant:** before anything else, run the standards gate — `templates/_standards-gate.md`,
in full. It is one question in the ordinary case, and it settles whose rules govern this build. It
comes first: the disclosure line, then the gate, then this template's house assumptions and its
warning, then the seven things the developer has to supply.

Both versions of this template run the gate — this one and the rules-based method
(`audit-logging-lite-scaffold`). They differ in method, not in whose rules govern them.

---

## To the AI assistant building this

**Everything above is the specification.** Build a system that satisfies *What you end up with*,
passes every entry under *How you validate the template's output*, and holds every line under
*The same behavior every time, not the same structure*. How you do that is yours to decide, within
*Free to choose alternatives*.

- **Read every file in `standards/` and apply it.** Naming, the audit column names, error handling and
  query style all come from there and never from this file. Lines marked **[your standards]** in
  *The same behavior every time, not the same structure* are outcomes that layer requires; they are
  stated here so the developer sees the
  whole bar in one place, and they move with the layer if a shop replaces it.
- **This file names no procedures, no modules and no tables on purpose.** How the work divides, what
  it is called, and where things live are declared free above. Do not import a decomposition from
  anywhere else, and do not treat the count of anything in this file as a count of procedures to write.
- **Ask for the seven things under *Information and conditions you need to supply*,** one at a time, through the interactive
  selection control where the answer is a choice and as a plain question where it is a name. Two of
  them are gates: a database in real use with no backup stops the build, and permission to change the
  tables is asked immediately before anything is changed.
- **Never infer an answer that belongs to the developer** — not from what the database looks like, not
  from reasoning that makes an answer seem obvious. Where a check exists to answer a question, run the
  check at the point the sequence calls for it rather than working the answer out yourself from the
  tables.
- **Surface both house assumptions and the warning in the front matter** and get the developer's
  answer before building.
- **The error-handling report comes from `standards/error-handling.md`.** One constraint on the answer
  is not a preference and is stated as an invariant above: nothing in the recording path interrupts
  the person editing.
- **Whatever proves an attachment is this system's own work must be written by the build itself.** A
  mark inside the attachment is the only thing that survives every naming convention. Verified: such a
  mark survives being loaded and read back unchanged.
- **The build record reports against *How you validate the template's output*** — a completed check list, not a narrative.
  Anything you verified along the way goes there rather than into the conversation.

## Extra options

*Named optional extensions, none of them filled in for an engagement.*

- **A user identity function the database already has**, where one exists that knows the signed-in person
  by their real name rather than their Windows account. Offer it by name and let the developer decide.
  Choosing it moves the requirement about reaching every front end onto whatever holds that function.
- **A form for reading the log**, which is named above as something this does not do.
- **Removing old log rows on a schedule**, likewise.
