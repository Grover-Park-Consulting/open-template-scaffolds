---
step: 04-data-macros
title: "Build the Data Macros"
platform_facts: [data-macro-rules, vba-import-xml-entities, table-in-use]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Data Macros

**The tables exist from step 3: do not rebuild them here.** Attach the Data Macros to the audited
tables, as the facts below say. *Facts about the platform*, below, is context for this step.

### Facts about the platform

**Nothing in this section is part of the specification.** It is here so a build knows what Access will
and will not do, and so a developer reading it can see why the design looks as it does. Everything the
build has to produce is stated in *What you end up with*, in the checks under *How you validate the
template's output*, and in the list under *The same behavior every time, not the same structure* — and
is stated there in full. A requirement that appears to live only here is a defect in this template: it
belongs in one of those three sections, and finding one is worth reporting.

**These read as open problems; they are not.** Every one of them is already solved, either in
`templates/_materialization.md` (the library-wide Data Macro reference — the UTF-16 write
requirement, the replaces-not-merges behavior, the namespace rewrite on export, and the Long Text
`[Old]` limitation are all rules stated there) or in the working code of the paired scaffold,
`audit-logging-lite-scaffold.md` (which shows the actual staging mechanism for a Long Text value,
the comment-based ownership marker, and the System-table read boundary). This section names the
facts; it does not re-solve them.

- **Data Macros.** Access can attach automatic behaviour to a table itself using Data Macros. This behaviour occurs regardless of who makes
  the change to data and how they make it. It is the only mechanism in Access that a person editing a table cannot bypass,
  and the only one that still runs when Access is refusing to run code.
- **The ordinary VBA programming interface cannot programmatically create a Data Macro.** They can
  only be written programmatically as an XML document and loaded into the table. They can be read back out the same way.
- That XML has to be written in the two-byte character encoding, **UTF-16**. Written any other way it
  loads without complaint but does nothing.
- Loading XML for a table's Data Macros **replaces the Data Macros already attached to that table**. It never merges. Two separate
  loads means the second wipes out the first. Everything a table needs has to be built into one XML document.
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
- **Access cannot hand the previous contents of a long-text field to its own Data Macro, and
  asking for it stops the macro.** This is a limitation in Data Macros. The data types affected
  are the long ones — the memo type and anything built on it. **The macro does not return a
  blank and carry on: it abandons the rest of its work at that line, silently, and nothing is
  reported.** So a macro written to read that value does not merely lose the value; every action
  after it is skipped too. This shapes the design for long-text fields: the old value has to be
  captured a moment before the change, held somewhere, and retrieved for the audit log after the
  change — and that holding place is for long-text values, nothing else.
- **A hyperlink field is a long-text field.** Access reports it as one and it needs the same handling,
  although a developer sees "Hyperlink" in the table designer. A list that refers to a hyperlink field as a long text field names
  something the developer cannot easily find. The two are told apart by a flag on the field rather than by its type. Hyperlink fields are audited the same way ordinary long-text fields are audited.
- **Attachment fields and calculated fields cannot usefully be audited.** A calculated field cannot be
  edited directly. Only its components are ever changed. Therefore, calculated fields themselves are not auditable.
  **A calculated field is told apart from an ordinary one by its stored formula holding something, never by
  whether it can be asked for one.** Every field answers that question, and an ordinary field answers it with
  nothing — so a build that treats being able to ask as the answer marks every field in every table as one
  that can never be audited, and reports that it succeeded, because each step did. The Data Macros go on,
  the settings table fills up, and not one switch can ever be turned on.
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
