---
step: 02-limits
title: "What the template does not do"
platform_facts: []
route: both
when: when you present the design
---

**Who reads this:** the developer whose database this is, and the AI assistant, which states these limits when it presents the design.

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
  **The same thing happens when you take the audit logging back off, and it is worth knowing before you
  do.** Access holds a table's automatic behaviour as one whole, so removing this template's part of it
  clears the rest along with it — including whatever was there before the build. Where a table filled its
  own four stamping columns, it stops filling them once the audit logging is removed. The table still
  works and still accepts records; what stops, without saying so, is the stamping. The file written
  before the replacement is what puts it back.
- **Tidying up the files it writes.** Each time the build attaches Data Macros to a table, or removes
  them, it first writes out what was there into a file beside the database. That happens on every
  rebuild, not only the first, and nothing ever removes those files. A database whose audited tables are
  rebuilt a few times gathers hundreds of them. They are small, and deleting them is safe once you no
  longer want what they hold — but that is your decision to make, and nothing prompts you to make it.
- **Recognising Data Macros written by a version of this system that predates the mark it now writes.**
  Adding a mark to Data Macros was added to a later version of the template. Data Macros added in an earlier version
  are treated as somebody else's work and left in place.
- **Preventing two people building this into the same database at once.** The whole run assumes one
  person with everything else closed. That's not under the control of the template. You are responsible for how you work.
