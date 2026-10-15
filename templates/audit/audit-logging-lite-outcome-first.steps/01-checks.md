---
step: 01-checks
title: "How you validate the template's output"
platform_facts: []
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

You can validate that the template produced the results we promised. Perform each of these checks. Successfully completing these checks indicates the template ran as expected. Each validation check can be expressed as a yes/no question. **Validate by asking: Does this happen or not?**

**Do these on a copy, and stop if you do not have one.** The checks add records, change fields and
delete records in your own tables — that is what they test, and there is no version of them that
leaves your data alone. Make a copy of the database once auditing has been built into it, run every
check on the copy, and keep your working file out of it. Where your database is split into two files,
copy both and keep them together.

1. **A table that cannot be audited is reported as not auditable.**
   Some tables cannot be audited: those whose primary key is text, those whose primary key is made
   up of more than one field (composite keys), those which have no primary key, and those whose
   primary key is a number that cannot be stored as a Long Integer, such as a Large Number, a
   Decimal or a Double.
   - Before anything is changed, the build shows you your tables, each marked auditable or not
     auditable
   - Each table marked not auditable comes with the reason
   - If every table can be audited, the list says so
   - **A table whose primary key is a Replication ID is the one case you decide.** You are offered
     that choice under *Free to choose alternatives*, and only where your database has at least one
     such table
     - where you chose to audit them, they are marked auditable like any other table
     - where you did not, they are marked not auditable, with that as the reason, and they are
       left alone
2. **A change in one field is recorded as one row.**
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
3. **Only what changed is recorded.**
   - Change two fields of one record at once
   - Open the log to confirm
     - it has exactly two rows for that change
   - Move to a different record
   - Change nothing
   - Move off that record
   - Open the log to confirm
     - it has no new rows
4. **A new record is recorded.**
   - Add a record to a table
   - Open the log to confirm
     - there is a row for each field that was filled in
     - each row is marked as a creation
     - the "before" column is empty
     - each row is stamped with the time and your name
5. **A deleted record leaves its contents behind.**
   - Delete a record
   - Open the log to confirm
      - there is a row for every audited field the deleted record held, **including the fields that
        were empty**
      - marked as a deletion
      - showing what each field contained
   - The record is gone; however, what it held is still readable in the audit log.
   - **An empty field gets a row of its own here on purpose.** A row saying the field held nothing
     is a record that it held nothing. No row at all says only that nothing was written, which reads
     the same whether the field was empty or the deletion was never fully accounted for. The record
     is gone, so there is nothing left to check it against, and the log says it outright instead.
   - **Adding a record works the other way** — rows only for the fields you filled in — because that
     record is still there to be looked at.
6. **Long text survives being deleted.**
   - Repeat *A deleted record leaves its contents behind*, on a record with a long-text field which had
     something in it
   - Open the log to confirm
     - the log holds the full previous contents of the long-text field, not a blank and not a truncation of it
7. **Long text survives being changed.**
   - Edit a long-text field to something different and save
   - Open the log to confirm
     - the log shows the whole of what was in that long-text field before the edit
     - the log shows the whole of what is in that field after the edit
8. **Stamping columns fill themselves.**
   - If your tables also carry their own stamping columns, confirm they were not changed.
   - Add a record without touching those four columns
     - the "created" pair fills in
     - whether the "last changed" pair fills in as well is set by your own rules about those four
       columns, not by this template — see the note below
   - Change any field of that record
     - the "last changed" pair updates
     - the "created" pair does not
   - **Do this on a record that was created before today, if you have one.** A record you added a
     moment ago and changed straight away holds the same time in both pairs, because both happened
     in the same second. That is what you see whether the build is right or wrong, so it proves
     nothing either way. On an older record the difference is plain: the "created" pair stays where
     it was, and only the "last changed" pair moves. If every record is new, wait a minute between
     adding one and changing it.
   - **Two results are correct here, and which one you get is your rules' decision.** Where those
     rules say the last-changed pair stays empty until the first change — which is what the rules
     that came with this library say — a new record fills the created pair only. Where they say all
     four fill on creation, all four fill. **You are told which one your rules produce before
     anything is built**, so you can say at that point if you wanted the other.
9. **A table for which you switched off audit logging is not logged, but stamping still works.**
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
     - the same columns fill in as in *Stamping columns fill themselves* — switching audit logging
       off changes nothing about the stamping
10. **Turn logging on again.**
    - In that table switch a field back on
    - Build again
    - Change that field
      - Open the log to confirm
        - the change is recorded
    - Nothing you did earlier has to be undone first.
11. **Editing through a form is recorded the same way as editing in a table.**
    - Open a form bound to an audited table
    - Edit a record through it
    - Open the log to confirm
      - the log rows are the same as in *A change in one field is recorded as one row*
    - If your database is split, do this with a form in the front end.
12. **A bulk update is recorded.**
    - Run an update query that changes several records at once
    - Open the log to confirm
      - every changed field of every changed record is in the log
13. **The way back out works.**
    - Use a copy of the database you can throw away
    - Remove the audit logging Data Macros from the tables
    - Confirm
      - the tables still accept records
      - a file holding what was there before has been written
      - a table which had Data Macros of your own — unrelated to this template's Data Macros
        - if the build was stopped, those Data Macros are intact
        - if the build ran, those Data Macros are saved to a file you can read
14. **When a build is blocked, it stops rather than leaving your tables half-finished.**
    An open table or form in the database to which you are adding audit logging blocks the template
    - Open the database
    - Leave a table or form open
    - Start the build
    - Confirm that the build
      - stops
      - names what is open
      - changes nothing
15. **The recording happens entirely in the Data Macro, not in code the macro calls.**
    This one you check by looking rather than by editing anything.
    - Open one of the audited tables in design view
    - Open its After Insert Data Macro
    - Confirm you can see the Data Macro actions that write the log row — create a row, set its
      field values
    - A macro whose only action calls a function, and does nothing else, has moved the
      recording into code. You will not see that pattern in any Data Macro
    - Then open the table that holds values between a change and the moment after it
      - it holds values from long-text fields only
      - one holding ordinary fields as well is the same fault seen from the other side: a
        build that cannot read old values in the macro has to copy everything out first
16. **Your existing who-and-when columns keep being filled.**
    This one applies only if your tables already carry those four columns.
    - Add a record to an audited table and save it
    - Look at the record's "added by" and "added on" columns
      - they hold your name and the time you saved it
    - Change a field of that record and save it
      - the "changed by" and "changed on" columns hold your name and the time
    - Switch auditing off for every field of that table, then change a field again and save
      - the "changed by" and "changed on" columns are still filled
      - the log has no new rows for that table
17. **Auditing cannot be got around by moving the database somewhere Access does not trust.**
    This is the check that tells a correct build apart from one that has moved the recording into code.
    - Pick an audited table that has no long-text field being audited
    - Copy the database to a folder you have not told Access to trust — a new folder on the desktop
      will do. Where your database is split into two files, copy both and keep them together
    - Open that copy. Access shows a warning bar across the top, and runs none of the database's code
    - Change one field of one record in that table and save it
    - One of two things happens, and either is correct
      - the save is refused, naming something it cannot use, or
      - the save goes through and the log holds the row, exactly as in *A change in one field is
        recorded as one row*
    - Which of the two you get depends on where the name recorded against a change comes from. A name
      the database has to work out in code cannot be worked out here, so the save is refused before it
      happens. A name that needs no code lets the save through, and the row appears
    - What must not happen is a save that goes through and leaves nothing in the log
    - Delete the copy when you are done
18. **A long-text value is never quietly recorded as blank.**
    - In the same copy used for *Auditing cannot be got around by moving the database somewhere Access
      does not trust*, with code still not running, change a long-text field that had something in
      it, and save
    - One of two things happens, and either is correct
      - the save reports a failure, or
      - the log holds the full previous contents
    - What must not happen is a change that saves without complaint and leaves a log row showing that
      field as empty when it was not
    - Delete the copy when you are done
