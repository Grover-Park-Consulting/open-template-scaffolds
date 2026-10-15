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

Perform each of these checks against a copy of your database. **Validate by asking: does this happen
or not?**

1. **An ordinary error is recorded correctly.**
   - Force an error in a procedure that calls the logger
   - Confirm the log holds that error's number, its description, the module and procedure it
     happened in, when it happened, and who was running the code
2. **An oversized description is shortened, not rejected.**
   - Force an error whose description is longer than the log's description column can hold
   - Confirm the log holds a shortened version of it and the write succeeded
3. **A procedure with no line numbers logs that plainly, not blank.**
   - Force an error in a procedure that carries no line numbers
   - Confirm the log's line column shows the known "no line" value for that procedure, not an empty
     column
4. **A procedure with line numbers logs the line that actually failed.**
   - Force an error on a specific numbered line
   - Confirm the log names that line, not the procedure's first line or a default
5. **Errors are never overwritten or lost.**
   - Force two different errors, one after the other
   - Confirm the log holds a separate row for each, with each one's own detail intact
6. **Choosing the table records to the table.**
   - Where you chose to record to a table, force an error
   - Confirm a row appears in the table
7. **Choosing the fallback still records when the table cannot be reached.**
   - Where you chose the table with a file fallback, make the table unreachable (rename it, or, in a
     split database, remove the link) and force an error
   - Confirm the entry appears in the file instead
   - Confirm the person at the keyboard sees only what you chose them to see — not a report that the
     table failed, unless you chose the full-detail option
8. **The reference number shown matches the row that was written.**
   - Where you chose to show a reference number, force an error and note the number shown
   - Confirm that number identifies the row just written, not a different one
9. **Running the build again does not duplicate anything.**
   - Run whatever created the log table and the logging code a second time
   - Confirm there is still exactly one log table, and no second copy of the logging code
10. **The logger cannot itself cause a new failure, even when every recording route fails at once.**
    - Make both the table and the file destination unreachable at the same time (rename the table and
      point the file path somewhere that cannot be written to)
    - Force an error
    - Confirm the application does not crash and no new, unhandled error appears
    - Confirm the person at the keyboard is told the problem could not be recorded, and asked to note
      what they were doing
