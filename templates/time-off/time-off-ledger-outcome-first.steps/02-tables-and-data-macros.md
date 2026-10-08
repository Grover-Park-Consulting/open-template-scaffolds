---
step: 02-tables-and-data-macros
title: Build the tables and the Data Macros the chosen routes need
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, data-macro-rules, vba-import-xml-entities, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Tables and Data Macros

Open the copy the build goes into, as the facts below say. Build the tables `time-off-ledger-schema`
defines, where they are not there already, and the Data Macros the routes the developer chose for
Business Rule 1 and the posting checks call for. Read any Data Macro a table already carries before
writing to it, and merge.

**Build the tables first, then their Data Macros, in this step, and confirm both exist before you
fetch step 3.** A Data Macro needs only the name of a function it calls, not its code, so write the
call now and leave the function to step 3. If you build the tables with a procedure, that procedure
is this step's work and builds tables only: write none of the application's own code here (the
rules, the posting, anything run when the database opens). That code is step 3's, and step 3
carries the facts it depends on.

