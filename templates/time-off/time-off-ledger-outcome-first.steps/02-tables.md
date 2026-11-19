---
step: 02-tables
title: Build the tables
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Tables

Open the database the developer supplied, as the facts below say. Build the tables `time-off-ledger-schema`
defines, where they are not there already. Fetch the parts of that schema you need with `get_part`:
its checks (step 1), its tables (step 2) and Data Macros (step 3). Its other parts repeat what you already hold: skip
them.

**Build the tables only, and confirm they exist before you fetch step 3.** If you build the tables
with a procedure, that procedure is this step's work and builds tables only: write none of the
application's own code here (the rules, the posting, anything run when the database opens). That
code is step 4's, and step 4 carries the facts it depends on.
