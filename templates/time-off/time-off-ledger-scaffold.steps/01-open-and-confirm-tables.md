---
step: 01-open-and-confirm-tables
title: Open the copy and confirm the tables are there
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Open the copy and confirm the tables

Open the copy of the database the code goes into, as the facts below say. Confirm that the five
tables under *Prerequisites* in the master exist, with the fields `time-off-ledger-schema` gives them.
This scaffold builds no tables. Where they are missing, build them first: fetch
`time-off-ledger-schema` with `get_template` (never read its file) and build them from it, as the
facts below say. Confirm they exist before you fetch step 2.

