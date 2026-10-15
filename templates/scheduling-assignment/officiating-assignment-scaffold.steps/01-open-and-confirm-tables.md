---
step: 01-open-and-confirm-tables
title: "Open the database and confirm the tables are there"
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, table-in-use, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Open the database and confirm the tables

Open the database the code goes into, as the facts below say. Confirm that the tables under
*Prerequisites* in the master exist, with the fields `officiating-assignment-schema` gives them. This scaffold builds
no tables. Where they are missing, build them first: fetch `officiating-assignment-schema` with `get_template` (never read
its file), then its checks (step 1) and tables (step 2) with `get_part`, and build them from it, as
the facts below say. Its other parts repeat what you already hold: skip them. Confirm the tables
exist before you fetch step 2.

