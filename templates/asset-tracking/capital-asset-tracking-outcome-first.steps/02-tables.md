---
step: 02-tables
title: "Build the tables the design needs"
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, table-in-use, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Tables

Open the database the developer supplied, as the facts below say. Build the tables `capital-asset-tracking-schema`
defines, where they are not there already. Fetch the parts of `capital-asset-tracking-schema` you need with `get_part`:
its checks (step 1) and its tables (step 2). Its other parts repeat what you already hold: skip them.
Change nothing before the developer has given permission.

