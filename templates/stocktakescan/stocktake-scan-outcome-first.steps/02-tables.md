---
step: 02-tables
title: "Build the tables the design needs"
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, table-in-use, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Tables

Open the database the developer supplied, as the facts below say. Build the tables `stocktake-schema`
defines and the new `Products` fields the front matter names, where they are not there already.
Fetch `stocktake-schema` with `get_template`, then its checks (step 1) and tables (step 2) with
`get_part`, and build from them; its other parts repeat what you already hold: skip them. Run its
checks with this template's, since they test the same tables.
Change nothing before the developer has given permission.

