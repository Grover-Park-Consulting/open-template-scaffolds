---
step: 02-tables
title: "Build the tables the design needs"
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, table-in-use, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Tables

Open the database the developer supplied, as the facts below say. Build the tables `stocktake-schema`
defines and the new `Products` fields the front matter names, where they are not there already.
Fetch `stocktake-schema` with `get_template` and its tables step (step 2) with `get_part`, and build
from them; run its checks (its step 1) with this template's, since they test the same tables.
Change nothing before the developer has given permission.

