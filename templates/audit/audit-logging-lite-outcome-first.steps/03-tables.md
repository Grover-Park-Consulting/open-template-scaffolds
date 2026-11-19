---
step: 03-tables
title: "Build the tables"
platform_facts: [open-existing-startup, dao-table-build, sql-server-ddl, table-in-use, sql-insert-truncation]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Tables

Open the database the developer supplied, as the facts below say. Build the log table, the settings
table and the staging table where they are not there already. Change nothing before the developer
has given permission.

**From the master's instructions to the AI assistant, for this step.** They bind as the master's do.

- **One thing in the fact `dao-table-build` is outside the narrowing *Free to choose alternatives*
  allows and does bind here: rule 6, about a field that receives values from other fields.** The
  log's before-and-after value columns are such a field — every audited field in the database writes
  into them — so they hold the widest text the engine has and they accept a value set to nothing at
  all. That is not a route decision and it is not open under *Free to choose alternatives*; a log
  that cannot record a field emptied is a log that loses the change it exists to record.

