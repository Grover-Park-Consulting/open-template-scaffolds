---
step: 03-code
title: Write and import the code
platform_facts: [mcp-module-import, mcp-line-numbers, vba-import-xml-entities, domain-function-transaction, row-lock-errors, recordset-append-crash, app-startup-autoexec]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Code

**The tables and Data Macros exist from step 2: do not rebuild them here.** Write and import only the
application's own procedures, including every function a Data Macro from step 2 calls by name.

Write and import the procedures the build needs, divided as the design-principles standard says,
within *Free to choose alternatives*. Each function a Data Macro calls goes into the back end and
every front end. Where the developer chose to post when the database opens, the open-time work
follows the startup-conventions standard and the fact on the open-time macro below.

