---
id: dao-builds-links-and-queries-without-access
mechanism: split-database
kind: guardrail
confidence: Observed
status: active
lesson: 'DAO alone, with no Access window, can create a new accdb that holds linked tables and saved queries (CreateTableDef with Connect and SourceTableName, then CreateQueryDef). It cannot create forms or modules. Build the file with DAO first, then add the forms and modules through the Access MCP server.'
applies_when: A build needs a new front-end file with links to a back end and saved queries.
evidence: A script run under PowerShell created a file with nine links and eight saved queries. Access then imported the forms and modules into it, the links resolved, and the file compiled.
environment:
- DAO.DBEngine.120 is registered for the 32-bit PowerShell only when Office is 32-bit; the 64-bit PowerShell reports the class as not registered. Last checked 2026-10-19.
sources:
- lessons-review-form 2026-10-19
confirmations: []
contradictions: []
---
