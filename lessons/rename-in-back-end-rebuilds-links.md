---
id: rename-in-back-end-rebuilds-links
mechanism: split-database
kind: guardrail
confidence: Observed
status: active
lesson: A link whose source table was renamed cannot be refreshed (3011, the old name is looked up), and SourceTableName cannot be set once the TableDef is in the collection (3268). Delete and recreate each front-end link.
applies_when: A table is renamed in the back end of a split database.
evidence: Refreshing the link raised 3011; setting SourceTableName raised 3268.
environment: []
sources:
- catalog-schema 09-22
confirmations: []
contradictions: []
---
