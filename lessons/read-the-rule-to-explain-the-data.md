---
id: read-the-rule-to-explain-the-data
mechanism: schema-migration
kind: guardrail
confidence: Observed
status: active
lesson: Read an existing validation rule before judging odd data. A uniform value across many rows may be what the rule required, not noise.
applies_when: Stored data looks odd during a migration and the table has a validation rule.
evidence: 1,773 rows held the same value because the old rule required it.
environment: []
sources:
- catalog-schema 09-22
confirmations: []
contradictions: []
---
