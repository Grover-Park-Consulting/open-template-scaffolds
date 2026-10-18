---
id: rename-orphans-validation-rule
mechanism: schema-migration
kind: guardrail
confidence: Observed
status: active
lesson: 'Renaming a field through DAO updates indexes and relationships but not a table validation rule. The failure appears later on an unrelated statement: error 3799 "Could not find field" naming the old field.'
applies_when: A field is renamed in a table whose validation rule names that field.
evidence: The index reported the new field name; ALTER TABLE ... ALTER COLUMN raised 3799 naming the old one.
environment: []
sources:
- catalog-schema 09-22
confirmations: []
contradictions: []
---

Potential lead, not part of the lesson and not applied: data-macro expressions and default-value expressions probably orphan the same way. Not tested.
