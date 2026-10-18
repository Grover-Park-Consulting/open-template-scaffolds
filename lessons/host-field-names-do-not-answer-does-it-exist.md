---
id: host-field-names-do-not-answer-does-it-exist
mechanism: host-adoption
kind: guardrail
confidence: Observed
status: active
lesson: '"Does the host already have this field" cannot be answered from field names. Two fields named like a product code and a quantity per unit held an unscannable code with an embedded space and descriptive text. Read the data.'
applies_when: Deciding whether an existing database already holds a field a template needs.
evidence: Fields that looked right by name held an unscannable code with an embedded space and descriptive text.
environment: []
sources:
- stocktake-schema 09-22
confirmations: []
contradictions: []
---
