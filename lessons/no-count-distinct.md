---
id: no-count-distinct
mechanism: ace-sql
kind: guardrail
confidence: Observed
status: active
lesson: 'ACE SQL has no COUNT(DISTINCT ...): "Syntax error (missing operator)". Use a grouped subquery.'
applies_when: A query needs a count of distinct values.
evidence: The engine refused COUNT(DISTINCT ...) with the syntax error above.
environment: []
sources:
- stocktake-schema 09-22
confirmations: []
contradictions: []
---
