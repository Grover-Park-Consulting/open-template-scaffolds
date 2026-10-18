---
id: select-without-from-fails
mechanism: ace-sql
kind: guardrail
confidence: Observed
status: active
lesson: 'A SELECT of scalar subqueries with no FROM fails: "Query input must contain at least one table or query". Run separate counts.'
applies_when: A statement selects scalar subqueries with no FROM clause.
evidence: The engine raised the error above. This is documented, long-standing behavior of the Access database engine's SQL, so it is as strong as an observation.
environment: []
sources:
- audit-logging-lite-schema 09-22
confirmations:
- stocktake-scan-outcome-first 2026-10-04
contradictions: []
---
