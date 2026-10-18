---
id: compact-resets-autonumber
mechanism: build-verification
kind: guardrail
confidence: Observed
status: active
lesson: Deleting test rows and compacting resets the AutoNumber seed; the first real row is ID 1.
applies_when: Test rows are deleted from a table with an AutoNumber key and the database is compacted.
evidence: The first real row after compacting was ID 1.
environment: []
sources:
- error-logging-schema 09-22
confirmations: []
contradictions: []
---
