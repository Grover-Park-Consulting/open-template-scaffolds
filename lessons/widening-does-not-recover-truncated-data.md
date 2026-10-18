---
id: widening-does-not-recover-truncated-data
mechanism: schema-migration
kind: guardrail
confidence: Observed
status: active
lesson: Widening a Text column does not restore text already truncated; re-derive from the untruncated source.
applies_when: A Text column that has already truncated data is widened.
evidence: 727 values at exactly 50 characters before; 906 values over 50 characters after re-deriving from the Memo source.
environment: []
sources:
- catalog-schema 09-22
confirmations: []
contradictions: []
---
