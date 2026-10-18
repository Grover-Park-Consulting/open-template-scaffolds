---
id: failure-value-must-not-be-a-meaningful-value
mechanism: error-handling
kind: guardrail
confidence: Observed
status: active
lesson: A swallowed error that returns 0 where 0 also means "unmatched" turns a referential-integrity failure into an ordinary unmatched record. Give a failure a value that is not also a result, or test for it.
applies_when: A procedure swallows an error and returns a value that its caller also treats as a normal result.
evidence: A good barcode with a broken session ID landed in the review queue as an unknown product.
environment: []
sources:
- stocktake-scan-scaffold 09-22
confirmations: []
contradictions: []
---
