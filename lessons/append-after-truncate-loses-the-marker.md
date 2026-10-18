---
id: append-after-truncate-loses-the-marker
mechanism: derived-values
kind: guardrail
confidence: Observed
status: active
lesson: A derivation that appends text after Left(..., n) can cut off the appended text. Truncate the base first, then append within the width.
applies_when: A derived value appends text after truncating a base value to a fixed width.
evidence: 3 of 1,929 rows lost the appended suffix.
environment: []
sources:
- catalog-outcome-first 09-23
confirmations: []
contradictions: []
---
