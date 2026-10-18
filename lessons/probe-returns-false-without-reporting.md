---
id: probe-returns-false-without-reporting
mechanism: error-handling
kind: guardrail
confidence: Inferred
status: active
lesson: A procedure whose failure is an answer (a file probe, "is this file ours") should return False from its handler without reporting. Reporting puts a support dialog in front of someone who picked the wrong file.
applies_when: A procedure's failure is an answer to a question the caller asked, not a fault.
evidence: Two builds returned False without reporting and passed all checks; the dialog consequence is reasoned, and the second build followed the first, so it is not counted.
environment: []
sources:
- app-startup-outcome-first 09-23
confirmations: []
contradictions: []
---
