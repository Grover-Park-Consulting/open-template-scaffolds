---
id: constants-after-procedures-do-not-compile
mechanism: vba-generation
kind: guardrail
confidence: Observed
status: active
lesson: 'Module-level Const appended after existing procedures fails to compile: "Only comments may appear after End Sub."'
applies_when: Code is appended to a module that already contains procedures.
evidence: The compile error above, on a Const placed after End Sub.
environment: []
sources:
- audit-logging-lite-outcome-first 09-22
confirmations: []
contradictions: []
---
