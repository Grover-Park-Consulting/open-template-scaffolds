---
id: checks-must-name-the-value
mechanism: build-verification
kind: guardrail
confidence: Inferred
status: active
lesson: 'A check that confirms only "refused" can pass by accident. Name which row and why. Check your own assertions: two checks reported problems that did not exist (a " WHERE " search against SQL stored with newlines; label attachment tested where existence was required).'
applies_when: A build check asserts that something was refused or found.
evidence: One check caught a tautology only because the arbitrary match happened to be an inactive row. The two builds failed differently, so they are not a confirmation.
environment: []
sources:
- officiating-assignment-outcome-first 09-24
- publication-form 09-23
confirmations: []
contradictions: []
---
