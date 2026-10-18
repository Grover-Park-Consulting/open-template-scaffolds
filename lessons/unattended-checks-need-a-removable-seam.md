---
id: unattended-checks-need-a-removable-seam
mechanism: build-verification
kind: guardrail
confidence: Observed
status: active
lesson: A check that turns on a file chooser cannot run unattended. A test seam (public flags that answer the chooser) makes it runnable but ships in the module and lets the application answer its own file chooser. Remove it before use.
applies_when: A check depends on a file chooser dialog and must run without a person.
evidence: The seam made the check run; it shipped in the module until removed. The scaffold build copied the seam and is not counted.
environment: []
sources:
- app-startup-outcome-first 09-23
confirmations: []
contradictions: []
---
