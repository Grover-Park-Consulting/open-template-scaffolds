---
id: identifiers-are-case-insensitive-and-shadow
mechanism: vba-generation
kind: guardrail
confidence: Observed
status: active
lesson: frm and FRM are the same identifier. A local variable shadowed a module constant, and DoCmd.OpenForm FRM passed an unset object (error 2498, no compile error, four failed runs). A helper V with a parameter v gave "Duplicate declaration" and stalled an unattended run about thirteen minutes.
applies_when: Generated VBA declares two names that differ only in case, across scopes or within one procedure.
evidence: Error 2498 with no compile error over four failed runs; a Duplicate declaration compile error that stalled an unattended run.
environment: []
sources:
- publication-form 09-23
confirmations:
- officiating-assignment-scaffold 09-26
contradictions: []
---
