---
id: form-events-need-event-procedure-property
mechanism: form-building
kind: guardrail
confidence: Observed
status: active
lesson: Writing Form_Open and similar into the form module does nothing until the form's OnOpen and similar properties say [Event Procedure].
applies_when: A form's event code is written into its module by a tool.
evidence: The form opened silently with an empty selector until the properties were set.
environment: []
sources:
- publication-form 09-23
confirmations: []
contradictions: []
---
