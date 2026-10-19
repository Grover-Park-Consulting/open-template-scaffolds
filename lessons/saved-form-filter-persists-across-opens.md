---
id: saved-form-filter-persists-across-opens
mechanism: form-building
kind: guardrail
confidence: Observed
status: active
lesson: 'Access saves the filter last applied to a form and re-applies it at the next open, while the control that applied it shows its default. When a control applies the form''s filter, have Form_Open turn the filter off and set the control to its default.'
applies_when: A form applies a filter from a combo box or a button, and the form is saved after use.
evidence: A saved form reopened showing only one approval while its filter box showed All, until Form_Open set FilterOn to False and the box to All.
environment: []
sources:
- lessons-review-form 2026-10-18
confirmations: []
contradictions: []
---
