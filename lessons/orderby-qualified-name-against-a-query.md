---
id: orderby-qualified-name-against-a-query
mechanism: form-building
kind: guardrail
confidence: Observed
status: active
lesson: A form bound to a query with OrderBy set to a table-qualified name prompts "Enter Parameter Value". It is invisible in Design view and fatal in Normal view. Order by the query's own column.
applies_when: A form is bound to a query and its OrderBy names a table-qualified column.
evidence: The parameter prompt appeared in Normal view and not in Design view.
environment:
- Form built through an Access MCP server's form tool. Last checked 2026-09-23.
sources:
- publication-form 09-23
confirmations: []
contradictions: []
---
