---
id: form-text-quotes-are-backslash-escaped
mechanism: form-building
kind: guardrail
confidence: Observed
status: active
lesson: 'In the text Access reads for a form or report (LoadFromText, SaveAsText), a quote inside a property value is written \" and a backslash \\. A doubled quote ("") is not an escape: Access reads it as the end of one string and the start of the next, drops both, and stores the value without them, with no error. Every property that holds text with quotes is affected, not only IIf: ControlSource, DefaultValue, ValidationRule, RowSource. Write quotes as \" in generated form or report text.'
applies_when: A tool or script generates form or report text, or edits exported text, and a property value contains a quote.
evidence: 'Reproduced in a fresh database. One ControlSource written with "" loaded as =IIf(1=1,Enter,) and with \" loaded exactly; DefaultValue, ValidationRule and a combo box RowSource behaved the same. Access''s own export writes \". The generator''s text was written with "" and every quote in the built form was lost. The same literals set from VBA or DAO kept their quotes. With \", the real form loaded, survived an export and re-import, and survived VBA injection.'
environment:
- Access 16.0 (32-bit Office), form text loaded as UTF-16 by LoadFromText, called by an Access MCP server that passes the text through unchanged. Last checked 2026-10-25.
sources:
- northwind-performance-review 10-16
- quote-reproduction 10-25
confirmations: []
contradictions: []
---
