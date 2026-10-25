---
id: a-reported-write-does-not-prove-the-value
mechanism: build-verification
kind: guardrail
confidence: Observed
status: active
lesson: 'A tool that reports success for writing a value proves the call ran, not that the value is the one intended. When a value passes through a layer with its own quoting rules (generated text, a script argument, a loader), that layer can drop characters and report nothing. After the build, read back every property that holds code or text with delimiters (a control source, default value, validation rule, row source, caption) from the saved object and compare it with the intended string, character for character. Then open the form and read what the control shows. Checks that call procedures never see a control bound to a broken expression.'
applies_when: A build writes property values by import, script or tool, and a later step reports success.
evidence: 'A form imported with "imported" held a control source with every quote removed. 76 checks passed, all through procedures; none read a control. In a test the broken expression returned an empty value and no error, so nothing on screen marked it. A second person found it by inspecting the finished database.'
environment:
- Access 16.0 form import through an Access MCP server. Last checked 2026-10-25.
sources:
- northwind-performance-review 10-16
- quote-reproduction 10-25
confirmations: []
contradictions: []
---
