---
id: sql-tool-reports-ok-for-a-refused-statement
mechanism: build-verification
kind: guardrail
confidence: Observed
status: active
lesson: The SQL tool of an Access MCP server reported "ok, 0 rows" for a statement the engine had refused, and showed no error number. A check that reads that report cannot tell a refusal from an empty result. Prove a refusal by counting the rows before and after, or run the statement through a procedure that executes with fail-on-error and returns the engine's error number.
applies_when: A check reads the result of SQL run through an Access MCP server's SQL tool.
evidence: Refusals were first read from the report, then proved by row counts, and a harness function returned the real numbers (3314, 3315).
environment:
- SQL run through an Access MCP server's SQL tool; may differ by server and version. Last checked 2026-10-04.
sources:
- stocktake-scan-outcome-first 10-04
confirmations: []
contradictions: []
---
