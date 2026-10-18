---
id: held-database-variable-under-com
mechanism: vba-generation
kind: guardrail
confidence: Observed
status: active
lesson: Under COM automation, CurrentDb.TableDefs("name") raised 3420 "Object invalid or no longer set". Holding the Database in a variable and calling TableDefs.Refresh first fixed it.
applies_when: VBA runs under COM automation from outside Access and reads a TableDef through CurrentDb.
evidence: Error 3420 on the CurrentDb form; the held-variable form with Refresh succeeded.
environment:
- COM automation from a shell, Access 2016; may not occur in an interactive session. Last checked 2026-09-26.
sources:
- officiating-assignment-scaffold 09-26
confirmations: []
contradictions: []
---
