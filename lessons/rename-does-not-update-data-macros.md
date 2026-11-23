---
id: rename-does-not-update-data-macros
mechanism: data-macro
kind: guardrail
confidence: Observed
status: active
lesson: Renaming a field, in the table designer or in code, leaves every data macro that names it still naming the old field, with no message, even with Name AutoCorrect on. The next insert fails with "The identifier '[old name]' could not be found" and saves nothing. After renaming a field, rewrite every data macro that names it before anything writes to the table.
applies_when: A field that any data macro reads or writes is renamed, on its own table or on another table that looks it up, in any event.
evidence: A field a Before Change macro read was renamed in the table designer and, separately, through DAO; a field it wrote was renamed in the designer. The macro still named the old field each time (exported text unchanged byte for byte), no message appeared, and every insert, from the datasheet and from an INSERT statement, failed with the identifier error and saved no row. Nothing was written to the application log table. The test covered one Before Change macro on the renamed field's own table; that other tables' macros and other events behave the same is the developer's own observation from long use, not tested here.
environment:
- Access Current Channel version 2611, build 16.0.20625.20108, 32-bit; Name AutoCorrect all three options on. Last checked 2026-11-23.
sources:
- data-macro-rename 11-23
confirmations: []
contradictions: []
---
