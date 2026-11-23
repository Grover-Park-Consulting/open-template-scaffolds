---
id: existing-columns-do-not-say-what-fills-them
mechanism: host-adoption
kind: guardrail
confidence: Inferred
status: active
lesson: Before a build writes into columns an existing database already carries, find out what fills them today, table by table - the table's own data macro, a column default, or code such as a form, a save routine or an import. Their names do not say. Add to what already fills them, never write over it; where nothing does, ask.
applies_when: A build adds anything that writes values into columns an existing database already has, such as who-and-when columns.
evidence: A design read an existing database whose tables already carried who-and-when columns. Nothing in the names showed what filled them, and the generated table macro would have written over every value the database's own code set, including any set on purpose. Found by reading the design and code; no build ran.
environment: []
sources:
- audit-scaffold 11-22
confirmations: []
contradictions: []
---
