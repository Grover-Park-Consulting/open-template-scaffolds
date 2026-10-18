---
id: defaults-written-for-an-old-rule-go-untested
mechanism: build-verification
kind: guardrail
confidence: Observed
status: active
lesson: When a rule changes during a migration, a default written for the old rule is left alone by clearing stored data, and checks that supply every column never exercise it. Include one insert that omits the column and see what the table does. This lesson does not say what the default should be; that is the rule's question.
applies_when: A migration changes a business rule and the table keeps column defaults written for the old rule.
evidence: A leftover default refused every ordinary insert while every read-side check passed; only an insert supplying the minimum a user would supply found it.
environment: []
sources:
- catalog-outcome-first 09-23
confirmations: []
contradictions: []
---
