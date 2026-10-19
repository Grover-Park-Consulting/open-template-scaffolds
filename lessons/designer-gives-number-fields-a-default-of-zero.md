---
id: designer-gives-number-fields-a-default-of-zero
mechanism: field-properties
kind: consequence
triggers_when: The developer builds the tables by hand in the table designer AND the design makes a Number field nullable or requires it empty under a rule.
confidence: Inferred
status: active
lesson: A Number field created in the table designer gets a Default Value of 0 automatically. A design that makes such a field nullable, or requires it to be empty under a rule, is contradicted by that default unless the builder clears it. Where the developer builds the tables by hand from an approved design, warn them in the hand-off to check each Number field's Default Value against the design. The assistant cannot check the developer's tables and promises nothing about them. This lesson does not say what a default should be; that is the business rule's question.
applies_when: The developer creates tables by hand in the table designer from an approved design.
evidence: George's repeated observation over years of using the designer. Seen again on 2026-10-19, when two foreign-key Number fields in hand-built tables carried a Default Value of 0 the developer had not set, and had to be cleared. Not reproduced in any recorded assistant build, so Inferred. Checkable by adding a Number field in the designer and reading its Default Value.
environment:
- Tables built in the Access table designer, Access 2016. Fields created through DAO get no default unless the code sets one; the build records are consistent with this. Last checked 2026-10-02.
sources:
- George (not a build record)
confirmations:
- lessons-review-form 2026-10-19
contradictions: []
---
