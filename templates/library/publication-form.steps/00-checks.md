---
step: 00-checks
title: "Validating the build"
platform_facts: []
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

## Validating the build

Per `_template-schema.md` §9.6: the form baseline (checks 1–8), then one check per Feature and
optional feature of this template.

| # | Check |
|---|---|
| 1 | Each form in `new_forms`, and the record-source query, exists and opens in Form view with no error and no parameter prompt. |
| 2 | Every control in Layout exists with the name, type and binding listed; each data control has its label. |
| 3 | Each subform shows only the rows of the main form's current record, and a row added there is saved against that record. |
| 4 | A new record entered through the form, supplying every Required field, saves. Leaving one empty is refused, and the record stays on screen unsaved. |
| 5 | Each combo bound to a foreign key lists its lookup's rows and stores the key, not the text shown. |
| 6 | New, save, delete (which asks first), clear, and record navigation work as `form-conventions` describes. |
| 7 | Tab order follows the Layout order. |
| 8 | The code-behind compiles, and every handler names its error-handling shape. |
| 9 | Genre, letter and keyword each rebuild the selector list; `<All Publications>` appears in all three and shows every publication, ordered by sort title. Choosing one shows that one. |
| 10 | Each quick-add button adds a row in its dialog, and the new row comes back selected. |
| 11 | An invalid field is highlighted on save and cleared when you move off it. |
| 12 | Optional features, where enabled: the cover image shows for a record with a link and "no image" for one without; the audit summary shows who created and last changed the record. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
