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
| 9 | A new position seed row appears in the crew's position list with no change to the form. |
| 10 | Assigning the same official twice to a game, or an inactive official, is refused with a plain message before any engine error. |
| 11 | Saving a game whose home and away teams match, or play at different levels, is refused with the field highlighted. |
| 12 | Optional features, where enabled: the selected crew member's photo shows (blank, with no error, where the file is missing); the rate shown matches `GetApplicablePayRate`. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
