---
step: 01-checks
title: "Validating the build"
platform_facts: []
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

## Validating the build

Per `_template-schema.md` §4.2: the structural baseline instantiated against this schema's five
tables.

| # | Check |
|---|---|
| 1 | An ordinary insert succeeds on each of the five tables, supplying every `Required` field. |
| 2 | Every `Required` field on every table refuses a missing value; no `AllowZeroLength` fields are declared in this schema, so that half of the check does not apply here. Run it before any Data Macro is attached, as check 3 says, and supply the audit columns yourself while you do: once the macro is attached it fills them in before the check can see them missing, and check 6 is what tests that they are filled. |
| 3 | Each declared unique index refuses its duplicate: `tlkpTimeOffCategory.TimeOffCategoryName` and `TimeOffCategoryCode`; `tlkpEntryReason.EntryReasonName`; `tblAccrualSchedule` on (`TimeOffCategoryID`, `MinYearsOfService`); `tblTimeOffEntry` on `CorrectsTimeOffEntryID` and on `ReplacesTimeOffEntryID`. **Also confirm the other direction for those two indexes:** several `tblTimeOffEntry` rows with an empty `CorrectsTimeOffEntryID`, and several with an empty `ReplacesTimeOffEntryID`, are all accepted. **And confirm there is no unique index on the accrual period:** two Earned rows for the same employee, category and `AccrualPeriodStart` are both accepted by the table, because the paired template, not the table, decides when that is allowed. Run this check before any Data Macro is attached to `tblTimeOffEntry`, because once the paired template's checks are attached they refuse these rows before the table's own indexes are reached, and the indexes cannot then be seen working. |
| 4 | Every relationship refuses deleting a parent row that has children, as declared (all six are no cascade). Confirm each directly, including both `tblTimeOffEntry` self-references: an entry that has a Correction pointing at it, and one that has a replacement pointing at it, cannot be deleted. |
| 5 | The three seed rows in `tlkpEntryReason` (Earned, Taken, Correction) are present exactly as specified. **Not applicable** to `tlkpTimeOffCategory`'s sample rows: they are samples to replace, not declared in front-matter `seeds`, so their absence is not a defect. |
| 6 | The house audit columns stamp correctly on every table, per Standards Layer below. |
| 7 | An insert citing an `EmployeeID`, `TimeOffCategoryID`, `EntryReasonID`, `CorrectsTimeOffEntryID` or `ReplacesTimeOffEntryID` that does not exist is refused. |
| 8 | An insert with `FirstName`, `LastName`, `TimeOffCategoryName`, `EntryReasonName` or `TimeOffEntryNote` longer than its declared width is refused, not silently truncated, on every route this build writes by. Where the build writes by SQL, it checks the value against the field's definition first. A SQL `INSERT` the build does not make is not tested here (`_template-schema.md` §4.2 item 8). |
| 9 | `Description` is present on every field of every built table, matching this template's own Purpose & rules text. |
| 10 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming what already exists, never a bare "duplicate object" error. |
| 11 | The two validation rules refuse on every route. `TimeOffHours` of 0 and of 4.1 are refused and 4.25 and -4.25 are accepted; `HoursPerPeriod` of 0 and of -1 are refused and 5 is accepted. Make each attempt by a SQL `INSERT`, by a recordset append, and by a SQL `UPDATE` of a good row, and confirm each refusal shows the rule's own text. Run it before any Data Macro is attached, so the rule, not the macro, is what refused. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
