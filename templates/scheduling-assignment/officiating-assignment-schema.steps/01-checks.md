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

Per `_template-schema.md` §4.2 — the structural baseline instantiated against this schema's ten
tables.

| # | Check |
|---|---|
| 1 | An ordinary insert succeeds on each of the ten tables, supplying every `Required` field. |
| 2 | Every `Required` field on every table refuses a missing value; no `AllowZeroLength` fields are declared in this schema, so that half of the check doesn't apply here. |
| 3 | Each declared unique index refuses its duplicate: `tlkpPlayLevel.PlayLevelName`/`PlayLevelCode`; `tlkpOfficialPosition.PositionName`/`PositionCode`; `tblVenue.VenueName`; `tblTeam` on (`PlayLevelID`, `TeamName`); `tblGameOfficial` on both (`GameID`, `OfficialPositionID`) and (`GameID`, `OfficialID`); `tblPositionRate` on (`PlayLevelID`, `OfficialPositionID`, `EffectiveDate`); `tblAppSetting.SettingName`. |
| 4 | `tblGame → tblGameOfficial` cascade-deletes as declared. Every other relationship does **not** — in particular, `tblOfficial → tblGameOfficial` refuses deleting an official who has assignment history, rather than silently dropping the history; confirm this negative case directly, since it is the one relationship in this schema where "no cascade" is the safety property, not a default. |
| 5 | `tlkpOfficialPosition`'s two declared seed rows (Plate, Base) and `tblAppSetting`'s one (`OfficialPhotoFolder` = `Images\`) are present exactly as specified. **Not applicable** to `tlkpPlayLevel`'s sample rows (10U/12U/14U/16U/18U) — those are documented as samples to replace, not declared in front-matter `seeds`, so their absence is not a defect. |
| 6 | The house audit columns stamp correctly on every table, per Standards Layer below. |
| 7 | A `tblGameOfficial` insert citing a `GameID`, `OfficialID`, or `OfficialPositionID` that doesn't exist is refused. |
| 8 | An insert with `TeamName`, `VenueName`, or another `Text(n)` field longer than its declared width is refused, not silently truncated. Test it on the route this build actually inserts by: a bound form or recordset refuses an over-long value, while a SQL `INSERT` stores the first n characters without raising an error, so where the build inserts by SQL the build itself must refuse the value, or the entry records that it does not. |
| 9 | `Description` is present on every field of every built table, matching this template's own Purpose & rules text. |
| 10 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming what already exists — never a bare "duplicate object" error. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.
