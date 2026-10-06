---
template: time-off-ledger-schema
title: Time Off Ledger — Table Schema
domain: time-off
type: table-schema
version: 0.3.6
status: draft
standards_layer: [audit-columns, naming-conventions, error-handling]
new_tables:
  - tlkpTimeOffCategory
  - tlkpEntryReason
  - tblEmployee
  - tblAccrualSchedule
  - tblTimeOffEntry
seeds:
  - tlkpEntryReason.Earned
  - tlkpEntryReason.Taken
  - tlkpEntryReason.Correction
platform_facts: [target-file, dao-table-build, sql-server-ddl, sql-insert-truncation, data-macro-rules, domain-function-transaction, row-lock-errors, recordset-append-crash]
house_assumptions:
  - "tblTimeOffEntry — the balance is never stored; it is the sum of TimeOffHours for an employee and a time off category, computed whenever it is needed (Business Rule 2). A stored balance is an Extra Option"
  - "tblTimeOffEntry.TimeOffHours — hours are kept in quarter-hour steps, so sums are exact; an entry whose hours are not a multiple of a quarter hour is refused (Business Rule 8). A shop that tracks finer fractions changes the field's type and lifts that refusal"
  - "tlkpTimeOffCategory.AccrualPeriodMonths — accrual periods are counted from each employee's HireDate in blocks of this many months; AccrualPeriodStart is the first day of each block, so an annual category accrues on the hire anniversary"
  - "tblEmployee.AccrualStartDate — hours are earned only for periods that start on or after this date, so an employee who was already working when the ledger was started is not credited every period since their hire date at once. Empty means count from HireDate"
  - "tblTimeOffEntry — time taken is recorded one row per day, not as a start and end date, so every calculation and the overlap check (Business Rule 9) work on single dates. This makes more rows; a date-range design is an Extra Option. Nothing here assumes a five-day working week: a day is taken or it is not, and a request spanning several days is turned into rows by whatever posts it, which is where the shop's own working week applies"
  - "tblTimeOffEntry.ReplacesTimeOffEntryID — a wrong entry is replaced by posting the corrected one first and then cancelling the wrong one, the two together or not at all (Business Rule 10). A cancelled entry still holds its accrual period and its day, which is why the table has no unique index on either"
  - "tblEmployee — a minimal employee table is created here; most databases already have one, and hooking into it is an Extra Option"
---

# Time Off Ledger — Table Schema

**Status last determined:** 2026-10-05.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Tables for tracking **time off** (vacation, sick leave and similar) as a **ledger**: a list of
entries, one for each time hours are earned, taken or corrected. An employee's balance is not
written down anywhere. It is added up from the entries whenever somebody asks for it, so it can
never disagree with the entries it came from.

**What a ledger means here.** Every change to anyone's time off is a new row. A row that has been
posted is never edited and never deleted. A mistake is fixed in one of two ways. An entry that
should not exist at all is cancelled by a second row that offsets it. An entry that is wrong but
should be there is **replaced**: the corrected entry is posted first, and then the wrong one is
cancelled, the two together or not at all. The history of what happened, and when, stays whole.

**Time earned follows length of service.** How many hours an employee earns in each accrual
period (the stretch of time over which hours are earned) depends on how many full years they have
worked, using a schedule the developer fills in. Whether the schedule also depends on an
employee's role is an Extra Option.

This is a greenfield template: it creates its whole schema and hooks into no existing tables.

## Prerequisites

This template hooks into no existing tables. What does have to be settled before you build is
**which file the tables go in**.

**Where these tables live.** These templates are designed for a **split database**, the normal
shape for Access applications, especially those used by several people. One file holds the tables
(the **back end**, usually on a shared network drive, never on OneDrive, Dropbox or any other
file-syncing cloud folder, which corrupts a shared Access back end). Each person runs their own copy
of a second file holding the forms, reports and code (the **front end**), whose tables are *links*
pointing at the back end. A single-file database, one .accdb holding everything, is an acceptable
choice for one user, and everything here works there too: create everything in that one file and
ignore the distinction.

All five tables below, two lookups and three entities, are created in the **back end** and linked
into each front end.

## Entities

### tlkpTimeOffCategory

Grain: one row per kind of time off.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `TimeOffCategoryID` | AutoNumber | PK | Surrogate key |
| `TimeOffCategoryName` | Text(30) | Required | e.g. Vacation, Sick |
| `TimeOffCategoryCode` | Text(4) | Nullable | Short code for reports |
| `AccrualPeriodMonths` | Byte | Required | Months in one accrual period; must be at least 1 (1 is monthly, 12 is annual) |

Indexes: PK on `TimeOffCategoryID`; unique on `TimeOffCategoryName`; unique on
`TimeOffCategoryCode`.

Seed rows (samples, replace with the organization's real categories): Vacation (12 months), Sick
(1 month), Personal (12 months).

### tlkpEntryReason

Grain: one row per reason an entry can exist.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `EntryReasonID` | AutoNumber | PK | Surrogate key |
| `EntryReasonName` | Text(30) | Required | Earned, Taken or Correction |

Indexes: PK on `EntryReasonID`; unique on `EntryReasonName`.

Seed rows (declared): Earned, Taken, Correction.

### tblEmployee

Grain: one row per employee.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `EmployeeID` | AutoNumber | PK | Surrogate key |
| `FirstName` | Text(50) | Required | |
| `LastName` | Text(50) | Required | |
| `HireDate` | Date/Time | Required | The date length of service is counted from |
| `AccrualStartDate` | Date/Time | Nullable | Hours are earned only for accrual periods that start on or after this date (Business Rule 4). Empty means count from `HireDate`. Set it to the day the ledger was started for anyone already employed then |
| `InactiveDate` | Date/Time | Nullable | Empty means the employee is still employed. A date means they stopped being active on that day (Business Rule 6) |

Indexes: PK on `EmployeeID`; index on `LastName`.

Derived (not stored): completed years of service on any given date, from `HireDate`.

### tblAccrualSchedule

Grain: one row per time off category and length-of-service band. Together the rows for one
category are its schedule.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `AccrualScheduleID` | AutoNumber | PK | Surrogate key |
| `TimeOffCategoryID` | Long | FK → tlkpTimeOffCategory, Required | |
| `MinYearsOfService` | Integer | Required | The band starts at this many completed years; must be 0 or more |
| `HoursPerPeriod` | Double | Required | Hours earned in each accrual period while in this band. A validation rule on the field, `>0`, with the text "Hours per period must be more than zero.", refuses anything else, because an entry is never for zero hours (Business Rule 8). A band that earns nothing is left out of the schedule, and the run reports its periods as not posted (Business Rule 3) |

Indexes: PK on `AccrualScheduleID`; unique on (`TimeOffCategoryID`, `MinYearsOfService`); FK index
on `TimeOffCategoryID`.

### tblTimeOffEntry

Grain: one row per posted event: hours earned, hours taken, or a correction. This is the ledger.
Time taken is recorded **one row per day**: a week of vacation is five Taken rows, one for each
day. Five is the usual working week, not a rule: the table holds one row for each day actually
taken, so a four-day week or a rota works the same way.

**This makes the table grow faster than a start-and-end design would.** An employee who takes 20
days off a year adds 20 rows a year; 500 employees add 10,000 rows a year, plus the Earned
entries. An Access back end holds that comfortably, and what you get for it is that every sum and
every overlap check works on single dates. A screen that takes a request for several days can
still accept a start and an end and post one row for each day. If you would rather store one row
per request, see *Date-range entries* under Extra Options before you build.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `TimeOffEntryID` | AutoNumber | PK | Surrogate key |
| `EmployeeID` | Long | FK → tblEmployee, Required | |
| `TimeOffCategoryID` | Long | FK → tlkpTimeOffCategory, Required | |
| `EntryReasonID` | Long | FK → tlkpEntryReason, Required | |
| `EntryDate` | Date/Time | Required | The date the entry takes effect. For Taken, the day off. For Earned, the first day of the accrual period it covers. For a Correction, the date of the entry it cancels, so the cancellation sits on the day it cancels |
| `TimeOffHours` | Double | Required | Signed: hours earned are positive, hours taken are negative; never 0, and always a multiple of a quarter hour (Business Rule 8). A validation rule on the field refuses both: `<>0 And Int([TimeOffHours]*4)=[TimeOffHours]*4`, with the text "Hours must be a multiple of a quarter hour, and not zero." The database engine applies it on every route, whether a form, a recordset or a SQL statement. Whether the sign suits the reason is a rule about the whole row and is the paired template's to enforce |
| `AccrualPeriodStart` | Date/Time | Nullable | Filled on every Earned entry, and only on an Earned entry: the first day of the accrual period it covers (Business Rule 4). An Earned entry without it is refused. That the date really is the start of one of the employee's periods is checked by the posting procedures; the table does not check it for an entry inserted directly |
| `CorrectsTimeOffEntryID` | Long | FK → tblTimeOffEntry, Nullable | Filled only on a Correction: the entry it cancels (Business Rule 7). Named for its role because the table cannot hold two columns called `TimeOffEntryID` |
| `ReplacesTimeOffEntryID` | Long | FK → tblTimeOffEntry, Nullable | Filled only on an Earned or Taken entry that replaces a wrong one: the entry it replaces (Business Rule 10) |
| `TimeOffEntryNote` | Text(255) | Nullable | Why, in the person's own words. Required on a Correction that cancels hours earned (Business Rule 7) |

Indexes: PK on `TimeOffEntryID`; unique on `CorrectsTimeOffEntryID` and unique on
`ReplacesTimeOffEntryID`, each **leaving out rows where the column is empty** (most rows leave it
empty, and any number of them must coexist). Between them they mean an entry can be cancelled once
and replaced once, and the table itself refuses a second one even when two are posted at the same
moment. Also: index on (`EmployeeID`, `TimeOffCategoryID`, `AccrualPeriodStart`), **not unique**,
and index on (`EmployeeID`, `TimeOffCategoryID`, `EntryDate`); FK indexes on `TimeOffCategoryID`
and `EntryReasonID`.

**Why no unique index on the accrual period.** A cancelled entry stays in the table and still holds
its period. A unique index would refuse the corrected entry that is meant to take its place. "One
active Earned entry per period" (Business Rule 4) cannot be written as an index, because whether an
entry is active depends on whether a Correction exists, so the paired outcome-first template
enforces it.

Derived (not stored): the balance, as the sum of `TimeOffHours` for an employee and a time off
category (Business Rule 2). Also derived: whether an entry is **active**, which means no Correction
cancels it.

## Relationships

- `tlkpTimeOffCategory (1) → (∞) tblAccrualSchedule` on `TimeOffCategoryID`: no cascade
- `tlkpTimeOffCategory (1) → (∞) tblTimeOffEntry` on `TimeOffCategoryID`: no cascade
- `tlkpEntryReason (1) → (∞) tblTimeOffEntry` on `EntryReasonID`: no cascade
- `tblEmployee (1) → (∞) tblTimeOffEntry` on `EmployeeID`: no cascade (an employee with entries
  cannot be deleted; the ledger is never deleted, Business Rule 1)
- `tblTimeOffEntry (1) → (∞) tblTimeOffEntry` on `CorrectsTimeOffEntryID`: no cascade; the column
  is nullable and carries a unique index, so an entry has at most one Correction
- `tblTimeOffEntry (1) → (∞) tblTimeOffEntry` on `ReplacesTimeOffEntryID`: no cascade; the column
  is nullable and carries a unique index, so an entry is replaced at most once

## Business Rules

1. **The ledger only grows.** A posted `tblTimeOffEntry` row is never edited and never deleted,
   whatever route a change comes in by. A mistake is fixed by cancelling the entry (Rule 7) or by
   replacing it (Rule 10). The paired outcome-first template states how this is enforced.
2. **The balance is added up, never stored.** An employee's balance in a category is the sum of
   `TimeOffHours` over that employee's entries in that category, and is computed whenever it is
   needed. No table holds it.
3. **Hours earned follow length of service.** For an accrual period, the applicable
   `tblAccrualSchedule` row is the one for the category with the greatest `MinYearsOfService` not
   above the employee's completed years of service on `AccrualPeriodStart`. If no row applies,
   nothing is posted and the gap is reported, never treated as zero hours. A schedule that also
   depends on role is an Extra Option.
4. **Hours are earned once per period, and every period that is due gets posted.** An employee has
   at most one active Earned entry for a category and accrual period; the one exception is a
   replacement waiting to cancel the entry it replaces (Rule 10). Accrual periods are counted from
   `HireDate` in blocks of `AccrualPeriodMonths` months, and `AccrualPeriodStart` is the first day
   of the block. Only periods that start on or after the employee's `AccrualStartDate` (their
   `HireDate` where that is empty) and on or before today earn hours, and an employee who has
   stopped working earns nothing for a period that starts after their `InactiveDate` (Rule 6). A
   period that was missed is posted the next time hours are posted. **A period that already has an
   Earned entry, cancelled or not, counts as posted when hours that are due are posted**, so a
   cancellation made to take hours away is not quietly undone by the next run. To give the hours
   again, someone posts a new Earned entry for that period on purpose. Two entries for one period
   posted at the same moment cannot both succeed.
5. **No entry can take the balance below zero.** An entry's effect on the balance is its own hours,
   except that a replacement counts only for its hours less the hours of the entry it replaces
   (Rule 10). An entry whose effect lowers the balance, which means time taken and also a Correction
   that cancels hours earned, is refused when the balance after it would be below zero. Replacing
   96 hours with 64 lowers the balance by 32 and is judged on that. If more time has been taken than
   the corrected hours allow, the replacement is refused as a whole and nothing is left behind,
   whichever of its two writes the refusal comes at. The shortfall stays visible, and the time taken
   is dealt with first.
   Two such entries posted at the same moment cannot both succeed. Allowing a negative balance is
   an Extra Option.
6. **An employee who has stopped working earns and takes nothing afterwards.** When an employee's
   `InactiveDate` is filled, no entry is dated after it. An empty `InactiveDate` means the
   employee is still employed.
7. **A correction cancels exactly one entry, and an entry is cancelled once.** A Correction entry
   names the one entry it cancels in `CorrectsTimeOffEntryID`, with hours equal and opposite to
   that entry's. The entry named belongs to the same employee and the same category as the
   Correction, and is not itself a Correction. The cancelled entry itself is left as it was, and no
   second Correction may name it. **A Correction that cancels hours earned carries a note saying
   why**, so that taking hours away from an employee is never anonymous.
8. **Hours are signed, never zero, and in quarter-hour steps.** An Earned entry has positive
   hours; a Taken entry has negative hours; a Correction has whatever sign offsets the entry it
   names. An entry whose hours are zero, or not a multiple of a quarter hour, is refused by a
   validation rule on the field, so that sums are exact and the refusal holds on every route.
9. **One category per day.** An employee has at most one active Taken entry for any one
   `EntryDate`, so vacation cannot be booked in the middle of sick leave, and two entries cannot
   claim the same day under different categories. An entry is active unless a Correction cancels
   it, so a cancelled day can be booked again. A part day is one entry with part of a day's hours.
   The one exception is a replacement waiting to cancel the entry it replaces (Rule 10). Two
   entries posted at the same moment for the same day cannot both succeed. The paired outcome-first
   template states how this is enforced.
10. **A wrong entry is replaced by posting the corrected one first and then cancelling the wrong
    one, and the two happen together or not at all.** A replacement is a new Earned or Taken entry
    that names, in `ReplacesTimeOffEntryID`, the entry it replaces. The entry named belongs to the
    same employee and the same category, has the same reason (Earned replaces Earned, Taken replaces
    Taken), is still active, has not been replaced before, and is not a Correction. A replacement
    for an Earned entry covers the same accrual period. While the replacement waits to cancel the
    entry it replaces, that entry is not counted against Rules 4 and 9 (two entries for one period
    or one day), and Rule 5 counts the replacement for only its hours less the hours it replaces.
    The replacement is finished when a Correction cancels the entry it replaces. If anything stops
    the pair before then, because the cancellation is refused or an error occurs, the replacement is
    not left behind: both rows are written, or neither is. A replacement posted on its own, with no
    cancellation following it in the same step, is refused, because it would leave the wrong entry
    and its replacement both counting.

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

## Standards Layer

- **Audit columns:** every table above additionally carries the audit set supplied by the active
  `audit-columns.md`; audit fields never appear in this template's field tables. These columns
  record who created or changed a row and when. They are separate from the ledger, which records
  what happened to time off.
  On `tblTimeOffEntry`, `ModifiedDate` and `ModifiedBy` stay empty for good: no entry is ever
  changed (Business Rule 1), so nothing ever stamps them. That is by design. Check 6 confirms on that
  table that the creation pair is filled and the change pair is empty.
- **Naming conventions:** this template is written in the OTS default style (`tbl` and `tlkp`
  prefixes, `[Entity]ID` keys, qualified field names). A practice with different conventions builds
  the same entities under its own `naming-conventions.md` without editing this template.
- **Error handling:** any VBA generated alongside takes its error pattern from the active
  `error-handling.md`.

## Extra Options

*Named optional extensions, none of them filled in for an engagement; the filled copy is saved to the developer's
own library, not committed here.*

- **Role-based accrual.** Adds a role lookup, a role on each employee, and a role column on
  `tblAccrualSchedule`, so each role (for example hourly, salaried, management) has its own
  schedule by length of service. The tenure rule (Business Rule 3) stays; only which schedule's
  bands it reads changes. Accrual uses the employee's current role when the entry is posted.
- **Role history.** On top of role-based accrual: a dated record of each employee's roles, so
  accrual uses the role held on the day the period began.
- **Allow going below zero.** Lifts Business Rule 5. When chosen, the build also asks how an
  overdraw is handled, and the answer is recorded with the build. Candidate answers: allow it and
  flag it, allow it only with a recorded reason, or allow it up to a stated limit. This is also the
  way to accept a replacement that leaves time taken exceeding the corrected hours, where the
  developer decides that is what happened.
- **Stored balance.** Keeps a balance on a row per employee and category, updated in the same step
  as every entry, with a check that it equals the sum of the entries. It reverses Business Rule 2.
  It suits a very large ledger or a shop that wants the stored figure as a control total. It is not
  needed for correctness: the entries are the record.
- **Hook into an existing employee table.** Uses the database's own employee table in place of
  `tblEmployee`.
- **Date-range entries.** A Taken entry holds a start date and an end date instead of one row per
  day. There are far fewer rows. In exchange, the one-category-per-day rule (Business Rule 9)
  becomes a check that two date ranges do not overlap, and cancelling part of a range needs a rule
  of its own, which this option adds.
- **Holidays excluded from time taken.** Counts only working days when a request spans several days.
- **Carry-over and expiry.** Limits how much of a balance carries into a new year, and when unused
  hours lapse.

## Parked / future considerations (not in this design)

- **Several balances on one screen.** Reporting across categories and employees is queries over the
  ledger and is left to the developer.
- **Approval of requests.** Who may post Taken entries, and whether a manager approves them first.
- **Opening balances.** Time already earned before the ledger was started. One way is an Earned
  entry for an accrual period before the employee's `AccrualStartDate`; this design does not say
  how an existing balance should be entered, and a shop moving over from another system decides.
- **Replacing a Correction.** A Correction cannot itself be corrected or replaced. To reinstate an
  entry that was cancelled by mistake, post a new entry; the history shows both.
