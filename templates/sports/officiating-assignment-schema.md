---
template: sports-officiating-assignment-schema
title: Sports Officiating Assignment — Table Schema
domain: sports
type: table-schema
version: 0.1.0
status: draft
standards_layer: [audit-columns, naming-conventions, error-handling]
new_tables:
  - tlkpPlayLevel
  - tlkpOfficialPosition
  - tblVenue
  - tblManager
  - tblTeam
  - tblOfficial
  - tblGame
  - tblGameOfficial
  - tblPositionRate
  - tblAppSetting
seeds:
  - tlkpOfficialPosition.Plate
  - tlkpOfficialPosition.Base
  - tblAppSetting.OfficialPhotoFolder
house_assumptions:
  - tblGameOfficial — crew-by-junction; the position set is seed data in tlkpOfficialPosition, so a bigger crew is a new seed row, never a schema change
  - tblPositionRate — pay rates are effective-dated by (PlayLevelID, PositionID); the applicable rate is the row with the latest EffectiveDate on or before the game date
  - tblManager — a team manager is a normalized entity with minimal contact fields; richer contact management is an Extra Option
  - tblGame — a game's play level is derived from its teams (both teams share one PlayLevel); it is not stored on the game
---

# Sports Officiating Assignment — Table Schema

## Intent

Tables for assigning **officials** (umpires, referees) to **games** by **position**, with
**effective-dated pay rates** by play level and position. Covers the recurring league-season
workload: who works which game, in what role, and what they're owed for it.

Shaped from a contributed real-world officiating back-end (genericized; schema only, no data).
The source hardcoded exactly two officiating slots as columns on the game (`Plate`, `Base`) and
then needed a UNION query to un-pivot them back into rows — the clearest possible case for the
**assignment junction** at this template's heart: positions become seed data, crew size becomes
a data decision, and the un-pivot query becomes a plain SELECT.

This is a greenfield template — it creates its whole schema and hooks into no host tables.

## Entities

### tlkpPlayLevel

Grain: one row per play level (age bracket, division tier).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `PlayLevelID` | AutoNumber | PK | Surrogate key |
| `PlayLevelName` | Text(30) | Required | Display name |
| `PlayLevelCode` | Text(4) | Nullable | Short code for rosters/schedules |

Indexes: PK on `PlayLevelID`; unique on `PlayLevelName`; unique on `PlayLevelCode`.

Seed rows (samples — replace with the league's real levels): 10U, 12U, 14U, 16U, 18U.

### tlkpOfficialPosition

Grain: one row per officiating position a game can require.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `PositionID` | AutoNumber | PK | Surrogate key |
| `PositionName` | Text(30) | Required | e.g. Plate, Base |
| `PositionCode` | Text(4) | Nullable | Short code |

Indexes: PK on `PositionID`; unique on `PositionName`; unique on `PositionCode`.

Seed rows (samples): Plate (P), Base (B). Growing the crew — a third base umpire, a field
referee — is a new seed row here plus assignment rows, never a schema change (Business Rule 1).

### tblVenue

Grain: one row per playing venue.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `VenueID` | AutoNumber | PK | Surrogate key |
| `VenueName` | Text(60) | Required | |

Indexes: PK on `VenueID`; unique on `VenueName`.

*(Reconstructed: the source's game table carried a venue FK pointing at a table that didn't
exist in the file.)*

### tblManager

Grain: one row per team manager. One manager may manage several teams.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ManagerID` | AutoNumber | PK | Surrogate key |
| `FirstName` | Text(50) | Required | |
| `LastName` | Text(50) | Required | |
| `CellPhone` | Text(25) | Nullable | |
| `EmailAddress` | Text(100) | Nullable | |

Indexes: PK on `ManagerID`; index on `LastName`.

*(Normalized from a bare text column on the source's team table, so team forms can offer a
manager combo/list box instead of retyped names.)*

### tblTeam

Grain: one row per team.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `TeamID` | AutoNumber | PK | Surrogate key |
| `TeamName` | Text(60) | Required | Unique within its play level |
| `DivisionName` | Text(40) | Nullable | |
| `PlayLevelID` | Long | FK → tlkpPlayLevel, Required | |
| `ManagerID` | Long | FK → tblManager, Nullable | |

Indexes: PK on `TeamID`; unique on (`PlayLevelID`, `TeamName`) — the same team name may recur
across levels; FK indexes on `PlayLevelID`, `ManagerID`.

### tblOfficial

Grain: one row per official.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `OfficialID` | AutoNumber | PK | Surrogate key |
| `FirstName` | Text(50) | Required | |
| `LastName` | Text(50) | Required | |
| `BirthDate` | Date/Time | Nullable | Age is derived from this, never stored (Business Rule 8) |
| `StreetAddress` | Text(100) | Nullable | Extend with City/Region fields per your house address convention |
| `PostalCode` | Text(10) | Nullable | |
| `CellPhone` | Text(25) | Nullable | |
| `HomePhone` | Text(25) | Nullable | |
| `EmailAddress` | Text(100) | Nullable | |
| `OfficialIsActive` | Boolean | Required | Default Yes; only active officials take assignments (Business Rule 3) |
| `PhotoFileName` | Text(100) | Nullable | File **name** only — the folder comes from `tblAppSetting` (Business Rule 9) |

Indexes: PK on `OfficialID`; index on `LastName`.

Derived (not stored): age, from `BirthDate`.

### tblGame

Grain: one row per scheduled game.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `GameID` | AutoNumber | PK | Surrogate key |
| `GameStart` | Date/Time | Required | Date **and** time in one field |
| `GameEnd` | Date/Time | Nullable | When present, must be later than `GameStart` (Business Rule 7) |
| `HomeTeamID` | Long | FK → tblTeam, Required | |
| `AwayTeamID` | Long | FK → tblTeam, Required | Must differ from `HomeTeamID` (Business Rule 4) |
| `VenueID` | Long | FK → tblVenue, Nullable | |

Indexes: PK on `GameID`; index on `GameStart`; FK indexes on `HomeTeamID`, `AwayTeamID`,
`VenueID`.

Derived (not stored): the game's play level, from its teams (Business Rule 5).

### tblGameOfficial

Grain: one row per official-position-game assignment — the junction that replaces the source's
hardcoded per-position columns.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `GameOfficialID` | AutoNumber | PK | Junction surrogate key per standards |
| `GameID` | Long | FK → tblGame, Required | |
| `OfficialID` | Long | FK → tblOfficial, Required | |
| `PositionID` | Long | FK → tlkpOfficialPosition, Required | |

Indexes: PK on `GameOfficialID`; unique on (`GameID`, `PositionID`) — one official per position
per game; unique on (`GameID`, `OfficialID`) — one position per official per game (both =
Business Rule 2); FK index on `OfficialID`.

### tblPositionRate

Grain: one row per (play level, position, effective date) — pay rates with effective dating,
formalized from the source's `EffDate` idea.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `PositionRateID` | AutoNumber | PK | Surrogate key |
| `PlayLevelID` | Long | FK → tlkpPlayLevel, Required | |
| `PositionID` | Long | FK → tlkpOfficialPosition, Required | |
| `PayRate` | Currency | Required | Must be ≥ 0 |
| `EffectiveDate` | Date/Time | Required | Rate applies from this date until a later row supersedes it |

Indexes: PK on `PositionRateID`; unique on (`PlayLevelID`, `PositionID`, `EffectiveDate`);
FK index on `PositionID`.

### tblAppSetting

Grain: one row per named application setting. Standalone by design — no relationships.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `AppSettingID` | AutoNumber | PK | Surrogate key |
| `SettingName` | Text(50) | Required | |
| `SettingValue` | Text(255) | Nullable | |
| `SettingDescription` | Text(255) | Nullable | |

Indexes: PK on `AppSettingID`; unique on `SettingName`.

Seed row: `OfficialPhotoFolder` = `Images\` — the folder holding official photos, relative to
the application folder unless given as an absolute path (Business Rule 9). *(This table exists
because the source displayed photos through a hardcoded local file path baked into a form
control — configuration belongs in data, not in controls.)*

## Relationships

- `tlkpPlayLevel (1) → (∞) tblTeam` on `PlayLevelID` — no cascade
- `tlkpPlayLevel (1) → (∞) tblPositionRate` on `PlayLevelID` — no cascade
- `tlkpOfficialPosition (1) → (∞) tblPositionRate` on `PositionID` — no cascade
- `tlkpOfficialPosition (1) → (∞) tblGameOfficial` on `PositionID` — no cascade
- `tblManager (1) → (∞) tblTeam` on `ManagerID` — no cascade; FK nullable
- `tblVenue (1) → (∞) tblGame` on `VenueID` — no cascade; FK nullable
- `tblTeam (1) → (∞) tblGame` on `HomeTeamID` — no cascade
- `tblTeam (1) → (∞) tblGame` on `AwayTeamID` — no cascade
- `tblGame (1) → (∞) tblGameOfficial` on `GameID` — cascade delete (an assignment has no
  meaning without its game)
- `tblOfficial (1) → (∞) tblGameOfficial` on `OfficialID` — no cascade (deleting an official
  with assignment history should be refused, not silently propagated)
- `tblAppSetting` — standalone; deliberately unrelated

## Business Rules

1. **Crew by junction.** A game's officiating crew is its `tblGameOfficial` rows. The position
   set is seed data in `tlkpOfficialPosition`; changing crew size or composition is a data
   change, never a schema change.
2. **Assignment uniqueness.** One official per position per game, and one position per official
   per game — enforced by the junction's two unique indexes, and checked in a friendly way
   before insert (see the paired scaffold's `ValidateAssignment`).
3. **Only active officials.** An assignment may only reference an official whose
   `OfficialIsActive` is Yes.
4. **A team cannot play itself.** `HomeTeamID` ≠ `AwayTeamID`.
5. **Game level is derived.** Both teams of a game share one `PlayLevelID`; the game's level is
   read through its teams, not stored on the game.
6. **Applicable pay rate.** For a given assignment, the rate is the `tblPositionRate` row for
   the game's play level and the assignment's position with the **latest `EffectiveDate` on or
   before the game date**. No matching row = no rate (surface a warning, don't guess).
7. **Game times are ordered.** When `GameEnd` is present, it is later than `GameStart`.
8. **Age is derived.** Never store an age; compute it from `BirthDate` when needed.
9. **Photos are file names plus one folder setting.** `tblOfficial.PhotoFileName` holds only
   the file name; the folder comes from the `OfficialPhotoFolder` row of `tblAppSetting`
   (relative to the application folder unless absolute). The folder is created and ensured at
   startup, and the photo picker copies the chosen file into it — so the stored name always
   resolves. No paths are ever hardcoded in controls or code.

## Standards Layer

- **Audit columns** — every table above additionally carries the audit set supplied by the
  active `audit-columns.md`; audit fields never appear in this template's field tables. (The
  source's per-table timestamp columns are retired in the audit set's favor.)
- **Naming conventions** — this template is written in the GPC house style (`tbl`/`tlkp`
  prefixes, `[Entity]ID` keys, qualified field names). A practice with different conventions
  regenerates the same entities under its own `naming-conventions.md` without editing this
  template.
- **Error handling** — any VBA generated alongside (see the paired
  `sports-officiating-assignment-scaffold`) takes its error pattern from the active
  `error-handling.md`.

## Extra Options

*Empty in the base template. Filled per engagement; the filled copy is saved to the developer's
own library, not committed here.*

- **Official qualifications** — an official × play level qualification/certification junction,
  limiting which levels an official may work. *(On its own domain merits; nothing in the source
  carried this.)*
- **Season / schedule** — a `tblSeason` grouping games, with per-season team registration.
- **Richer manager/contact management** — addresses, roles, one person managing and
  officiating.
- **Compensation ledger** — worked-game payment tracking (assignment × rate snapshot → amount
  owed/paid), if pay must survive later rate edits.

## Parked / future considerations (not in this design)

- **Provenance notes.** The source back-end contributed the domain and several good instincts
  (effective-dated rates) alongside teaching examples the template deliberately reverses:
  hardcoded position columns (→ the junction), a stored text age (→ derived from `BirthDate`),
  three coexisting image mechanisms with a hardcoded path (→ `PhotoFileName` +
  `tblAppSetting`), an orphaned venue FK (→ `tblVenue` reconstructed), and a leftover testing
  field the owner confirmed for deletion.
- **Person supertype.** Officials and managers share contact shape; deliberately **not**
  unified — a person supertype over-models a template of this size.
- **Game status workflow** (scheduled / played / cancelled / forfeited) — add as a lookup +
  FK when an engagement needs it; touches compensation rules.
