---
template: sports-officiating-assignment-form
title: Sports Officiating Assignment — Game Assignment Form
domain: sports
type: form-spec
version: 0.1.0
status: draft
implements: sports-officiating-assignment-schema
record_source: qryGame_frm
standards_layer: [form-conventions, naming-conventions]
new_forms: [frmGame_Assignment, sfrmGame_Crew]
---

# Sports Officiating Assignment — Game Assignment Form

## Intent

A default entry/edit form for scheduling a game and staffing its officiating crew
(`sports-officiating-assignment-schema`). The main form edits the game; the crew subform edits
the assignment junction — the UI expression of the schema's central idea, that a crew is rows,
not columns. This is a **default layout, not a finished design** — structurally and
functionally complete, but unstyled ("ugly but correct" is a pass); house design defaults and
any reusable form framework are deferred to the standards layer, which this template **names,
not redefines**.

## Prerequisites

| Object | Role |
|---|---|
| `sports-officiating-assignment-schema` (`implements`) | The tables this form edits |
| `qryGame_frm` | The form's record source over `tblGame` |
| `sports-officiating-assignment-scaffold` | The paired scaffold; the subform's validation calls its `ValidateAssignment`, the optional features call `GetAppSetting` / `GetApplicablePayRate` |
| `form-conventions.md` | House design defaults + the named reusable patterns (validation highlights, quick-add, audit display) |

## Layout

Regions, each with an ordered control inventory. Arrangement = region + order; positioning is
the materialization step's job (a default stacked layout, not pixel-placed).

### Form Header — the game

| Control | Type | Bound to | Notes |
|---|---|---|---|
| `txtGameStart` | Textbox | `GameStart` | Date+time; required |
| `txtGameEnd` | Textbox | `GameEnd` | Must be later than start when present (schema Business Rule 7) |
| `cboHomeTeamID` | Combo | `HomeTeamID` | → `tblTeam`; shows team + level |
| `cboAwayTeamID` | Combo | `AwayTeamID` | → `tblTeam`; must differ from home (Business Rule 4); same level (Business Rule 5) |
| `cboVenueID` | Combo | `VenueID` | → `tblVenue`; quick-add per `form-conventions` |

### Detail — the crew

| Control | Type | Bound to | Notes |
|---|---|---|---|
| `sfrmGame_Crew` | Subform | `tblGameOfficial` | Linked on `GameID`; continuous. Rows: `cboPositionID` (→ `tlkpOfficialPosition`) + `cboOfficialID` (→ `tblOfficial`, active only, "LastName, FirstName") |
| `imgOfficialPhoto` | Image | — (driven by `PhotoFileName`) | **Optional feature 1** — selected crew member's photo |
| `txtApplicableRate` | Textbox | — (computed) | **Optional feature 2** — read-only rate for the selected crew row |

*Hidden/internal (present, not user-facing):* `GameID` (PK, main form); `GameOfficialID` (PK,
subform rows).

### Form Footer

| Control | Type | Bound to | Notes |
|---|---|---|---|
| `cmdNew` / `cmdSave` / `cmdDeleteGame` + record nav | Button | — | from `form-conventions` |

## Features

1. **Crew as rows** — the subform edits `tblGameOfficial` directly: pick a position, pick an
   official, done. Adding a third crew slot is a `tlkpOfficialPosition` seed row appearing in
   the combo — no form change (Business Rule 1).
2. **Friendly assignment validation** — the subform's before-update calls the paired
   scaffold's `ValidateAssignment`, refusing duplicates and inactive officials with a plain
   message before the junction's unique indexes raise an engine error (Business Rules 2, 3).
   Surfaced with the **validation-highlights** pattern from `form-conventions.md`.
3. **Game-level checks** — home ≠ away (Business Rule 4) and matching play levels (Business
   Rule 5) validated on the main form's save, same highlight pattern.
4. **CRUD + navigation** — add / save / delete / clear + record navigation, from
   `form-conventions`.

**Optional features enabled in this form** *(engagement-specific — not baseline features):*

- **Official photo** — `imgOfficialPhoto` shows the selected crew member's photo on the
  subform's row change: folder from `GetAppSetting("OfficialPhotoFolder")` + the official's
  `PhotoFileName` (Business Rule 9). Skip it and nothing else changes.
- **Applicable rate display** — `txtApplicableRate` shows the selected crew row's pay via
  `GetApplicablePayRate` (Business Rule 6): read-only, the scaffold working visibly in the UI.

## Standards Layer

- **Form conventions** (`form-conventions.md`) — control-name prefixes
  (`txt`/`cbo`/`cmd`/`sfrm`/`img`), the standard button set + footer nav, tab order, default
  control type per data type, default sizing, and the named reusable patterns this form invokes
  (validation highlights, quick-add, audit display).
- **Naming** (`naming-conventions.md`) — form / subform / record-source names.
- **Error handling** — any code-behind follows the active `error-handling.md` (shared with the
  paired scaffold).

## Materialization

The form-spec materializes as importable Access form text (`LoadFromText`) with a default
stacked layout and the code-behind wired to the paired `vba-scaffold` procedures, or is built
live via the Access MCP `create_form` / `create_control` tools. The markdown is the source of
truth; see `_materialization.md` for the mapping rules.

## Extra Options

*Empty in the base template. Filled per engagement; the filled copy is saved to the developer's
own library, not committed here.*

- **Photo loader** — a `cmdLoadPhoto` file-dialog picker on the Official manage form that
  stores the chosen file's **name** into `PhotoFileName` (the folder stays a
  `tblAppSetting` concern). *(The contributed source's one keeper procedure is this pattern's
  ancestor.)*
- **Schedule view** — a read-only continuous games list with crew-completeness flags.
- **Official manage form** — a full entry/edit form for `tblOfficial` (mirror the library's
  publication-form pattern).

## Parked / future considerations (not in this design)

- **Assignment by drag / grid** — a crew board (officials × games) is a different UI paradigm;
  out of scope for a default layout.
- **Styling / polish** — out of scope by design (the adopter's pass).
