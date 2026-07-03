---
template: library-catalog-publication-form
title: Library Catalog — Publication Entry Form
domain: library
type: form-spec
version: 0.1.0
status: draft
implements: library-catalog-schema
record_source: qryPublication_frm
standards_layer: [form-conventions, naming-conventions]
new_forms: [frmPublication_Edit, sfrmPublication_Creator, sfrmPublication_Genre]
---

# Library Catalog — Publication Entry Form

## Intent

A default entry/edit form for the catalog's publications (`library-catalog-schema`). It provides a
working, feature-complete layout for cataloguing a title and navigating the collection through a
**layered record selector**. This is a **default layout, not a finished design** — structurally and
functionally complete, but unstyled; the adopter applies their own polish ("ugly but correct" is a
pass). House design defaults and the reusable form framework are deferred to the standards layer:
this template **names** the framework pieces it relies on, it does not redefine them.

## Prerequisites

| Object | Role |
|---|---|
| `library-catalog-schema` (`implements`) | The tables this form edits |
| `qryPublication_frm` | The form's record source over `tblPublication`; its `WHERE` is rewritten by the selector |
| `form-conventions.md` | House design defaults **+** the named reusable patterns (layered selector, quick-add, validation highlights, audit display) |
| Forms framework (deferred) | `RefreshSQLWhere`, the alpha-filter routine, `TempVars` wrappers, and the audit/image/validation helpers — supplied by the host, **named not redefined** here |

## Layout

Regions, each with an ordered control inventory. Arrangement = region + order; positioning is the
materialization step's job (a default stacked layout, not pixel-placed).

### Form Header — the layered record selector

*(See Features for the behavior. The alpha selector is a reusable framework component, named here, not defined by this template.)*

| Control | Type | Bound to | Notes |
|---|---|---|---|
| `cboFilterGenre` | Combo | — (unbound) | filters the selector list by genre |
| `sfrmAlphaSelect` | Subform | — | A–Z + All (`*`) + non-alpha (`#`); framework component |
| `txtKeywordFilter` | Textbox | — (unbound) | keyword(s), comma-separated |
| `cmdGoKeyword` | Button | — | runs the keyword search |
| `cboSelectPublication` | Combo | — (unbound) | the selector; its row source is rebuilt by the filters, its selection rewrites the form |
| `txtSelectedCount` | Textbox | — (=expression) | "*n* items" — the count of the current selector list |

### Detail

| Control | Type | Bound to | Notes |
|---|---|---|---|
| `txtPublicationTitle` | Textbox | `PublicationTitle` | Memo; multi-line |
| `cboPublisherID` | Combo | `PublisherID` | → `tblPublisher`; `cmdAddPublisher` quick-add |
| `cmdAddPublisher` | Button | — | quick-add a publisher |
| `sfrmPublication_Creator` | Subform | `tblPublicationCreator` | M:N creators; `cmdAddAuthor` quick-add |
| `cmdAddAuthor` | Button | — | quick-add a creator |
| `sfrmPublication_Genre` | Subform | `tblPublicationGenre` | M:N genres; `cmdAddGenre` quick-add |
| `cmdAddGenre` | Button | — | quick-add a genre |
| `cboMediaTypeID` | Combo | `MediaTypeID` | → `tlkpMediaType` |
| `cboMediaConditionID` | Combo | `MediaConditionID` | → `tlkpMediaCondition` |
| `cboShelfID` | Combo | `ShelfID` | → `tlkpShelf` |
| `txtYearPublished` | Textbox | `YearPublished` | |
| `txtEdition` | Textbox | `Edition` | |
| `txtPrinting` | Textbox | `Printing` | |
| `txtCatalogNumber` | Textbox | `CatalogNumber` | |
| `txtISBN` | Textbox | `ISBN` | |
| `txtPages` | Textbox | `Pages` | |
| `txtListPrice` | Textbox | `ListPrice` | Currency |
| `chkMultiVolumeSet` | Checkbox | `MultiVolumeSet` | bound; the *Volume-toggle UI behavior* is Parked (see below) |
| `txtVolume` | Textbox | `Volume` | |
| `txtNumberOfVolumes` | Textbox | `NumberOfVolumes` | default 1 |
| `txtPublicationComments` | Textbox | `PublicationComments` | Memo; multi-line |
| `txtPublicationInternalComments` | Textbox | `PublicationInternalComments` | Memo; multi-line |
| `imgCoverImage` + `lblNoImage` | Image / Label | — (driven by `CoverImageLink`) | cover display (framework `DisplayImage`) |

*Hidden/internal (present, not user-facing):* `PublicationID` (PK), `PublicationSortTitle`
(system-maintained sort key — the selector orders on it), `CoverImageLink` (drives the cover image).
*Excluded:* `ConfidenceLevel` (a source-provenance artifact the schema parks out).

### Form Footer

| Control | Type | Bound to | Notes |
|---|---|---|---|
| `cmdNew` / `cmdSave` / `cmdDeletePub` / `cmdClearForm` + record nav | Button | — | from `form-conventions` |

## Features

1. **Layered record selector ("All or One")** — *the headline; defers to the framework.* Three filter
   inputs rebuild the selector's row source, and the selector's selection rewrites the form:
   - **(a) Genre + Alpha** — `cboFilterGenre` and `sfrmAlphaSelect` drive the genre/alpha row-source
     query. The chosen letter rides in a **`TempVar`** (`Alpha1stChar`), per `query-style.md`, so the
     filter is decoupled from any one control.
   - **(b) Keyword** — `txtKeywordFilter` + `cmdGoKeyword` swap the selector to a dynamically-built
     keyword query (title + comments `LIKE`, comma-separated terms).
   - **(c) Select** — choosing in `cboSelectPublication` rewrites the form's record-source `WHERE`
     (via the framework's `RefreshSQLWhere`): one publication, or **All** when the `<All Publications>`
     entry is chosen, ordered by `PublicationSortTitle`.
   - **Generalization:** the `<All Publications>` entry is made **consistent across all three filter
     modes** (genre, alpha, keyword), so "All or One" holds wherever it makes sense — not just in the
     keyword path.
2. **Quick-add lookups** — `cmdAddPublisher` / `cmdAddAuthor` / `cmdAddGenre`: open the matching
   manage form in add mode (dialog), return the new ID via a `TempVar`, then requery + select it. One
   pattern, deferred to the framework.
3. **Validation highlights** — invalid fields are flagged on save and cleared on move (a **core**
   framework pattern, `form-conventions.md`).
4. **CRUD + navigation** — add / save / delete / clear + record navigation, from `form-conventions`.

**Optional features enabled in this form** *(engagement-specific — `form-conventions.md` Optional
patterns; not baseline form features):*

- **Cover image** — `imgCoverImage` / `lblNoImage` display `CoverImageLink` via the framework's
  `DisplayImage`. Not every catalog has cover images.
- **Audit display** — a per-record audit summary, the form-side surfacing of the `audit-columns`
  standard. Fully supported here; non-mandatory.

## Standards Layer

- **Form conventions** (`form-conventions.md`) — control-name prefixes (`txt`/`cbo`/`chk`/`cmd`/`sfrm`),
  the standard button set + footer nav, tab order, default control type per data type (combo for an FK,
  checkbox for Boolean, multi-line for Memo), default sizing, and the **named reusable patterns** the
  form invokes (layered selector, quick-add, validation highlights, audit display).
- **Naming** (`naming-conventions.md`) — form / subform / record-source names.
- **Forms framework** (host-supplied, named not redefined) — `RefreshSQLWhere`, the alpha-filter
  routine, the `TempVars` wrappers, and the audit / image / validation helpers. *(The layered-selector
  engine is a tracked future `vba-scaffold` candidate.)*

## Materialization

The form-spec materializes as importable Access form text (`LoadFromText`) with a default stacked
layout and the code-behind wired in — event handlers calling the framework helpers (and any paired
`vba-scaffold` procedures). The markdown → Access-text mapping is **defined and hand-validated** as
part of this template type's proof; the actual generator is built in the MCP phase (B3). Alternatively,
the form is built live via the Access MCP `create_form` / `create_control` tools.

## Extra Options

*Empty in the base template. Filled per engagement; the filled copy is saved to the developer's own
library, not committed here.*

- **Read-only card view** — a `frmPublication_View` variant for browsing without edit.
- **Cover image from clipboard / file picker** — set `CoverImageLink` without hand-typing a path.
- **Per-copy holdings subform** — only when the schema's per-copy `Holding` extra option is adopted.

## Parked / future considerations (not in this design)

- **`MultiVolumeSet` → `Volume` UI toggle** — the schema's Business Rule #3 companion UI behavior
  (enable/require `Volume` from the checkbox). **Deferred pending client feedback (~3–4 months); it is
  not in the production copy either.** The data-level validation remains the table's; only the form's
  enable/disable affordance is parked.
- **Layered-selector `vba-scaffold`** — the selector engine (`RefreshSQLWhere`, `FilterPubTitles`, the
  alpha subform, the row-source swap) is a reusable pattern that belongs in its own `vba-scaffold`. This
  form-spec names it as a deferred dependency; building that scaffold is tracked as a future step.
- **MRU "recently edited"** — a most-recently-used list of titles; a framework feature, off by default.
- **Styling / polish** — out of scope by design (the adopter's pass).
