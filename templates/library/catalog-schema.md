---
template: library-catalog-schema
title: Library Publication Catalog — Table Schema
domain: library
type: table-schema
version: 0.1.0
status: draft
standards_layer:
  - audit-columns
  - naming-conventions
new_tables:
  - tblPublication
  - tblCreator
  - tblPublisher
  - tblPublicationCreator
  - tblPublicationGenre
  - tlkpGenre
  - tlkpMediaType
  - tlkpMediaCondition
  - tlkpBookcase
  - tlkpShelf
house_assumptions:
  - tblPublication — one record per title, not per physical copy. NumberOfVolumes counts duplicate copies (when MultiVolumeSet is false) or volumes in a set (when true); individual copies are not separately tracked or identified. A library that tracks individual physical copies would instead add a per-copy Holding entity — see Extra Options.
---

# Library Publication Catalog — Table Schema

## Intent

A **standalone catalog** for a non-circulating reference library: books and other publications,
their creators and publishers, subject genres, and a physical shelf location. Unlike the Northwind
stocktake template, this one **builds on nothing existing** — it is a complete, self-contained schema
(no `extends`, no host tables).

The model is **title-centric**: one row per title, never per physical copy. Duplicate copies and
multi-volume sets are both expressed through a count (`NumberOfVolumes`) disambiguated by a boolean
(`MultiVolumeSet`), rather than by per-copy records. That is a deliberate **house assumption**
(declared in front-matter) suited to a collection that does not sticker, barcode, or individually
track its items; a library that needs per-copy identity adds the Holding extension in Extra Options.

*(Shaped from a real reference collection and genericized; collection-specific provenance fields are
documented under Parked, not carried into the template.)*

## Entities

Naming follows the **GPC house style** (`tbl` for entity/junction tables, `tlkp` for lookups); a
different practice regenerates the same entities under its own conventions (see Standards Layer).
Audit columns are supplied by the standards layer and are intentionally absent below.

### tblPublication

Grain: one row per published **title** held in the collection.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `PublicationID` | AutoNumber | PK | Surrogate key |
| `PublicationTitle` | Memo | Required | The full title, untruncated — Memo handles long historical/rare titles that exceed 255 chars |
| `PublicationSortTitle` | Text(255) | Required | Indexable/sortable form of the title; **auto-derived** from `PublicationTitle` (leading noise words removed) on insert and refreshed on title edit. System-maintained, not user-entered — see Business Rules |
| `CatalogNumber` | Text(255) | Nullable | Optional title-level accession/catalog number (not a per-copy number — see house assumption) |
| `ISBN` | Text(50) | Nullable | ISBN, where known |
| `Edition` | Text(255) | Nullable | Edition statement |
| `Printing` | Text(255) | Nullable | Printing statement |
| `YearPublished` | Integer | Nullable | Year of publication |
| `Pages` | Long | Nullable | Page count |
| `PublisherID` | Long | FK → tblPublisher, Nullable | Publisher |
| `MediaTypeID` | Long | FK → tlkpMediaType, Nullable | Format / media type |
| `MediaConditionID` | Long | FK → tlkpMediaCondition, Nullable | Physical condition |
| `ShelfID` | Long | FK → tlkpShelf, Nullable | Shelf location |
| `Volume` | Long | Nullable | Which volume this record represents; governed by the MultiVolumeSet rule |
| `NumberOfVolumes` | Long | Required | Copy count (single-volume work) or volume count (set); defaults to 1; see MultiVolumeSet rule |
| `MultiVolumeSet` | Boolean | Required | False = `NumberOfVolumes` counts duplicate copies; True = a multi-volume set; defaults to False |
| `ListPrice` | Currency | Nullable | List/catalog price |
| `CoverImageLink` | Text(255) | Nullable | URL or path to a cover image |
| `PublicationComments` | Memo | Nullable | Public-facing notes |
| `PublicationInternalComments` | Memo | Nullable | Internal/staff notes |

Indexes: PK on `PublicationID`; non-unique index on `PublicationSortTitle` (the sort key);
non-unique indexes on `PublisherID`, `MediaTypeID`, `MediaConditionID`, `ShelfID` (FKs).

### tblCreator

Grain: one row per creator (author, editor, or other contributor).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `CreatorID` | AutoNumber | PK | Surrogate key |
| `CreatorLastName` | Text(100) | Required | Family name, or the single name of a corporate/anonymous creator |
| `CreatorFirstName` | Text(100) | Nullable | Given name |
| `CreatorMiddleName` | Text(100) | Nullable | Middle name(s) |

Indexes: PK on `CreatorID`; non-unique index on `CreatorLastName`.

### tblPublisher

Grain: one row per publisher.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `PublisherID` | AutoNumber | PK | Surrogate key |
| `PublisherName` | Text(255) | Required | Publisher name |

Indexes: PK on `PublisherID`.

### tblPublicationCreator

Junction — many-to-many between publications and creators. Grain: one row per
(`PublicationID`, `CreatorID`) pair.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `PublicationCreatorID` | AutoNumber | PK | Surrogate key |
| `PublicationID` | Long | FK → tblPublication, Required | The publication |
| `CreatorID` | Long | FK → tblCreator, Required | The creator |

Indexes: PK on `PublicationCreatorID`; **unique index on (`PublicationID`, `CreatorID`)** — a creator
links to a publication at most once; non-unique index on `CreatorID` (FK).

### tblPublicationGenre

Junction — many-to-many between publications and genres. Grain: one row per
(`PublicationID`, `GenreID`) pair.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `PublicationGenreID` | AutoNumber | PK | Surrogate key |
| `PublicationID` | Long | FK → tblPublication, Required | The publication |
| `GenreID` | Long | FK → tlkpGenre, Required | The genre |

Indexes: PK on `PublicationGenreID`; **unique index on (`PublicationID`, `GenreID`)**; non-unique
index on `GenreID` (FK).

### tlkpGenre

Lookup — subject genres.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `GenreID` | AutoNumber | PK | Surrogate key |
| `GenreName` | Text(255) | Required | Genre label |
| `SortOrder` | Long | Nullable | Display order |

Indexes: PK on `GenreID`.

### tlkpMediaType

Lookup — format / media type (e.g. hardcover, paperback, periodical).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `MediaTypeID` | AutoNumber | PK | Surrogate key |
| `MediaTypeName` | Text(255) | Required | Media-type label |
| `SortOrder` | Long | Nullable | Display order |

Indexes: PK on `MediaTypeID`.

### tlkpMediaCondition

Lookup — physical condition grades (e.g. fine, good, fair, poor).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `MediaConditionID` | AutoNumber | PK | Surrogate key |
| `MediaConditionName` | Text(255) | Required | Condition label |
| `SortOrder` | Long | Nullable | Display order |

Indexes: PK on `MediaConditionID`.

### tlkpBookcase

Lookup — physical bookcases; parent of `tlkpShelf`.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `BookcaseID` | AutoNumber | PK | Surrogate key |
| `BookcaseName` | Text(50) | Required | Bookcase label |
| `BookcaseDescription` | Text(150) | Nullable | Free-text description |

Indexes: PK on `BookcaseID`.

### tlkpShelf

Lookup — physical shelves; each belongs to one bookcase.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ShelfID` | AutoNumber | PK | Surrogate key |
| `ShelfName` | Text(50) | Required | Shelf label |
| `ShelfDescription` | Text(100) | Nullable | Free-text description |
| `BookcaseID` | Long | FK → tlkpBookcase, Required | Owning bookcase |

Indexes: PK on `ShelfID`; non-unique index on `BookcaseID` (FK).

## Relationships

- `tblPublisher (1) → (∞) tblPublication` on `PublisherID` — no cascade
- `tlkpMediaType (1) → (∞) tblPublication` on `MediaTypeID` — no cascade
- `tlkpMediaCondition (1) → (∞) tblPublication` on `MediaConditionID` — no cascade
- `tlkpShelf (1) → (∞) tblPublication` on `ShelfID` — no cascade
- `tlkpBookcase (1) → (∞) tlkpShelf` on `BookcaseID` — no cascade (location hierarchy)
- `tblPublication (1) → (∞) tblPublicationCreator` on `PublicationID` — **cascade delete** (creator links are meaningless without the publication)
- `tblCreator (1) → (∞) tblPublicationCreator` on `CreatorID` — no cascade
- `tblPublication (1) → (∞) tblPublicationGenre` on `PublicationID` — **cascade delete**
- `tlkpGenre (1) → (∞) tblPublicationGenre` on `GenreID` — no cascade

## Business Rules

1. **One record per title** — the catalog is title-centric (see the declared house assumption). A
   title with multiple duplicate copies remains one row.
2. **Sort title is auto-derived** — `PublicationSortTitle` is maintained automatically from
   `PublicationTitle`: leading noise words (`The`, `A`, `An`, …) are stripped and the result stored
   (truncated to 255). It is set when a record is created and refreshed whenever the title is edited;
   it is system-maintained, not user-entered. *(The derivation mechanism is platform-specific — an
   Access data macro, a SQL Server trigger, or compute-on-write — and is supplied by the
   implementation / standards layer, not specified here.)*
3. **Volume / set validation** *(record-level validation rule — Access table `Validation Rule` or a
   SQL `CHECK` constraint):*
   - `NumberOfVolumes` is always required and defaults to **1**.
   - When `MultiVolumeSet = True`: `Volume` is **required** and must fall in **1 … NumberOfVolumes**,
     and `NumberOfVolumes` must be **≥ 2** (a set has more than one volume).
   - When `MultiVolumeSet = False`: `Volume` must be **Null**.
   - *(The companion UI behavior — enabling/disabling the `Volume` control by `MultiVolumeSet` — is a
     form concern, deferred to a `form-spec` template.)*
4. **No duplicate creator or genre links** — enforced by the unique indexes on
   (`PublicationID`, `CreatorID`) and (`PublicationID`, `GenreID`).
5. **Deleting a publication** cascades to its creator and genre links, never to the looked-up
   creators, publishers, genres, or locations.

## Standards Layer (supplied externally, not in this template body)

- **Audit columns** — `CreatedDate`, `CreatedBy`, `ModifiedDate`, `ModifiedBy`, `AccessTS` on every
  table, supplied by the standards layer (see `standards/audit-columns.md`). *The source collection
  had no audit columns; the standards layer supplies them — a clean demonstration of the split.*
- **Naming conventions** — this template follows the GPC `tbl`/`tlkp` house style and the
  qualified-field rules (see `standards/naming-conventions.md`). A different practice regenerates the
  same entities under its own conventions without editing this template.

## Extra Options (engagement-specific — stubs)

*Empty in the base template. Filled per engagement; the filled copy is saved to the developer's own
library, not committed here.*

- **Per-copy holdings** — for libraries that *do* track individual physical copies: add a `tblHolding`
  entity (one row per copy: own catalog/accession number, condition, location, acquisition date),
  child of `tblPublication`. This replaces the title-level copy count with true per-copy records (the
  alternative to this template's declared one-record-per-title assumption).
- **Circulation / lending** — a standard loan model (borrowers, loans, due dates, returns). A
  well-understood pattern; omitted from the core only because the source collection is
  non-circulating. Straightforward to add when the library lends.
- **Subject / keyword tagging** — controlled-vocabulary tagging of publications for discovery.
  Genuinely valuable, but the governance model (a *curated* vocabulary, synonym control, preventing
  vocabulary sprawl, approved/banned terms) is the hard part and needs deliberate design before use —
  not a bolt-on. *(Parked as a known-thorny future feature.)*
- **Creator role** — distinguish author / editor / translator / illustrator on `tblPublicationCreator`
  by adding a `CreatorRoleID` (FK to a `tlkpCreatorRole` lookup).

## Parked / future considerations (not in this design)

- **Source provenance (excluded fields).** This template was shaped from a real reference collection
  whose records were assembled by matching volunteer-supplied title/author lists against the Google
  Books API. The source carried collection-specific provenance fields — a match-confidence score
  (`ConfidenceLevel`), the matched Google publisher string (`PublisherGoogle`), and a separate
  capture date (`DateCaptured`) — all artifacts of that one-time, API-assisted build process. They
  are documented here but **excluded** from the generic template. (A general "external-source
  cataloging" extension — capturing match confidence and source provenance when a catalog is built
  from a third-party bibliographic API — could be a future extra option if the pattern recurs.)
