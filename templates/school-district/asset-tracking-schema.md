---
template: school-district-asset-tracking-schema
title: School District Capital Asset Tracking
domain: school-district
type: table-schema
version: 0.2.0
status: draft
standards_layer: [audit-columns, naming-conventions, error-handling, query-style]
new_tables: [tblAsset, tblAssetHistory, tblInventoryAuditSession, tblInventoryAuditScan, tblSite, tblRoom, tblDepartment, tblCustodian, tlkpAssetCategory, tlkpAssetStatus, tlkpFundingSource, tlkpDepreciationMethod, tlkpHistoryChangeType, tlkpScanResult]
house_assumptions:
  - "tblDepartment — modeled as district-wide, independent of Site, because Site and Department are documented as separate, non-nested reporting axes; a practice that treats departments as site-scoped must regenerate this entity"
  - "tblAsset.VendorName / PurchaseOrderNumber — kept as plain text fields rather than a full vendor entity, deferring vendor master-data modeling to a future template"
  - "tlkpDepreciationMethod — seeded with only Straight-Line; the schema carries a method FK for extensibility, but no alternate depreciation calculation is implemented"
---

# School District Capital Asset Tracking

## Intent

Tracks every capital asset above a district's capitalization threshold — computers, vehicles,
land, cameras, furniture, HVAC equipment, network equipment — from acquisition through disposal.
Each asset carries a unique barcode, a current location (site + room), an owning department, an
accountable custodian, a funding source, depreciation inputs, and warranty dates. A standing
movement/change log gives a complete audit history independent of the host's row-level audit
columns, and a barcode-scan inventory-verification flow supports periodic physical counts by
site, reconciling what's on file against what's actually found.

This schema is built entirely new — it needs no existing tables and doesn't connect into an
existing database. (In database terms: a *greenfield* schema.)

## Prerequisites

This template hooks into no existing tables — it builds its whole schema from nothing. What does
have to be settled before you build is **which file the tables go in**.

**Where these tables live.** These templates are designed for a **split database** — the normal
shape for Access applications, especially those in multi-user environments: one file holds the
tables (the **back end**, usually on a shared network drive — never on OneDrive, Dropbox, or any
other file-syncing cloud folder, which corrupts a shared Access back end), and each person runs
their own copy of a second file holding the forms, reports, and code (the **front end**), whose
tables are *links* pointing at the back end. A single-file database — one .accdb holding
everything — is an acceptable choice for one user, and everything here works there too: create
everything in that one file and ignore the distinction.

All fourteen tables below — eight entities and six lookups — are created in the **back end** and
linked into each front end. A district running inventory counts at several sites at once is a
multi-user application by definition: each site's counting station is its own front end, all
writing scans into the one shared back end.

## Entities

### tblAsset

Grain: one row per physical capital asset.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `AssetID` | AutoNumber | PK | Surrogate key |
| `AssetBarcode` | Text(50) | Required | The asset tag scanned for inventory verification; unique |
| `AssetCategoryID` | Long | FK → tlkpAssetCategory, Required | Computers / vehicles / land / cameras / furniture / HVAC equipment / network equipment |
| `AssetDescription` | Text(255) | Required | Short identifying description |
| `Manufacturer` | Text(150) | Nullable | |
| `ModelNumber` | Text(100) | Nullable | |
| `SerialNumber` | Text(100) | Nullable | Manufacturer serial; distinct from `AssetBarcode` |
| `RoomID` | Long | FK → tblRoom, Required | Current location; site is derived via `tblRoom.SiteID`, not duplicated here |
| `DepartmentID` | Long | FK → tblDepartment, Required | Owning department |
| `CustodianID` | Long | FK → tblCustodian, Required | Person accountable for the asset |
| `FundingSourceID` | Long | FK → tlkpFundingSource, Required | General fund / bond measure / categorical-grant / donation, etc. |
| `AcquisitionDate` | Date/Time | Required | When the asset was placed in service |
| `AcquisitionCost` | Currency | Required | Must exceed the district's capitalization threshold (Business Rule 1) |
| `VendorName` | Text(255) | Nullable | See `house_assumptions` — plain text, no vendor entity |
| `PurchaseOrderNumber` | Text(50) | Nullable | |
| `DepreciationMethodID` | Long | FK → tlkpDepreciationMethod, Required | Straight-line by default; see `house_assumptions` |
| `UsefulLifeYears` | Integer | Required | Drives the depreciation schedule |
| `SalvageValue` | Currency | Required | Floor for computed book value; default 0 |
| `WarrantyStartDate` | Date/Time | Nullable | |
| `WarrantyExpirationDate` | Date/Time | Nullable | |
| `AssetStatusID` | Long | FK → tlkpAssetStatus, Required | Active / under repair / surplus / missing / disposed |
| `DisposalDate` | Date/Time | Nullable | Required when `AssetStatusID` = Disposed (Business Rule 5) |
| `DisposalMethod` | Text(255) | Nullable | Sold / scrapped / donated / traded in |
| `AssetNotes` | Memo | Nullable | |

Indexes: PK on `AssetID`; unique index on `AssetBarcode`; FK indexes on `AssetCategoryID`,
`RoomID`, `DepartmentID`, `CustodianID`, `FundingSourceID`, `DepreciationMethodID`, `AssetStatusID`.

Derived (not stored): `AnnualDepreciation = (AcquisitionCost − SalvageValue) / UsefulLifeYears`;
`CurrentBookValue` — straight-line from `AcquisitionDate`, floored at `SalvageValue`, frozen once
`UsefulLifeYears` has elapsed; `AssetAgeYears`.

### tblAssetHistory

Grain: one row per recorded change to an asset's location, custodian, department, funding source,
or status — the standing audit/movement trail, distinct from the row-level audit columns deferred
to the standards layer.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `AssetHistoryID` | AutoNumber | PK | Surrogate key |
| `AssetID` | Long | FK → tblAsset, Required | The asset that changed |
| `HistoryChangeTypeID` | Long | FK → tlkpHistoryChangeType, Required | Location / custodian / department / funding source / status |
| `ChangeDate` | Date/Time | Required | When the change took effect |
| `PreviousValue` | Text(255) | Nullable | Prior value (room name, custodian name, status, etc.) |
| `NewValue` | Text(255) | Required | New value |
| `ChangeReason` | Text(255) | Nullable | Free text — e.g. "room reassignment," "staff turnover" |

Indexes: PK on `AssetHistoryID`; FK index on `AssetID`.

### tblInventoryAuditSession

Grain: one row per physical-inventory verification event at a site.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `InventoryAuditSessionID` | AutoNumber | PK | Surrogate key |
| `SiteID` | Long | FK → tblSite, Required | School site being verified |
| `AuditDate` | Date/Time | Required | |
| `ConductedByCustodianID` | Long | FK → tblCustodian, Nullable | Staff member running the count |
| `AuditNotes` | Memo | Nullable | |

Indexes: PK on `InventoryAuditSessionID`; FK index on `SiteID`.

### tblInventoryAuditScan

Grain: one row per barcode scan recorded within an inventory-audit session.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `InventoryAuditScanID` | AutoNumber | PK | Surrogate key |
| `InventoryAuditSessionID` | Long | FK → tblInventoryAuditSession, Required | Owning session |
| `ScannedBarcode` | Text(50) | Required | Raw scanned value, before resolution |
| `AssetID` | Long | FK → tblAsset, Nullable | Resolved asset; null if the barcode matched nothing on file |
| `ScanResultID` | Long | FK → tlkpScanResult, Required | Matched / unexpected / duplicate (Business Rule 6) |
| `ScannedRoomID` | Long | FK → tblRoom, Nullable | Room the asset was physically found in; flags a misplaced asset when it differs from `tblAsset.RoomID` |
| `ScannedOn` | Date/Time | Required | |

Indexes: PK on `InventoryAuditScanID`; FK indexes on `InventoryAuditSessionID`, `AssetID`.

### tblSite

Grain: one row per school site / district facility.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `SiteID` | AutoNumber | PK | Surrogate key |
| `SiteName` | Text(150) | Required | |
| `SiteCode` | Text(20) | Required | District-assigned site code; unique |

Indexes: PK on `SiteID`; unique index on `SiteCode`.

### tblRoom

Grain: one row per room/space within a site.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `RoomID` | AutoNumber | PK | Surrogate key |
| `SiteID` | Long | FK → tblSite, Required | Owning site |
| `RoomNumber` | Text(20) | Required | |
| `RoomName` | Text(100) | Nullable | |

Indexes: PK on `RoomID`; unique index on (`SiteID`, `RoomNumber`); FK index on `SiteID`.

### tblDepartment

Grain: one row per district department. See `house_assumptions` — modeled as district-wide,
not site-scoped.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `DepartmentID` | AutoNumber | PK | Surrogate key |
| `DepartmentName` | Text(100) | Required | Unique |

Indexes: PK on `DepartmentID`; unique index on `DepartmentName`.

### tblCustodian

Grain: one row per staff member who can be assigned asset custody or conduct an inventory audit.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `CustodianID` | AutoNumber | PK | Surrogate key |
| `CustodianFirstName` | Text(75) | Required | |
| `CustodianLastName` | Text(75) | Required | |
| `CustodianEmail` | Text(255) | Nullable | |
| `DepartmentID` | Long | FK → tblDepartment, Nullable | Home department |
| `CustodianActive` | Boolean | Required | Default True |

Indexes: PK on `CustodianID`; FK index on `DepartmentID`.

### Lookup tables

Trivial, uniform lookups (`<name>ID` + descriptor + `SortOrder`), grouped here per §4 of the
canonical format — each row is still its own discrete table.

| Table | Seed rows |
|---|---|
| `tlkpAssetCategory` | Computers; Vehicles; Land; Cameras; Furniture; HVAC Equipment; Network Equipment |
| `tlkpAssetStatus` | Active; Under Repair; Surplus; Missing; Disposed |
| `tlkpFundingSource` | General Fund; Bond Measure; Categorical/Grant; Donation |
| `tlkpDepreciationMethod` | Straight-Line |
| `tlkpHistoryChangeType` | Location; Custodian; Department; Funding Source; Status |
| `tlkpScanResult` | Matched; Unexpected; Duplicate |

## Relationships

- `tblSite (1) → (∞) tblRoom` on `SiteID` — restrict (a site can't be deleted while rooms exist)
- `tblRoom (1) → (∞) tblAsset` on `RoomID` — restrict
- `tblDepartment (1) → (∞) tblAsset` on `DepartmentID` — restrict
- `tblCustodian (1) → (∞) tblAsset` on `CustodianID` — restrict
- `tlkpFundingSource (1) → (∞) tblAsset` on `FundingSourceID` — restrict
- `tlkpAssetCategory (1) → (∞) tblAsset` on `AssetCategoryID` — restrict
- `tlkpAssetStatus (1) → (∞) tblAsset` on `AssetStatusID` — restrict
- `tlkpDepreciationMethod (1) → (∞) tblAsset` on `DepreciationMethodID` — restrict
- `tblAsset (1) → (∞) tblAssetHistory` on `AssetID` — cascade delete (history is meaningless without the asset; in practice assets are disposed, never hard-deleted, so this rarely fires)
- `tlkpHistoryChangeType (1) → (∞) tblAssetHistory` on `HistoryChangeTypeID` — restrict
- `tblSite (1) → (∞) tblInventoryAuditSession` on `SiteID` — restrict
- `tblCustodian (1) → (∞) tblInventoryAuditSession` on `ConductedByCustodianID` — restrict
- `tblInventoryAuditSession (1) → (∞) tblInventoryAuditScan` on `InventoryAuditSessionID` — cascade delete
- `tblAsset (1) → (∞) tblInventoryAuditScan` on `AssetID` — restrict
- `tlkpScanResult (1) → (∞) tblInventoryAuditScan` on `ScanResultID` — restrict
- `tblRoom (1) → (∞) tblInventoryAuditScan` on `ScannedRoomID` — restrict
- `tblDepartment (1) → (∞) tblCustodian` on `DepartmentID` — restrict

## Business Rules

1. **Capitalization threshold** — `tblAsset.AcquisitionCost` must exceed the district's
   capitalization floor (default $500) at save time.
2. **Barcode uniqueness** — every asset carries a barcode; `tblAsset.AssetBarcode` is enforced
   unique by index.
3. **Standing audit trail** — any change to `tblAsset.RoomID`, `CustodianID`, `DepartmentID`,
   `FundingSourceID`, or `AssetStatusID` writes a row to `tblAssetHistory` capturing the prior and
   new values and the change type. This is independent of, and in addition to, the row-level
   audit columns deferred to the standards layer (§6).

   **This template does not decide *how* that row gets written, and in a split database the two
   ways differ in what they catch.** Choose deliberately:

   - **A Before Change data macro** attaches to `tblAsset` itself, in the back end. It fires for
     every change to the table however it was made — a form, a query, an import, a bulk update,
     someone typing directly into the table. **If that macro calls a VBA function** (to get the
     current user's name, for instance), that function must exist in **every front end** as well as
     the back end, because a macro fired by an edit made through a link looks for the function in
     *that person's* front end. A missing copy shows up as a failed save, not a missing history
     row. This is the one placement mistake that testing on a single file can never reveal.
   - **Form or VBA code in the front end** fires only for changes made through the form that
     carries it. Anything else — an import, a query, a direct table edit, a second form somebody
     writes later — changes the asset and writes no history row. The trail is silently incomplete,
     and it looks correct until someone audits it.

   For a district asset register, where the point of the trail is accountability, the data macro is
   the stronger choice. If you use front-end code instead, know that its coverage stops at the form.
4. **Depreciation, computed not stored** — `CurrentBookValue` is derived at query time from
   `AcquisitionDate`, `UsefulLifeYears`, and `SalvageValue` using the method named by
   `DepreciationMethodID`; it is floored at `SalvageValue` and frozen once `UsefulLifeYears` has
   elapsed. Only Straight-Line is seeded (see `house_assumptions`).
5. **Disposal** — `DisposalDate` is required once `AssetStatusID` = Disposed. Disposed assets are
   retained, never deleted, so the audit history and any prior inventory-scan history stay intact.
6. **Inventory verification** — an `tblInventoryAuditScan` row resolves `ScannedBarcode` to
   `AssetID`. `ScanResultID` = Matched when the asset is on file as expected at the session's
   site; Unexpected when the asset's on-file room belongs to a different site; Duplicate on a
   repeat scan of the same barcode within the same session. Assets expected at a site (Active
   status, `tblRoom.SiteID` = the session's `SiteID`) with no Matched scan in the session are
   reportable as missing — this is a query against the session, not a stored flag.
7. **No unassigned assets** — `RoomID`, `DepartmentID`, `CustodianID`, and `FundingSourceID` are
   required on every asset; an asset is never on file without a location, owning department,
   accountable custodian, and funding source.
8. **Reporting axes are direct joins** — site (via `tblRoom.SiteID`), room, department, and
   category are each a single-join lookup off `tblAsset`, so reporting and inventory verification
   by site, room number, department, and asset category require no derived or multi-hop joins.

## Standards Layer

- **Audit columns** — `AddedBy` / `AddedOn` / `ModifiedBy` / `ModifiedOn` on every `tbl`/`tlkp`
  table, supplied by the host's audit convention. Not present in the field tables above; see
  Business Rule 3 for the separate, template-defined movement/change history, which is not a
  substitute for these.
- **Naming conventions** — this template follows GPC's `tbl`/`tlkp` prefix policy and
  field-qualification rules (`Status` → `<Entity>StatusID`, `Notes` → `<Entity>Notes`). A
  practice on a different naming convention regenerates the same entities under its own policy
  without editing this file.
- **Error handling** — any VBA generated against this schema (history-stamping logic, scan
  resolution, depreciation calculation) uses the house `errHandler` / global-error pattern.
- **Query style** — SQL behind any reporting queries or VBA recordsets follows the house
  aliasing/qualification and formatting conventions.

## Parked / future considerations

- Vendor master-data table (replacing the plain-text `VendorName`/`PurchaseOrderNumber` fields).
- Multi-method depreciation (declining-balance, units-of-production) beyond the seeded
  Straight-Line row.
- Site-scoped departments, for practices that organize departments per school rather than
  district-wide.
- A paired `vba-scaffold` for barcode-scan resolution and history-stamping, and a paired
  `form-spec` for asset entry/edit and inventory-audit data capture.

## Extra Options

*Empty in the base template. Filled per client engagement.*
- District-specific capitalization threshold, if different from $500.
- GASB 34 / state-reporting field extensions, if the district requires fields beyond this core set.
