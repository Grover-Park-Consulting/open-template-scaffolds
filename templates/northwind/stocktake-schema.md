---
template: northwind-stocktake-schema
title: Northwind Scanned Stocktake — Table Schema
domain: northwind
type: table-schema
version: 0.1.0
status: draft
extends: Northwind (Access Developer Edition)
requires_tables:
  - Products
  - Employees
  - SystemSettings
requires_fields:
  - Products.ProductID
  - Products.SKUBarCode
  - Products.QuantityInPackage
standards_layer:
  - audit-columns
  - naming-conventions
  - error-handling
new_tables:
  - StockTakeSession
  - StockTakeCount
  - StockTakeScan
  - StockTakeStatus
  - StockTakeCountMethod
  - ScanStatus
  - ProductShrinkageAllowance
  - RemediationStatus
seeds:
  - SystemSettings.DefaultAllowableShrinkageRate
house_assumptions:
  - "StockTakeCount.CountedQuantity stored, not derived — a reconciled count is a durable audit fact business decisions rely on; it must not change if scan detail is later edited or archived. Stored despite being derivable; the alternative is to compute it on demand."
---

# Northwind Scanned Stocktake — Table Schema

## Intent

Add a **real, scan-driven stocktake** to a Northwind-derived database. Base Northwind's
`StockTake` records only a pseudo count for demo purposes — aggregate quantities hand-typed after the
fact, with no audit trail of how they were derived. This template replaces that with a
three-level structure that supports **two count methods over one schema**:

- **Manual count (Level 1):** a counter records a quantity per product directly.
- **Scanned count (Level 2):** a counter scans each product's barcode; individual scans roll
  up into the per-product count automatically.

The schema connects into (grafts onto) an existing Northwind `Products` table and is entirely PC-based
(Access tables; a laptop plus an attached handheld scanner provides warehouse mobility).
A mobile, browser-based interface is out of scope for this template.

## Prerequisites — hooks into the existing schema

This template does not stand alone; it extends an existing Northwind database. The generator
must confirm these exist and wire the new tables to them:

> **Out of the box, Northwind Dev doesn't have two of the fields this template relies on.**
> `Products.SKUBarCode` and `Products.QuantityInPackage` exist only in modified copies of Northwind —
> they were added for the scan workflow this template was shaped from. If your copy doesn't have
> them, **they need to be added to `Products` as part of building this template** — unless you've
> already added them yourself (possibly under different names, which the design should then use
> instead). The `check_compatibility` tool reports exactly which required pieces your database
> already has.

| Existing object | Used as | Notes |
|---|---|---|
| `Products.ProductID` (AutoNumber PK) | Parent of every count line | The primary connection (graft) point |
| `Products.SKUBarCode` (Memo) | Scan-resolution target | A scanned code is matched against this to resolve `ProductID`. **Standards/implementation note:** a Memo cannot be indexed; for production scan performance the standards layer may call for an indexed Text barcode field. The template depends on the field but does not alter `Products`. |
| `Products.QuantityInPackage` (Long) | Package-scan multiplier | When a package barcode is scanned, units added = `QuantityInPackage` (see Business Rules) |
| `Employees.EmployeeID` (AutoNumber PK) | Who conducted the session | `StockTakeSession.ConductedByEmployeeID` FK |
| `SystemSettings` (key/value) | Default allowable shrinkage rate | Seed row `DefaultAllowableShrinkageRate`; follows the host `[percent*1000]` convention used by `TaxRate` (e.g. `"50"` = 0.05 = 5%). Per-product values in `ProductShrinkageAllowance` override it. |

## Entities

Naming follows the **Northwind house style (no `tbl`/`tlkp` prefixes)** to fit cleanly alongside
the host database. GPC field-qualification rules still apply (no bare reserved/ambiguous nouns:
`Status` → `StockTakeStatusID`/`ScanStatusID`, `Notes` → `SessionNotes`). Audit columns are
supplied by the standards layer (see Standards Layer) and are intentionally absent from the
field lists below.

### StockTakeSession — one stocktake event

Grain: one row per stocktake conducted (a date, optionally a location, a status, a person).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `StockTakeSessionID` | AutoNumber | PK | Surrogate key for the stocktake event |
| `StockTakeDate` | Date/Time | Required | When the count was conducted |
| `StockTakeStatusID` | Long | FK → StockTakeStatus, Required | Workflow state (Open → In Progress → Counted → Reconciled → Closed) |
| `ConductedByEmployeeID` | Long | FK → Employees, Nullable | Who ran the session (hook into existing `Employees`) |
| `SessionNotes` | Memo | Nullable | Free-text notes for the event |

Indexes: PK on `StockTakeSessionID`.

### StockTakeCount — one line per product per session

Grain: exactly one row per (`StockTakeSessionID`, `ProductID`). Holds the result of the count for that
product, regardless of which method produced it.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `StockTakeCountID` | AutoNumber | PK | Surrogate key for the count line |
| `StockTakeSessionID` | Long | FK → StockTakeSession, Required | Owning stocktake event |
| `ProductID` | Long | FK → Products, Required | Product being counted (hook into existing `Products`) |
| `StockTakeCountMethodID` | Long | FK → StockTakeCountMethod, Required | How this line was counted (Manual / Scan) |
| `ExpectedQuantity` | Long | Nullable | System on-hand snapshotted when the session opened |
| `CountedQuantity` | Long | Nullable | The counted result. Manual: entered directly. Scan: maintained as `SUM(StockTakeScan.ScanQuantity)` for this line |
| `RemediationStatusID` | Long | FK → RemediationStatus, Required | Outcome of the shrinkage reality check; defaults to None. Set to Flagged when variance exceeds the effective allowable shrinkage rate (logic in the coding section) |

Indexes: PK on `StockTakeCountID`; **unique index on (`StockTakeSessionID`, `ProductID`)** — enforces one count
line per product per session; non-unique index on `ProductID` (FK); non-unique index on
`RemediationStatusID` (FK).

Derived (not stored): `VarianceQuantity = CountedQuantity − ExpectedQuantity` — computed in
queries/reports to avoid staleness.

### StockTakeScan — individual scans rolling up into a count line

Grain: one row per physical scan. Present only for scanned counts (Level 2).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `StockTakeScanID` | AutoNumber | PK | Surrogate key (replaces the Dataverse GUID key entirely) |
| `StockTakeCountID` | Long | FK → StockTakeCount, Required | The count line this scan contributes to — the structural rollup link |
| `ScanCode` | Text(255) | Required | Raw barcode string as scanned; resolved against `Products.SKUBarCode` |
| `ScanQuantity` | Long | Required | Units this scan adds to the count line (see package rule) |
| `ScanStatusID` | Long | FK → ScanStatus, Required | Valid / Unmatched / Duplicate |
| `ScannedOn` | Date/Time | Required | Timestamp of the scan |

Indexes: PK on `StockTakeScanID`; non-unique index on `StockTakeCountID` (FK).

### ProductShrinkageAllowance — per-product shrinkage tolerance (admin-managed)

Grain: at most one row per product, holding a non-default allowable shrinkage tolerance.
A 1:1 extension of `Products` — it adds stocktake-specific configuration without altering the
host table. Products without a row inherit the `SystemSettings` default.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ProductID` | Long | PK + FK → Products | Shared primary key (1:1 with `Products`) |
| `AllowableShrinkageRate` | Single | Required | Allowable shortfall as a **fraction** (`0.0500` = 5%) |

Indexes: PK on `ProductID` (also the FK to `Products`).

### Lookup tables (Northwind-style, mirroring `OrderStatus`/`TaxStatus`)

| Table | Rows (seed) |
|---|---|
| `StockTakeStatus` | Open; In Progress; Counted; Reconciled; Closed |
| `StockTakeCountMethod` | Manual; Scan |
| `ScanStatus` | Valid; Unmatched; Duplicate |
| `RemediationStatus` | None; Flagged; Under Review; Resolved |

## Relationships

New (within this template):
- `StockTakeSession (1) → (∞) StockTakeCount` on `StockTakeSessionID` — cascade delete
- `StockTakeCount (1) → (∞) StockTakeScan` on `StockTakeCountID` — cascade delete
- `StockTakeStatus (1) → (∞) StockTakeSession` on `StockTakeStatusID`
- `StockTakeCountMethod (1) → (∞) StockTakeCount` on `StockTakeCountMethodID`
- `ScanStatus (1) → (∞) StockTakeScan` on `ScanStatusID`
- `RemediationStatus (1) → (∞) StockTakeCount` on `RemediationStatusID`

Hooks into existing Northwind schema:
- `Products (1) → (∞) StockTakeCount` on `ProductID` — **no cascade** (never delete count
  history when a product changes)
- `Products (1) → (0..1) ProductShrinkageAllowance` on `ProductID` — cascade delete (the
  tolerance is pure config for that product, meaningless without it)
- `Employees (1) → (∞) StockTakeSession` on `ConductedByEmployeeID` — no cascade

## Business Rules

1. **One count line per product per session** — enforced by the unique index on
   (`StockTakeSessionID`, `ProductID`). A product is counted at most once per event, by one method.
2. **Scan resolution** — on each scan, match `ScanCode` against `Products.SKUBarCode` to find
   `ProductID`. Find or create the `StockTakeCount` line for (`StockTakeSessionID`, `ProductID`), then
   insert the `StockTakeScan` against that `StockTakeCountID`. Unmatched codes are stored with
   `ScanStatusID = Unmatched` and no `StockTakeCountID` resolution (held for review).
3. **Rollup** — for scanned lines, `StockTakeCount.CountedQuantity = SUM(ScanQuantity)` across
   the line's scans, maintained as scans are added or removed.
   **Why stored, not derived:** `CountedQuantity` is deliberately a stored value, not a figure
   computed on demand. A reconciled stocktake count is a durable audit fact — business decisions
   are made from it, and accountability is lost if it later changes. Storing it fixes the official
   number even if scan detail is edited or archived after reconciliation, and gives every report one
   uniform read path across both count methods (manual entry writes it directly; scans roll up into
   it). The trade-off — a stored aggregate can drift from its source if the maintaining process fails
   — is accepted deliberately and is the maintainer's responsibility (a data macro or the count-rollup
   procedure keeps it current). Where durability isn't needed and scan detail is permanent, computing
   it on demand is the cleaner choice. (Declared in `house_assumptions`.)
4. **Package scanning** — if a scanned code denotes a package rather than a unit, `ScanQuantity`
   for that scan = `Products.QuantityInPackage`. Unit scans contribute 1 (or the entered count).
5. **Expected quantity** — `ExpectedQuantity` is snapshotted from the system's computed on-hand
   at the moment the session opens, so variance reflects the count against a fixed baseline.
6. **Variance** — computed as `CountedQuantity − ExpectedQuantity` in queries/reports; not stored.
7. **Effective shrinkage rate** — for a product, use `ProductShrinkageAllowance.AllowableShrinkageRate`
   (a fraction) if a row exists; otherwise fall back to `SystemSettings.DefaultAllowableShrinkageRate`
   ÷ 1000 (the `[percent*1000]` host convention). Both normalize to a fraction before comparison.
8. **Shrinkage reality check** *(logic deferred to the coding section; schema support only)* — when a
   count line shows a shortfall, compute shortfall fraction = `(ExpectedQuantity − CountedQuantity) /
   ExpectedQuantity`. If it exceeds the effective allowable rate, set `RemediationStatusID = Flagged`
   for review; otherwise leave it `None`. Shrinkage = damage, misplacement, or theft.

## Standards Layer (supplied externally, not in this template body)

The following are deliberately **omitted** here and contributed by the developer's standards
layer, so the same template produces house-conforming output for any practice:

- **Audit columns** — `AddedBy`, `AddedOn`, `ModifiedBy`, `ModifiedOn` on every new table,
  maintained by the Northwind data-macro audit pattern (NorthwindFeatures #30).
- **Naming conventions** — table/field prefix policy. *This template honors Northwind's
  no-prefix house style; a GPC-standard practice would instead apply `tbl`/`tlkp`.* **Northwind
  is itself the worked illustration of why standards must be user-customizable: the team developing the
  Northwind Templates agreed to a generic naming convention — a different practice would
  regenerate these same entities under its own conventions without touching this template.**
- **Error handling** — the generalized `errHandler` / `GlblErrMsg` pattern for any VBA produced
  alongside this schema.

## Extra Options (engagement-specific — stub)

*Empty in the base template. Filled per client engagement; the filled copy is saved to the
developer's own library, not committed here.*

- **Multi-location / warehouse** — add `StockTakeLocationID` (FK to a `StockTakeLocation`
  lookup) to `StockTakeSession` for businesses counting across multiple sites.
- **Cloud + mobile** — migrate these tables to a cloud database and add a mobile PowerApps
  scanner interface (the original delivery model; out of scope for the PC-based template).
- **Package-count detail** — add an explicit `PackageCount` to `StockTakeScan` if package and
  unit scans must be reported separately rather than normalized into `ScanQuantity`.
- **Category-level shrinkage tolerance** — set allowable shrinkage on `ProductCategories` (e.g.
  produce vs. hardware) as an alternative or fallback tier between the per-product override and
  the system default. Often more realistic than per-product for large catalogs.

## Parked / future considerations (not in this design)

- **PowerApps logic derivation** — the original PowerApps app exists and may contain scan-
  processing logic not yet reflected in any VBA; revisit to derive it if/when the mobile path
  is pursued.
- **Admin count override** — allowing an administrator to override a counted quantity for
  administrative reasons.
- **Barcode label printing** — a forms/reports concern (printing product labels in a barcode
  font); belongs to the forms release, not this schema template.
