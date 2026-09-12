---
template: northwind-stocktake-schema
title: Northwind Scanned Stocktake — Table Schema
domain: northwind
type: table-schema
version: 0.5.1
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
  - ProductVarianceAllowance
  - RemediationStatus
seeds:
  - SystemSettings.DefaultAllowableShortageRate
  - SystemSettings.DefaultAllowableOverageRate
  - SystemSettings.DuplicateScanWindowSeconds
house_assumptions:
  - "StockTakeCount.CountedQuantity stored, not derived — a reconciled count is a durable audit fact business decisions rely on; it must not change if scan detail is later edited or archived. Stored despite being derivable; the alternative is to compute it on demand."
  - "RemediationStatus records that a count line needs review, not which direction (shortage or overage) tripped it. The signed VarianceQuantity (Business Rule 6) already answers that at review time, by its sign — a practice wanting the direction stored on the count line itself, rather than read from a query, changes this."
  - "ProductVarianceAllowance renames the table formerly called ProductShrinkageAllowance, now that it holds an overage tolerance as well as a shortage one — 'shrinkage' names loss specifically and would misname the overage column. A practice already using the old name changes this back."
---

# Northwind Scanned Stocktake — Table Schema

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

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

**Where these tables live.** These templates are designed for a **split database** — the normal
shape for Access applications, especially those in multi-user environments: one file holds the
tables (the **back end**, usually on a shared network drive — never on OneDrive, Dropbox, or any
other file-syncing cloud folder, which corrupts a shared Access back end), and each person runs
their own copy of a second file holding the forms, reports, and code (the **front end**), whose
tables are *links* pointing at the back end. A single-file database — one .accdb holding
everything — is an acceptable choice for one user, and everything here works there too: create
everything in that one file and ignore the distinction.

Unlike the greenfield templates in this library, this one grafts onto a database that already
exists, so the placement is decided for you: **the new tables go in whichever file already holds
`Products`** — the back end, if the host is split. This is not a preference. Access cannot enforce
a relationship between tables in two different files, and `StockTakeCount` takes an enforced
relationship to `Products` (see Relationships). Split them across files and the relationship simply
cannot be created.

This template does not stand alone; it extends an existing Northwind database. The generator
must confirm these exist and wire the new tables to them:

> **Out of the box, Northwind Dev doesn't have two of the fields this template relies on.**
> `Products.SKUBarCode` and `Products.QuantityInPackage` exist only in modified copies of Northwind —
> they were added for the scan workflow this template was shaped from. If your copy doesn't have
> them, **they need to be added to `Products` as part of building this template** — unless you've
> already added them yourself (possibly under different names, which the design should then use
> instead). The `check_compatibility` tool reports exactly which required pieces your database
> already has.
>
> **If the host is split, adding those fields is a back-end change**, made once in the file that
> holds `Products`. A front end does not pick up a new field on its own — its link still describes
> the table as it was. Refresh the links in every front end (Access's Linked Table Manager) or the
> new fields simply won't appear there, and code referring to them fails with "item not found in
> this collection."

| Existing object | Used as | Notes |
|---|---|---|
| `Products.ProductID` (AutoNumber PK) | Parent of every count line | The primary connection (graft) point |
| `Products.SKUBarCode` (Memo) | Scan-resolution target | A scanned code is matched against this to resolve `ProductID`. **Standards/implementation note:** a Memo cannot be indexed; for production scan performance the standards layer may call for an indexed Text barcode field. The template depends on the field but does not alter `Products`. |
| `Products.QuantityInPackage` (Long) | Package-scan multiplier | When a package barcode is scanned, units added = `QuantityInPackage` (see Business Rules) |
| `Employees.EmployeeID` (AutoNumber PK) | Who conducted the session | `StockTakeSession.ConductedByEmployeeID` FK |
| `SystemSettings` (key/value) | Default allowable variance rates | Seed rows `DefaultAllowableShortageRate` and `DefaultAllowableOverageRate`; both follow the host `[percent*1000]` convention used by `TaxRate` (e.g. `"50"` = 0.05 = 5%). Per-product values in `ProductVarianceAllowance` override either one independently. |
| `SystemSettings` (key/value) | Duplicate-scan detection window | Seed row `DuplicateScanWindowSeconds`; a plain integer count of seconds (e.g. `"120"`), not the `[percent*1000]` convention above — this key holds a duration, not a rate. See Business Rule 2. |

## Entities

Naming follows the **Northwind house style (no `tbl`/`tlkp` prefixes)** to fit cleanly alongside
the host database. OTS field-qualification rules still apply (no bare reserved/ambiguous nouns:
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
| `RemediationStatusID` | Long | FK → RemediationStatus, Required | Outcome of the variance reality check; defaults to None. Set to Flagged when the shortfall or the overage exceeds its effective allowable rate (logic in the coding section). Does not record which direction tripped it — see Business Rule 8 |

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
| `StockTakeCountID` | Long | FK → StockTakeCount, Nullable | The count line this scan contributes to — the structural rollup link. Null for an Unmatched scan (Business Rule 2), which never resolves to a count line. |
| `ScanCode` | Text(255) | Required | Raw barcode string as scanned; resolved against `Products.SKUBarCode` |
| `ScanQuantity` | Long | Required | Units this scan adds to the count line (see package rule) |
| `ScanStatusID` | Long | FK → ScanStatus, Required | Valid / Unmatched / Duplicate |
| `ScannedOn` | Date/Time | Required | Timestamp of the scan |

Indexes: PK on `StockTakeScanID`; non-unique index on `StockTakeCountID` (FK).

### ProductVarianceAllowance — per-product variance tolerance (admin-managed)

Grain: at most one row per product, holding non-default allowable variance tolerances — a
shortage tolerance, an overage tolerance, or both. A 1:1 extension of `Products` — it adds
stocktake-specific configuration without altering the host table. A product with no row, or a
row with one of the two rates left blank, inherits the matching `SystemSettings` default for
that rate; the two directions fall back independently of each other.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ProductID` | Long | PK + FK → Products | Shared primary key (1:1 with `Products`) |
| `AllowableShortageRate` | Single | Nullable | Allowable shortfall as a **fraction** (`0.0500` = 5%). Blank falls back to `SystemSettings.DefaultAllowableShortageRate` |
| `AllowableOverageRate` | Single | Nullable | Allowable overage as a **fraction** (`0.0500` = 5%). Blank falls back to `SystemSettings.DefaultAllowableOverageRate` |

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
- `Products (1) → (0..1) ProductVarianceAllowance` on `ProductID` — cascade delete (the
  tolerance is pure config for that product, meaningless without it)
- `Employees (1) → (∞) StockTakeSession` on `ConductedByEmployeeID` — no cascade

## Business Rules

1. **One count line per product per session** — enforced by the unique index on
   (`StockTakeSessionID`, `ProductID`). A product is counted at most once per event, by one method.
2. **Scan resolution** — on each scan, match `ScanCode` against `Products.SKUBarCode` to find
   `ProductID`. Find or create the `StockTakeCount` line for (`StockTakeSessionID`, `ProductID`), then
   check the scans already on that `StockTakeCountID` for one with the same `ScanQuantity` and a
   `ScannedOn` within `SystemSettings.DuplicateScanWindowSeconds` of the new scan. If one is found,
   the new scan is inserted with `ScanStatusID = Duplicate`; otherwise `Valid`. A duplicate is still
   inserted and still resolves to the count line — it is excluded from the rollup (Business Rule 3),
   not discarded. Unmatched codes are stored with `ScanStatusID = Unmatched` and no `StockTakeCountID`
   resolution (held for review) — the duplicate check does not apply to them, since they never reach a
   count line.
3. **Rollup** — for scanned lines, `StockTakeCount.CountedQuantity = SUM(ScanQuantity)` across
   the line's scans **where `ScanStatusID = Valid`**, maintained as scans are added or removed. A scan
   marked `Duplicate` (Business Rule 2) is never counted twice.
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
7. **Effective variance rates** — for a product, use `ProductVarianceAllowance.AllowableShortageRate`
   (a fraction) if a row exists and that column holds a value; otherwise fall back to
   `SystemSettings.DefaultAllowableShortageRate` ÷ 1000 (the `[percent*1000]` host convention). The
   same resolution applies independently to `AllowableOverageRate` / `DefaultAllowableOverageRate` —
   a product can override one direction's tolerance without overriding the other. Both normalize to a
   fraction before comparison.
8. **Variance reality check** *(logic deferred to the coding section; schema support only)* — using
   the signed `VarianceQuantity` from Business Rule 6, compare **quantities, never fractions:**
   - Where `VarianceQuantity` is negative (a shortfall): flag when
     `ExpectedQuantity − CountedQuantity  >  effective shortage rate × ExpectedQuantity`.
   - Where `VarianceQuantity` is positive (an overage): flag when
     `CountedQuantity − ExpectedQuantity  >  effective overage rate × ExpectedQuantity`.

   Either comparison holding sets `RemediationStatusID = Flagged` for review; otherwise it stays
   `None`. **`RemediationStatus` does not record which direction tripped it** — the sign of
   `VarianceQuantity`, read at review time, already answers that, so nothing is stored redundantly
   (declared in `house_assumptions`). Shortfall = damage, misplacement, or theft; overage = a
   receiving, return, or count error that inflated the figure.

   **Why quantities, not the shortfall/overage fraction used in earlier drafts of this rule.** The
   natural way to state a tolerance is as a percentage — "flag a shortfall over 5% of what was
   expected" — which reads as *fraction exceeds rate*: `(ExpectedQuantity − CountedQuantity) /
   ExpectedQuantity > rate`. That is exactly the earlier form of this rule, and it divides by
   `ExpectedQuantity` — undefined the moment a count line's expected quantity is zero. **Multiplying
   both sides of that comparison by `ExpectedQuantity` produces the form given above: the same
   comparison, for every count line where `ExpectedQuantity > 0`, with no division anywhere.** (The
   direction of the inequality is unaffected — `ExpectedQuantity` is positive whenever it appears in
   a denominator here, so multiplying by it never flips which side is larger.) A comparison that
   was never dividing has nothing to break when `ExpectedQuantity` reaches zero: the right-hand side
   of each line above becomes `rate × 0 = 0`, and the comparison still runs.
   **What that means at zero is the correct answer on its own terms, not a worked-around edge
   case.** `CountedQuantity` cannot be negative, so where `ExpectedQuantity = 0`,
   `CountedQuantity ≥ ExpectedQuantity` always — the shortfall line can never hold, only the overage
   line can, and it reduces to `CountedQuantity − 0 > 0`: **flag any nonzero count.** That is the
   right outcome stated in its own right — a count line where the system expected nothing and a
   counter found something is not a rounding error near a percentage threshold, it is the most
   notable thing this check can find, and this form flags it without inventing a substitute value
   for `ExpectedQuantity` to divide by.

## Standards Layer (supplied externally, not in this template body)

The following are deliberately **omitted** here and contributed by the developer's standards
layer, so the same template produces house-conforming output for any practice:

- **Audit columns** — `AddedBy`, `AddedOn`, `ModifiedBy`, `ModifiedOn` on every new table,
  maintained by the Northwind data-macro audit pattern (NorthwindFeatures #30).
- **Naming conventions** — table/field prefix policy. *This template honors Northwind's
  no-prefix house style; the OTS default would instead apply `tbl`/`tlkp`.* **Northwind
  is itself the worked illustration of why standards must be user-customizable: the team developing the
  Northwind Templates agreed to a generic naming convention — a different practice would
  build these same entities under its own conventions without touching this template.**
- **Error handling** — the `errHandler` / `Cleanup` pattern for any VBA produced
  alongside this schema.

## Extra Options (engagement-specific — stub)

*Named optional extensions, none of them filled in for an engagement; the filled copy is saved to the
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
