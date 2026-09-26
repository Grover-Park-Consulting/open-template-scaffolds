---
template: stocktake-schema
title: Scanned Stocktake — Table Schema
domain: stocktakescan
type: table-schema
version: 0.8.1
status: review
extends: Northwind (Access Developer Edition)
requires_tables:
  - Products
  - Employees
  - SystemSettings
requires_fields:
  - Products.ProductID
new_fields:
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
  - "RemediationStatus — records that a count line needs review, not which direction (shortage or overage) tripped it. The signed VarianceQuantity (Business Rule 6) already answers that at review time, by its sign — a practice wanting the direction stored on the count line itself, rather than read from a query, changes this."
  - "ProductVarianceAllowance — renames the table formerly called ProductShrinkageAllowance, now that it holds an overage tolerance as well as a shortage one — 'shrinkage' names loss specifically and would misname the overage column. A practice already using the old name changes this back."
---

# Scanned Stocktake — Table Schema

**Status last determined:** 2026-09-26.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Add a **real, scan-driven stocktake** to a Northwind-derived database. Base Northwind's
`StockTake` records only a pseudo count for demo purposes — aggregate quantities hand-typed after the
fact, with no audit trail of how they were derived. This template replaces that with a
three-level structure that supports **two count methods over one schema**:

**Confirm with the developer whether the host's "on-hand" computation reads `StockTake`.**
Business Rule 5 below snapshots "the system's computed on-hand" into every baseline line's
`ExpectedQuantity`. If that computation is itself derived from `StockTake` — rather than from
`Order Details`, a running-total mechanism, or something else — building this template stops
anyone from entering new `StockTake` rows, and the on-hand figure it depends on goes stale from
that point on. Ask before building, rather than assuming either answer: what `StockTake` is left
for once this template is in place (left alone, migrated, or retired) is the developer's call.

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

> **This template adds two fields to `Products`, and that is the only change it makes to a table
> you already have.** A scanned stocktake needs somewhere to hold the code that is scanned and, for
> products that ship in packages, how many units a package holds. Northwind Dev has neither out of
> the box, so **building this template creates them:**
>
> - `SKUBarCode` — Short Text, 50 characters, optional, with a **unique** index on it
> - `QuantityInPackage` — Long, optional
>
> **The index is the point of the Short Text field.** Every scan resolves a code by looking it up in
> this field, which makes it the most frequently read field in the whole design, and a field Access
> cannot index makes every one of those lookups read the entire product list. A barcode is short —
> 50 characters is generous for one — so keeping the field small enough to index gives nothing up.
>
> **The index must be unique, not merely present.** A barcode identifies one product; without a
> unique index, two products could carry the same code and scan resolution would silently pick
> whichever one its query happens to return first, with nothing refusing the second product's value
> or flagging the collision. Access permits any number of Nulls in a unique index, so this does not
> conflict with the field staying optional.
>
> **If you have already added your own barcode or package-quantity field, say so** — possibly under
> different names, which the design then uses instead of creating new ones. The one thing the design
> needs of an existing barcode field is that it can carry an index.
>
> **If the host is split, adding those fields is a back-end change**, made once in the file that
> holds `Products`. A front end does not pick up a new field on its own — its link still describes
> the table as it was. Refresh the links in every front end (Access's Linked Table Manager) or the
> new fields simply won't appear there, and code referring to them fails with "item not found in
> this collection."

| Existing object | Used as | Notes |
|---|---|---|
| `Products.ProductID` (AutoNumber PK) | Parent of every count line | The primary connection (graft) point |
| `Products.SKUBarCode` (Short Text, 50, unique index) | Scan-resolution target | A scanned code is matched against this to resolve `ProductID`. **Created by this template where the host doesn't already have it** — see the note above. The index is required rather than an optimization: scan resolution is the most frequent read in the design, and uniqueness is required so a code never resolves to more than one product |
| `Products.QuantityInPackage` (Long) | Package-scan multiplier | When a package barcode is scanned, units added = `QuantityInPackage` (see Business Rules). **Created by this template where the host doesn't already have it** |
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
| `StockTakeCountMethodID` | Long | FK → StockTakeCountMethod, Required | How this line was counted (Manual / Scan) — `Pending` at baseline creation, before either has happened (Business Rule 5) |
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
| `StockTakeCountMethod` | Pending; Manual; Scan |
| `ScanStatus` | Valid; Unmatched; Duplicate |
| `RemediationStatus` | None; Flagged; Under Review; Resolved |

## Relationships

New (within this template):
- `StockTakeSession (1) → (∞) StockTakeCount` on `StockTakeSessionID` — cascade delete
- `StockTakeCount (1) → (∞) StockTakeScan` on `StockTakeCountID` — cascade delete
- `StockTakeStatus (1) → (∞) StockTakeSession` on `StockTakeStatusID` — no cascade
- `StockTakeCountMethod (1) → (∞) StockTakeCount` on `StockTakeCountMethodID` — no cascade
- `ScanStatus (1) → (∞) StockTakeScan` on `ScanStatusID` — no cascade
- `RemediationStatus (1) → (∞) StockTakeCount` on `RemediationStatusID` — no cascade

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

   **The window measures against the nearest existing scan, including one already marked
   `Duplicate`, not against the original `Valid` scan that started the chain.** A steady drip of
   scans on the same product and quantity, each arriving just inside the window of the one before
   it, therefore stays one count for as long as the drip continues — the window never re-anchors to
   when counting on that item actually began. This is accepted as-is: it protects against the
   scenario the rule exists for (an accidental double-scan seconds apart) and a real chain of scans
   that slow is unusual enough in practice not to warrant comparing against every prior scan instead
   of the nearest one. A practice that wants the window anchored to the first `Valid` scan instead
   changes this rule.
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
5. **Expected quantity, and the baseline it belongs to** — opening a session **creates a count
   line for every product the session covers**, each one carrying `ExpectedQuantity` snapshotted from
   the system's computed on-hand at that moment. Variance is then measured against a baseline that is
   fixed for the life of the session and does not move as stock transactions carry on around it.
   **`StockTakeCountMethodID` is `Pending` on a baseline line** — `StockTakeCountMethod` is
   Required, and at the moment a baseline line is created neither a manual entry nor a scan has
   happened yet for that product, so neither existing value is true. The first thing that actually
   counts the product — a manual entry or the first scan resolving to it — updates the line to
   `Manual` or `Scan` at that point.

   **Creating those lines when the session opens, rather than as scans arrive, is what this rule is
   for.** A stocktake exists to find the difference between what the system believes is on hand and
   what is actually on the floor, and the largest difference it can find is a product the system
   expected to have and the counters found none of. A build that creates a count line only when
   something is scanned cannot report that product at all: no scan, no line, so it appears in no
   variance report and the loss is invisible. With the lines created up front, that product ends the
   session counted zero against its expected quantity and is flagged by Business Rule 8 like any
   other shortfall.

   **That flag has to be set when the line is created, not left for a scan that may never come.**
   Business Rule 8 decides whether a count line is flagged by comparing `CountedQuantity` against
   `ExpectedQuantity` — but nothing runs that comparison on its own. A build that creates the
   baseline lines and moves on, leaving the flag to be set only when a scan for that product arrives,
   satisfies the letter of "a count line for every product" while missing the point of it: the one
   line this rule exists to protect is exactly the one that never gets a scan, so it would sit
   unflagged for the entire session. Evaluate Business Rule 8 against each line at the moment it is
   created, before the baseline step is considered finished.

   **Which products a session covers** is the engagement's to decide; the default is every product
   not marked discontinued. A product added to the catalog *after* the session opened has no line and
   gets one on first scan, carrying the on-hand of that moment — the same path that handles two
   counters reaching a new product at once (Business Rule 1).
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

   **This automatic evaluation only ever sets `None` or `Flagged`, and only where the line's
   current status is already one of those two.** `Under Review` and `Resolved` are a reviewer's own
   decision, made by a human outside this rule; a later scan against the same product (a recount,
   or a late duplicate resolving) re-runs this evaluation, and without this guard would silently
   revert that decision back to `Flagged` or `None` the moment the comparison above still holds or
   stops holding. Leave `Under Review` and `Resolved` alone — the reviewer, not a recount, is what
   moves a line off them.

   **Why quantities, not the shortfall/overage fraction used in earlier drafts of this rule.** The
   natural way to state a tolerance is as a percentage — "flag a shortfall over 5% of what was
   expected" — which reads as *fraction exceeds rate*: `(ExpectedQuantity − CountedQuantity) /
   ExpectedQuantity > rate`. That is exactly the earlier form of this rule, and it divides by
   `ExpectedQuantity` — undefined the moment a count line's expected quantity is zero. **Multiplying
   both sides of that comparison by `ExpectedQuantity` produces the form given above: the same
   comparison, for every count line where `ExpectedQuantity > 0`, with no division anywhere.**
   (Multiplying by a positive number leaves which side is larger unchanged, and a positive
   `ExpectedQuantity` is the only case the fraction form could be evaluated on at all.) A comparison that
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

   **A negative expected quantity is flagged every time, and that is deliberate.** Where the host's
   on-hand calculation can return a negative number — Northwind's can, when recorded sales run ahead
   of recorded receipts — a count line can open with a negative `ExpectedQuantity`. The overage
   comparison above then holds for any count at all, including a count of zero, so the line is
   flagged and somebody looks at it. That is the right answer on its own terms: the system's own
   figure was impossible before anybody counted anything. **Do not add a branch for this case**, and
   in particular do not guard the comparison with a test on `ExpectedQuantity` — that puts back the
   division this form exists to remove.

## Validating the build

Per `_template-schema.md` §4.2 — the structural baseline instantiated against this schema's eight
new tables, plus the two fields it grafts onto the existing `Products` table (check 7).

| # | Check |
|---|---|
| 1 | An ordinary insert succeeds on each of the eight new tables, supplying every `Required` field. |
| 2 | Every `Required` field on every new table refuses a missing value; no `AllowZeroLength` fields are declared in this schema, so that half of the check doesn't apply here. |
| 3 | `StockTakeCount`'s unique index on (`StockTakeSessionID`, `ProductID`) refuses the duplicate it names — a second count line for the same product in the same session. |
| 4 | `StockTakeSession → StockTakeCount` and `StockTakeCount → StockTakeScan` cascade-delete as declared. `Products → StockTakeCount` does **not** — deleting a product with count history is refused, never silently dropping the history — and `Products → ProductVarianceAllowance` **does** cascade, removing the tolerance row when its product goes. Confirm both the cascading and the non-cascading case directly; they sit on opposite sides of the same host table. |
| 5 | The three `SystemSettings` seed rows (`DefaultAllowableShortageRate`, `DefaultAllowableOverageRate`, `DuplicateScanWindowSeconds`) are present with their specified values, in the host's existing `[percent*1000]`/plain-integer conventions as documented. |
| 6 | The audit columns the active standards layer supplies (see Standards Layer below) stamp correctly on every new table: who created the row and when, on insert; who last changed it and when, on each update; and the created pair left frozen on a later update. |
| 7 | **Graft-specific, not part of the generic baseline:** `Products.SKUBarCode` and `Products.QuantityInPackage` exist (created fresh, or confirmed against fields the developer already had), `SKUBarCode` carries its index and the index is **unique** — inserting or updating a second product to the same barcode as an existing one is refused — and, where the host is split, every front end's linked-table definition of `Products` was refreshed and shows both new fields. A front end that wasn't relinked is the specific, silent failure this template's own Prerequisites section warns about: code referencing either field fails with "item not found in this collection." |
| 8 | A `StockTakeCount` insert citing a `StockTakeSessionID` or `ProductID` that doesn't exist is refused; likewise a `StockTakeScan` insert citing a `StockTakeCountID` that doesn't exist. |
| 9 | An insert with `ScanCode` or another `Text(n)` field longer than its declared width is refused, not silently truncated. |
| 10 | `Description` is present on every field of every new table, matching this template's own Purpose & rules text — and on `Products.SKUBarCode`/`QuantityInPackage`, where the two grafted fields carry their own descriptions. |
| 11 | Running the table-build `Sub` a second time either re-runs cleanly or fails naming what already exists — never a bare "duplicate object" error, and never a second attempt to add `SKUBarCode`/`QuantityInPackage` to `Products` if they're already there. |

Report against this list exactly as `_template-schema.md` §12.2 states for every checklist in the
library: one entry per check, a literal `Result: PASSED` or `Result: NOT PASSED`.

## Standards Layer (supplied externally, not in this template body)

The following are deliberately **omitted** here and contributed by the developer's standards
layer, so the same template produces house-conforming output for any practice:

- **Audit columns** — the house default set (`CreatedDate`, `CreatedBy`, `ModifiedDate`,
  `ModifiedBy`) on every new table, maintained by the active standards layer's mechanism
  (`standards/audit-columns.md`). A host this template is grafted onto may already stamp
  who-and-when its own way — a Northwind-derived database, for one, commonly uses
  `AddedBy`/`AddedOn`/`ModifiedBy`/`ModifiedOn` via data macros. Check the target database's
  existing convention before naming any new audit column and match it; use the house names only
  where no existing convention is present.
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
