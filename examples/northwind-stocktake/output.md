# Worked Example — The Result (Northwind Stocktake)

What the AI produced from [`prompt.md`](prompt.md). **Standards applied:** Northwind-consistent —
no table-name prefixes, and the four `AddedBy / AddedOn / ModifiedBy / ModifiedOn` audit columns on
every new table, maintained by data macros (as in the rest of Northwind). **Declared house
assumptions:** none — this template declares none, so there was nothing to confirm. **Your
specifics, folded in:** shrinkage shown as a percentage and flagged above 5%, with a per-product
override for high-value items (the template's existing shrinkage tolerance, set to your 5% default).

This is presented for your review — approve it, or tell me what to change.

## The tables, at a glance

```mermaid
erDiagram
    PRODUCTS                ||--o{ STOCKTAKECOUNT            : "counted on each line"
    PRODUCTS                ||--o| PRODUCTSHRINKAGEALLOWANCE : "optional tolerance"
    EMPLOYEES               ||--o{ STOCKTAKESESSION          : "conducted by"
    STOCKTAKESESSION        ||--o{ STOCKTAKECOUNT            : "has count lines"
    STOCKTAKECOUNT          ||--o{ STOCKTAKESCAN             : "rolls up from scans"
    STOCKTAKESTATUS         ||--o{ STOCKTAKESESSION          : "status of"
    STOCKTAKECOUNTMETHOD    ||--o{ STOCKTAKECOUNT            : "counted by method"
    SCANSTATUS              ||--o{ STOCKTAKESCAN             : "status of"
    REMEDIATIONSTATUS       ||--o{ STOCKTAKECOUNT            : "review outcome"

    STOCKTAKESESSION {
        AutoNumber StockTakeSessionID PK
        DateTime StockTakeDate
        Long StockTakeStatusID FK
        Long ConductedByEmployeeID FK
    }
    STOCKTAKECOUNT {
        AutoNumber StockTakeCountID PK
        Long StockTakeSessionID FK
        Long ProductID FK
        Long StockTakeCountMethodID FK
        Long ExpectedQuantity
        Long CountedQuantity
        Long RemediationStatusID FK
    }
    STOCKTAKESCAN {
        AutoNumber StockTakeScanID PK
        Long StockTakeCountID FK
        Text ScanCode
        Long ScanQuantity
        Long ScanStatusID FK
        DateTime ScannedOn
    }
    PRODUCTSHRINKAGEALLOWANCE {
        Long ProductID PK-FK
        Single AllowableShrinkageRate
    }
```

*(Existing Northwind tables — `Products`, `Employees`, `SystemSettings` — are shown only where the
new tables connect to them; they aren't changed.)*

## Every field, table by table

### StockTakeSession — one stocktake event

One row per stocktake conducted (a date, a status, who ran it).

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `StockTakeSessionID` | AutoNumber | PK | Surrogate key for the stocktake event |
| `StockTakeDate` | Date/Time | Required | When the count was conducted |
| `StockTakeStatusID` | Long | FK → StockTakeStatus, Required | Workflow state (Open → In Progress → Counted → Reconciled → Closed) |
| `ConductedByEmployeeID` | Long | FK → Employees, Nullable | Who ran the session |
| `SessionNotes` | Memo | Nullable | Free-text notes for the event |
| `AddedBy` | Text | Required | *Standards layer — stamped on create (data macro)* |
| `AddedOn` | Date/Time | Required | *Standards layer — stamped on create (data macro)* |
| `ModifiedBy` | Text | Nullable | *Standards layer — stamped on each change* |
| `ModifiedOn` | Date/Time | Nullable | *Standards layer — stamped on each change* |

Indexes: PK on `StockTakeSessionID`.

### StockTakeCount — one line per product per session

Exactly one row per (`StockTakeSessionID`, `ProductID`) — the count result, whichever method
produced it.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `StockTakeCountID` | AutoNumber | PK | Surrogate key for the count line |
| `StockTakeSessionID` | Long | FK → StockTakeSession, Required | Owning stocktake event |
| `ProductID` | Long | FK → Products, Required | Product being counted |
| `StockTakeCountMethodID` | Long | FK → StockTakeCountMethod, Required | How this line was counted (Manual / Scan) |
| `ExpectedQuantity` | Long | Nullable | System on-hand snapshotted when the session opened |
| `CountedQuantity` | Long | Nullable | The counted result. Manual: entered directly. Scan: `SUM(StockTakeScan.ScanQuantity)` |
| `RemediationStatusID` | Long | FK → RemediationStatus, Required | Review outcome; defaults to None, set to Flagged when shrinkage exceeds the allowed rate |

*Plus the four standard audit columns (`AddedBy` / `AddedOn` / `ModifiedBy` / `ModifiedOn`).*

Indexes: PK on `StockTakeCountID`; **unique on (`StockTakeSessionID`, `ProductID`)**; index on
`ProductID`; index on `RemediationStatusID`.

Derived (not stored): `VarianceQuantity = CountedQuantity − ExpectedQuantity`; shrinkage % =
`(ExpectedQuantity − CountedQuantity) / ExpectedQuantity`.

### StockTakeScan — individual scans that roll up into a count line

One row per physical scan; present only for scanned counts.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `StockTakeScanID` | AutoNumber | PK | Surrogate key |
| `StockTakeCountID` | Long | FK → StockTakeCount, Required | The count line this scan adds to |
| `ScanCode` | Text(255) | Required | Raw barcode string, resolved against `Products.SKUBarCode` |
| `ScanQuantity` | Long | Required | Units this scan adds (packages use `Products.QuantityInPackage`) |
| `ScanStatusID` | Long | FK → ScanStatus, Required | Valid / Unmatched / Duplicate |
| `ScannedOn` | Date/Time | Required | Timestamp of the scan |

*Plus the four standard audit columns.* Indexes: PK on `StockTakeScanID`; index on `StockTakeCountID`.

### ProductShrinkageAllowance — per-product tolerance (your high-value override)

At most one row per product; a tighter tolerance that overrides the 5% default. Products without a
row inherit the default.

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `ProductID` | Long | PK + FK → Products | Shared key (1:1 with `Products`) |
| `AllowableShrinkageRate` | Single | Required | Allowed shortfall as a fraction (`0.0500` = 5%) |

*Plus the four standard audit columns.* Indexes: PK on `ProductID`.

### Lookup tables

| Table | Seed rows |
|---|---|
| `StockTakeStatus` | Open; In Progress; Counted; Reconciled; Closed |
| `StockTakeCountMethod` | Manual; Scan |
| `ScanStatus` | Valid; Unmatched; Duplicate |
| `RemediationStatus` | None; Flagged; Under Review; Resolved |

*Each lookup also carries the four standard audit columns.*

## How the tables connect

- `StockTakeSession → StockTakeCount` (cascade delete) · `StockTakeCount → StockTakeScan` (cascade delete)
- `Products → StockTakeCount` (**no** cascade — never lose count history) · `Products → ProductShrinkageAllowance` (cascade) · `Employees → StockTakeSession` (no cascade)
- Each lookup feeds its foreign key (status, method, scan status, remediation).

## The rules built in

1. One count line per product per session (enforced by the unique index).
2. A scan resolves `ScanCode` against `Products.SKUBarCode`, finds-or-creates the count line, and
   attaches to it; unmatched codes are parked as `Unmatched`.
3. Scanned lines keep `CountedQuantity = SUM(ScanQuantity)`.
4. A package scan adds `Products.QuantityInPackage`; a unit scan adds 1.
5. `ExpectedQuantity` is snapshotted when the session opens, so variance is measured against a
   fixed baseline.
6. **Shrinkage % = (Expected − Counted) ÷ Expected**, shown on the count line.
7. **Your 5% rule:** if shrinkage exceeds the product's allowed rate — the per-product override if
   one exists, otherwise the 5% default from `SystemSettings` — the line is **Flagged** for review.
   (The flagging logic lives in code; these tables hold everything it needs.)

---

> **Want GPC house style instead?** Swap the standards layer for the GPC default and regenerate —
> the same tables come back as `tblStockTakeCount`, `tlkpScanStatus`, and so on, with
> `CreatedDate / CreatedBy` audit columns. Nothing in the template changes; only the standards layer
> does. That's the split that lets one template serve every shop.
