---
template: northwind-stocktake-scan-scaffold
title: Northwind Scanned Stocktake — Scan-Processing VBA Scaffold
domain: northwind
type: vba-scaffold
version: 0.5.0
status: draft
extends: Northwind (Access Developer Edition)
implements: northwind-stocktake-schema
requires_tables:
  - Products
  - StockTakeSession
  - StockTakeCount
  - StockTakeScan
  - ScanStatus
  - RemediationStatus
  - ProductVarianceAllowance
  - SystemSettings
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
target_module: modStockTakeScan
new_procedures:
  - ProcessScan
  - ResolveScanCode
  - EnsureCountLine
  - DetectDuplicateScan
  - RecordScan
  - RefreshCountRollup
  - EvaluateVariance
---

# Northwind Scanned Stocktake — Scan-Processing VBA Scaffold

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Realize the scan-processing logic that the Northwind stocktake **table** template
(`northwind-stocktake-schema`) defers "to the coding section." This scaffold supplies the
**procedure skeletons** — signatures, recordset plumbing, control flow, and error-handling
structure — for resolving a scan to a product, recording it, rolling the scans up into the stored
count, and evaluating the variance. It does **not** write the domain logic itself: each procedure
marks where that goes, sourced from the table template's numbered **Business Rules**. House style
(the central error logger, how SQL is written) is deferred to the standards layer.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here (signature, plumbing, control flow, error structure).
- **`[STANDARDS]`** — house style, deferred to the standards layer (`error-handling.md`, `query-style.md`, `naming-conventions.md`).
- **`[BUSINESS LOGIC]`** — the domain rule you fill in, sourced from the table template's Business Rules.

## Prerequisites

| Object | Role |
|---|---|
| `northwind-stocktake-schema` tables | The scaffold runs against the tables that template creates (`StockTakeSession`/`StockTakeCount`/`StockTakeScan`, the lookups, `ProductVarianceAllowance`) |
| `Products.SKUBarCode`, `Products.QuantityInPackage` | Scan resolution + package quantity |
| `SystemSettings.DefaultAllowableShortageRate`, `SystemSettings.DefaultAllowableOverageRate` | Fallback variance rates |
| `SystemSettings.DuplicateScanWindowSeconds` | Duplicate-scan detection window |
| A central error logger | `error-handling.md` |

### Where this module goes in a split database

A **split database** is the normal shape for Access applications, especially those in multi-user
environments: one file holds the tables (the **back end**, usually on a shared network drive — never
on OneDrive, Dropbox, or any other file-syncing cloud folder, which corrupts a shared Access back
end), and each person runs their own copy of a second file holding the forms, reports, and code (the
**front end**), whose tables are *links* pointing at the back end. A single-file database — one
.accdb holding everything — is an acceptable choice for one user, and everything here works there
too; put everything in that one file and ignore the distinction.

| Module | Back end | Front end | Why |
|---|---|---|---|
| `modStockTakeScan` (this scaffold) | No | **Yes** | `ProcessScan` is called by whatever the counter is using to scan — a form on a laptop in the warehouse. The code belongs beside that form, in each counting station's own front end. |

**A stocktake is the multi-user case, not the exception.** Several people counting different aisles
at once is the ordinary way this gets used, and each of them is running a separate front end writing
into the one shared back end. Two consequences worth knowing before you fill in the business logic:

1. **Every scan is a round trip.** `ProcessScan` runs four operations against the back end for one
   barcode — resolve the code, find or create the count line, insert the scan row, update the
   rollup. On a single-file database that is invisible. Across a network link, at scanning speed, it
   is the thing that decides whether the app feels instant or laggy. Keep the work in SQL that the
   back end can do in one pass rather than opening recordsets and stepping through them.
2. **`EnsureCountLine` can race.** Two counters scanning the same product at the same moment can
   both find no count line and both create one — and the schema's unique index on
   (`StockTakeSessionID`, `ProductID`) will refuse the second insert with an engine error. Checking
   first does not prevent this; it only makes it rare. Trap the duplicate-key error and re-read the
   line that the other person just created, rather than assuming the check is sufficient.

**If you prefer a data macro for the rollup instead**, it attaches to `StockTakeCount` in the back
end and maintains `CountedQuantity` there — closer to the data, and it fires no matter which front
end (or which other tool) inserted the scan. That is a legitimate alternative to
`RefreshCountRollup`; the schema template's Business Rule 3 permits either.

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. **Every procedure ends with
the same `errHandler` block** — shown in full in `ProcessScan` and referenced thereafter, because
the VBE-reflection form (`error-handling.md`) is *identical* in every procedure by design. Line
numbers are deliberately absent (house-specific; see `error-handling.md`).

### ProcessScan — `Public Sub` (entry point)

```vba
Public Sub ProcessScan(ByVal lSessionID As Long, _
                       ByVal sScanCode As String, _
                       ByVal lScanQuantity As Long)
    ' [SCAFFOLD] Process one scan end to end for a session.
    Dim lProductID As Long
    Dim lCountID   As Long

    On Error GoTo errHandler

    lProductID = ResolveScanCode(sScanCode)
    If lProductID = 0 Then
        ' [BUSINESS LOGIC #2] unmatched code: record for review, no count line
        RecordScan 0, sScanCode, lScanQuantity
        GoTo Cleanup
    End If

    lCountID = EnsureCountLine(lSessionID, lProductID)
    ' [BUSINESS LOGIC #2] RecordScan resolves Valid vs. Duplicate itself, via DetectDuplicateScan
    RecordScan lCountID, sScanCode, lScanQuantity
    RefreshCountRollup lCountID
    EvaluateVariance lCountID

Cleanup:
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] error reporting comes from the standards layer.
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Sub
```

*(`scanStatusValid` / `scanStatusUnmatched` / `scanStatusDuplicate` resolve to `ScanStatus` seed rows — wired per engagement.)*

### ResolveScanCode — `Private Function` → `Long`

```vba
Private Function ResolveScanCode(ByVal sScanCode As String) As Long
    ' [SCAFFOLD] Resolve a raw scan code to a ProductID; 0 = unmatched.
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #2] match sScanCode against Products.SKUBarCode
    ' >>> resolution query, written per query-style.md <<<
    sSql = vbNullString

    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then ResolveScanCode = rs!ProductID

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see ProcessScan)
    Resume Cleanup
End Function
```

### EnsureCountLine — `Private Function` → `Long`

```vba
Private Function EnsureCountLine(ByVal lSessionID As Long, _
                                 ByVal lProductID As Long) As Long
    ' [SCAFFOLD] Find the count line for (session, product) or create it; return StockTakeCountID.
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #1] one count line per (StockTakeSessionID, ProductID)
    ' >>> lookup query for an existing line, per query-style.md <<<
    sSql = vbNullString
    Set rs = db.OpenRecordset(sSql, dbOpenDynaset)

    If rs.EOF Then
        ' [BUSINESS LOGIC] create the line (count method = Scan); set EnsureCountLine = new StockTakeCountID
    Else
        EnsureCountLine = rs!StockTakeCountID
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Function
```

### DetectDuplicateScan — `Private Function` → `Boolean`

```vba
Private Function DetectDuplicateScan(ByVal lCountID As Long, _
                                     ByVal lScanQuantity As Long) As Boolean
    ' [SCAFFOLD] True if an existing scan on lCountID matches this one closely enough to be
    '            the same physical item scanned twice. Called only when lCountID <> 0 — the
    '            duplicate check does not apply to unmatched scans (Business Rule 2).
    Dim db     As DAO.Database
    Dim rs     As DAO.Recordset
    Dim sSql   As String
    Dim lWindow As Long

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #2] lWindow = SystemSettings.DuplicateScanWindowSeconds (plain integer
    '            seconds — not the [percent*1000] convention used elsewhere in SystemSettings).
    ' >>> read the setting, per query-style.md <<<
    lWindow = 0

    ' [BUSINESS LOGIC #2] an existing StockTakeScan on lCountID with the same ScanQuantity and a
    '            ScannedOn within lWindow seconds of Now() makes this scan a duplicate.
    ' >>> lookup query, per query-style.md <<<
    sSql = vbNullString
    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    DetectDuplicateScan = Not rs.EOF

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Function
```

### RecordScan — `Private Function` → `Long`

```vba
Private Function RecordScan(ByVal lCountID As Long, _
                            ByVal sScanCode As String, _
                            ByVal lScanQuantity As Long) As Long
    ' [SCAFFOLD] Insert one StockTakeScan row; return the new StockTakeScanID.
    '            lCountID = 0 for an unmatched scan (no count line).
    Dim db           As DAO.Database
    Dim lScanStatusID As Long

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #2] resolve the status: Unmatched when there is no count line, else
    '            Duplicate when DetectDuplicateScan says so, else Valid. A duplicate is still
    '            inserted here — RefreshCountRollup (Business Rule 3) is what excludes it.
    If lCountID = 0 Then
        lScanStatusID = scanStatusUnmatched
    ElseIf DetectDuplicateScan(lCountID, lScanQuantity) Then
        lScanStatusID = scanStatusDuplicate
    Else
        lScanStatusID = scanStatusValid
    End If

    ' [BUSINESS LOGIC #4] a package scan adds Products.QuantityInPackage; a unit scan adds 1.
    '            ScannedOn = Now(); ScanStatusID = lScanStatusID.
    ' >>> insert the scan row, per query-style.md; set RecordScan = new StockTakeScanID <<<

Cleanup:
    On Error Resume Next
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Function
```

### RefreshCountRollup — `Private Sub`

```vba
Private Sub RefreshCountRollup(ByVal lCountID As Long)
    ' [SCAFFOLD] Recompute the stored rollup for one count line.
    Dim db As DAO.Database

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #3] StockTakeCount.CountedQuantity = SUM(StockTakeScan.ScanQuantity) for
    '            lCountID, counting only scans where ScanStatusID = scanStatusValid — a scan
    '            marked Duplicate is excluded, never counted twice (stored, not derived — see the
    '            table template's house_assumptions).
    ' >>> update query, per query-style.md <<<

Cleanup:
    On Error Resume Next
    Set db = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Sub
```

### EvaluateVariance — `Private Sub`

```vba
Private Sub EvaluateVariance(ByVal lCountID As Long)
    ' [SCAFFOLD] Set the remediation flag for one count line after its rollup.
    Dim db As DAO.Database

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #7,#8] VarianceQuantity = CountedQuantity - ExpectedQuantity (signed).
    '            Compare QUANTITIES, never a fraction — do not write ExpectedQuantity into a
    '            denominator anywhere in this procedure. See the table template's Business Rule 8
    '            for why: the fraction form (variance / ExpectedQuantity > rate) divides by zero
    '            the moment a count line's ExpectedQuantity is 0. Multiplying both sides of that
    '            comparison by ExpectedQuantity gives the same result wherever ExpectedQuantity > 0
    '            and never divides at all, which is the form below.
    '            Where VarianceQuantity < 0 (shortfall): effective rate =
    '            ProductVarianceAllowance.AllowableShortageRate if that row/column holds a value,
    '            else SystemSettings.DefaultAllowableShortageRate / 1000; flag when
    '            (ExpectedQuantity - CountedQuantity) > effective rate * ExpectedQuantity.
    '            Where VarianceQuantity > 0 (overage): same resolution against
    '            AllowableOverageRate / DefaultAllowableOverageRate; flag when
    '            (CountedQuantity - ExpectedQuantity) > effective rate * ExpectedQuantity.
    '            Either comparison holding sets RemediationStatusID = Flagged, else None.
    '            At ExpectedQuantity = 0 this needs no special case: CountedQuantity can't be
    '            negative, so the shortfall line can never hold, and the overage line reduces to
    '            "CountedQuantity > 0" — any nonzero count where none was expected gets flagged,
    '            which is the correct answer, not a worked-around one.
    '            Do not write which direction tripped it anywhere — RemediationStatus does
    '            not record that (table template's house_assumptions); a reviewer reads the sign
    '            of VarianceQuantity instead.
    ' >>> read values, compute, update RemediationStatusID, per query-style.md <<<

Cleanup:
    On Error Resume Next
    Set db = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Sub
```

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure plus the error-reporting call and
  line-number policy come from `error-handling.md`, which ranks three options and says when each
  fits. The `MsgBox` block shown is option 3, which needs nothing installed; a practice with its own
  logger substitutes it at the call site, and may number lines or not.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house query
  standard (aliasing, where querydefs live, formatting, safe criteria).
- **Naming** — procedure, variable, and parameter names follow `naming-conventions.md`.

## Extra Options

*Named optional extensions, none of them filled in for an engagement; the filled copy is saved to the
developer's own library, not committed here.*

- **Batch / session transaction** — wrap a whole counting session's scans in one transaction (the
  `error-handling.md` transaction guard).
- **Unmatched-scan review queue** — route `scanStatusUnmatched` scans to a review surface instead of
  leaving them parked.

## Parked / future considerations (not in this design)

- **Package-vs-unit disambiguation** — *how* a scan is known to be a package vs. a unit (the rule
  that drives Business Rule #4) is itself undecided in the table template; lives here when resolved.
- **Indexed barcode field** — the table template flags `Products.SKUBarCode` (Memo) as unindexable;
  a production build wants an indexed barcode field for `ResolveScanCode` performance.
