---
template: northwind-stocktake-scan-scaffold
title: Northwind Scanned Stocktake — Scan-Processing VBA Scaffold
domain: northwind
type: vba-scaffold
version: 0.1.0
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
  - ProductShrinkageAllowance
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
  - RecordScan
  - RefreshCountRollup
  - EvaluateShrinkage
---

# Northwind Scanned Stocktake — Scan-Processing VBA Scaffold

## Intent

Realize the scan-processing logic that the Northwind stocktake **table** template
(`northwind-stocktake-schema`) defers "to the coding section." This scaffold supplies the
**procedure skeletons** — signatures, recordset plumbing, control flow, and error-handling
structure — for resolving a scan to a product, recording it, rolling the scans up into the stored
count, and evaluating shrinkage. It does **not** write the domain logic itself: each procedure
marks where that goes, sourced from the table template's numbered **Business Rules**. House style
(the central error logger, how SQL is written) is deferred to the standards layer.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here (signature, plumbing, control flow, error structure).
- **`[STANDARDS]`** — house style, deferred to the standards layer (`error-handling.md`, `query-style.md`, `naming-conventions.md`).
- **`[BUSINESS LOGIC]`** — the domain rule you fill in, sourced from the table template's Business Rules.

## Prerequisites

| Object | Role |
|---|---|
| `northwind-stocktake-schema` tables | The scaffold runs against the tables that template creates (`StockTakeSession`/`StockTakeCount`/`StockTakeScan`, the lookups, `ProductShrinkageAllowance`) |
| `Products.SKUBarCode`, `Products.QuantityInPackage` | Scan resolution + package quantity |
| `SystemSettings.DefaultAllowableShrinkageRate` | Fallback shrinkage rate |
| A central error logger | `error-handling.md` (GPC default: `codearchive.GlblErrMsg`) |

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
        RecordScan 0, sScanCode, lScanQuantity, scanStatusUnmatched
        GoTo Cleanup
    End If

    lCountID = EnsureCountLine(lSessionID, lProductID)
    RecordScan lCountID, sScanCode, lScanQuantity, scanStatusValid
    RefreshCountRollup lCountID
    EvaluateShrinkage lCountID

Cleanup:
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] house-specific central logger, shown as a demo
    Call codearchive.GlblErrMsg(iLn:=Erl, _
        sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
        sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), bLog:=True)
    Resume Cleanup
    Resume
End Sub
```

*(`scanStatusValid` / `scanStatusUnmatched` resolve to `ScanStatus` seed rows — wired per engagement.)*

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

### RecordScan — `Private Function` → `Long`

```vba
Private Function RecordScan(ByVal lCountID As Long, _
                            ByVal sScanCode As String, _
                            ByVal lScanQuantity As Long, _
                            ByVal lScanStatusID As Long) As Long
    ' [SCAFFOLD] Insert one StockTakeScan row; return the new StockTakeScanID.
    '            lCountID = 0 for an unmatched scan (no count line).
    Dim db As DAO.Database

    On Error GoTo errHandler
    Set db = CurrentDb

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

    ' [BUSINESS LOGIC #3] StockTakeCount.CountedQuantity = SUM(StockTakeScan.ScanQuantity) for lCountID
    '            (stored, not derived — see the table template's house_assumptions).
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

### EvaluateShrinkage — `Private Sub`

```vba
Private Sub EvaluateShrinkage(ByVal lCountID As Long)
    ' [SCAFFOLD] Set the remediation flag for one count line after its rollup.
    Dim db As DAO.Database

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #7,#8] effective rate = ProductShrinkageAllowance.AllowableShrinkageRate if a row
    '            exists, else SystemSettings.DefaultAllowableShrinkageRate / 1000.
    '            shortfall = (ExpectedQuantity - CountedQuantity) / ExpectedQuantity.
    '            if shortfall > effective rate, set RemediationStatusID = Flagged, else None.
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

- **Error handling** — the `errHandler`/`Cleanup` structure plus the central logger and line-number
  policy come from `error-handling.md`. The `GlblErrMsg` call shown is the GPC default
  (house-specific); a forked practice substitutes its own logger, and may number lines or not.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house query
  standard (aliasing, where querydefs live, formatting, safe criteria).
- **Naming** — procedure, variable, and parameter names follow `naming-conventions.md`.

## Supplementals

*Empty in the base template. Filled per client engagement; the filled copy is saved to the
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
