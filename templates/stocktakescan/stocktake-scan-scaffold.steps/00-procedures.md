---
step: 00-procedures
title: "The procedures this template writes"
platform_facts: []
route: both
---

**Who reads this:** the AI assistant, drafting the design.

## The procedures

These are this template's procedures as its code steps write them. Draft the design spec from them:
give each procedure the declaration below and say what it does. Draft no signature or behaviour of
your own. A change to either is the developer's decision: ask it before you present the design.

### OpenStockTakeSession

```vba
Public Function OpenStockTakeSession(Optional ByVal dtStockTakeDate As Variant, _
                                     Optional ByVal lConductedByEmployeeID As Long = 0) As Long
```

Create a stocktake session and take its baseline; return StockTakeSessionID. Creating the session, its count lines and their first variance check run inside a single transaction — a session that failed partway through its baseline is a session with an incomplete baseline, which no later scan repairs, so all three either land together or not at all.

### ProcessScan

```vba
Public Sub ProcessScan(ByVal lSessionID As Long, _
                       ByVal sScanCode As String, _
                       ByVal lScanQuantity As Long)
```

Process one scan end to end for a session, inside one transaction: the scan row, the rollup, and the variance evaluation either all land or none do.

Errors go to the shared error logger (`LogError`) only: a message box is ruled out in the scanning path. Where the logger is not installed, it is built first.

### ResolveScanCode

```vba
Private Function ResolveScanCode(ByVal sScanCode As String) As Long
```

Resolve a raw scan code to a ProductID; 0 = unmatched.

### EnsureCountLine

```vba
Private Function EnsureCountLine(ByVal lSessionID As Long, _
                                 ByVal lProductID As Long, _
                                 ByVal db As DAO.Database) As Long
```

Find the count line for (session, product) or create it; return StockTakeCountID.

### DetectDuplicateScan

```vba
Private Function DetectDuplicateScan(ByVal lCountID As Long, _
                                     ByVal lScanQuantity As Long, _
                                     ByVal db As DAO.Database) As Boolean
```

True if an existing scan on lCountID matches this one closely enough to be the same physical item scanned twice. Called only when lCountID <> 0 — the duplicate check does not apply to unmatched scans (Business Rule 2).

### RecordScan

```vba
Private Function RecordScan(ByVal lCountID As Long, _
                            ByVal sScanCode As String, _
                            ByVal lScanQuantity As Long, _
                            ByVal db As DAO.Database) As Long
```

Insert one StockTakeScan row; return the new StockTakeScanID. lCountID = 0 for an unmatched scan (no count line).

### RefreshCountRollup

```vba
Private Sub RefreshCountRollup(ByVal lCountID As Long, ByVal db As DAO.Database)
```

Recompute the stored rollup for one count line.

### EvaluateVariance

```vba
Private Sub EvaluateVariance(ByVal lCountID As Long, ByVal db As DAO.Database)
```

Set the remediation flag for one count line after its rollup.
