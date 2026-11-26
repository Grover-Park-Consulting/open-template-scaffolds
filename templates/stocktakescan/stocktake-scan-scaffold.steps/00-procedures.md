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
                                     Optional ByVal lngConductedByEmployeeID As Long = 0) As Long
```

Create a stocktake session and take its baseline; return StockTakeSessionID, or 0 when it could not be opened (the person has been told why). Creating the session, its count lines and their first variance check run inside a single transaction — a session that failed partway through its baseline is a session with an incomplete baseline, which no later scan repairs, so all three either land together or not at all.

Error handling (`error-handling.md`): reporting shape.

### ProcessScan

```vba
Public Function ProcessScan(ByVal lngSessionID As Long, _
                            ByVal strScanCode As String, _
                            ByVal lngScanQuantity As Long) As String
```

Process one scan end to end for a session, inside one transaction: the scan row, the rollup, and the variance evaluation either all land or none do.

Error handling (`error-handling.md`): reporting shape.

### ResolveScanCode

```vba
Private Function ResolveScanCode(ByVal strScanCode As String) As Long
```

Resolve a raw scan code to a ProductID; 0 = unmatched.

Error handling (`error-handling.md`): propagating shape.

### EnsureCountLine

```vba
Private Function EnsureCountLine(ByVal lngSessionID As Long, _
                                 ByVal lngProductID As Long, _
                                 ByVal db As DAO.Database) As Long
```

Find the count line for (session, product) or create it; return StockTakeCountID.

Error handling (`error-handling.md`): propagating shape.

### DetectDuplicateScan

```vba
Private Function DetectDuplicateScan(ByVal lngCountID As Long, _
                                     ByVal lngScanQuantity As Long, _
                                     ByVal db As DAO.Database) As Boolean
```

True if an existing scan on lngCountID matches this one closely enough to be the same physical item scanned twice. Called only when lngCountID <> 0 — the duplicate check does not apply to unmatched scans (Business Rule 2).

Error handling (`error-handling.md`): propagating shape.

### RecordScan

```vba
Private Function RecordScan(ByVal lngSessionID As Long, _
                            ByVal lngCountID As Long, _
                            ByVal strScanCode As String, _
                            ByVal lngScanQuantity As Long, _
                            ByVal db As DAO.Database) As Long
```

Insert one StockTakeScan row; return the new StockTakeScanID. lngCountID = 0 for an unmatched scan (no count line). lngSessionID is stored on every scan, so an unmatched scan still belongs to its session.

Error handling (`error-handling.md`): propagating shape.

### RefreshCountRollup

```vba
Private Sub RefreshCountRollup(ByVal lngCountID As Long, ByVal db As DAO.Database)
```

Recompute the stored rollup for one count line.

Error handling (`error-handling.md`): propagating shape.

### EvaluateVariance

```vba
Private Sub EvaluateVariance(ByVal lngCountID As Long, ByVal db As DAO.Database)
```

Set the remediation flag for one count line after its rollup.

Error handling (`error-handling.md`): propagating shape.
