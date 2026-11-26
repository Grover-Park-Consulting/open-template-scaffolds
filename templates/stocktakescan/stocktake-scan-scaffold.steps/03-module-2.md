---
step: 03-module-2
title: "Write and import the scan module, modStockTakeScan, part 2 of 2 (from DetectDuplicateScan)"
platform_facts: [mcp-module-import, mcp-line-numbers, vbe-line-continuation, domain-function-transaction, row-lock-errors, ace-update-aggregate]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures, continued

These complete `modStockTakeScan` from step 2. Each is a step inside `ProcessScan`'s transaction, so
each hands its errors back to `ProcessScan` (the propagating shape) and reports nothing itself. When
they are written, import the whole module.

### DetectDuplicateScan — `Private Function` → `Boolean`

```vba
Private Function DetectDuplicateScan(ByVal lngCountID As Long, _
                                     ByVal lngScanQuantity As Long, _
                                     ByVal db As DAO.Database) As Boolean
    Const ProcName As String = "DetectDuplicateScan"
    ' [SCAFFOLD] True if an existing scan on lngCountID matches this one closely enough to be
    '            the same physical item scanned twice. Called only when lngCountID <> 0 — the
    '            duplicate check does not apply to unmatched scans (Business Rule 2).
    ' [SCAFFOLD] db is the same Database object ProcessScan opened its transaction on — see
    '            EnsureCountLine's note above; the reason is the same one here.
    Dim rst     As DAO.Recordset
    Dim strSQL   As String
    Dim lngWindow As Long

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler

    ' [BUSINESS LOGIC #2] lngWindow = SystemSettings.DuplicateScanWindowSeconds (plain integer
    '            seconds — not the [percent*1000] convention used elsewhere in SystemSettings).
    ' >>> read the setting, per query-style.md <<<
    lngWindow = 0

    ' [BUSINESS LOGIC #2] an existing StockTakeScan on lngCountID with the same ScanQuantity and a
    '            ScannedOn within lngWindow seconds of Now() makes this scan a duplicate.
    ' >>> lookup query, per query-style.md <<<
    strSQL = vbNullString
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
    DetectDuplicateScan = Not rst.EOF

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step inside ProcessScan's
    '            transaction. Copy Err first, release only what this procedure opened, and
    '            hand the error back, so ProcessScan rolls back instead of committing.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set rst = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### RecordScan — `Private Function` → `Long`

```vba
Private Function RecordScan(ByVal lngSessionID As Long, _
                            ByVal lngCountID As Long, _
                            ByVal strScanCode As String, _
                            ByVal lngScanQuantity As Long, _
                            ByVal db As DAO.Database) As Long
    Const ProcName As String = "RecordScan"
    ' [SCAFFOLD] Insert one StockTakeScan row; return the new StockTakeScanID.
    '            lngCountID = 0 for an unmatched scan (no count line). lngSessionID is stored on
    '            every scan, so an unmatched scan still belongs to its session.
    ' [SCAFFOLD] db is the same Database object ProcessScan opened its transaction on — passed on
    '            to DetectDuplicateScan below rather than let it open its own CurrentDb.
    Dim lngScanStatusID As Long

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler

    ' [BUSINESS LOGIC #2] resolve the status: Unmatched when there is no count line, else
    '            Duplicate when DetectDuplicateScan says so, else Valid. A duplicate is still
    '            inserted here — RefreshCountRollup (Business Rule 3) is what excludes it.
    If lngCountID = 0 Then
        lngScanStatusID = scanStatusUnmatched
    ElseIf DetectDuplicateScan(lngCountID, lngScanQuantity, db) Then
        lngScanStatusID = scanStatusDuplicate
    Else
        lngScanStatusID = scanStatusValid
    End If

    ' [BUSINESS LOGIC #4] Business Rule 4: a package scan adds Products.QuantityInPackage; a unit
    '            scan adds 1. This signature has no parameter telling RecordScan which case applies
    '            — lngScanQuantity is the caller-supplied scan quantity, not a package/unit flag — so
    '            the rule as stated cannot be delivered from this code alone. How a scan is known to
    '            be a package vs. a unit is parked (see "Validating the build," check 5, and
    '            "Package-vs-unit disambiguation" below); settle it with the developer and extend
    '            this signature (or resolve it inside the function) before writing the insert.
    '            StockTakeSessionID = lngSessionID; ScannedOn = Now(); ScanStatusID = lngScanStatusID.
    ' >>> insert the scan row, per query-style.md; set RecordScan = new StockTakeScanID <<<

Cleanup:
    On Error Resume Next
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step inside ProcessScan's
    '            transaction. Copy Err first, release only what this procedure opened, and
    '            hand the error back, so ProcessScan rolls back instead of committing.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### RefreshCountRollup — `Private Sub`

```vba
Private Sub RefreshCountRollup(ByVal lngCountID As Long, ByVal db As DAO.Database)
    Const ProcName As String = "RefreshCountRollup"
    ' [SCAFFOLD] Recompute the stored rollup for one count line.
    Dim rst As DAO.Recordset

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler

    ' [BUSINESS LOGIC #3] StockTakeCount.CountedQuantity = SUM(StockTakeScan.ScanQuantity) for
    '            lngCountID, counting only scans where ScanStatusID = scanStatusValid — a scan
    '            marked Duplicate is excluded, never counted twice (stored, not derived — see the
    '            table template's house_assumptions).
    ' [BUSINESS LOGIC #5] The same UPDATE also sets StockTakeCountMethodID = Scan. A scan reaching
    '            this line means it is now being counted by scanning — true whether the line was
    '            Pending (its first scan) or already Scan (a later one on the same line); this
    '            procedure runs on every scan, so writing it unconditionally each time is simpler
    '            than checking first and gives the same result.
    ' [SCAFFOLD] Two steps, and the engine forces both of them.
    '            ONE: read the sum into a variable. ACE refuses an aggregate subquery in an
    '            UPDATE's SET clause (error 3073), so the sum cannot stay inside the UPDATE.
    '            TWO: write a plain UPDATE carrying that number as a literal.
    '            Read the sum with a recordset on db — the same Database object ProcessScan passed
    '            in, opened on the transaction's own Workspace — NEVER with DSum. A domain
    '            function runs on Access's own separate connection, outside the transaction
    '            ProcessScan opened, so it cannot see the StockTakeScan row
    '            RecordScan inserted moments earlier. The sum comes back as the previously committed
    '            total, this UPDATE succeeds with that wrong number, the transaction commits, and
    '            CountedQuantity runs exactly one scan behind for the life of the session. Nothing is
    '            raised and nothing is logged. See _materialization.md, "A domain function cannot see
    '            the work of the transaction it is called inside," for the measured behaviour and the
    '            general rule, and error-handling.md's "Transaction guard" for where db comes
    '            from.
    ' >>> SELECT SUM(...) into a snapshot recordset on db, then the UPDATE — both per query-style.md <<<

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step inside ProcessScan's
    '            transaction. Copy Err first, release only what this procedure opened, and
    '            hand the error back, so ProcessScan rolls back instead of committing.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set rst = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Sub
```

### EvaluateVariance — `Private Sub`

```vba
Private Sub EvaluateVariance(ByVal lngCountID As Long, ByVal db As DAO.Database)
    Const ProcName As String = "EvaluateVariance"
    ' [SCAFFOLD] Set the remediation flag for one count line after its rollup.
    ' [SCAFFOLD] db is the same Database object the caller's transaction is open on — read the
    '            row's current CountedQuantity/ExpectedQuantity with a recordset on db, never with
    '            DLookup. The same reasoning as RefreshCountRollup's note applies: this runs inside
    '            the caller's own still-open transaction (ProcessScan for an ordinary scan,
    '            OpenStockTakeSession for a freshly created baseline line), and a domain function
    '            would read the last committed value instead of what that transaction just wrote.

    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler

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
    '            Either comparison holding sets RemediationStatusID = Flagged, else None — but only
    '            where the line's current RemediationStatusID is already None or Flagged. Where it
    '            is Under Review or Resolved, a human reviewer put it there; this evaluation runs on
    '            every scan (including a recount), and without checking first would silently revert
    '            that decision. Read the current value before writing a new one, and skip the write
    '            entirely when it is Under Review or Resolved.
    '            At ExpectedQuantity = 0 this needs no special case: CountedQuantity can't be
    '            negative, so the shortfall line can never hold, and the overage line reduces to
    '            "CountedQuantity > 0" — any nonzero count where none was expected gets flagged,
    '            which is the correct answer, not a worked-around one.
    '            Where the host's on-hand figure can come back negative, a count line can carry a
    '            negative ExpectedQuantity. The overage line then holds for any count at all,
    '            including zero, and the line is flagged. Leave that alone — the system's own
    '            figure was impossible before anyone counted. Do NOT add a branch for it and do
    '            NOT guard these comparisons with a test on ExpectedQuantity: that is the division
    '            coming back by another route.
    '            Do not write which direction tripped it anywhere — RemediationStatus does
    '            not record that (table template's house_assumptions); a reviewer reads the sign
    '            of VarianceQuantity instead.
    ' >>> read values, compute, update RemediationStatusID, per query-style.md <<<

Cleanup:
    On Error Resume Next
    Exit Sub

errHandler:
    ' [STANDARDS - error-handling.md] propagating shape: a step inside ProcessScan's
    '            transaction. Copy Err first, release only what this procedure opened, and
    '            hand the error back, so ProcessScan rolls back instead of committing.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Sub
```
