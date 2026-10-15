---
step: 02-module
title: "Write and import the scan module, modStockTakeScan"
platform_facts: [mcp-module-import, mcp-line-numbers, vbe-line-continuation, domain-function-transaction, row-lock-errors, ace-update-aggregate]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. **Every procedure ends with
the same `errHandler` block** — shown in full in `ProcessScan` and referenced thereafter, because
the errHandler block's form (`error-handling.md`) is *identical* in every procedure by design. Line
numbers are deliberately absent (house-specific; see `error-handling.md`).

### OpenStockTakeSession — `Public Function` → `Long`

**Run this before anything else in a stocktake, and before any check.** It is what takes the
baseline: Business Rule 5 says every count line exists, carrying the expected quantity of the moment
the session opened, before a single code is scanned. Without it the only count lines that ever exist
are for products somebody scanned, and a product that has gone missing entirely — the most serious
thing a stocktake can find — leaves no row anywhere to report it.

```vba
Public Function OpenStockTakeSession(Optional ByVal dtStockTakeDate As Variant, _
                                     Optional ByVal lConductedByEmployeeID As Long = 0) As Long
    ' [SCAFFOLD] Create a stocktake session and take its baseline; return StockTakeSessionID.
    '            ONE, TWO and THREE below run inside a single transaction — a session that failed
    '            partway through its baseline is a session with an incomplete baseline, which no
    '            later scan repairs, so all three either land together or not at all.
    Dim ws       As DAO.Workspace
    Dim db       As DAO.Database
    Dim rs       As DAO.Recordset
    Dim sSql     As String
    Dim bInTrans As Boolean

    On Error GoTo errHandler
    ' [STANDARDS — error-handling.md, "Transaction guard"] the Database used for every write and
    '            every read below must come from the same Workspace the transaction is begun on.
    '            A transaction covers what is written through its own Workspace; a Database taken
    '            from a different workspace is outside it.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)          ' same workspace as BeginTrans

    ws.BeginTrans
    bInTrans = True

    ' [BUSINESS LOGIC #5] ONE: insert the StockTakeSession row (status = Open; the date defaults to
    '            today where none was passed; ConductedByEmployeeID left null where 0 was passed)
    '            and read back its StockTakeSessionID into OpenStockTakeSession.
    ' >>> session insert, per query-style.md <<<

    ' [BUSINESS LOGIC #5] TWO: create one StockTakeCount line for every product the session covers —
    '            by default every product not marked discontinued — each carrying
    '            ExpectedQuantity = the host's computed on-hand for that product AT THIS MOMENT,
    '            CountedQuantity = 0, count method = Pending (neither a manual entry nor a scan has
    '            happened for this product yet — see the table template's Business Rule 5),
    '            RemediationStatusID = None.
    '            The whole point of the rule is that this happens once, here. Taking the figure
    '            later, as each scan arrives, measures every product against a different moment and
    '            leaves an unscanned product with no line at all.
    ' [SCAFFOLD] Where the host computes on-hand with a VBA function rather than a query, this is a
    '            loop over the product list calling it per product, not a single INSERT ... SELECT.
    '            Which it is depends on the host; both satisfy the rule. Read the host's on-hand
    '            before the transaction opens, not inside it: a host function may itself write and
    '            then read back through a domain function, which cannot see the transaction's work.
    ' >>> baseline creation, per query-style.md <<<

    ' [BUSINESS LOGIC #5] THREE: evaluate EACH line just created, immediately, in this same
    '            procedure — do not leave this for ProcessScan. A product nobody ever scans never
    '            reaches ProcessScan, so if evaluation is left to it, that line's RemediationStatusID
    '            stays at None for the entire session no matter how large the shortfall. At this
    '            point CountedQuantity is 0 for every line, so this is just Business Rule 8 run
    '            against each line's own ExpectedQuantity, the same call ProcessScan makes later.
    ' [SCAFFOLD] Pass db through to EvaluateVariance — do not let it open its own CurrentDb. Every
    '            line it reads here is one TWO just wrote, inside this same still-open transaction.
    ' >>> call EvaluateVariance(lNewCountID, db) once per newly created StockTakeCountID <<<

    ws.CommitTrans
    bInTrans = False

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing: Set ws = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see ProcessScan), plus the
    '            transaction guard's own rollback: an error partway through ONE, TWO or THREE must
    '            not leave a session row with some, but not all, of its baseline lines.
    If bInTrans Then ws.Rollback: bInTrans = False
    Resume Cleanup
End Function
```

### ProcessScan — `Public Sub` (entry point)

```vba
Public Sub ProcessScan(ByVal lSessionID As Long, _
                       ByVal sScanCode As String, _
                       ByVal lScanQuantity As Long)
    ' [SCAFFOLD] Process one scan end to end for a session, inside one transaction: the scan row,
    '            the rollup, and the variance evaluation either all land or none do. See "Point of
    '            the transaction, and where it stops" below the code for what this buys and what it
    '            deliberately does not cover.
    Dim ws         As DAO.Workspace
    Dim db         As DAO.Database
    Dim lProductID As Long
    Dim lCountID   As Long
    Dim bInTrans   As Boolean

    On Error GoTo errHandler
    ' [STANDARDS — error-handling.md, "Transaction guard"] db comes from the same Workspace the
    '            transaction is begun on — see that section for why a db from a different workspace
    '            compiles cleanly and still produces a silently wrong number.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)          ' same workspace as BeginTrans

    ws.BeginTrans
    bInTrans = True

    lProductID = ResolveScanCode(sScanCode)
    If lProductID = 0 Then
        ' [BUSINESS LOGIC #2] unmatched code: record for review, no count line
        RecordScan 0, sScanCode, lScanQuantity, db
        GoTo Commit
    End If

    lCountID = EnsureCountLine(lSessionID, lProductID, db)
    ' [BUSINESS LOGIC #2] RecordScan resolves Valid vs. Duplicate itself, via DetectDuplicateScan
    RecordScan lCountID, sScanCode, lScanQuantity, db
    ' [SCAFFOLD] RefreshCountRollup and EvaluateVariance both take db, not CurrentDb, and neither
    '            reads its numbers with a domain function (DSum/DLookup/DMax). A domain function
    '            runs on Access's own separate connection, outside the transaction this procedure
    '            just began, so it cannot see the StockTakeScan row RecordScan inserted moments ago
    '            — see each procedure's own note, and _materialization.md, "A domain function
    '            cannot see the work of the transaction it is called inside."
    RefreshCountRollup lCountID, db
    EvaluateVariance lCountID, db

Commit:
    ws.CommitTrans
    bInTrans = False

Cleanup:
    Set db = Nothing: Set ws = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] error reporting comes from the standards layer, plus the
    '            transaction guard's own rollback: a scan that failed partway through must not leave
    '            a scan row, a rollup, or a flag written without the other two.
    ' [BUSINESS LOGIC — error-logging-outcome-first.md, "How this is built"] Option 2 (a message
    '            box) is not available in this procedure: nothing in the scanning path may interrupt
    '            a counter mid-scan with something they cannot act on. Option 1 (LogError) is the
    '            only choice here — where it is not yet installed, build error-logging-scaffold.md
    '            first rather than falling back to a message box in ProcessScan.
    If bInTrans Then ws.Rollback: bInTrans = False
    LogError MODULE_NAME, PROC_NAME, Erl
    Resume Cleanup
    Resume
End Sub
```

**Point of the transaction, and where it stops.** Wrapping one scan's writes in a transaction makes
that scan atomic — the scan row, the rollup, and the flag either all commit or none do, so a failure
partway through never leaves `CountedQuantity` out of step with the scan record it was computed from.
It says nothing about the scan before it or the scan after it; each call to `ProcessScan` opens and
closes its own transaction, and nothing here holds one open across scans, across a counting session,
or across a batch. That wider scope is deliberately not built here — see the Batch / session
transaction Extra Option below for what changes if a developer wants to widen it, and why this
template stops at one scan.

*(`scanStatusValid` / `scanStatusUnmatched` / `scanStatusDuplicate` resolve to `ScanStatus` seed rows — wired per engagement.)*

### ResolveScanCode — `Private Function` → `Long`

```vba
Private Function ResolveScanCode(ByVal sScanCode As String) As Long
    ' [SCAFFOLD] Resolve a raw scan code to a ProductID; 0 = unmatched.
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    ' [STANDARDS — error-handling.md, "Transaction guard"] CurrentDb is fine here: ProcessScan's
    '            transaction never writes to Products, so this read has none of the transaction's
    '            work to see. Pass db in instead if a build ever adds a write against Products
    '            inside the same transaction, so every read and write in it uses one Database.
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
                                 ByVal lProductID As Long, _
                                 ByVal db As DAO.Database) As Long
    ' [SCAFFOLD] Find the count line for (session, product) or create it; return StockTakeCountID.
    ' [SCAFFOLD] db is passed in from ProcessScan's own transaction — do not Set db = CurrentDb
    '            here. This runs inside that transaction, and a line created at THREE below must be
    '            visible to RecordScan/RefreshCountRollup/EvaluateVariance later in the same call,
    '            which only holds if every one of them is reading and writing through the one
    '            Database object the transaction was begun on.
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler

    ' [BUSINESS LOGIC #1] one count line per (StockTakeSessionID, ProductID)
    ' [SCAFFOLD] In an ordinary session the line is already here — OpenStockTakeSession created one
    '            for every product the session covers. The create branch below is for a product
    '            added to the catalog after the session opened, and it is where two counters can
    '            collide. It is NOT the place the baseline gets taken; a build that leans on it for
    '            that has no line for any product nobody scanned. See Business Rule 5.
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
    Set rs = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Function
```

### DetectDuplicateScan — `Private Function` → `Boolean`

```vba
Private Function DetectDuplicateScan(ByVal lCountID As Long, _
                                     ByVal lScanQuantity As Long, _
                                     ByVal db As DAO.Database) As Boolean
    ' [SCAFFOLD] True if an existing scan on lCountID matches this one closely enough to be
    '            the same physical item scanned twice. Called only when lCountID <> 0 — the
    '            duplicate check does not apply to unmatched scans (Business Rule 2).
    ' [SCAFFOLD] db is the same Database object ProcessScan opened its transaction on — see
    '            EnsureCountLine's note above; the reason is the same one here.
    Dim rs     As DAO.Recordset
    Dim sSql   As String
    Dim lWindow As Long

    On Error GoTo errHandler

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
    Set rs = Nothing
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
                            ByVal db As DAO.Database) As Long
    ' [SCAFFOLD] Insert one StockTakeScan row; return the new StockTakeScanID.
    '            lCountID = 0 for an unmatched scan (no count line).
    ' [SCAFFOLD] db is the same Database object ProcessScan opened its transaction on — passed on
    '            to DetectDuplicateScan below rather than let it open its own CurrentDb.
    Dim lScanStatusID As Long

    On Error GoTo errHandler

    ' [BUSINESS LOGIC #2] resolve the status: Unmatched when there is no count line, else
    '            Duplicate when DetectDuplicateScan says so, else Valid. A duplicate is still
    '            inserted here — RefreshCountRollup (Business Rule 3) is what excludes it.
    If lCountID = 0 Then
        lScanStatusID = scanStatusUnmatched
    ElseIf DetectDuplicateScan(lCountID, lScanQuantity, db) Then
        lScanStatusID = scanStatusDuplicate
    Else
        lScanStatusID = scanStatusValid
    End If

    ' [BUSINESS LOGIC #4] Business Rule 4: a package scan adds Products.QuantityInPackage; a unit
    '            scan adds 1. This signature has no parameter telling RecordScan which case applies
    '            — lScanQuantity is the caller-supplied scan quantity, not a package/unit flag — so
    '            the rule as stated cannot be delivered from this code alone. How a scan is known to
    '            be a package vs. a unit is parked (see "Validating the build," check 5, and
    '            "Package-vs-unit disambiguation" below); settle it with the developer and extend
    '            this signature (or resolve it inside the function) before writing the insert.
    '            ScannedOn = Now(); ScanStatusID = lScanStatusID.
    ' >>> insert the scan row, per query-style.md; set RecordScan = new StockTakeScanID <<<

Cleanup:
    On Error Resume Next
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Function
```

### RefreshCountRollup — `Private Sub`

```vba
Private Sub RefreshCountRollup(ByVal lCountID As Long, ByVal db As DAO.Database)
    ' [SCAFFOLD] Recompute the stored rollup for one count line.
    Dim rs As DAO.Recordset

    On Error GoTo errHandler

    ' [BUSINESS LOGIC #3] StockTakeCount.CountedQuantity = SUM(StockTakeScan.ScanQuantity) for
    '            lCountID, counting only scans where ScanStatusID = scanStatusValid — a scan
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
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Sub
```

### EvaluateVariance — `Private Sub`

```vba
Private Sub EvaluateVariance(ByVal lCountID As Long, ByVal db As DAO.Database)
    ' [SCAFFOLD] Set the remediation flag for one count line after its rollup.
    ' [SCAFFOLD] db is the same Database object the caller's transaction is open on — read the
    '            row's current CountedQuantity/ExpectedQuantity with a recordset on db, never with
    '            DLookup. The same reasoning as RefreshCountRollup's note applies: this runs inside
    '            the caller's own still-open transaction (ProcessScan for an ordinary scan,
    '            OpenStockTakeSession for a freshly created baseline line), and a domain function
    '            would read the last committed value instead of what that transaction just wrote.

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
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Sub
```
