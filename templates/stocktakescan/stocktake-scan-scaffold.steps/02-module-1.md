---
step: 02-module-1
title: "Write the scan module, modStockTakeScan, part 1 of 2"
platform_facts: [vbe-line-continuation, domain-function-transaction, row-lock-errors]
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
    Dim booInTrans As Boolean

    On Error GoTo errHandler
    ' [STANDARDS — error-handling.md, "Transaction guard"] the Database used for every write and
    '            every read below must come from the same Workspace the transaction is begun on.
    '            A transaction covers what is written through its own Workspace; a Database taken
    '            from a different workspace is outside it.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)          ' same workspace as BeginTrans

    ws.BeginTrans
    booInTrans = True

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
    booInTrans = False

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing: Set ws = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see ProcessScan), plus the
    '            transaction guard's own rollback: an error partway through ONE, TWO or THREE must
    '            not leave a session row with some, but not all, of its baseline lines.
    If booInTrans Then ws.Rollback: booInTrans = False
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
    Dim booInTrans   As Boolean

    On Error GoTo errHandler
    ' [STANDARDS — error-handling.md, "Transaction guard"] db comes from the same Workspace the
    '            transaction is begun on — see that section for why a db from a different workspace
    '            compiles cleanly and still produces a silently wrong number.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)          ' same workspace as BeginTrans

    ws.BeginTrans
    booInTrans = True

    lProductID = ResolveScanCode(sScanCode)
    If lProductID = 0 Then
        ' [BUSINESS LOGIC #2] unmatched code: record for review, no count line
        RecordScan lSessionID, 0, sScanCode, lScanQuantity, db
        GoTo Commit
    End If

    lCountID = EnsureCountLine(lSessionID, lProductID, db)
    ' [BUSINESS LOGIC #2] RecordScan resolves Valid vs. Duplicate itself, via DetectDuplicateScan
    RecordScan lSessionID, lCountID, sScanCode, lScanQuantity, db
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
    booInTrans = False

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
    If booInTrans Then ws.Rollback: booInTrans = False
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

**The module continues in step 3.** Import it once, when step 3's procedures are written.
