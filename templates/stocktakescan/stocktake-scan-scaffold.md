---
template: northwind-stocktake-scan-scaffold
title: Northwind Scanned Stocktake — Scan-Processing VBA Scaffold
domain: stocktakescan
type: vba-scaffold
version: 0.8.3
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
  - OpenStockTakeSession
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

## Validating the build

**To the AI assistant.** This template and `northwind-stocktake-scan-outcome-first` promise the
identical result — the same stocktake behaviour, produced by a different route — and each says so
where it names the other. That promise is what gets checked, not which route built it, so there is
one checklist for both rather than two.

**Run every entry under `northwind-stocktake-scan-outcome-first.md`'s "How you validate the
template's output" against this build, on a copy, exactly as that template requires.** Do this
whether you generated the code yourself or handed the developer the files to import — the checks
read the database this build produced, not the procedures that produced it. Report against that same
numbered list in the build record: one entry per check, what was done and what was observed. Passed
and not passed are the only outcomes, including where the first method to run a check hits an
obstacle — see `_template-schema.md` §12.2 for the full rule, the `Result: PASSED` / `Result: NOT
PASSED` line every entry opens with, and what to do before settling for a soft result.

**Do not devise your own list in place of that one.** Reading the procedures below and working out
a plausible set of checks from them is easy to do and produces a list that tests what this code
does. The list above tests what the developer was promised, which is a different thing and the only
one of the two that is validation. So a build has passed when it has passed those checks, and a
report saying validation passed means those checks and no others — name the list you ran, so the
developer can see which one it was.

Seven things follow from this being procedure skeletons rather than an open route.

- **Drive every check through `ProcessScan`.** Each one is written as something the developer does
  with a scanner; here the equivalent is a call to `ProcessScan` with a session, a code, and a
  quantity. Drive them that way rather than inserting rows into `StockTakeScan` by hand — an insert
  made directly bypasses `RecordScan`, `RefreshCountRollup` and `EvaluateVariance`, which is the
  whole of what is being checked.

- **Compile the host's VBA project before running any of them, and record that you did.** The checks
  read a database, and code that will not compile never reaches it — so an uncompiled build fails
  every check on the list at once, with the wrong cause attached to each of them. This is not a
  formality. Generated VBA can be correct in its logic and still not compile, because DAO puts
  similarly-named members on different objects: transactions belong to the `Workspace`
  (`DBEngine.Workspaces(0).BeginTrans`), not to the `Database` object the rest of the code is
  holding, and a `Database` has no `BeginTrans` at all. A compile is the only thing that catches
  that class of mistake, and it catches it in seconds.

- **Checks 1, 3 and 4 are the rollup checks, and they are what catches a stale read.** Check 1 fails
  when the first scan against a count line leaves `CountedQuantity` at zero; check 4 fails when the
  rollup runs behind the scans instead of with them. Both are failure modes of `RefreshCountRollup`
  reading through a domain function from inside the transaction `ProcessScan` opened — see that
  procedure's `[SCAFFOLD]` note for why, and `_materialization.md` for the measured behaviour. **If
  the Batch / session transaction Extra Option was taken, run all three again afterwards**, and
  check 3 in particular: widening the transaction moves `DetectDuplicateScan`'s read inside it.

- **Check 12 tests the outcome of the `EnsureCountLine` race, and this template names the
  mechanism.** `OpenStockTakeSession` pre-creates a count line for every product the session covers
  (Business Rule 5), so for those products two counters both find a line already there and neither
  creates one — the check confirms that. The race is still reachable and `EnsureCountLine` still has
  to handle it: a product added to the catalog after the session opened has no line, and two counters
  can reach it at the same moment. There the check confirms the duplicate-key refusal was turned into
  "use the line the other counter just created" — trapped and re-read, not surfaced to the counter as
  a failure. **Run it against a product with no line**, or it tests nothing; which path the check
  exercised is part of what the entry records.

- **Check 13 is the one that catches a missing baseline, and it is the check this route is most
  likely to fail.** The skeletons below are a scan path: everything in them starts from a scan
  arriving, and a build that lets `EnsureCountLine` do all the count-line creating passes checks 1
  through 12 without ever opening a session properly. Check 13 asks about a product nobody scanned,
  which that build has no row for at all. `OpenStockTakeSession` is in this scaffold for exactly this
  reason — run it before any check, and give check 13 a product you deliberately leave unscanned.

- **Check 5 needs Business Rule 4, which this template marks without deciding.** `RecordScan`'s
  `[BUSINESS LOGIC #4]` marker says a package scan adds `Products.QuantityInPackage` and a unit scan
  adds 1; *how* a scan is known to be one or the other is parked, here and in the paired table
  template both. Settle it with the developer before the checks run, not during them. Where the
  engagement decides there are no package barcodes at all, check 5 has no input to run against, and
  its entry records that decision and names whose it was.

- **Confirm a raised error reaches the logger at every frame it passes through, as a fourteenth
  entry.** The `errHandler` block sits in every procedure by design, so an error raised deep in the
  call chain should be logged by the procedure that raised it *and* by every procedure it passes
  through on the way out, and leave no scan row behind. **Raise one deliberately rather than waiting
  for one:** calling `ProcessScan` with a `StockTakeSessionID` that does not exist makes the count-line
  insert violate referential integrity, which raises from inside `EnsureCountLine` — three frames of
  log entries for the one error, and a scan count unchanged before and after. This is this route's own
  behaviour rather than anything the outcome-first list promises, so it is recorded after the
  thirteen, not folded into them. Adding to that list is allowed; substituting for it is not.

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
    '            every read below must come from the same Workspace the transaction is begun on —
    '            never from CurrentDb, which is a different connection and does not see this
    '            transaction's own uncommitted work.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)          ' NOT CurrentDb - see error-handling.md

    ws.BeginTrans
    bInTrans = True

    ' [BUSINESS LOGIC #5] ONE: insert the StockTakeSession row (status = Open; the date defaults to
    '            today where none was passed; ConductedByEmployeeID left null where 0 was passed)
    '            and read back its StockTakeSessionID into OpenStockTakeSession.
    ' >>> session insert, per query-style.md <<<

    ' [BUSINESS LOGIC #5] TWO: create one StockTakeCount line for every product the session covers —
    '            by default every product not marked discontinued — each carrying
    '            ExpectedQuantity = the host's computed on-hand for that product AT THIS MOMENT,
    '            CountedQuantity = 0, count method = Scan, RemediationStatusID = None.
    '            The whole point of the rule is that this happens once, here. Taking the figure
    '            later, as each scan arrives, measures every product against a different moment and
    '            leaves an unscanned product with no line at all.
    ' [SCAFFOLD] Where the host computes on-hand with a VBA function rather than a query, this is a
    '            loop over the product list calling it per product, not a single INSERT ... SELECT.
    '            Which it is depends on the host; both satisfy the rule.
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
    '            transaction is begun on, never from CurrentDb — see that section for why a wrong
    '            source here compiles cleanly and still produces a silently wrong number.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)          ' NOT CurrentDb - see error-handling.md

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
    If bInTrans Then ws.Rollback: bInTrans = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
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

    ' [BUSINESS LOGIC #4] a package scan adds Products.QuantityInPackage; a unit scan adds 1.
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
    ' [SCAFFOLD] Two steps, and the engine forces both of them.
    '            ONE: read the sum into a variable. ACE refuses an aggregate subquery in an
    '            UPDATE's SET clause (error 3073), so the sum cannot stay inside the UPDATE.
    '            TWO: write a plain UPDATE carrying that number as a literal.
    '            Read the sum with a recordset on db — the same Database object ProcessScan passed
    '            in, opened on the transaction's own Workspace — NEVER with DSum, and never on a
    '            fresh CurrentDb. Either of those runs on Access's own separate connection, outside
    '            the transaction ProcessScan opened, so it cannot see the StockTakeScan row
    '            RecordScan inserted moments earlier. The sum comes back as the previously committed
    '            total, this UPDATE succeeds with that wrong number, the transaction commits, and
    '            CountedQuantity runs exactly one scan behind for the life of the session. Nothing is
    '            raised and nothing is logged. See _materialization.md, "A domain function cannot see
    '            the work of the transaction it is called inside," for the measured behaviour and the
    '            general rule, and error-handling.md's "Transaction guard" for the db-vs-CurrentDb
    '            half of the same mistake.
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
    '            Either comparison holding sets RemediationStatusID = Flagged, else None.
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

- **Batch / session transaction** — widen `ProcessScan`'s one-transaction-per-scan (built into this
  scaffold by default) to cover a whole counting session's scans in one transaction instead (the
  `error-handling.md` transaction guard). This template deliberately stops at one scan — the library
  is *open*, so widening scope from here, if a developer wants it, is theirs or their own AI
  assistant's to build against the same rule this scaffold already follows. **Taking this option
  changes which reads are safe, so re-check every one of them.** With one transaction per scan,
  `RefreshCountRollup` and `EvaluateVariance` are the only procedures here that read a table their own
  transaction has already written. With one transaction per session, every scan after the first reads
  scan rows and count lines that the same still-open transaction wrote — so `DetectDuplicateScan` and
  `EnsureCountLine` must come off domain functions as well, or they silently stop seeing the session's
  own work. Same rule, more call sites: see `_materialization.md`, "A domain function cannot see the
  work of the transaction it is called inside."
- **Unmatched-scan review queue** — route `scanStatusUnmatched` scans to a review surface instead of
  leaving them parked.

## Parked / future considerations (not in this design)

- **Package-vs-unit disambiguation** — *how* a scan is known to be a package vs. a unit (the rule
  that drives Business Rule #4) is itself undecided in the table template; lives here when resolved.
- **Indexed barcode field** — the table template flags `Products.SKUBarCode` (Memo) as unindexable;
  a production build wants an indexed barcode field for `ResolveScanCode` performance.
