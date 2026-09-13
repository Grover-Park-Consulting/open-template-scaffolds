---
template: northwind-stocktake-scan-scaffold
title: Northwind Scanned Stocktake — Scan-Processing VBA Scaffold
domain: northwind
type: vba-scaffold
version: 0.7.1
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

## Validating the build

**To the AI assistant.** This template and `northwind-stocktake-scan-outcome-first` promise the
identical result — the same stocktake behaviour, produced by a different route — and each says so
where it names the other. That promise is what gets checked, not which route built it, so there is
one checklist for both rather than two.

**Run every entry under `northwind-stocktake-scan-outcome-first.md`'s "How you validate the
template's output" against this build, on a copy, exactly as that template requires.** Do this
whether you generated the code yourself or handed the developer the files to import — the checks
read the database this build produced, not the procedures that produced it. Report against that same
numbered list in the build record: one entry per check, what was done and what was observed, passed
or not passed. An entry with neither outcome is a check that was not run, and the record is not
complete until it has one.

**Do not devise your own list in place of that one.** Reading the procedures below and working out
a plausible set of checks from them is easy to do and produces a list that tests what this code
does. The list above tests what the developer was promised, which is a different thing and the only
one of the two that is validation. So a build has passed when it has passed those checks, and a
report saying validation passed means those checks and no others — name the list you ran, so the
developer can see which one it was.

Six things follow from this being procedure skeletons rather than an open route.

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
  mechanism.** Where the build creates a count line on demand, the check confirms the duplicate-key
  refusal was turned into "use the line the other counter just created" — trapped and re-read, not
  surfaced to the counter as a failure. Where the build instead pre-creates every count line when
  the session opens, the race is gone by construction, and the check confirms that: two counters
  scanning the same product both find a line already there and neither creates one. Run it either
  way. Which of the two shapes the build used is part of what the entry records.

- **Check 5 needs Business Rule 4, which this template marks without deciding.** `RecordScan`'s
  `[BUSINESS LOGIC #4]` marker says a package scan adds `Products.QuantityInPackage` and a unit scan
  adds 1; *how* a scan is known to be one or the other is parked, here and in the paired table
  template both. Settle it with the developer before the checks run, not during them. Where the
  engagement decides there are no package barcodes at all, check 5 has no input to run against, and
  its entry records that decision and names whose it was.

- **Confirm a raised error reaches the logger at every frame it passes through, as a thirteenth
  entry.** The `errHandler` block sits in every procedure by design, and `EnsureCountLine` raises
  when a product has no count line. That should leave a log entry from the procedure that raised
  *and* from `ProcessScan`, roll the transaction back, and leave no scan row behind. This is this
  route's own behaviour rather than anything the outcome-first list promises, so it is recorded
  after the twelve, not folded into them. Adding to that list is allowed; substituting for it is
  not.

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
    Dim rs As DAO.Recordset

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #3] StockTakeCount.CountedQuantity = SUM(StockTakeScan.ScanQuantity) for
    '            lCountID, counting only scans where ScanStatusID = scanStatusValid — a scan
    '            marked Duplicate is excluded, never counted twice (stored, not derived — see the
    '            table template's house_assumptions).
    ' [SCAFFOLD] Two steps, and the engine forces both of them.
    '            ONE: read the sum into a variable. ACE refuses an aggregate subquery in an
    '            UPDATE's SET clause (error 3073), so the sum cannot stay inside the UPDATE.
    '            TWO: write a plain UPDATE carrying that number as a literal.
    '            Read the sum with a recordset on db — NEVER with DSum. A domain function runs on
    '            Access's own separate connection, outside the transaction ProcessScan opened, so
    '            it cannot see the StockTakeScan row RecordScan inserted moments earlier. The sum
    '            comes back as the previously committed total, this UPDATE succeeds with that wrong
    '            number, the transaction commits, and CountedQuantity runs exactly one scan behind
    '            for the life of the session. Nothing is raised and nothing is logged. See
    '            _materialization.md, "A domain function cannot see the work of the transaction it
    '            is called inside," for the measured behaviour and the general rule.
    ' >>> SELECT SUM(...) into a snapshot recordset on db, then the UPDATE — both per query-style.md <<<

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
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
  `error-handling.md` transaction guard). **Taking this option changes which reads are safe, so
  re-check every one of them.** With one transaction per scan, `RefreshCountRollup` is the only
  procedure here that reads a table its own transaction has already written. With one transaction per
  session, every scan after the first reads scan rows and count lines that the same still-open
  transaction wrote — so `DetectDuplicateScan` and `EnsureCountLine` must come off domain functions
  as well, or they silently stop seeing the session's own work. Same rule, more call sites: see
  `_materialization.md`, "A domain function cannot see the work of the transaction it is called
  inside."
- **Unmatched-scan review queue** — route `scanStatusUnmatched` scans to a review surface instead of
  leaving them parked.

## Parked / future considerations (not in this design)

- **Package-vs-unit disambiguation** — *how* a scan is known to be a package vs. a unit (the rule
  that drives Business Rule #4) is itself undecided in the table template; lives here when resolved.
- **Indexed barcode field** — the table template flags `Products.SKUBarCode` (Memo) as unindexable;
  a production build wants an indexed barcode field for `ResolveScanCode` performance.
