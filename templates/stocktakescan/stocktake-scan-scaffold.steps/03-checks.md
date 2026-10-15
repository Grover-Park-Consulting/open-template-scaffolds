---
step: 03-checks
title: "Run the checks"
platform_facts: [mcp-run-procedure, mcp-file-release, row-lock-errors, domain-function-transaction]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

The checklist this step runs is in `stocktake-scan-outcome-first`, in its step that holds "How you
validate the template's output". Fetch it with `get_part` (the `parts` list of that template's
master names the step) before running anything.

## Validating the build

**To the AI assistant.** This template and `stocktake-scan-outcome-first` promise the
identical result — the same stocktake behaviour, produced by a different route — and each says so
where it names the other. That promise is what gets checked, not which route built it, so there is
one checklist for both rather than two.

**Run every entry under `stocktake-scan-outcome-first.md`'s "How you validate the
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

- **Check 14 ("a scan that fails partway through leaves nothing behind") also confirms a raised
  error reaches the logger at every frame it passes through.** The `errHandler` block sits in every
  procedure by design, so an error raised deep in the call chain should be logged by the procedure
  that raised it *and* by every procedure it passes through on the way out, and leave no scan row
  behind. **Raise one deliberately rather than waiting for one:** calling `ProcessScan` with a
  `StockTakeSessionID` that does not exist makes the count-line insert violate referential
  integrity, which raises from inside `EnsureCountLine` — one log entry per frame the error passes
  through (three in the standard shape; a different number where the host's own error handler is
  in use, and that number is a result to record, not a target), and a scan count unchanged before
  and after. The multi-frame logging is this route's own
  behaviour rather than anything the outcome-first list promises, so record it as part of running
  check 14, not as a separate entry — the list itself stays the fourteen it already is.
