---
template: stocktake-scan-scaffold
title: Scanned Stocktake — Scan-Processing VBA Scaffold
domain: stocktakescan
type: vba-scaffold
version: 0.11.1
status: stable
extends: Northwind (Access Developer Edition)
implements: stocktake-schema
requires_tables:
  - Products
  - StockTakeSession
  - StockTakeCount
  - StockTakeScan
  - ScanStatus
  - RemediationStatus
  - ProductVarianceAllowance
  - SystemSettings
platform_facts: [target-file, sql-insert-truncation, domain-function-transaction, row-lock-errors, ace-update-aggregate]
steps: [01-open-and-confirm-tables, 02-module, 03-checks]
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
warnings:
  - These procedures write real rows into the stocktake tables as scans come in, and a run that goes
    wrong leaves part of a count behind with no way to undo it. A build against a database in real
    use is preceded by a backup copy of the file, and the developer is asked for one before anything
    is changed.
related:
  - "capital-asset-tracking-outcome-first — a similar process for a different purpose:
    reconciling a table of items against barcode scans. The purpose of asset-tracking is to account
    for **fixtures and equipment** rather than **sale inventory**. The scan-resolution logic differs
    between the two and they solve similar, but different, problems, so they do not share a
    scan-resolution mechanism."
---

# Scanned Stocktake — Scan-Processing VBA Scaffold

**Status last determined:** 2026-09-26.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Realize the scan-processing logic that the Northwind stocktake **table** template
(`stocktake-schema`) defers "to the coding section." This scaffold supplies the
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
| `stocktake-schema` tables | The scaffold runs against the tables that template creates (`StockTakeSession`/`StockTakeCount`/`StockTakeScan`, the lookups, `ProductVarianceAllowance`) |
| `Products.SKUBarCode`, `Products.QuantityInPackage` | Scan resolution + package quantity |
| `SystemSettings.DefaultAllowableShortageRate`, `SystemSettings.DefaultAllowableOverageRate` | Fallback variance rates |
| `SystemSettings.DuplicateScanWindowSeconds` | Duplicate-scan detection window |
| A central error logger | `error-handling.md` |

### Ask before building

**Where this is a database already in use, ask for a backup copy before changing anything.** This
template alters the file it is built into. Two questions, in this order:

1. *"Is this a database you already use, or a new one you're trying this out on?"*
2. Where it is one they already use: *"Make a copy of the file before I start. Say when it's done and
   the build continues; say there's no copy and it stops."*

**Stop and build nothing if they say there is no copy.** A copy made before anything changes is the
only way back.

A new database, or one they are trying this out on, needs no copy — the question ends at step 1.

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

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure plus the error-reporting call and
  line-number policy come from `error-handling.md`, which ranks two options and says when each
  fits. The `MsgBox` block shown is option 2, which needs nothing installed; a practice with its own
  logger substitutes it at the call site, and may number lines or not.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house query
  standard (aliasing, where querydefs live, formatting, safe criteria).
- **Naming** — procedure, variable, and parameter names follow `naming-conventions.md`.

**To the AI assistant: once the build is reported finished, and only then, mention each entry under
`related` in the front matter** (`_template-schema.md` §7.1) — one line per entry, what it is and
why. This is not part of the build, never a gate, and never read before this point.

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