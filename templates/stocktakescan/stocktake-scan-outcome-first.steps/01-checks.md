---
step: 01-checks
title: "How you validate the template's output"
platform_facts: [row-lock-errors, domain-function-transaction]
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

Perform each of these checks against a copy of your database with the tables already built.
**Validate by asking: does this happen or not?**

1. **A barcode that matches a product resolves to that product.**
   - Scan a barcode you know belongs to a product in your catalog
   - Confirm a count line exists for that product in the session, showing that scan's quantity
2. **A barcode that matches nothing is held for review, not lost.**
   - Scan a code that matches no product
   - Confirm it appears in the scan record, marked as unmatched
   - Confirm no product's counted quantity changed because of it
3. **Scanning the same item twice, moments apart, counts once.**
   - Scan a product's barcode
   - Scan the same barcode again immediately, for the same quantity
   - Confirm both scans appear in the scan record
   - Confirm the product's counted quantity reflects only the first
4. **Two separate scans, spaced apart, both count.**
   - Scan a product's barcode
   - Wait past the duplicate-detection window you supplied
   - Scan the same barcode again, for the same quantity
   - Confirm the product's counted quantity now reflects both
5. **A package barcode adds the package quantity.**
   - Scan a product's package barcode, where one exists
   - Confirm the counted quantity increased by that product's package quantity, not by one
6. **A shortfall beyond the allowed tolerance is flagged.**
   - Count a product to fewer than expected, by more than its allowable shortage rate
   - Confirm that product's count line is flagged for review
7. **A shortfall within the allowed tolerance is not flagged.**
   - Count a product to fewer than expected, by less than its allowable shortage rate
   - Confirm that product's count line is not flagged
8. **An overage beyond the allowed tolerance is flagged.**
   - Count a product to more than expected, by more than its allowable overage rate
   - Confirm that product's count line is flagged for review
9. **An overage within the allowed tolerance is not flagged.**
   - Count a product to more than expected, by less than its allowable overage rate
   - Confirm that product's count line is not flagged
10. **The flag does not say which direction, but the variance does.**
    - Flag one count line by shortfall (check 6) and a different one by overage (check 8)
    - Confirm both count lines show the same flag, with nothing on either one distinguishing them
    - Confirm the counted and expected quantities on each still tell you which is which
11. **A count found where none was expected is always flagged.**
    - Find a count line whose `ExpectedQuantity` is zero, or set one up on a product with no
      system on-hand quantity at the moment a session opens
    - Count any nonzero quantity for it
    - Confirm that count line is flagged for review, however small the count
12. **Two counters, one product, the same moment, do not collide.**
    - **Pick the product deliberately: it must have no count line yet.** Opening a session
      pre-creates a count line for every product the baseline covers (Business Rule 5), so an
      ordinary in-scope product already has its line before either counter scans it — there is no
      race left to collide on, and the check passes trivially even on a build with no race recovery
      at all. Use a product added to the catalog after the session opened, which the baseline never
      saw and which genuinely has no line until the first scan creates one.
    - With that product, this needs two people, or two sessions open at once, counting it for the
      first time in the same stocktake within a second or two of each other
    - Confirm both scans succeed
    - Confirm they land against one count line for that product, not two
    - This is the check named in the warning about concurrent counters — it is worth trying
      deliberately, not just trusting the design
    - **Where the tool building this cannot produce two literally simultaneous processes against
      the file** — a single-threaded automation tool, or a machine that opens Access exclusively so
      a second process cannot reach it while a call is in flight — that is not evidence this check
      cannot run. Open a second, independent connection to the same file (a second DAO or ADODB
      connection, not the one the build already holds), and use it to insert the competing row
      between the first connection's own check and its own insert. That forces the engine to raise
      the real collision error rather than asking you to infer one from documentation, and lets you
      confirm the recovery path returns the other connection's row, not a duplicate. See
      `_template-schema.md` §12.2 for the rule this follows
13. **A product nobody scans is still counted, and shows up as missing.**
    - Open a stocktake and count your way through some but not all of the products, exactly as a
      real one would go
    - Pick a product you know the system expected stock of and that nobody scanned
    - Confirm it has a count line in that session anyway, counted zero against what was expected
    - Confirm that line is flagged for review
    - Confirm its count method reads as not-yet-counted, not as Manual or Scan — neither has
      happened for this product
    - **This check is the one that fails when the baseline was never taken, and it is also the one
      that fails when the baseline was taken but never evaluated.** A build that creates a count line
      only when a scan arrives passes checks 1 through 12 and fails this one, because the product it
      is asking about is invisible to it. A build that creates the line up front but leaves it to be
      flagged only when a scan for that product arrives also passes checks 1 through 12 — the line
      exists — and still fails this one, because it stays unflagged for a product that never gets a
      scan. Both are the same underlying gap: something about the never-scanned product never runs
14. **A scan that fails partway through leaves nothing behind.**
    - Force a scan to fail after it has started but before it would normally finish — set up a
      condition that makes one of the scan's own writes fail, for example a session that is already
      closed, or a count line's foreign key pointing at a session that no longer exists, so the
      failure lands between the point the scan is accepted and the point every one of its writes
      would ordinarily land
    - Confirm no `StockTakeScan` row was left behind for that attempt
    - Confirm the count line's `CountedQuantity` and `RemediationStatusID` are exactly what they were
      before the attempt, not partially updated
    - This is the check named in *the same behavior every time* as "a single scan's writes are
      atomic" — a build with nothing tying a scan's writes together will pass every check above and
      fail only this one, because a partial failure there leaves some of a scan's writes landing and
      others not
