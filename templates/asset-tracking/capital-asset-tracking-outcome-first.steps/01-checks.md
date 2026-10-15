---
step: 01-checks
title: "How you validate the template's output"
platform_facts: []
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

**Do these on a copy, and stop if you do not have one.** The checks add, change, and delete records in
your own tables — that is what they test, and there is no version of them that leaves your data alone.
Make the copy once the build is finished, run every check on it, and keep your working file out of it.
Where your database is split into two files, copy both and keep them together.

1. **A new asset below the threshold is refused.**
   - Add a new asset with an acquisition cost below the district's threshold — do this once through a
     form or the table directly, and once through an append/import query
   - Confirm neither save succeeds
2. **A new asset exactly at the threshold is refused.**
   - Add a new asset with an acquisition cost exactly equal to the threshold
   - Confirm the save does not succeed — the rule is "must exceed," not "must reach"
3. **A new asset above the threshold succeeds.**
   - Add a new asset with an acquisition cost above the threshold
   - Confirm it saves normally
4. **A grandfathered asset can still be edited.**
   - Find or create an asset whose acquisition cost is at or below the district's *current* threshold
     (representing one capitalized before the threshold was last raised)
   - Change a field on it that is not the acquisition cost — its room, say
   - Confirm the save succeeds and is not blocked by the cost check
5. **Lowering an existing asset's cost below the threshold is refused.**
   - Take an asset whose cost is above the threshold and edit its acquisition cost down to at or below
     the threshold
   - Confirm the save does not succeed
6. **A tracked field change is recorded as one history row.**
   - Change one of `RoomID`, `CustodianID`, `DepartmentID`, `FundingSourceID`, or `AssetStatusID` on
     one asset
   - Open the history to confirm it has one new row for that asset naming the change type, the prior
     value, the new value, and when it happened
7. **An untracked field change writes no history row.**
   - Change `AssetNotes` (or another field not named in Business Rule 3) on an asset
   - Confirm no new row appears in the history
8. **Computed book value moves with its inputs.**
   - Note an asset's current book value
   - Change its acquisition cost, useful life, or salvage value
   - Confirm the book value shown changes accordingly, without running any separate recalculation step
9. **Computed book value is floored and freezes.**
   - Take an asset whose useful life has already elapsed
   - Confirm its book value equals its salvage value, not a negative number
   - Wait, or adjust the acquisition date, so more time passes
   - Confirm the book value has not changed further
10. **Disposal without a date is refused.**
    - Set an asset's status to Disposed and leave the disposal date empty
    - Confirm the save does not succeed, whatever the route
11. **A disposed asset keeps its history.**
    - Dispose of an asset that has prior history rows and prior inventory-scan rows
    - Confirm the asset, its history, and its scan history are all still present and readable
12. **A matched scan resolves as Matched.**
    - Start an inventory-audit session for a site
    - Scan the barcode of an asset on file, expected at that site
    - Confirm the scan resolves as Matched
13. **An out-of-site scan resolves as Unexpected.**
    - In the same session, scan the barcode of an asset whose on-file room belongs to a different site
    - Confirm the scan resolves as Unexpected
14. **A repeated scan resolves as Duplicate.**
    - Scan the same barcode a second time in the same session
    - Confirm the second scan resolves as Duplicate and the first is unchanged
15. **A misplaced asset is identifiable without altering its record.**
    - Scan an asset's barcode while recording a room different from its on-file room
    - Confirm the scan records the room it was actually found in
    - Confirm the asset's own `RoomID` is unchanged
16. **A missing asset is identifiable by query, not by a stored flag.**
    - Close a session having left at least one site-expected, Active asset unscanned
    - Confirm nothing on that asset's own record marks it missing
    - Confirm a query against the session identifies it as expected and unmatched
17. **Running the build again does not duplicate anything.**
    - Run whatever attaches the Data Macros and builds the supporting objects a second time
    - Confirm there is still exactly one copy of each — no duplicated Data Macro actions, no second
      settings row, no second copy of any generated procedure
18. **When a build is blocked, it stops rather than leaving tables half-finished.**
    - Leave a table or form open in the database
    - Start the build
    - Confirm it stops, names what is open, and changes nothing
