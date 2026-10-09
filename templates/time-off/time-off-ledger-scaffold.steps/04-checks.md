---
step: 04-checks
title: Run the checks
platform_facts: [mcp-run-procedure, mcp-file-release, row-lock-errors, domain-function-transaction]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

The checklist this step runs is in `time-off-ledger-outcome-first`, in its step that holds "How you
validate the template's output". Fetch it with `get_part` (the `parts` list of that template's master
names the step) before running anything.

## Validating the build

**To the AI assistant.** This template and `time-off-ledger-outcome-first` promise the identical
result by different routes. That promise is what is checked, so there is one checklist for both.

**Run every entry under that template's "How you validate the template's output" against this build,
on a copy, as it requires, and report against the same numbered list in the build record.** Do not
devise a list from reading the procedures below: that tests what this code does, not what the
developer was promised. Name the list you ran.

Seven things follow from this being skeletons rather than an open route.

- **Drive every check through the six posting procedures**, never by inserting into
  `tblTimeOffEntry` by hand, except where a check says to make a direct insert. A direct insert
  bypasses the lock and the transaction, which is most of what is being checked.
- **Compile the VBA project before running any check, and record that you did.** Code that will not
  compile fails every check at once with the wrong cause attached to each.
- **Leave `CheckHook` empty in the delivered database.** The checks that need a competing writer
  (outcome-first checks 8, 10, 17, 31 and 32): run them on a check copy whose `CheckHook` you have filled in
  to place the competing row at the point the check names: `afterCheck:Earned`, `afterCheck:Taken`,
  `afterCheck:Correction` after a check has passed, and `betweenHalves` between the two writes of a
  replacement. Record what the copy changed.
- **The competing writer must be a second process you started and identified.** Observed in the
  2026-10-05 trial: a second connection opened inside the same process is not independent, because
  `CurrentDb` belongs to the default workspace the first writer is using, and a file that an Access
  MCP server holds open exclusively cannot be opened by a second process at all. Close that session,
  start both processes yourself on a check copy, and identify each by process ID, as
  the method `access-gate` requires. Where you cannot, record the check as not tested.
- **Run the same checks once with `LockEmployee` emptied**, on a copy, for the checks that force a
  collision, so that you know each can fail. Record that it did.
- **Check 30 forces a failure between the two writes of a replacement.** Force it from `CheckHook`
  at `betweenHalves`, in `modTimeOffPosting`, never inside `TimeOffEntryCheck`. An error raised
  there reaches the engine as an unhandled error and can stop the run with a dialog.
- **Checks 2 and 23 carry their own conditions** (the route chosen for Business Rule 1, and whether
  hours are posted by the developer, on open, or on a schedule). State them in the entry.
