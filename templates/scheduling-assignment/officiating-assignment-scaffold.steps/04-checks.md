---
step: 04-checks
title: "Run the checks"
platform_facts: [mcp-run-procedure, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

The checklist this step runs is in `officiating-assignment-outcome-first`, in its step that holds "How you
validate the template's output". Fetch it with `get_part` (the `parts` list of that template's master
names the step) before running anything.

## Validating the build

**To the AI assistant.** This template and `officiating-assignment-outcome-first` reach the
same result by different routes, and each says so where it names the other. That result is what gets
checked, not which route built it, so there is one checklist for both rather than two.

**Run every entry under `officiating-assignment-outcome-first.md`'s "How you validate the
template's output" against this build, on a copy, exactly as that template requires.** Do this
whether you generated the code yourself or handed the developer the files to import — the checks read
the database this build produced, not the procedures that produced it. Report against that same
numbered list in the build record: one entry per check, what was done and what was observed. Passed
and not passed are the only outcomes, including where the first method to run a check hits an
obstacle — see `_template-schema.md` §12.2 for the full rule, the `Result: PASSED` / `Result: NOT
PASSED` line every entry opens with, and what to do before settling for a soft result.

**Check 2 is the one entry that turns on which route built the database, and that list already says
so.** It tests that an inactive official is refused on a direct edit to `tblGameOfficial`, which only
the Data Macro route reaches. This template is the VBA route named under that file's *Business Rule
3*: `ValidateAssignment` covers whatever calls it and nothing else. Run the check anyway, record what
you observed, name the route this build took, and report the entry the way check 2's own text
directs — it is a disclosed trade-off of this route, not something this build got wrong.

**Do not devise your own list in place of that one.** Reading the procedures below and working out a
plausible set of checks from them is easy to do and produces a list that tests what this code does.
The list above tests what the developer was promised, which is a different thing and the only one of
the two that is validation. So a build has passed when it has passed those checks, and a report
saying validation passed means those checks and no others — name the list you ran, so the developer
can see which one it was.

Four things follow from this being procedure skeletons rather than an open route.

- **Run `EnsureGameValidationRule` before checks 3, 4, 8, 9 and 10.** All five test a game-level rule
  "whatever route you use," one of them through an append or import query — and nothing in VBA
  reaches a save made that way. The table-level Validation Rule does, which is what that procedure
  puts in place. It is a one-time setup call rather than part of the assignment path, so a build that
  never runs it fails all five at once with the wrong cause attached to each.

- **Close everything before running `EnsureGameValidationRule`, and confirm `ListOpenObjects` reports
  it clean if it stops you.** Altering `tblGame` needs the table to itself; something left open fails
  with the engine's own error 3420, which names nothing about what is open. `ListOpenObjects` exists
  to name it instead — run it yourself if the procedure stops, rather than guessing what to close.

- **Drive checks 1 and 2 through `AssignOfficial`, not by inserting rows into `tblGameOfficial`.** An
  insert made directly bypasses `ValidateAssignment`, which on this route is the whole of what those
  two checks examine.

- **Compile the host's VBA project before running any of them, and record that you did.** The checks
  read a database, and code that will not compile never reaches it — so an uncompiled build fails
  every check on the list at once, with the wrong cause attached to each of them. A compile catches
  that in seconds.
