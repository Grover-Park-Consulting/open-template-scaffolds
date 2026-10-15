---
step: 05-checks
title: "Run the checks"
platform_facts: [mcp-run-procedure, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

The checklist this step runs is in `error-logging-outcome-first`, in its step that holds "How you
validate the template's output". Fetch it with `get_part` (the `parts` list of that template's master
names the step) before running anything.

## Validating the build

**To the AI assistant.** This template and `error-logging-outcome-first` promise the identical result
— the same error logging, produced by a different route — and each says so where it names the other.
That promise is what gets checked, not which route built it, so there is one checklist for both
rather than two.

**Run every entry under `error-logging-outcome-first.md`'s "How you validate the template's output"
against this build, on a copy, exactly as that template requires.** Do this whether you generated the
code yourself or handed the developer the files to import — the checks read the database this build
produced, not the procedures that produced it. Report against that same numbered list in the build
record: one entry per check, what was done and what was observed. Passed and not passed are the only
outcomes, including where the first method to run a check hits an obstacle — see
`_template-schema.md` §12.2 for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every
entry opens with, and what to do before settling for a soft result.

**Do not devise your own list in place of that one.** Reading the procedures below and working out a
plausible set of checks from them is easy to do and produces a list that tests what this code does.
The list above tests what the developer was promised, which is a different thing and the only one of
the two that is validation. So a build has passed when it has passed those checks, and a report
saying validation passed means those checks and no others — name the list you ran, so the developer
can see which one it was.

Three things follow from this being working code shaped by a wizard rather than an open route.

- **Record the six wizard answers alongside the checks.** Checks 6, 7 and 8 each read differently
  depending on what the developer chose — where errors are recorded, whether there is a file
  fallback, and what the person at the keyboard is shown. The entry for each of those three is only
  interpretable next to the answer it was run against, so the build record carries both.

- **Drive every check through `LogError`.** Each one is written as something that happens when an
  error is raised; here the equivalent is a call to `LogError` from a handler. Drive them that way
  rather than appending rows to `tblErrorLog` by hand — a row written directly bypasses
  `WriteErrorToTable`, `WriteErrorToFile` and `ShowErrorToUser`, which is the whole of what is being
  checked.

- **Compile the host's VBA project before running any of them, and record that you did.** The checks
  read a database, and code that will not compile never reaches it — so an uncompiled build fails
  every check on the list at once, with the wrong cause attached to each of them. This matters more
  here than elsewhere: the wizard's alternatives ship as commented blocks, and moving which block is
  live is exactly the kind of edit that leaves a project that no longer compiles.
