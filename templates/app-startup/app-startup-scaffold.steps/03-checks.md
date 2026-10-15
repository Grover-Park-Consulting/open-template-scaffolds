---
step: 03-checks
title: "Run the checks"
platform_facts: [mcp-run-procedure, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

The checklist this step runs is in `app-startup-outcome-first`, in its step that holds "How you
validate the template's output". Fetch it with `get_part` (the `parts` list of that template's master
names the step) before running anything.

## Validating the build

**To the AI assistant.** This template and `app-startup-outcome-first` promise the identical result —
the same startup behaviour, produced by a different route — and each says so where it names the
other. That promise is what gets checked, not which route built it, so there is one checklist for
both rather than two.

**Run every entry under `app-startup-outcome-first.md`'s "How you validate the template's output"
against this build, on a copy, exactly as that template requires.** Do this whether you generated the
code yourself or handed the developer the files to import — the checks read the application this
build produced, not the procedures that produced it. Report against that same numbered list in the
build record: one entry per check, what was done and what was observed. Passed and not passed are the
only outcomes, including where the first method to run a check hits an obstacle — see
`_template-schema.md` §12.2 for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every
entry opens with, and what to do before settling for a soft result.

**Do not devise your own list in place of that one.** Reading the procedures below and working out a
plausible set of checks from them is easy to do and produces a list that tests what this code does.
The list above tests what the developer was promised, which is a different thing and the only one of
the two that is validation. So a build has passed when it has passed those checks, and a report
saying validation passed means those checks and no others — name the list you ran, so the developer
can see which one it was.

Three things follow from this being procedure skeletons rather than an open route.

- **Drive every check through `Startup`.** Each one is written as something the developer does by
  opening the application; here the equivalent is the `AutoExec` path into `Startup`. Drive them that
  way rather than calling `EnsureBackEndLink` or `RelinkAllTables` directly — a direct call skips the
  order `Startup` imposes, and that order is part of what is being checked.

- **Compile the host's VBA project before running any of them, and record that you did.** The checks
  read a running application, and code that will not compile never runs — so an uncompiled build
  fails every check on the list at once, with the wrong cause attached to each of them. A compile
  catches that in seconds.

- **Check 9 is always performable on this route, so run it.** It asks whether the application
  reconnects on its own when it already knows where the data file went, which depends on the
  remembered location being readable. The outcome-first list leaves where that location lives free,
  and names this scaffold's `USysLocalSetting` as one worked answer — the answer this route takes.
  A build that wrote the path into the code instead could not perform check 9; this one can, so
  recording it as anything other than run is wrong here.
