---
step: 11-checks
title: "Run the checks"
platform_facts: [mcp-run-procedure, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

The checklist this step runs is in `audit-logging-lite-outcome-first`, under "How you validate the
template's output". Fetch it from that template through the template library MCP server (its master,
or the step its `parts` list names) before running anything.

## Validating the build

**To the AI assistant.** This template and `audit-logging-lite-outcome-first` promise the identical
result — the same audit trail, produced by a different route — and each says so where it names the
other. That promise is what gets checked, not which route built it, so there is one checklist for
both rather than two.

**Run every entry under `audit-logging-lite-outcome-first.md`'s "How you validate the template's
output" against this build, on a copy, exactly as that template requires.** Do this whether you
generated the code yourself or handed the developer the files to run — the checks read the database
this build produced, not the procedures that produced it. Report against that same numbered list in
the build record: one entry per check, what was done and what was observed. Passed and not passed are
the only outcomes, including where the first method to run a check hits an obstacle — see
`_template-schema.md` §12.2 for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every
entry opens with, and what to do before settling for a soft result.

Two things follow from this being working code rather than a generated route:

- **The checks that distinguish a Data-Macro recording from one that runs in VBA** — *The recording
  happens entirely in the Data Macro, not in code the macro calls* and *Auditing cannot be got
  around by moving the database somewhere Access does not trust* — apply here exactly as written.
  Complete, working VBA is not an exemption from either: it is this scaffold's `Build*` functions
  that write the XML the table loads, and the check confirms what actually landed on the table, not
  what the generator intended.
- **Where a check's wording assumes house_assumptions the developer changed** — the identity source,
  chiefly — read the check against whichever function this build actually used. A check that says
  "your name" means whoever `AuditUser()`, the host's own identity function, or `CurrentUser()`
  resolves to, not a specific one of the three.
