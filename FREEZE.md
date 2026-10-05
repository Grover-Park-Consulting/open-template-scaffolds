# Human-Facing Text Freeze

**Who reads this:** any AI assistant working in this repository, before editing any of the files
listed below. **Anyone else:** this records a standing rule George put in place on the library's
own wording — read it to see what it means for text you might otherwise expect to be touched
freely.

## The rule

The files listed under **Frozen files** carry a soft freeze, set by George on 2026-08-26 after a
file-by-file readability pass. His own words at the time: *"I am very happy with the current
human-facing language we are using now. I believe, in fact, that we should freeze these files and
only allow revisions if you or I find an error that has to be corrected."*

**What this prohibits:** rewording, tightening, "improving," or retroactively restyling any of the
text in these files. If a style decision is made later — a better term, a cleaner sentence shape, a
new convention — **it does not travel backwards into these files.** Do not propose an edit to a
frozen file on the grounds that it now reads better, is more consistent with something else, or
could be tightened.

**What this allows:** correcting an outright error. George: *"we reserve the right to correct an
outright error if one ever surfaces."* The reason the freeze exists at all — in his words, *"It
means I can't have a brain fart and change something tomorrow"* — binds him as much as it binds
any assistant. A passing remark is not a lift. **Get explicit sign-off on the specific correction,
with the error stated, before editing.** Do not infer approval from silence, from the shape of a
task you were given, or from the fact that the fix seems obviously right.

**The line between an error and a preference:**

- **Error — worth raising.** Broken markup that changes what renders. An instruction that sends a
  reader somewhere wrong. A claim the files themselves don't support. A name that contradicts the
  same thing's name elsewhere. A cross-reference pointing at the wrong step or section. A typo.
- **Preference — do not raise it.** Shorter phrasing. A word you'd have chosen differently.
  Restructuring a paragraph. A voice inconsistency no reader would actually trip over.

**Two byte-identical blocks live inside the frozen set** — a fix to either must be applied to every
copy in the same pass, or not at all:

1. The shared header block across the eight `standards/` files, from `**Who reads this:**` through
   the closing line that names `README.md` as listing all seven — `standards/README.md` carries the
   same block minus that closing line, which would otherwise link to itself.
2. Everything from `### Instructions to the AI` to end of file, in both
   `prompts/BuildNewTables-StartHere.md` and `examples/northwind-stocktake/prompt.md`.

## Frozen files

The 2026-08-26 pass covered these 18 files:

- `README.md`
- `WELCOME.md`
- `standards/audit-columns.md`
- `standards/design-principles.md`
- `standards/error-handling.md`
- `standards/form-conventions.md`
- `standards/naming-conventions.md`
- `standards/query-style.md`
- `standards/startup-conventions.md`
- `standards/README.md`
- `templates/_template-schema.md`
- `prompts/BuildNewTables-StartHere.md`
- `mcp-server/README.md`
- `mcp-server/server.py` — name lines only, not the whole file
- `mcp-server/setup.ps1` — name lines only, not the whole file
- `templates/errors/error-logging-scaffold.md`
- `examples/northwind-stocktake/prompt.md`
- `examples/northwind-stocktake/output.md`

**`CLAUDE.md` is not in this list.** It is project instructions, not part of the readability pass,
and has never needed a lift.

**The three `templates/audit/` files are not in this list either**, and whether they should be was
never settled explicitly — one lift was granted for them anyway on 2026-09-07, on the asker's own
reading that the freeze covered them. Treat them as outside the freeze unless George says otherwise
for a specific edit.

## Lifts already granted

Every lift so far was granted for a specific, stated edit — never as a standing exception. Recorded
so an assistant doesn't mistake old grounds for new permission:

- 2026-09-04 (×2): one addition to `templates/_template-schema.md` §1, one addition to
  `standards/error-handling.md` (a new closing section). Additions only, not rewording.
- 2026-09-07: `templates/audit/audit-logging-lite-outcome-first.md` and
  `audit-logging-lite-schema.md` — the Replication ID design, including a genuine reword of two
  entries, not just an addition. (These two files sit inside the unsettled audit-templates question
  above.)
- 2026-09-21 (×3 across two sessions same day): a `README.md` table-row rewrite (whole-row replace,
  justified by the row actively misleading the reader — not by preference); an addition to
  `templates/errors/error-logging-scaffold.md`; and a whole-row rewrite of
  `templates/_template-schema.md` §2's `warnings` row, granted because the *existing* wording was
  shown to cause a real misreading, not because a better wording was proposed.

- 2026-10-02 (×1): `standards/error-handling.md`. The "Substituting your own logger" instructions
  directed the very merge that produced a handler whose `Resume Cleanup` never ran (a re-raising
  host handler never returns). Replaced with a logger-or-handler test and a section on accepting a
  host's own handler, plus one sentence in "What a conforming build looks like". A correction of
  an error, not a rewording.

- 2026-10-02 (×1): `templates/_template-schema.md` §4.2 item 8. The wording "refused, not silently
  truncated" hid a route-dependent behavior: a bound form or recordset refuses an over-long value, and
  a SQL `INSERT` stores the first n characters without an error, which a build record showed passing
  the check on one route and failing it on the other. One sentence added, naming the route; the same
  sentence added to the five schema checks that repeat the rule. An addition, not a rewording.

- 2026-10-02 (×1): `standards/error-handling.md`, "Line numbering". Two bullets added: number only
  statements that can fail (a numbered declaration makes `Erl` report the declaration's number,
  shown by a probe), and read `Erl` on the handler's first statement (George's own long
  experience). The standard said nothing on either, and a build numbered a `Dim`. An addition, not a
  rewording.

- 2026-10-05 (×1): `standards/error-handling.md`, "Transaction guard" and conformance check 8. The
  stated reason (writes through `CurrentDb` escape a transaction begun on `Workspaces(0)`) was
  contradicted by a three-repeat test on 2026-10-04: coverage follows the workspace, not the
  `Database` object. Reason rewritten as an observed account; the `ws.Databases(0)` instruction
  stands. The same claim was corrected in `templates/stocktakescan/stocktake-scan-scaffold.md` and
  `stocktake-scan-outcome-first.md`, neither of which is frozen. A correction of an error, not a
  rewording.

- 2026-10-05 (×1): `README.md`, the Access MCP server paragraph. One addition: a sixth, optional
  ability (opening an existing database with its startup skipped), disclosed because the library now
  relies on it by default and a server without it falls back to asking the developer. The five
  required abilities are unchanged. The matching sentence was added to `CLAUDE.md`, which is not
  frozen. An addition, not a rewording.

- 2026-10-05 (×1): `standards/error-handling.md`, "Line numbering". One bullet added: the
  `Cleanup:` and `errHandler:` lines may be left unnumbered, and `Erl` then still reports the failing
  line. The section's skeleton numbers those lines and said nothing on leaving them out, while a
  build that left them unnumbered worked. An addition, not a rewording.

**None of these lifts generalize.** A lift granted for one file, one row, or one addition is not
evidence the freeze is loosening generally — get a fresh ruling for the next edit, every time.

## If you think you've found an error in a frozen file

1. State what's wrong and why it's an error, not a preference, using the test above.
2. Show the specific text — quote it, don't summarize it or point at a section number.
3. Wait for explicit approval before editing. Silence, a broad task assignment, or the fix seeming
   obvious are not approval.
