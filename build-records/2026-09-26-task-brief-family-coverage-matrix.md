# Task Brief: Template Family Coverage Matrix

**Who reads this:** the AI assistant picking up this task (currently DeepSeek, via OpenCode).
**Before anything else, read `build-records/2026-09-26-claude-review-of-deepseek-frontmatter-pass.md`
in full, and re-read `FREEZE.md`.** That review covers your last two tasks and both times found the
same problem: an edit to `templates/_template-schema.md`, a frozen file, made without asking — the
second time as an undisclosed side effect of a task that didn't call for touching it at all. This
task gives you no reason to open that file, or any of the other 17 frozen files. If partway through
you find yourself about to touch one of them, stop and ask instead of proceeding.

## The task

From your own consistency review (`build-records/2026-09-26-0611-consistency-conformity-review.md`),
the **"Template families are uneven"** finding: only `scheduling-assignment` ships a full set
(schema + outcome-first + scaffold + form); every other domain is missing one or more pieces, and
nothing states what actually exists.

**Build a coverage matrix — a new file, not an edit to any existing one.** Create
`templates/COVERAGE.md`. For each of the 7 domains, list which of the four template types
(table-schema, vba-scaffold, form-spec, outcome-first) currently exist, and which don't. A simple
table is enough:

| Domain | table-schema | outcome-first | vba-scaffold | form-spec |
|---|---|---|---|---|
| ... | ... | ... | ... | ... |

This is a **factual inventory, not a judgment call.** Don't decide whether a missing piece *should*
exist, don't propose filling any gap, and don't editorialize about which domains are "more
complete." State what's there and what isn't. If a domain has a template type under a name that
doesn't obviously match the type (check `type:` in each template's front matter, not just the
filename), use the front matter, not the filename, to decide which column it belongs in.

**What you may touch:** nothing except the new `templates/COVERAGE.md` file itself. No template
file, no test file, no `CLAUDE.md`, no `README.md`. If you think the matrix should be linked from
somewhere else in the repo, say so in your build record and let George decide — don't add the link
yourself.

## Verification

- `python mcp-server/run_validate.py` → confirm still 21/21 (you shouldn't have touched anything
  the validator checks, so this is a check that you didn't).
- `git status` / `git diff --stat` → confirm the only change is the new file.

## When you're done

Write your build record as before, in `build-records/`, dated and labeled. State explicitly, as its
own line, whether you touched any of the 18 files in `FREEZE.md` — "no" should be true this time,
and say so plainly rather than leaving it to be inferred from a file list.
