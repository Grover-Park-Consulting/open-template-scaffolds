# Task Brief: Prompt Naming Alignment + Missing-Prompt Check

**Who reads this:** the AI assistant picking up this task (currently DeepSeek, via OpenCode).
**Before starting, re-confirm you've read `FREEZE.md` and your last review**
(`build-records/2026-09-26-claude-review-of-deepseek-coverage-matrix.md`). One file in `prompts/`
— `prompts/BuildNewTables-StartHere.md` — is on the frozen list; every other file in that folder is
not. This task has no reason to touch the frozen one.

## The task has two parts

### Part 1 — naming alignment

From your own consistency review, prompt filenames use different casing/words than the slugs they
correspond to (e.g. `StocktakeScan-OutcomeFirst-StartHere.md` vs slug `stocktake-scan-outcome-first`;
`OfficiatingAssignment-*` vs `officiating-assignment-*`). Since your first task renamed several
slugs, some of these prompt filenames now refer to names that no longer exist at all —
`AssetTracking-OutcomeFirst-StartHere.md` still says `AssetTracking`, but the template it starts is
now `capital-asset-tracking-outcome-first`.

Bring every prompt filename's domain/type words into alignment with the **current** template slug
it corresponds to (read `templates/_COVERAGE.md` for the current, correct set of slugs per domain
and type — it was built for exactly this kind of lookup). Use PascalCase-with-hyphens as the
existing prompt files already do (e.g. `CapitalAssetTracking-OutcomeFirst-StartHere.md`), not the
slug's own kebab-case — match the convention the *other* prompt filenames already use, not the
template slug's literal casing.

**Check every reference to a prompt's old filename after renaming it** — anywhere in `README.md`
(read-only: it's frozen, so if a reference there is now wrong, name it and stop rather than editing
it), `CLAUDE.md`, or any template's own body that names a paired prompt.

### Part 2 — the "missing prompt" finding, check before touching

Your own review flagged: "no scaffold prompt for `asset-tracking` (only outcome-first)." Check this
against `templates/_COVERAGE.md` before doing anything about it. If `asset-tracking` has no
`vba-scaffold` template at all, there is nothing for a scaffold prompt to start, and the finding was
a false positive — report that, and do not create a prompt for a template that doesn't exist.
If some other domain genuinely has a shipped template with no corresponding prompt, name it in your
build record, but **do not write the missing prompt yourself** — that's new human-facing content,
which is a different kind of task than a rename, and George should see what's actually missing
before anyone writes it.

## What you may touch

Files in `prompts/` other than `BuildNewTables-StartHere.md`, by rename. Any file that references a
prompt's old filename, by updating the reference to the new filename — except frozen files, where
you name the stale reference instead of editing it.

## Verification

- `python mcp-server/run_validate.py` → confirm still 21/21 (prompts aren't validated by this tool,
  so this is really confirming you didn't touch anything it does check).
- Grep the repo for every renamed prompt's old filename — confirm nothing still points at it, or if
  something does and it's in a frozen file, that it's named in your build record instead of edited.

## When you're done

Build record as before, in `build-records/`. State plainly, as its own line: whether you touched
any of the 18 files in `FREEZE.md` (should be no), and whether Part 2 found a real gap or confirmed
the original finding was a false positive.
