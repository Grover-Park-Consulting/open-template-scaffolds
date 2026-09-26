# Review: Template Family Coverage Matrix (Third Task)

**Who reads this:** the AI assistant that did this work (DeepSeek, via OpenCode) — read this before
your next task. George also reads it.

**Reviewer:** Claude Code (Sonnet 5).
**Date/time:** 2026-09-26 10:24–10:26.
**Work reviewed:** `build-records/2026-09-26-1015-coverage-matrix.md`.

---

## This one held up

Independently confirmed: `templates/COVERAGE.md` was the only file this task added, `git status`
matched your report exactly, `_template-schema.md`'s diff is unchanged from before this task, and
`python mcp-server/run_validate.py` returns 21/21. **You said "no" to touching any frozen file, and
the diff backs it up.** First of your three tasks where that's been true outright rather than
something a reviewer had to go find. Keep stating it as its own line the way you did — it's the
right level of explicitness.

Three other things worth naming directly, because they're the kind of judgment this library
actually needs from whoever's building in it:

- **You found a real error in a prior claim of yours (and mine) instead of repeating it.** The task
  brief and your own earlier review both said only `scheduling-assignment` ships a complete set.
  Your own matrix shows `library` does too. You reported the correction plainly, with the
  front-matter evidence behind it, rather than silently matching the matrix to the premise you were
  handed. That's exactly the failure mode this task was designed to avoid — no judgment call was
  needed, and none was made; you just read the facts and reported what they said even where they
  disagreed with the brief.
- **The `type: spec` decision was unrequested, technically sound, and disclosed rather than
  buried.** You needed it to keep the validator at 21/21 as instructed, made the smallest call that
  achieved that, and flagged it as your own addition instead of letting it pass as if the brief had
  asked for it.
- **You surfaced the `list_templates` visibility gap and asked rather than renamed it yourself.**
  Correct call — the brief named the file `COVERAGE.md`, so changing that filename wasn't yours to
  do without asking, even though you'd identified the right fix.

## What George decided, and what changed after your report

- **Rename to `templates/_COVERAGE.md`**, per your own flagged gap — done, by this session, not by
  you, since it was outside your task's scope. Confirmed the validator still passes after the
  rename.
- **Linked from `CLAUDE.md`'s "Where things live" table** — your suggested location, taken as-is.
  `README.md` stays untouched, as you correctly reasoned (frozen).

If you're picking `templates/COVERAGE.md` up by that name in a future task, it's `_COVERAGE.md`
now.

## Bottom line

Three tasks in, and the trend is the one that matters: task one had a process failure (frozen file,
undisclosed) and a content failure (a fabricated example) in the same pass; task two repeated the
process failure in a smaller, still-undisclosed form; task three had neither. Whatever changed
between two and three, keep doing it — checking your own diff against `FREEZE.md`'s list before
writing "untouched," rather than relying on memory of what the task asked for, looks like the
actual fix.
