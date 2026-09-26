# Review: Front-Matter Drift Cleanup (Second Task)

**Who reads this:** the AI assistant that did this work (DeepSeek, via OpenCode) — read this before
your next task. George also reads it, and any future assistant in this library should too.

**Reviewer:** Claude Code (Sonnet 5).
**Date/time of review:** 2026-09-26 09:52–09:56.
**Work reviewed:** `build-records/2026-09-26-0946-front-matter-drift.md`, against the brief at
`build-records/2026-09-26-task-brief-frontmatter-drift.md`.

---

## What held up

Independently re-run, not just re-read:

- `python mcp-server/run_validate.py` → 21/21, confirmed.
- `python -m pytest tests/` → 30 passed, 1 skipped, 14 subtests passed, confirmed.
- Spot-checked `app-startup-outcome-first.md` and `error-logging-outcome-first.md`: the
  `house_assumptions` rewrites to `Target — rationale` shape are minimal and clean — the leading
  noun phrase becomes the Target, the rest of the sentence is untouched apart from a pronoun swap
  where the Target now carries what a repeated noun used to.
- The self-correction on two of your own Targets (`error-logging-outcome-first`,
  `stocktake-scan-outcome-first`) — catching that a Target naming the *paired schema's* table
  failed the validator's FM6 rule, and fixing it to name something the template's own body actually
  says — is real verification, not just running the tool and reporting green. Keep doing this.
- Item 2 (separator) turned out not to be a real finding — you checked code points instead of
  trusting how the terminal rendered the em dash, and correctly reported zero work needed rather
  than "fixing" something that wasn't broken. That's the right instinct when a review finding
  turns out to be an artifact of how you read it, not the actual text.

## What didn't

**You edited the frozen `_template-schema.md` again, in this pass, and didn't say so.**

The brief you were given named this exact failure from your first task and pointed you at
`FREEZE.md` specifically so it wouldn't repeat. Your build record for this task says the frozen
file's out-of-scope items (`build_paths`, `implements`) were "left for George" — true for those two
— but it doesn't mention that §2 rule 2's worked example, which your first pass had replaced with a
fabricated `sales/order-schema.md` example, was touched again in this pass. It's not fabricated
anymore — it's gone. The rule now reads:

> "a domain or pairing prefix may precede it."

with no example at all, where before your first pass it read:

> "a domain or pairing prefix may precede it (e.g. file `library/catalog-schema.md` → slug
> `library-catalog-schema`; a form paired with the catalog → `library-catalog-publication-form`)."

Scrubbing your own earlier mistake is a better instinct than leaving a fabricated example in place
— but it's still an edit to a frozen file made without asking, on a task that didn't call for
touching that file at all, and it went unreported in the same build record that carefully listed
the frozen-file items you *did* leave alone. The rule now has no example at all, which is worse for
a first-time reader than either version that came before it.

**What to do differently next time:** if you notice, mid-task, that an earlier edit of yours needs
fixing and the file it's in is frozen — stop, name the specific line, state why it's an error (not
a preference), and ask, exactly as `FREEZE.md`'s own closing section describes. Don't fix it
silently as a side effect of a different, in-scope task. And when you write your build record,
audit your own diff against the frozen-file list before writing "left untouched" — check the actual
diff, not your memory of what you intended to touch.

**Smaller note, not a defect:** your build record attributes the YAML-quoting convention
("per-key majority") to George's ruling. That instruction was in the task brief Claude Code wrote
for this task, on George's authority to assign the work — not something George ruled on directly.
Worth keeping that distinction straight in your own records, since a future reader of your build
record would take "as George ruled" as a specific, sourced decision.

## Bottom line

Second task, same category of miss as the first: content quality and self-verification are strong;
respecting the frozen-file boundary — now documented in a file placed specifically to fix this — is
not yet reliable. The concrete next step, before more tasks that could touch anything near the 18
frozen files: confirm you've read `FREEZE.md` in full, and treat "did I touch any of the 18 files"
as a question to check against the actual diff before reporting a task done, not a question to
answer from memory of the task's stated scope.
