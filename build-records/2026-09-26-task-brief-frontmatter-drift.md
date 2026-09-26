# Task Brief: Front-Matter Format Drift Cleanup

**Who reads this:** the AI assistant picking up this task (currently DeepSeek, running via
OpenCode). George assigned it after reviewing your first task (the naming-drift pass) with Claude
Code — read `build-records/2026-09-26-claude-review-of-deepseek-naming-pass.md` first if you
haven't; it explains why this brief is more explicit than the last one was.

## The two things that went wrong last time — don't repeat either

1. **You edited `templates/_template-schema.md` without flagging that it's under a text freeze.**
   That file is one of 18 covered by `FREEZE.md`, now checked into the repo root — **read it before
   touching any frozen file.** If this task requires a change to a frozen file, stop and describe
   the specific change and why it's a genuine error (not a preference) instead of making the edit.
   George rules on frozen-file changes before they happen, not after.
2. **You replaced a real, checkable example in that file with a fabricated one** (`sales/order-schema.md`,
   a domain that doesn't exist in this library). Every example you write anywhere in this library —
   not just in frozen files — should point at something a reader can actually open and verify. If you
   need an illustrative example and no real file fits, say so and ask, rather than inventing one that
   looks real but isn't.

## What good looked like from your first pass, worth repeating

- Validator stayed 21/21 after your changes.
- The full `mcp-server` test suite passed.
- No dangling references anywhere to anything you renamed or removed.
- Your direction on stripping organization-specific names (`northwind`, `school-district`) matched
  what George wanted, even though he hadn't told you the reasoning — keep applying that same judgment
  to naming questions that come up in this task.

Confirm the same three things before you report this task done: validator, test suite, and a
sweep for anything left pointing at the old form.

## The task

From your own review (`build-records/2026-09-26-0611-consistency-conformity-review.md`), the
**"Front-matter format drift"** section — fix these three, and only these three, in this pass:

1. **`house_assumptions` entries are unevenly formed.** The schema (§2, *not* frozen for this
   purpose — front-matter *content* inside template files is not the same thing as the frozen
   *rule text* in `_template-schema.md` itself; you are not editing the schema file for this item)
   requires the shape `Target — rationale`. Bring every `house_assumptions` entry across all 21
   templates into that shape: a leading Target, then an em dash (`—`), then the rationale. You
   already found one instance where the Target is missing (`app-startup-outcome-first`'s first
   entry begins "The locations of any folders…" with no Target named). Find the rest.
2. **Separator inconsistency.** Some entries use ` — ` (em dash), others ` - ` (hyphen). Standardize
   on the em dash, matching the documented convention.
3. **YAML quoting inconsistency.** Some `house_assumptions`/`warnings`/`related` entries are quoted
   strings, others unquoted multiline blocks. Pick the convention that's already dominant across the
   21 templates (check which one most already use) and bring the minority into line with it — don't
   invent a new convention.

**Explicitly out of scope for this pass** — found in the same review, left for later:

- The undocumented `build_paths:` key (needs a schema-file change, which touches a frozen file —
  hold this one for George to rule on directly).
- The `implements` key's under-documentation in §2 (same reason).
- Standards-gate scope, template-family completeness matrix, `status: draft` metadata, prompt
  naming/coverage — none of these are front-matter formatting; leave them for separate tasks.

## When you're done

Write your own build record the way you did for the first review — author, date/time, scope, what
you changed and why, validator/test results — and put it in `build-records/` at the root, same as
your consistency review and your naming-drift pass. Claude Code will review it the same way this one
was reviewed, and that review will go in `build-records/` too.
