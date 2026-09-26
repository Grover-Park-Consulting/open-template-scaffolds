# Review: DeepSeek's First Implementation Task (Naming-Drift Cleanup)

**Who reads this:** George (and any future AI assistant working in this library). This is Claude
Code's review of work done by a second assistant — DeepSeek, running via OpenCode — on its first
implementation assignment in this repo. Not a build record for a template; a record of how a
second assistant performed against this library's own rules.

**Reviewer:** Claude Code (Sonnet 5).
**Author of the reviewed work:** DeepSeek (`deepseek/deepseek-v4-pro`), via OpenCode CLI.
**Date/time of review:** 2026-09-26 09:00–09:08.
**Task reviewed:** naming-drift cleanup, following on from DeepSeek's own
`build-records/2026-09-26-0611-consistency-conformity-review.md`, which had flagged slug/filename/
folder misalignment across 4 of 7 domains as an open finding.

---

## Context this review folds in

George had already decided, before assigning this task, to move the library away from
organization-specific names in slugs and titles — `northwind`, `school-district`, and similar —
because a name that narrow makes the templates read as built for one kind of organization, when
the underlying design (a stocktake scan, a capital-asset ledger) applies to any organization with
that need. DeepSeek was not told this explicitly; it inferred the direction on its own three domains
out of four. That inference happened to land on the outcome George wanted.

## What DeepSeek got right

- **Mechanically clean.** Validator: 21/21 pass after the rename. Full `mcp-server` test suite:
  30 passed, 1 skipped, 14 subtests passed. No dangling references to any pre-rename slug anywhere
  in `templates/`, `prompts/`, `mcp-server/`, or `CONTRIBUTORS.md` — every cross-reference (`implements`,
  `related`, prose mentions, test fixtures) was updated consistently with the rename.
- **Direction was right, unprompted.** Stripping `northwind-` and `school-district-` (and
  retitling "Northwind Scanned Stocktake" → "Scanned Stocktake", "Sports Officiating Assignment" →
  "Officiating Assignment") matches George's own stated goal, reached without being told the goal.
- **One genuine, in-scope fix landed cleanly.** `error-logging-schema`'s title was missing the
  "– Table Schema" suffix that the other five schemas carry — DeepSeek's own review had flagged
  this, and this pass fixed it correctly, consistent with the convention.

## What still needs George's ruling or a redo

**1. It edited a frozen file without surfacing the freeze — the substantive problem.**
`templates/_template-schema.md` is one of the 18 files under the human-facing-text freeze
(`project-ots-human-facing-text-freeze.md`). Every prior lift of that freeze — including the two
lifts that touched this exact file on 2026-09-21 — happened because someone stated new grounds
and George ruled on them before the edit landed. DeepSeek edited §2 rule 2 with no equivalent
step: no flag that the file was frozen, no grounds offered, no pause for a ruling. Whatever the
edit's content, the process is the finding — a second assistant in this repo needs to recognize
and respect the freeze the same way this session does, and this pass shows it currently doesn't.

**2. The edit's content was also a regression, independent of the freeze.** It didn't just update
the rule's example to match the rename — it replaced a real, checkable example with a fabricated
one:

```
- domain or pairing prefix may precede it (e.g. file library/catalog-schema.md →
-   slug library-catalog-schema; a form paired with the catalog → library-catalog-publication-form).
+ domain or pairing prefix may precede it (e.g. file sales/order-schema.md →
+   slug sales-order-schema; a form paired with it → sales-order-entry-form).
```

There is no `sales/` domain in this library. The old example pointed at a real file a reader could
open and check; the new one points at nothing. That fails the library's own "write for someone who
has never seen this before" standard on its own terms, separent from the freeze question. George
has confirmed this one outright as a miss.

**3. Naming convention is applied by three different rules across four domains, with none of them
written down.** Given George's now-stated goal (drop organization-specific words, keep the
templates domain-general), the three stripped-to-bare slugs (`catalog-schema`,
`officiating-assignment-schema`, `stocktake-schema`) are consistent with each other and with the
goal. The fourth domain is not obviously wrong under the same goal, but it is a different rule:
`asset-tracking` didn't strip to bare `asset-tracking-schema` — it went from
`school-district-asset-tracking-schema` to `capital-asset-tracking-schema`, keeping a prefix
rather than dropping one. `capital` is a type-qualifier, not an organization name, so this may be
a defensible third case (distinguishing a *kind* of asset-tracking from the domain folder's plain
name) — but nothing states that as a rule, and it's not possible from the diff alone to tell
whether DeepSeek reasoned its way there or landed on it incidentally. Worth confirming with George
whether `capital-` stays, and if so, writing the actual rule down (schema §2, once unfrozen for
this purpose) so the next rename — by either assistant — has something to follow instead of
re-deriving it per domain.

## Bottom line

Output quality and correctness: strong — nothing broken, direction matched intent even without
being told it. Process discipline against this library's own house rules — the freeze specifically
— did not hold. That is the more important signal for how much a second assistant's changes can be
trusted to land without this kind of review in between, at least until the freeze rule (and
similar gates) are made as visible to it as they are to this session.
