# Template Maturity Criteria

**Who reads this:** the AI assistant, deciding or checking a template's `status:` value; and George,
ruling on any case this file doesn't settle on its own.

**What this decides:** the difference between `status: draft`, `status: review`, and `status: stable`
in a template's front-matter (`_template-schema.md` §2). A template's structure passing `validate` is
not this — `validate` checks the document is well-formed; this file decides whether the thing it
describes has actually been built and shown to work.

## 1. The currency rule — read this first, it governs everything below

**A build record only counts as evidence for a template's current status if it verifies the template
as it exists now — not as it existed when the record was written.** Templates get revised: a
checklist gets added, a section gets rewritten, a defect gets fixed. A record written before a change
does not become evidence for the template after that change, no matter how thorough the record was at
the time.

**Concretely, for a template with a `## Validating the build` checklist:** a build record only counts
if it tested *that* checklist — the same items, the same count, as the checklist currently in the
template file. If the checklist has changed since the record was written — including a checklist
being added where none existed before — the record does not count. It does not count as a pass, and
it does not count as a fail either. It counts as **no current evidence**, a status distinct from
both, and it says nothing about whether the template would pass today. Find out by running the
current checklist, not by assuming.

## 2. The two branches

**If the template's own body has a `## Validating the build` checklist** (numbered checks, results
reported per `_template-schema.md` §12.2 as `Result: PASSED` / `Result: NOT PASSED`):

- `stable` requires at least one *current* (§1) build record showing every checklist item
  `Result: PASSED`. No partial credit, no inference from unrelated work going well.

**If the template genuinely has no such checklist today** (not: had none when it was last built, but
has one now — that case is §1, "no current evidence"):

- `stable` requires at least one *current* build record that explicitly reports the build finished
  without errors, against the template as it exists now.

## 3. A currently open defect disqualifies on its own

A checklist item reported `Result: NOT PASSED` and never re-verified after a fix keeps the template at
`draft`, regardless of how much other work has happened on the template since, and regardless of
whether that other work looks like it would plausibly have fixed the same thing. The defect gets its
own fix and its own re-verification. Don't fold a known failure into a general sense that the
template has improved.

## 4. `draft` vs `review`

- `draft`: the default. No qualifying evidence yet, under §1–3.
- `review`: `validate` passes and at least one build record exists, but it doesn't meet the `stable`
  bar under §2 — either because the only evidence is stale (§1), or the evidence is ambiguous (a
  checklist item whose result isn't a clean pass or fail — report this case rather than rounding it
  to one side). **A known, currently unresolved `NOT PASSED` item does not qualify a template for
  `review`** — §3 governs that case and it stays `draft` until fixed and re-verified.
- `stable`: meets §2 in full, with no open disqualifying defect under §3.

## 5. Recording the check

**Note the date a template's status was last determined, in the template itself** — a line reading
`Status last determined: YYYY-MM-DD` near its `## Validating the build` section (or, for a
`table-schema` template with no checklist, near the top of the file). Put the date there, not a
pointer to the specific build record used as evidence: `build-records/<slug>/` is a developer's own,
untracked local history (see `build-records/README.md`) — the file it names may not exist for the
next person who opens this template, on a clone that has never run a build of its own. The date
alone is what travels with the repo and stays true for everyone who reads it.

A `status: stable` (or `review`) claim with no date attached is exactly the kind of unverifiable
claim this file exists to prevent.
