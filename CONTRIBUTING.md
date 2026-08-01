# Contributing to Open Template Scaffolds

## How this works

Welcome, and thanks for thinking about contributing. We're building this library together,
and a little shared care keeps it useful for everyone who adopts it.

The library is **fork-and-own**: you're free to fork it, swap the `standards/` layer for your
own house rules, and never look back; there's no ongoing dependency on us. But many of us will also
want to *contribute back* a template we found useful, and that's what this document is about.

A quick note on roles:

- **You / the contributor** — the person submitting a template. Every "you" here means this person.
- **We / the library** — the curator-maintainer (George Hepworth today; hopefully a small team later).
- **The adopter** — the colleague who forks and uses a published template. The adopter is never "you."

Most of us will wear more than one of these hats over time — contributing a template one week
and adopting someone else's the next. That double role is exactly why the guidance below leans
toward keeping each published template as house-neutral as it can be: the more neutral it is, the
more readily the next colleague drops it into their own practice.

## What you can contribute

Right now: **table-schema templates** — structured, AI-readable definitions of a domain's tables,
relationships, and business rules. New business domains are especially welcome; the whole point of
the library is breadth of domain coverage that an adopter can lean on at project kickoff.

Two other template types — **`vba-scaffold`** and **`form-spec`** — aren't open for contribution
yet. Their formats are still being proven by hand, and we don't want submissions built against an
unproven format. A *complete* library will carry all three types, so they're firmly on the
roadmap — just not ready to accept yet.

## How a template must be structured

Every contributed template must:

- **Conform to the canonical format** in [`templates/_template-schema.md`](templates/_template-schema.md)
  and **pass `validate()`**. That spec is the contract; if a template passes, it's structurally
  correct by definition.
- **Live at `templates/<domain>/<name>.md`** — a kebab-case filename matching the template's
  `template` slug, under its domain folder.
- **Include the required sections** the format calls for (Intent, Standards Layer, Extra Options,
  plus the table-schema sections: Prerequisites, Entities, Relationships, Business Rules).
- **Be genericized** — no client-confidential or proprietary schema. Use a recognizable public
  domain or invented data. (Our own first template took a real engagement and recast it onto the
  public Northwind sample; that's the discipline.)
- **Be written for a reader who has never seen the system it describes.** Every term is defined
  where it first appears, or it isn't used — no undefined jargon, no coined label the reader has
  to look up. Prefer the plain name for a thing over the internal one. Where a rule has a visible
  consequence, state the consequence too: what the reader will actually see happen, not only the
  rule. If a sentence would send a first-time reader to a search engine or to another file, rewrite
  it. This is a review criterion, not a preference — real testing has repeatedly shown that where a
  template states a rule without its consequence, the reader concludes the system is broken.

`validate()` is **format-only**: it checks that a template is internally well-formed. It does not
open any database, and it does not judge whether a template *fits* a given application — that
always stays the adopter's call.

## Versioning a change

Every template carries a `version` in its front matter — a [semantic version](https://semver.org)
(`major.minor.patch`), counted per template and unrelated to any version number the library as a
whole might have.

**Change what a template produces, and bump its `version` in the same commit.**

- **Patch** (`0.1.0` → `0.1.1`) — a correction that doesn't change what an adopter has to do: a
  fixed typo in generated code, a clarified rule, a warning made plainer.
- **Minor** (`0.1.0` → `0.2.0`) — anything an adopter would want to know about: a new or renamed
  field, a changed default, an added business rule, a new required build step, a new section.
- **Major** (`0.1.0` → `1.0.0`) — a redesign that an existing build can't simply absorb.

Changes that alter nothing an adopter does — reflowing a paragraph, fixing punctuation — need no
bump. When you're unsure, bump the patch; an unnecessary bump costs nothing, and a missing one costs
somebody an afternoon.

**Why we ask for this.** A template is copied out of this library and lived with. Once it has been,
the `version` line is the *only* thing distinguishing the copy somebody is holding from the current
one — the file otherwise looks identical. This library has already paid for that twice: two separate
rounds of testing re-reported a bug that had been fixed weeks earlier, because the stale copy under
test and the corrected one in the repository both read `version: 0.1.0, status: draft`, and nothing
on either page said otherwise. The bump is one line and it is the whole defense.

The `status` field (`draft` / `review` / `stable`) is about maturity, not change, and moves on its
own schedule. It is not a substitute for a version bump.

## Handling house-specific bits

A little care with *house-specific* assumptions pays off: the more neutral each published template,
the more readily the next colleague can adopt it. This isn't a gate to clear — it's how we keep
each other's templates easy to use.

**Send these as-is**

- **Domain content** — anything any competent practice would model the same way. That's the point
  of a template; send it freely.
- **Rendering choices** — the naming style, prefixes, and type encodings a template has to commit
  to just to exist. Fine to include — just call them out in your `## Standards Layer` section so an
  adopter knows their AI can regenerate them under their own conventions. Forget to add them and there's no
  harm — we'll add the one-line note together.

**When your template carries something specific to your house** — a *modeling* decision where
another practice would reasonably do it differently (not just rename it) — let's give it a good
home, picking whichever of the following options keeps the shared template most adoptable:

1. **Private** — if your template doesn't really need it, keep it in your own copy. No need to
   send it upstream; it reaches no adopter.
2. **Optional** — if your variant needs it but the base template stands fine without it, park it in
   `## Extra Options` as a named, opt-in extra. It travels with the template but stays dormant
   until a colleague wants it.
3. **Declared** — if it's genuinely load-bearing and can't be lifted out, keep it and name it in
   the `house_assumptions:` front-matter list. It ships, and adopters see it coming and can keep it as is or
   rework it to fit their own house-specific standards.

**The one thing worth catching together** — a house *preference* dressed up as *domain content*. It
slips in looking universal, and then a colleague who swaps in their own standards quietly gets
output that doesn't fit. If we spot one in review we'll flag it and sort out which home above it
belongs in. Nothing gets "rejected" — it just gets another pass so it lands well for everyone
downstream.

**Why this order** — the three homes differ by how much a colleague inherits when
they adopt a template:
- Private: the adopter doesn't see it at all.
- Optional: the adopter chooses only the parts that fit their goal; leaving the rest out costs the base template nothing.
- Declared: the assumption is built into how the template works, so the adopter inherits it by necessity.

Reaching for the lowest tier that fits keeps every published template as close to house-neutral as
it can be. Our goal is that someone else's template feels like it was written for *your* shop.

**A named tool is a house-specific bit too.** The add-in you run over your code, the utility that
does the tedious part — naming any of them turns a convenience into what reads as a requirement, and
a colleague who doesn't have it reads a door closed to them. Describe what it does instead: "a tool
that adds line numbers to your code", not the name of the one on your machine. This covers everything
that ships, including the standards layer in this repo; the copy you keep for your own shop is yours
and may name whatever you like. Easy to miss and easy to fix — if it slips through, we'll catch it
together in review.

## Sibling templates, and the right to consolidate

As the library grows, more than one of us will encounter the same kind of task — exporting Access
data to a text file or a spreadsheet, say, where there are easily a dozen sensible ways to organize
the work. When submitted templates turn out to be **siblings** — same job, different cut — the library may
fold them into a single, stronger template rather than carry overlapping variants.
We reserve that right as curators (George today; a small team down the road), because the value
to an adopter is *one* clear, well-shaped answer per task, not a maze of almost-the-same options.

**Help us spot siblings.** If you think your template might overlap with an existing one, perhaps a
different take on the same task, say so when you submit it: a line in the pull request, or a note
in the template's `## Intent`. Flagging a possible sibling yourself spares everyone the duplicate-hunting.

**Provide useful enhancements or variations.** If you spot a template for which you have an enhancement
in your own library, or for which you have a different implementation, please share it and indicate which
one you want to contribute to. As above, include a line in the pull request or a note in the template's `## Intent`.

If you'd like to take a hand in shaping a merged result, we'd genuinely welcome it — you know your variant's
reasoning better than anyone. If you'd rather not spend the effort, that's fine too: the
consolidation goes ahead, and your contribution **keeps its credit** in the template's history
either way.

## How review works

Two gates, in this order:

1. **Automated** — CI runs `validate()` on the templates your pull request touches. A red check
   blocks merge. Because `validate()` is format-only, this runs without any database and without
   secrets.
2. **Human** — we review for domain correctness, clean standards separation, no proprietary data,
   and fit with the rest of the library. Format is the machine's job; judgment is ours.

Our ethos is **revise, don't reject**. If something needs another pass, we'll say what and why and
work it out with you — the goal is to get your template landing well, not to turn submissions away.

For now there's **one maintainer** (George Hepworth). Trusted reviewers will be invited to join later as the
library grows, but that's a future bonus, not a promise.

## Submitting

- Open a **pull request** against the canonical repository:
  [`github.com/Grover-Park-Consulting/open-template-scaffolds`](https://github.com/Grover-Park-Consulting/open-template-scaffolds).
- **Sign off your commits** with `git commit -s` — the [Developer Certificate of Origin](https://developercertificate.org/)
  sign-off, affirming you have the right to contribute the work.
- By submitting, you agree your contribution is licensed under the project's **MIT license**
  (inbound = outbound).
- If your template might be a **variant** of an existing one, flag it (see *Sibling templates* above).
- A **worked example** — a prompt plus the output it produces — is welcome but **never required**.
  What we're after is a *workable template* (a replicable scaffold), not a *working example* (a
  one-off demo). The template itself is the deliverable.

Thanks for helping build something the next developer can lean on.
