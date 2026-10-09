---
template: _method
title: Open Template Scaffolds — Method (how a run is conducted)
domain: _meta
type: spec
version: 0.1.0
status: draft
---

# Method — how a run is conducted

**Who reads this:** the AI assistant running a template, and contributors maintaining the library.

One home for method, the fourth kind of knowledge beside domain decisions (templates), house choices
(`standards/`) and platform facts (`_materialization.md`). Each section carries a `<!-- method: id -->`
marker. The template library MCP server chooses the sections a run needs from what the template
contains and the route (design or build), and delivers each where it is used: `get_method` at the
opening, the master (`get_template`) with what the template's type adds, the part holding a
`## Wizard` with the wizard method, and the `build-method` part on the build route
(`get_method` with `stage="build"` where no template fits). The text was moved
here verbatim on 2026-10-06 from `CLAUDE.md`, `_template-schema.md` and `_materialization.md`;
section numbers such as §10.4 inside it refer to where the text came from.

## Contents

- `run-opening` — delivered every run
- `design-review` — delivered every run
- `house-assumptions-and-warnings` — delivered every run
- `between-questions` — delivered every run
- `checklist-rule` — delivered every run
- `build-record` — delivered every run
- `related-after-finish` — delivered when the template has `related`
- `wizard` — delivered with the part that holds the template's `## Wizard`
- `explore-options` — delivered `type: outcome-first`
- `staged-procedures` — delivered `type: vba-scaffold`
- `design-only-handover` — delivered design route
- `build-route` — delivered build route
- `access-gate` — delivered build route
- `quiet-build` — delivered build route
- `runbook` — delivered build route
- `build-records-accumulate` — delivered build route

## run-opening
<!-- method: run-opening -->

*Delivered: every run.*

### Before the first question of any run — say what the tool can change, and where this run ends

**The first message the developer reads from you in any run contains the text below, said once,
before the first question.** The first question of any kind is the standards gate's, which every
run asks: say it immediately before the gate. The gate's text arrives as the `standards-gate` part,
or from `get_standards` on a from-scratch run; run the gate from that text.
It applies whatever the run is: a
template with a wizard, a template without one, a from-scratch design, or a plain-words request that
has matched nothing yet. This is the only place the text is defined.

> *"How much of my work you see go past depends on the assistant you're using and where you run it; that changes nothing about what you get, or about your decisions arriving as questions."*

**If the start-up form printed "parked on the taskbar",** add, in your own words:

> *"The window you saw at the start waits on your taskbar for the whole run. Click it there, or ask
> me, to read it again."*

**If the start-up form could not run** (no PowerShell, no desktop), also say the substance of `start/start-form-text.md` in your own words, once, in the same message. The form carries the rest of the opening when it runs; this line is the part it cannot carry.

**Why it is said at all.** Rules 2 and 4 of `between-questions` bind your output: a build going to plan
produces nothing between its first message and its last, and nothing needing action is buried in
between. Neither reaches the tool you are running inside, which may narrate on its own account or show
every file as it is written — and the developer cannot tell which of the two is talking. This line does not fix that. It tells them the variation is real, is not a fault, and changes nothing
about the result or about where their decisions are. **The library's entry documents say the same
thing to anyone who reads them; this text exists because two of the four ways a run starts touch no
file at all.**

**Say where this run ends, in the same breath.** The route decides what the developer has at the
end, so they learn it before they answer anything rather than after.

- **An Access MCP server is connected and the developer has not asked for a design only.** Say, in
  your own words, that the run can end either way and that they choose when the design is approved:
  with the approved design only, which they build themselves, or with a build, created and run in a
  copy of their database and checked before they see it. The design only is always open to them;
  the build is open because the tool to do it is connected. The choice is a wizard step (see "After
  approval — building it").
- **No Access MCP server is connected.** The run ends at the approved design. Say so now, in your
  own words, to this effect:

  > *"So you know where this ends up: I can design the whole thing with you, the tables, the fields,
  > how they connect, and the reasoning behind each choice, and you approve it before we're done.
  > What I can't do from here is build it in your database, because I have no way to open it and run
  > the code. You'll do that part yourself, the way you'd build anything else. There are tools that
  > would give me that ability; if you want to know what to look for, ask and I'll tell you."*

  **Say it plainly and without apology. It is a complete outcome, not a degraded one** — a design
  under the standards layer, approved by them, is what a competent developer needs. Do not name a
  product; describe what the missing tool does, so a reader who has none learns what to look for
  rather than that this was not written for them.

- **The developer asked for a design only.** The run is on the design route by their choice, even with
  an Access MCP server connected, and ends at the approved design. Say so now, in your own words, to
  this effect:

  > *"You've asked for the design only, so we finish with the design you approve, and you build from
  > it yourself."*

  It is a complete outcome, not a build cut short; do not offer the build route unless they ask.

**On the design route, the approval question says again where the run ends,** in your own words:
approving the design ends the run, and the developer builds from it. Where no Access MCP server is
connected, add that building it here would need a tool that can open and run things in their
database.

**Why they are told this first.** Without an Access MCP server you cannot run what you wrote, so
nothing you hand over has been executed by anything. A developer who learns that at the end has
spent the whole session believing they were getting tables.


## design-review
<!-- method: design-review -->

*Delivered: every run.*

After the template is loaded:

0. **Fetch the design parts before drafting anything.** The master names them in `parts`: the
   standards the design follows, the platform facts it depends on, and any step marked for the
   design. Follow each answer's `next` until it reaches the build, or ends.
1. **Apply the standards to everything you produce** — naming conventions, audit columns, and the
   error-handling pattern — plus the field-qualification rules (no bare reserved or ambiguous nouns;
   PK = `[Entity]ID`; a FK takes the referenced PK's name). These come from `standards/`, **never**
   from the template body.
2. **Honor the template.** Its entities, fields, relationships, and Business Rules are decisions
   already made. Carry them through unless the developer overrides a specific point.

   **Templates are complementary, not walled off.** A template's own body covers its domain, but a
   mechanism it uses — a Data Macro event, an error-handling pattern, a wizard shape — may already
   be documented more fully in another template, or in a prior build record for this one. Before
   implementing a cross-cutting mechanism from scratch, check `build-records/` (method
   `build-records-accumulate`) and any template whose domain overlaps for how it was already worked out, and
   say so when you draw on one.
3. **Surface every `house_assumptions` entry** the template declares in its front-matter. List them
   and ask the developer to confirm or override before you finalize. **Surface every `warnings`
   entry the same way** — those are conditions the build must act on, not suggestions (e.g. Data
   Macros cannot audit Long Text fields): state each one, get the developer's answer to whatever it
   says must be checked, and branch the build accordingly.
4. **Fold in the developer's specifics** and any named extras from the template's
   `## Extra Options` section.
5. **Don't invent domain content** beyond the template and what the developer supplied. If something
   is genuinely undetermined, ask.
6. **Present two parts for review:** a `mermaid` `erDiagram` (tables, keys, cardinality, and the
   connections into any existing host tables), then field-table detail (`| Field | Type | Key / Req | Purpose & rules |`)
   with indexes, derived values, and the standards-supplied audit columns. It is never final until the
   developer says so.

## house-assumptions-and-warnings
<!-- method: house-assumptions-and-warnings -->

*Delivered: every run.*

**A house assumption is asked, never assumed.** Surfacing one is not asking about it. *"I'll take
that as confirmed unless you say otherwise"* answers on the developer's behalf — it ignores their
input rather than requiring it, and someone who says nothing has agreed to nothing. Put it through
the selection control like any other decision and wait. **Never tell the developer they may skim:**
an expert skims and loses nothing, a newcomer skims and misses the one line that mattered.


## between-questions
<!-- method: between-questions -->

*Delivered: every run.*

**Every question to the developer goes through the interactive selection control** — the one that
renders each option as something they click. That covers every decision, not only wizard steps:
approving the design, confirming a field you proposed, where to save a file. A question written as
prose, a table, or a list they answer by typing promises a choice and delivers an essay question.

### 10.4 What the AI assistant says outside a step

§10.3 governs the step itself. Everything else said during a wizard — before the first question,
between two steps, before the build begins, and while it runs — has no specified shape, and
unspecified space is where ordinary explaining habits reassert themselves. Four rules govern it.
**The first two fix opposite problems, and neither is a rule about being brief.**

Nothing here is written in a template file. These are run-time rules: what the AI assistant says as the
wizard runs, composed in the conversation and never authored anywhere.

**1. Between two steps, name what was recorded and what is being asked next.**

One line for each, in the developer's words:

> *"Errors will go to a table, with a text file as the fallback. Step 4 asks where that table lives."*

Not:

> *"Step 4."*

A step number says where a question sits in a list, not what it is. A developer four questions in has
no other confirmation that the answer they clicked registered. **This rule makes what is said between
steps longer, not shorter** — that is what it is for. Where the wizard branches, or a step is asked
twice, this is where that is said.

**2. Before the first question, and before the build, say only what the developer must act on.**

These are the two moments with the most to report — the template that was matched and why, the
build-wide warnings, the house assumptions, what was found on opening the files — and the least use
for it. At the first, the developer has chosen nothing yet. At the second, they have chosen
everything and are waiting. Say four things at these moments and nothing else:

- Anything they must answer or confirm — asked as a question, never stated in prose (rule 3).
- Anything that changes what they do next.
- The disclosure line (defined in `run-opening`), before the first question only.
- Where this run ends: before the first question, and again when the design is approved. A build in
  their database, or the approved design and nothing further. The method `run-opening` carries the
  wording and the reason. It qualifies under the second item above, since it changes what the
  developer does when the run finishes.

Everything else — what was checked, what was found, what it meant — put in **the build record**.

**This rule does not govern the design presented for approval.** The diagram and field detail are the
deliverable the whole workflow exists to produce; they are not narration, and they are not shortened.
The rule governs the prose around them.

**The build record is always written**, and delivered as a file alongside the artifact. Without it
this rule deletes the detail rather than routing it, and the disclosure line promises something the
format does not keep. **`templates/_materialization.md`, "The build record", defines what it is
called, where it goes, and what belongs in it.**

**3. A house assumption is asked, never assumed.**

A template's `house_assumptions` entries are surfaced before the first question. **Surfacing is not
asking.** *"I'll take that as confirmed unless you say otherwise"* states the assumption and then
answers it on the developer's behalf — it **ignores their input rather than requiring it**, and the
developer who says nothing has not agreed to anything. It is put through the selection control like
any other decision, and the build waits for the answer.

This is the §10.7 trap one level up, and worse: a preferred choice at least appears in a question the
developer is looking at. A self-confirmed assumption appears in prose they were free to skim — and an
expert skims and loses nothing, while a newcomer skims and misses the one line that mattered.

**The disclosure line** is defined in `run-opening`, which is delivered on every run and is the only place its text lives. Say it as that section directs, before the first question.

**Do not tell the developer they may skim.** Knowing which paragraph is safe to skip is what
experience buys: an expert skims and loses nothing, a newcomer skims and misses the one line that
mattered. The disclosure line gives an anchor instead — everything you must act on is in a question.


## checklist-rule
<!-- method: checklist-rule -->

*Delivered: every run.*

**The build-record checklist rule — stated once here, canonical for every template type that carries
a checklist, `vba-scaffold` included. A template of either type points to this rule rather than
restating it.** The build record reports against the template's own checklist section, one entry per
check, each saying what was done and what was observed — a completed check list, not a narrative.
Passed and not passed are the only outcomes; there is no third. A check the template itself marks not
applicable is not an exception: confirm the condition it names (for example, that no seed rows are
declared) and record `Result: PASSED` with what you confirmed. A check tests what the build made,
through the routes the build made. Invalid data is the right probe there; sent through a route the
build neither made nor can close, it shows only what the engine does, and that belongs in
`_materialization.md` as a fact, not in a check. Each entry opens with a literal
`Result: PASSED` or `Result: NOT PASSED`, so the verdict is visible at a glance rather than buried
inside a sentence, followed by what was done and what was observed. A check is passed when every line
under it was observed; an entry that is not passed says what was done and what stopped it. An entry
with neither outcome is a check that was not run, and the record is not complete until it has one.

**An obstacle to one method of running a check is not evidence the check cannot run — it is evidence
that method cannot.** "Partially verified," "verified by review instead," "the structural guarantee
holds," or any other soft middle ground is not a third outcome; it is not passed, full stop, unless a
genuinely different method is tried first. A tool boundary, a locked resource, or a platform default
that blocks the first approach does not license a weaker one in its place — look for a structurally
different method that still exercises the real condition before settling for less. A second,
independent connection to the resource under test, opened separately from whatever route is blocked,
can often force the real condition (a genuine engine-raised error, not an inferred one) even where the
first route a builder reaches for cannot. Exhaust that kind of alternative before a check is recorded
as anything but passed or not passed.

**A check that forces two writers to collide counts only if it was shown able to fail.** Run it once
more on a copy with the protection against two at once switched off, and record what happened. If it
failed, the check stands. If it did not fail, record what stopped it: where the engine's own lock on
the first writer's unsaved rows refused the second writer, the check shows that nothing got between
the two, not which protection stopped it, and it passes on that statement. Where nothing named stopped
it, the method could not have shown a failure, and the check is not passed. Observed 2026-10-05: a
second connection opened inside the same Access process failed this test for two checks of three,
because the second writer read through the first writer's connection, and it can never show a
writer that is held back and then finishes, because both share one thread. A second Access process,
started on a check copy and identified by its process ID as the method `access-gate` requires, gave controls that
failed or were stopped by a named engine lock.


## related-after-finish
<!-- method: related-after-finish -->

*Delivered: when the template has `related`.*

### 7.1 The `related` front-matter key

Some templates solve problems that naturally follow from, or lead into, another template or a
standards file — building `app-startup` leaves an application with nowhere for unhandled errors
to go; building `error-logging` makes little sense without knowing what `error-handling.md`
already decided. `related` is how a template names that, without turning it into route
specification or a design-time input.

**It is advisory, and it is timed.** A `related` entry is never read while a template is being
designed or built, and it never changes what gets built. The AI assistant surfaces it exactly
once, after the build is reported finished — wherever that template's own closing instruction to
the AI assistant already lives (`## To the AI assistant building this` for `outcome-first`; the
equivalent build-completion point for other types). One line per entry: what it is, why it might
be worth considering next. **Not a gate, not a recommendation to act on now** — the developer
either takes it up in a future session or doesn't.

**This is what keeps it from becoming the thing `_template-schema.md` §12.4 and the library's own
"providing extraneous information opens gates best left closed" lesson warn against.** A mechanism
read during design can leak into the build; a mechanism read only after the build is already
finished cannot.

**Maintenance is a review step, not a mechanical rule.** When a new template is added, check
existing templates for a plausible `related` entry pointing at it, and add entries in both
directions deliberately where one makes sense — never automatically, and never on every pair just
because two templates share a domain word. A missing or stale entry is a defect to notice and fix
when found, the same way an undeclared `house_assumptions` entry is; nothing enforces it
mechanically beyond `validate`'s format/existence check (§2, rule 7).

**Write the rationale for the developer reading it, not as a terse instruction to a machine.**
Say why it might be worth their time, in plain, warm prose — "worth adding once X is built: you
now have Y, and this gives Z somewhere to go" reads as a person explaining something; "you now have
Y; consider Z" reads as AI-speak. Where two templates solve a similar-looking problem for different
purposes (e.g. two domains that both reconcile a table of items against barcode scans), say what's
actually shared and say plainly that the mechanism isn't — don't let the resemblance imply a
crossover that isn't there. The `Target — rationale` dash is the one structural exception to
keeping this dash-free; nothing else in the sentence should lean on one.

---


## wizard
<!-- method: wizard -->

*Delivered: when the template has a `## Wizard`.*

**Running a template's wizard.** The OTS Wizard is how a build gets from the developer's request
to an artifact that meets it. The request arrives underdetermined; the template and the standards
layer settle what they already carry, and the wizard settles the remainder, one decision at a
time, before anything is built. It works by interaction, not by formula: you ask, the developer
decides, and the build proceeds on what they settle between them. Some templates write their
questions out in a `## Wizard` section: a short run of one-question steps, each with a plainly
named preferred choice and a *Tell me more* block holding the reasoning and the warnings
(`templates/_template-schema.md` §10). It is named for the Access wizards these developers already
know, but nothing of it runs inside the database: build no form, install nothing to run it, and
leave no artifact behind. Ask the steps yourself, in conversation.

**A wizard of more than three steps opens with the entry question** (§10.6): *"This takes n
questions. Do you want to answer them, or shall I just build it?"* Ask it **even when the
developer's instruction was imperative** — "find a template and run it" is exactly the case it
exists for, and it costs them one click instead of seven questions. `Just build it` still stops at
every step that has no preferred choice, and you state the preferred choices before acting on them.

**Every step names a preferred choice, and this library never calls it a "default"** (§10.7). The
word means two things — *the one we'd point at first*, and *what happens when nobody chooses* — and
in a file you are reading in order to act, the second meaning wins and the question stops being
asked. **A preferred choice becomes the answer only when the developer declines to choose or tells
you to get on with it. Never on your own initiative**, however obvious the answer looks.

**Ask every step through the interactive selection control** — the one that renders each option as
something the developer clicks. **A step written out as prose, a markdown table, or a list they
have to answer by typing is a failed step**, however good its content: it promises a choice and
delivers an essay question. One step per ask, never two.

`Tell me more about <topic>` is **always the last option** — it is the only way the developer can
reach it, so it must be clickable. Choosing it shows the explanation and then **asks the same step
again, unchanged**. Never write "say tell me more if you want…" — that is the failure this rule
exists to prevent. `Go back to the previous question` is an option on every step after the first,
wherever there is room. The control takes at most four options, so a step offers at most three
substantive answers plus *Tell me more*.

Otherwise: name the preferred choice plainly — **never by emphasis, and never with a recommendation
attached** — and **when the developer changes an earlier answer, discard every answer after it**
and resume forward from there.


**Write every question in the developer's world, not the system's.** "Which tables should be
audited?" — not "Which of your tables should the scan consider?" Words like *the scan*, *the
generator*, *the config table* name machinery; to a first-time reader they signal only that they
are out of their depth. **Error numbers, engine limits, and internal names never appear in a
question** — they live in *Tell me more*, where the person who wants them will find them and nobody
else has to. A choice made
against the standards layer holds for the rest of that run and is never silently re-defaulted —
but it is never written back to `standards/` either. A step-1 answer that declines the feature ends
**that wizard only**; ask any other wizard in the same template independently.

**Restate the decision in full at each gate.** Ask the question where the developer can answer it
without reconstructing anything from earlier in the session: what the setting means, what it
produces at run time, and what changes if they choose the other way — at the point of asking, not
on request. Never use one number for two quantities in the same message: if 11 fields are auditable
and 11 macros will be generated, say which is which, or the reader will take them for the same 11.


### 10.3 How a step is asked — the mechanism

**Every step is put to the developer through the interactive selection control** — the one that
renders each option as something they click. **A step rendered as prose, a markdown table, or a
list the developer has to answer by typing is a failed step**, however good its content: it asks
them to compose an answer where they were promised a choice.

What this file holds is the **source** for that control, not the thing shown. The `**Ask:**` line
becomes the question, each option row becomes a clickable option with its short description
underneath, and the `**Preferred:**` line is stated with them.

Four rules follow from it:

- **One step per ask.** Never two steps in one control, even where the second seems to follow.
- **`Tell me more about <topic>` is always the last option**, on every step. There is no other way
  for the developer to reach it — they must be able to click it. Choosing it shows the
  *Tell me more* text and then **asks the same step again, unchanged**, so the explanation costs
  them nothing but a click.
- **`Go back to the previous question` is an option on every step after the first**, wherever there
  is room for it alongside the substantive options and *Tell me more*.
- **The control takes at most four options.** A step therefore carries **at most three substantive
  answers plus *Tell me more***. A decision with more than three natural answers is either two
  decisions, or has two answers that should be one — resolve it in the template. Never resolve it
  by dropping *Tell me more*, and never by silently cutting an option (rule 8).


### 10.5 Rules

1. **One decision per step.** A step that asks two things is two steps.
2. **The `Ask:` line is one short question, in the developer's words.** No clause explaining why it
   is being asked, no naming of the machinery behind it. "Which tables should be audited?" — not
   "Which of your tables should the scan consider for auditing? This is the one boundary decided in
   code, and everything finer-grained is a switch you flip in a table afterwards." Words like *the
   scan*, *the generator*, *the config table*, *the boundary* mean nothing to someone meeting this
   for the first time; what they convey is that they are out of their depth, and the likeliest
   response is to stop using it. **If the question needs a second sentence, that sentence belongs in
   *Tell me more*.**
3. **One name per thing, from the first step to the last.** Once something has been named — a file,
   a folder, a setting, a table, a step — it keeps that name in every question, every option, and
   every *Tell me more*. No synonyms, no switch to the more technical term later, no shortening
   after first use. **A second name for something already named is a defect even when both names
   are correct**: a new word signals a new thing, so the reader stops to work out what the
   difference is and finds none. That pause costs more than the repetition would have. **Where the
   plain name and the precise name compete, use the plain one** — a reader who feels talked down to
   is annoyed and keeps going; a reader who is not sure two words mean one thing has already lost
   the thread, and may not know they lost it. If the precise name is genuinely needed, it replaces
   the plain one from first use. **The exception is a term the developer probably already owns** —
   *referential integrity*, *cascade delete* — **or a plain word already taken by something else in
   this material**, where the paraphrase misdirects rather than merely under-informs. There, name
   both once, plain sentence first and the technical term marked as such: *"nothing in the database
   enforces that reference — the database term for this is referential integrity."* After the
   pairing the plain name carries on alone; pairing is a definition given once, and a synonym
   appearing later is still a defect. Where the plain word is taken, the pairing belongs **in the
   question**, not in *Tell me more* — a reader who does not open *Tell me more* has already taken
   the wrong meaning.
4. **A short description says what the option *is*, in one line — never why it is better.** No
   bolding, no ordering by preference, no "recommended". Every comparison lives in *Tell me more*.
   A description may carry a consequence the developer needs *at the moment of choosing* ("any Data
   Macros those tables already have are replaced"), but never the reasoning behind it.
5. **Error numbers, engine limits, version caveats, and internal names never appear outside *Tell
   me more*.** Someone who meets "error 3870" or "`Application.LoadFromText`" in
   a question they are being asked to answer learns one thing: this was not written for them. Put
   it one click away, where the person who wants it will find it and nobody else has to.
6. **Every step names a preferred choice — never a "default".** See §10.7. The `**Preferred:**`
   line is the only signal a reader gets about which option the library would point at first, and
   it is enough: no bolding, no "(Recommended)", no argument. Where the standards layer answers the
   question, the preferred choice is that answer; where the standards layer is silent, it is the
   template's own and the line says so. It may follow an earlier answer, in which case the line says
   which step it follows.

   **Say where it came from in plain words — never as a file name or a section number.** Rule 5
   forbids an internal name in a question, and a `Preferred:` line is part of the question. So write
   *"the naming style these templates follow"*, not `standards/naming-conventions.md` §1.1. **A line
   that cites a file forces whoever reads it aloud to invent a paraphrase**, and the paraphrase is
   then unreviewed: one such line produced *"from **your** naming conventions"* in a live run —
   claiming the developer had authored a file they had never seen. Give the spoken wording in the
   template and there is nothing to invent.

   **Avoid the possessive entirely.** *"Your standards"* is wrong for anyone who has not adopted a
   layer; *"the house standard"* assumes a house the reader may not have; *"the template's"* is
   inaccurate, since the template follows the layer rather than defining it. *"The standards these
   templates follow"* claims nothing about whose they are.
7. **A confirmation step has no preferred choice.** Where a step asks the developer to attest to
   something rather than to prefer something — that they have a backup, that a list the build will
   act on is correct — write `**Preferred:** none` and say why: nothing the library picks can stand
   in for the developer's own word.
8. **Options are re-ranked, never removed.** A choice the library ranks last is still offered, in
   the same plain form as the others.
9. ***Tell me more* stays closed until asked for** and gives one or two facts that might tip the
   choice — drawn from the standards files and the template's own description, not restated from
   them, and not exhaustive.
10. **Warnings live at the step they belong to.** A front-matter `warnings` entry that governs one
    decision is surfaced inside that step's *Tell me more*; one that governs the whole build is
    surfaced before step 1. This is the point of the format: the warnings are not less visible, they
    are visible where they are actionable.
11. **A choice made against the standards layer holds for that run** — carried forward to every
    later step, never quietly reverted, and never written back to the standards files. The next run
    starts from the standards again. Flexibility within limits.
12. **Going back is always available.** Every step after the first offers it, and the developer may
    name any earlier step at any time. **Changing an answer discards every answer after it** and the
    wizard resumes forward from the changed step — so a revised decision can never leave a stale one
    standing behind it.
13. **A wizard of more than three steps opens with the entry question** (§10.6), which is where the
    developer chooses whether to answer every step or have the preferred choices used. It is never
    an option inside Step 1.
14. **Ending early ends one wizard, not the run.** Where a step-1 answer declines the whole feature,
    that wizard stops; ask any other wizard in the same template independently.
15. **§8.4's facilitation rules apply in full.** Never infer the answer to a step, present one step
    at a time, and restate the decision at the gate so it can be answered without reconstructing
    anything from earlier in the session.


### 10.6 The entry question

**A wizard of more than three steps opens with one question before Step 1**, asked through the same
selection control as every other step:

> **Ask:** This takes *n* questions. Do you want to answer them, or shall I just build it?

| Option | Short description |
|---|---|
| `Ask me the questions` | Go through them one at a time. |
| `Just build it` | I use the preferred choice at each step, and only stop where a step needs something from you. |

**Preferred:** `Ask me the questions`.

It exists because an instruction to proceed — "find a template and run it" — is not permission to
put seven questions in front of someone. The entry question costs them one, and it is the only
place the wizard interposes itself between the instruction and the build.

Six rules govern it:

- **Asked once, before Step 1, and never again.** It is not an option inside Step 1, and no later
  step re-opens it.
- **`Just build it` cannot skip a confirmation step** (rule 7). A step with no preferred choice has
  nothing to fall back on, and passing one silently would answer for the developer on exactly the
  questions they were meant to answer. Say up front how many of those remain.
- **State the preferred choices before acting on them** — the answer being used at each skipped
  step, in a short list. `Just build it` authorizes known answers; it is not consent to be
  surprised.
- **A preferred choice that contradicts what the developer asked for is not a preferred choice on
  that run — ask the step, and say why you are asking it.** Preferred choices are written into a
  template before anyone has said what they want, so a request can arrive that one of them directly
  contradicts. A developer who asks for the feature on the database they already have has ruled out
  the step whose preferred choice builds a set of sample tables to try it on; using it anyway is
  precisely the surprise the rule above forbids. This applies to one step at a time — the rest of
  `Just build it` stands.
- **Ask it even when the developer sounded impatient.** Especially then: an imperative instruction
  is what this question is for, and answering it takes one click.
- ***n* is this template's own count** — the steps in its `## Wizard` section that apply to this
  run, a number the developer could arrive at from the file. A run can turn up questions no
  template carries: a build route where a connected tool offers one, a file that has to be made
  writable first, a step re-asked under the rule above. Don't fold those into *n* and don't try to
  predict them. Ask each where it arises and say it is one more than the number given at the start.


### 10.7 "Preferred choice", not "default" — and why the word matters

**A wizard step names a *preferred choice*. This library does not use the word "default" for it,
anywhere, deliberately.**

"Default" carries two meanings and nothing in the word says which is meant:

- **the choice we would point at first** — a recommendation, which still has to be offered; and
- **what happens when nobody chooses** — a fallback that fires on its own.

Written into a file that an AI reads and acts on, the second meaning wins. A step labelled
`Default:` reads as standing permission to skip the question, and the question stops being asked.
That is not a hypothetical: it is how `standards/error-handling.md` came to say *"emit option 3…
and say so"* and how a build came to pick its own error-handling option and announce the result to
a developer who had never been asked.

**The preferred choice becomes the answer in exactly two situations, and both are an act by the
developer:**

1. They decline to choose — "you pick", "whatever you think".
2. They ask to get on with it — the entry question's *"just build it"* answer (§10.6).

**It never becomes the answer on the AI assistant's initiative.** No amount of obviousness, data
shape, or convenience converts a preferred choice into a decision nobody made.

> **A note for anyone writing a template.** This distinction does not arise when you write code:
> a default parameter value simply *is* the fallback, and no reader expects otherwise. It arises
> the moment your reader is an agent that will act on what you wrote. Words that are precise in a
> function signature turn ambiguous in an instruction, and the ambiguity resolves toward *action*,
> because acting is what the reader is there to do. When in doubt, name the act you want and the
> act you don't.

`validate` does not check §10 at all — the format is proven by hand first, exactly as the three
template types were (see the scope note at the top of this file). `templates/errors/error-logging-scaffold.md`
is the worked example.

**Two different gaps sit inside that, and only one of them closes.** §10.2 and §10.5 describe things
that are in the file — a `### Step n —` heading, an `**Ask:**` line, a `**Preferred:**` line, a
two-column option table, a `<details>` block per step — and a checker could assert every one of them.
**§10.4 cannot be checked here at any point**, because nothing it governs is in a file: the line
between two steps is composed in the conversation, the disclosure line is spoken, and the build
record is written into the adopter's own folder. A green `validate` run says nothing about §10.4
either way, and would look identical if the rules were never followed. Until a checker exists, the
only thing enforcing §10 is the AI assistant reading it and a person noticing afterwards.

---


## explore-options
<!-- method: explore-options -->

*Delivered: `type: outcome-first`.*

### 12.5 The `Explore options` step

An outcome-first template leaves choices to whoever builds it and says which ones — the audit
template lists them under *Free to choose alternatives*. Absent this step, the build makes those
choices itself and records each in the build record. **The `Explore options` step lets the developer
have them laid out first.**

**Ask it once,** after the last of the template's questions and before you present the design,
through the selection control (§10.3), on any outcome-first template that lists open choices:

| Option | Short description |
|---|---|
| `Build it as specified` | The build makes the open choices itself and records each one in the build record. |
| `Explore options` | Before the design, the AI assistant lays out each open choice with its alternatives, and you pick. Takes longer, by about one question per choice. |

**Preferred:** `Build it as specified`.

**What `Explore options` produces.** For each open choice with more than one workable route: the
alternatives, what each costs, and which of the template's checks each passes. A route the AI
assistant would not have taken on its own belongs in the comparison if it passes the checks — that is
the step's purpose. The developer picks per choice, one choice per step. Each pick is recorded in the
build record and holds for the rest of the run, and the picks are restated in the design presented
for approval, never left to the transcript. Every route picked still passes every check the template
carries; the step adds no check and relaxes none.

**What it never reopens.** A mechanism the specification names under §12.4 is part of the outcome,
not an open choice, and the step offers no alternative to it. A developer who asks for one is asking
for a different template, or for the from-scratch path with its warning, and is told so in those
words.


## staged-procedures
<!-- method: staged-procedures -->

*Delivered: `type: vba-scaffold`.*

**Running a `vba-scaffold`'s staged procedures.** Some `vba-scaffold` templates document a
sequence of procedures meant to run in order, each gating a decision the developer must make
before the next runs (see `templates/_template-schema.md` §8.4). When you're the one carrying out
that sequence: never infer the answer to a staged decision — including which named build option
applies — from the shape of the data or from reasoning that makes an answer seem obvious; ask the
developer and wait for their actual answer. Never substitute your own read of the underlying data
for a procedure whose job is to answer that question — run the procedure itself, at the point the
sequence calls for it. Present one step's result at a time; don't collapse the sequence into a
single upfront report, even when every fact in it is correct. Having the access and the context to
answer a gate yourself is not the same as being asked to. Where the template is divided into steps,
answer every gate inside a step before you fetch the next part.


### 8.4 Staged execution and facilitation

Some `vba-scaffold` templates document procedures meant to run in a specific order, where each
step gates a decision the developer must make before the next one runs — picking among named
build options, reviewing a generated list before the next procedure acts on it, and the like.
When a template documents this kind of sequence, it must also state, alongside the sequence, a
facilitation rule for any assistant carrying out the steps on the developer's behalf:

- **Never infer the answer to a staged decision.** Not from the shape of the data, not from
  domain reasoning that makes an answer seem obvious. Ask the developer, and wait for their
  actual answer, even when it looks predictable.
- **Never substitute your own analysis for a procedure whose job is to answer the question.** If
  the sequence includes a check or scan procedure, run *that procedure*, at the point the
  sequence calls for it. Don't read the underlying data directly and report a conclusion in its
  place.
- **Present one step at a time.** Don't collapse a staged sequence into a single upfront report,
  even where every fact in it turns out correct — the sequence exists so the developer reviews
  and approves each gate, not just the end state.
- **Restate the decision in full at the gate.** Ask where the developer can answer without
  reconstructing anything from earlier in the session: what the setting means, what it produces
  at run time, and what changes if they choose the other way — at the point of asking, not on
  request. And never use one number for two quantities in the same message: if 11 fields are
  auditable and 11 macros will be generated, say which is which, because a reader will otherwise
  take them for the same 11.

This is in addition to — not a substitute for — a project's own standing rule that no edit happens
without explicit approval. It addresses a different failure mode: an assistant that has enough
context and initiative to *answer* a gate the developer was meant to answer, even where it never
touches a file. (See the `## Wizard` section of `templates/audit/audit-logging-lite-scaffold.md`
for a worked example of this note in place, next to the steps it governs. §10 is how a template
presents such a sequence to the developer; this section is the rule the AI assistant follows while
running it.)


## design-only-handover
<!-- method: design-only-handover -->

*Delivered: design route.*

**On the design route, the design is the whole deliverable,** whether no Access MCP server is
connected or the developer asked for a design only. Hand over the
approved design and a build record of the decisions behind it. **Generate no executable artifact:**
no VBA `Sub`, no `CREATE TABLE` DDL, no importable form text. Nothing that has never been run
crosses to the developer looking like a finished build. That is the failure this exists to prevent:
code whose first execution happens in the developer's own Access session, as an error, in front of
them, with no one but them to diagnose it.

**If they ask for the code anyway, give it to them.** Someone who asks for it by name, having been
told nothing has run it, is making their own call and is entitled to it. Put `UNVERIFIED` at the
head of every file handed over, and say in the message that nothing has executed it. Before writing
it, call `get_part` with part `build-method` and route `build`, and follow `next` through every build
part: the method, standards and platform facts for building arrive only there. What ends is
*offering* it as the deliverable, not their ability to have it.


## build-route
<!-- method: build-route -->

*Delivered: build route.*

### After approval — building it

The design is the first deliverable. **On the build route there is a second one.** Where this run
ends was said before the first question; if it was not, say it now, before this question.


**Where an Access MCP server is connected, the developer chooses how the run ends, as a wizard step**
(§10) asked when the design is approved:

> **Ask:** The design is approved. How do you want this run to end?
>
> | Option | Short description |
> |---|---|
> | `Build it` | I create it in a copy of your database and run it, so anything that fails gets fixed before you see it. |
> | `Design only` | We stop here. You get the approved design and a build record, and you build it yourself. No code is written. |
> | `Give me the code` | I write the code as files for you to import and run yourself. Nothing will have run it first. |
>
> **Preferred:** `Build it`.

**Two different servers, and only one of them can build anything.** This library ships **the
template library MCP server** (`mcp-server/`, registered by the `.mcp.json` at the root). It
reads templates and standards and **cannot create or change anything in a database**. Building
needs **an Access MCP server** — one whose tools open and modify an `.accdb`. This library does
not ship one. **The template library MCP server never satisfies this check:** if the only server
connected is that one, no Access MCP server is connected. Neither name is ever shortened to "the
MCP" — that phrase alone names both, which is how the two get confused.

**What counts as an Access MCP server is a short list of abilities, not a particular product.** It
has to open a database, import a code module, run a procedure by name, report back the error number
and description when one fails, and close and release the file. Anything that does those five things
serves this library; `README.md` states them for the developer. Judge what is connected by whether it
can do them, never by its name. A sixth ability is not required: opening an existing database with its
startup skipped (see `templates/_materialization.md`). A server without it still serves; the build
then asks the developer to switch startup off.

**Whoever has an Access MCP server connected installed it deliberately.** So the question above
names the connected server and asks; it does not explain what an MCP server is. There is no
reader who has one and does not know what it is.

**The two failures this sits between, both of which have actually happened.** Using an Access MCP
server without saying so leaves the developer watching objects appear in their database with no
idea another route existed. Asking an open-ended "how would you like me to build this?", with the
library's reasoning about adopters attached, is a gate that stops them for nothing — they
connected the server in order to have it used. **Say what you have, name the preferred choice,
give them one click to take the other.**


**If the Access MCP server drops mid-build, restoring it is yours to attempt, not a question to
hand the developer** — a tooling outage turned into a choice converts a trial of the template
into a trial of the plumbing. Reconnect and carry on if you can. Nothing in this library restores
it: it is registered in the developer's own AI client, not shipped here, so `mcp-server/setup.ps1`
has no bearing on it — that script sets up the template library MCP server.

**If you can't, the build pauses. It does not convert into a handoff.** Stop, and give the
developer the build record as it stands: every object created, every one not created, what state
the database is now in, and whether to retry once the server is back or restore their backup and
start again. Handing them code to finish a half-built database with is the worst version of the
file handoff rather than the safest — nobody on either side knows exactly what landed, and the
code has still never run. Tell them:

> The Access MCP server couldn't finish the build in this run, so I've stopped rather than leave
> you guessing. The build record lists what was created and what wasn't, and what state your
> database is in now. You can retry once the server is working again, or restore your backup and
> start clean. If the trouble persists, troubleshoot the server before retrying this template.

***Tell me more* on that step covers what each route does to the database** — not the entity
caveat below, which the developer can do nothing about and which is yours to handle silently:

> Building directly creates the modules and objects while you watch, and nothing is written until
> you choose it. It also means I run what I write, so anything that fails gets fixed before you
> see it. Taking the files means you import and run them yourself, at whatever pace you like, and
> it means you are the first thing that has ever run that code: anything wrong with it turns up as
> an error in front of you. Either way you approve the design first.


Then **ask which platform the tables are for**, and generate the matching artifact (keys,
relationships, indexes, lookup tables, and **seed rows** throughout):
- **Access (ACE) local tables** → a **VBA `Sub` using DAO**, built as the fact
  `dao-table-build` delivered with the build step says.
- **SQL Server** → `CREATE TABLE` DDL.
- The error-handling block in any generated VBA comes from the standards layer — a **dependency-free
  default** (a message box) unless the house `error-handling.md` specifies a central logger.

Apply the standards throughout, exactly as in the approved design.

### Building it part by part

**Fetch the template one part at a time, and do each step's work only after its part has arrived.**
Each answer's `next` names the call to make and when to make it; make it then, not earlier and not
later. Do one step at a time from that step's own answer: its text, the platform facts it uses, and
the standards in force. Do no step's work from memory or from another template: a step whose part you
have not fetched is a step you may not do, and its row is missing from the build record.

**A step's answer names the standards in force and the part that carried each one's text.** Where
you cannot see that text now (the conversation was summarized, or you would be quoting it from
memory), fetch that part again before writing anything it governs. What you remember of a standard
is not the standard.

## access-gate
<!-- method: access-gate -->

*Delivered: build route.*

**Before you build in a database this template alters, ask whether the developer has a backup of it.
If they say they have none, build nothing in that database until they have one.** This stops the
build and nothing else: the design, and any code you hand over for the developer to run, are not
held back by it.

**Before the first open of the target file, check for it being held by a process with no visible
lock file.** On Windows, a prior Access session can crash or hang and leave `MSACCESS.EXE` running
with the file open but no `.laccdb` beside it — the next exclusive-open attempt then fails with no
obvious cause. Check for this before attempting to open, not after a failed attempt: list
`MSACCESS.EXE` processes, and where one's command line names the target file (or an ambiguous
`-Embedding` instance can't be ruled out), ask the developer whether to end it before proceeding.
Never end a process without asking — an untitled instance is often the developer's own hung work,
not a stray one.

**At the start of every build, identify the Access process you will work in. Act on no other.** The
developer's own Access may be running beside yours, and every route to Access can reach it, because
Windows lists a running Access as something any program on the machine can attach to. Asking the
system for "an Access" does not let you name which process you get, so you never work that way.

1. List the `MSACCESS.EXE` process IDs before you begin.
2. Bring up the instance. Either the Access MCP server opens the file, or you start the process
   yourself with the file named on its command line and keep the process ID the launch returns.
3. List again. The new process ID, whose command line names the target file (or is an `-Embedding`
   instance where the server opened it), is your instance. Write it in the build record. Where you
   started the process, bind to the file and confirm the instance that answers reports the same ID.
4. Work only in that instance, and on a check copy whose name is unique to the run. Binding to a
   file confirms an instance; it does not establish one.

Do nothing to an instance you did not identify as yours: no writes, no quitting, no opening a file
in it. Something that needs a second Access process you cannot identify is not tested; say so in
the build record. A split design is the same step for each file you open yourself.

**Diagnostic VBA written to debug a build in progress must guard itself.** A throwaway probe run
live against the developer's own Access session can trigger the VBE's debugger if the machine is
set to break on all errors rather than unhandled ones — a setting outside this library's control,
and outside the developer's expectation. Automation can dismiss the resulting dialog but cannot
clear the paused state; only the developer can, by closing or resetting the project themselves,
which stops their build until they do. Wrap every such probe in its own resumable error handler
(`On Error Resume Next`, or equivalent) so it cannot trigger a break regardless of that machine's
setting — never rely on first setting the project's own error-trapping option, since a probe run
before that fix takes effect is exactly what causes this.

**Run nothing in Access that you cannot see into and clear without the developer.** Access freezes
silently, with no error returned to you, on a dialog, on a file another process holds, and on code
that runs at startup. Only the developer can clear it, and it stops the run until they do. Before
any action that could hit one:

1. Work only on a copy named for the run, never the developer's file. A hang or a forced close then
   costs nothing.
2. One owner per file. An Access MCP server session holds its file exclusively, so close the session
   before starting any process on that file. Test that the file is free before you launch; never
   launch to find out.
3. Start a dialog watcher before any process you launch. It reads and closes dialogs in your own
   identified processes only.
4. Give every process you launch a deadline. At the deadline, close it by process ID (yours only)
   and treat that copy as spoiled.
5. VBA you write for Access (checks, hooks, probes) must not be able to raise an unhandled error.
   Give it a handler throughout and no `Stop`.
6. Open anything that is not the developer's own open with startup skipped.

If you cannot do all of this for an action, do not run it. Record it as not tested in the build
record. This applies to every subagent you start: put it in the subagent's prompt.


## quiet-build
<!-- method: quiet-build -->

*Delivered: build route.*

§10.4 rules 1 to 3 and the disclosure line are delivered on every run, in the method
`between-questions`. Rule 4 is the build route's alone:

**4. While the build runs, do not narrate it.**

The questions are over and the developer is waiting for a result. Everything happening now is work
they already approved, so a running commentary on it reports progress to nobody: they cannot act on
it, cannot verify it, and cannot tell from it whether anything is going wrong. Rule 2 covers the
moment before the build; this covers the build itself, which is longer and where the habit is
strongest. Say three things between the last question and the finished artifact:

- **That it has started**, once, and what it will produce. Silence for several minutes is its own
  failure — this is the line that prevents it.
- **Anything that needs the developer to act** — a failure they have to clear, or something the
  build hit that no question covered. Always asked as a question, never narrated past.
- **That it is finished**: what was built, and where the build record is.

Every object created, every procedure run, every check that passed, every step that went exactly as
expected: put all of it in the build record. **When the build goes to plan, say nothing between its
first message and its last.**

**Progress commentary has a real audience, and it is not this one.** Someone developing or trialling
a template does want to watch each step land — they are reading for the template's behaviour, not
for their own database, and that is a different reader (see the three readers in `CLAUDE.md`). They
will say so. Absent that, the developer wants their tables, not a transcript of them being made.


## build-record
<!-- method: build-record -->

*Delivered: every run.*

### The build record

**Every build leaves a build record**: one file, written beside the artifact, saying how it was
built. It is not optional, and it is not a summary of the conversation. **This applies to every
template and every build** — whether a wizard ran or the developer answered everything at once,
whether the artifact is tables, code, or a form. **A design-route run leaves one too:** the design
is its artifact, and the record holds the decisions behind it. It is part of the deliverable, written
without being asked for, never offered as an extra.

It is also what `_template-schema.md` §10.4 promises the developer in your own words, and it is the
reason the detail can stay out of the messages they read while they are still deciding things.

**Name it `build-record-YYYY-MM-DD-HHMM.md`**, using the date and time the build started, so a
second build in the same folder never overwrites the first. Write it in the folder that holds the
artifact — beside the `.accdb`, not in the library. **On the design route there is no `.accdb` to
sit beside:** ask the developer, as a question, where their copy goes. Where a build touches two files, such as a front end and a back end,
one record covers both.

**Open the record with how the template arrived.** Write a table with one row for every part you
fetched, in the order you fetched it: the part; the version and `sha` its answer returned in
`served`; the platform fact ids it delivered; the time its answer arrived; and the first action you
took that the part governs, with its time. Where you did any of a part's work before that part
arrived, say so in its row, and name what you then checked again against its text. A part you never
fetched gets no row: leave the gap, because it is the evidence. If you read template files directly
instead of fetching them, say so here and name the files; do not fill the table in as if they had
been fetched.

**A build record does not end its life with the build it documents.** In addition to the developer's
own copy above, copy it into `build-records/<template-slug>/` in the library, named
`YYYY-MM-DD-<short-description>.md`. One folder per template — every build of that template lands in
the same folder, never a fresh one per run — because the folder's value is in what accumulates there,
for this template and for any other template that later needs the same mechanism (see the
method `build-records-accumulate`). Nothing about the record's content changes for this: it is
written exactly as the six parts below already require, narrated with its false starts intact. Only
its destination gains a second copy.

Six parts, in this order:

1. **What was built, and where.** Every object created, and which file it went into. The developer
   should be able to open the database and find each one.
2. **What was checked before building.** The state you found: what already existed, which references
   were present, whether the folder was trusted. This is the part that is worth nothing on the day
   and a great deal three months later, when something has changed and nobody remembers what it
   used to be.
3. **The decisions taken.** Every question the developer answered and what they chose, and any
   decision they handed back to you. Where a wizard ran, that is its steps and their answers.
4. **What was verified afterwards, and how.** The tests actually run, with their real results —
   never "tested and working". Say which paths were exercised, and name the ones that were not.
5. **Anything that did not match what the template said.** Divergences, surprises, and anything you
   worked around. This is the section a template author needs and nobody else will write.
6. **What is left for the developer to do.** Every follow-up the build could not complete, including
   anything the standards layer calls for that the build route could not deliver.

**Write it before you say the build is finished**, not when you are asked for it. A record written
later is written from memory, and the details worth keeping are the first ones to go.

**That is a deadline, not a cadence — write it as you go, not only at the end.** A record started
after the build is a reconstruction; the same six parts, filled in as each thing actually happens,
are a contemporaneous account. Confirmed by contrast: a build written up afterward, from a 90-minute
session, read thinner than builds written incrementally throughout.

**It records what happened, not what was meant to happen — good or bad, and whatever it reflects on
you.** A step that failed and was retried belongs in it. A test that was skipped belongs in it, named
as skipped. A build record in which everything went to plan is either untrue or not worth keeping.

**Two kinds of entry go missing, and they are the two worth most.** Anything the developer had to do
to unblock the run — a message they had to clear, a file they had to close. From where you sit that is
a moment of waiting; from where they sit the build stopped and demanded something, and it is the part
they will remember. Say what stopped, what they had to do, and how many times — not only a technical
note about the cause. And anything that went wrong because of how you worked rather than because of
their database: a file written in the wrong encoding, code that would not compile, a tool used the
wrong way. Both belong in part 5, with everything else that did not go as the template said. **Neither
is a new part of the record** — nothing here needs a seventh.

---


## runbook
<!-- method: runbook -->

*Delivered: build route.*

### The runbook — the two builds that leave the developer something to run

**When it applies:** any build that leaves a procedure for the developer to run — which the staged
`vba-scaffold` sequences do by design — and any code handed over headed `UNVERIFIED` because the
developer asked for it. Where you built everything directly and nothing is left to run, there is
nothing to write. **It is not the packaging for a run with no Access MCP server connected.** That
run ends at the approved design and generates no code at all, so there is nothing to write a runbook
about.

**A build the developer has to run any part of leaves a runbook**: one file, beside the code, saying
what to run and in what order. It is not optional, and it is not the build record. The build record
says what was done; the runbook says what they do next.

**Why it exists.** You know the order because you generated the code. They have a folder of files. A
procedure named `Three_…` tells them it is third; it does not tell them what the first two are,
which one needs an answer written into a table before the next will do anything, or which ones are
optional. **Naming a procedure is not documenting it**, and a developer who cannot tell an optional
tool from a required step either runs everything or runs nothing.

**Name it `runbook.md`**, and write it in the folder that holds the code.

Five parts:

1. **Before you start.** Everything that has to be true first, each one paired with what happens if
   it isn't: the folder trusted, the file writable, every object in the database closed, a backup
   taken. A precondition with no stated consequence gets skipped.
2. **Importing the files.** Which files, and where each one goes — in a split design, which belong
   in the back end and which have to be in every front end as well.
3. **What to run, in order.** Every procedure they run, numbered, each with: what it does, how to
   run it, what they should see when it worked, and what to do when it didn't. **Where a step needs
   them to go and look at something** — review rows in a table, set switches, check a value — that
   is its own numbered step, not a remark attached to another one.
4. **What is optional.** Procedures that exist but are no part of a normal run, and what each is
   for.
5. **How to tell it worked.** What they can open or run to confirm the result in their own database,
   without asking you.

**Write it in the second person and name every procedure exactly as it appears in the code.** A
runbook that says "run the setup procedures" has told them nothing they had not already guessed.

**Say it exists when you hand the files over.** A file in a folder nobody was told about is a file
nobody reads.

---


## build-records-accumulate
<!-- method: build-records-accumulate -->

*Delivered: build route.*

### Build records accumulate — read them before you build, add yours after

`build-records/` holds one folder per template slug (e.g.
`build-records/officiating-assignment-outcome-first/`), each containing every build record ever
produced for that template — never one build's private scratch, always a shared resource for the
next one, whatever template it's for.

**Before building**, check the folder for the template you're about to build, and check any other
template's folder for a mechanism you're about to reuse (see the complementary-templates rule in
the method `design-review`). Read what's there as a narrated account of what was tried, what failed,
and what held — not a rule list to apply blindly. If a fact it records no longer holds in your
environment, say so in your own record rather than assuming the old one is still current.

**As you build**, write the build record to the developer's own location as always, adding to it as
each thing happens rather than writing it up afterward.

**Once the build is finished**, copy the finished record once into `build-records/<template-slug>/`,
named `YYYY-MM-DD-<short-description>.md`. Create the template's folder if it doesn't exist yet. This
is not optional housekeeping — it is the whole point of the folder existing.
