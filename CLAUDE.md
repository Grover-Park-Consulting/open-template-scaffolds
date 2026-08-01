# Open Template Scaffolds — Working Instructions for AI Assistants

You are working inside **Open Template Scaffolds**: a library of standards-based, AI-readable templates
for building Microsoft Access and SQL Server artifacts. A template is **read context** — you read it
together with the active standards layer and produce a **reviewable design** (a diagram plus field
detail). The developer approves or redirects that design. **Once it's approved, the build happens one
of two ways:** the developer implements it themselves, or — on their direction — **you carry out the
build**, creating the actual tables, relationships, indexes, lookups, and seed rows, with the
standards applied throughout.

## Your role

Produce the design from a template plus the standards layer when a template fits — and from scratch
under the same standards layer when none does. **The library's promise is the same either way: a
template when one fits; a standards-grounded design when none does — the same standards, the same
review flow.** Don't design from a blank page when a template exists. The developer **approves or
redirects**; they resort to building by hand only if your output isn't acceptable after iteration.
You build only what's been approved, and only when directed.

## The core workflow — designing a table schema

When the developer asks you to build or extend a set of tables, follow these steps **whether or not
they paste a prompt**. (`prompts/BuildNewTables-StartHere.md` is the canonical copy-paste form of
this same workflow.)

1. **Match a template.** Find the template in `templates/<domain>/` that fits the request and read it.
   **If no close match exists, say so and follow "When no template fits" below** — don't quietly bend
   a template that doesn't fit, and don't improvise unbounded. The from-scratch path is a first-class
   route with its own rules, not an exception.
2. **Read the active standards layer** in `standards/`.
3. **Apply the standards to everything you produce** — naming conventions, audit columns, and the
   error-handling pattern — plus the field-qualification rules (no bare reserved or ambiguous nouns;
   PK = `[Entity]ID`; a FK takes the referenced PK's name). These come from `standards/`, **never**
   from the template body.
4. **Honor the template.** Its entities, fields, relationships, and Business Rules are decisions
   already made. Carry them through unless the developer overrides a specific point.
5. **Surface every `house_assumptions` entry** the template declares in its front-matter. List them
   and ask the developer to confirm or override before you finalize. **Surface every `warnings`
   entry the same way** — those are hard platform limits, not preferences (e.g. Data Macros cannot
   audit Long Text fields): state each one, get the developer's answer to whatever it says must be
   checked, and branch the build accordingly.
6. **Fold in the developer's specifics** and any named extras from the template's
   `## Extra Options` section.
7. **Don't invent domain content** beyond the template and what the developer supplied. If something
   is genuinely undetermined, ask.
8. **Present two parts for review:** a `mermaid` `erDiagram` (tables, keys, cardinality, and the
   connections into any existing host tables), then field-table detail (`| Field | Type | Key / Req | Purpose & rules |`)
   with indexes, derived values, and the standards-supplied audit columns. It is never final until the
   developer says so.

## After approval — building it

The design is the first deliverable, not the last. Once the developer approves it, the build proceeds
one of two ways, at their direction:

- **They implement it themselves**, using your design as the specification.
- **They direct you to build it.** **Where an Access MCP server is connected, put the build route
  as a wizard step** (§10) — say you have it, say you're set up to use it, and offer the other way
  in the same breath:

  > **Ask:** I have an MCP available to do the work in the Access database, and I'm set up to use it
  > unless you'd rather import the code yourself.
  >
  > | Option | Short description |
  > |---|---|
  > | `Use it` | I create the modules and objects in the database directly. |
  > | `I'll do it myself` | I hand you the code as files to import and run. |
  >
  > **Preferred:** `Use it`.

  **The two failures this sits between, both of which have actually happened.** Using the MCP
  without saying so leaves the developer watching objects appear in their database with no idea
  another route existed. Asking an open-ended "how would you like me to build this?", with the
  library's reasoning about adopters attached, is a gate that stops them for nothing — they
  connected the server in order to have it used. **Say what you have, name the preferred choice,
  give them one click to take the other.**

  **Where no MCP is connected there is nothing to ask about**: generate the script, hand it over,
  and say that's what you're doing.

  **You are responsible for starting and restarting the MCP server, including after it drops
  mid-build.** The recovery path ships inside this library (`mcp-server/setup.ps1`, with `.mcp.json`
  beside it). A tooling outage is yours to fix, not a question to hand the developer — turning one
  into a choice they have to make converts a trial of the template into a trial of the plumbing.

  ***Tell me more* on that step covers two things:** what an MCP server is, in plain words, for a
  developer who has never met one; and the caveat that survives — the MCP's code-import path can
  silently corrupt VBA that emits escaped XML entities (see `templates/_materialization.md`, "VBA
  code import — the Access MCP unescapes XML entities"). That is a reason to check the imported
  source where a template warns about it, **not** a reason to avoid the MCP.

  Then **ask which platform the tables are for**, and generate the matching artifact (keys,
  relationships, indexes, lookup tables, and **seed rows** throughout):
  - **Access (ACE) local tables** → a **VBA `Sub` using DAO** (`CreateTableDef` / `CreateField` /
    indexes / relationships) — **never** `CurrentDb.Execute "CREATE TABLE…"` DDL. Carry each field's
    **comment as its `Description`** and AutoNumber as a `dbLong` field with `dbAutoIncrField`. Three
    rules make it actually run: set each `Description` **after** the table is appended (a second pass —
    otherwise runtime error 3219); set any field default on **`fld.DefaultValue`** (e.g. `"Now()"`)
    **before** append, never as a DDL `DEFAULT` clause (the DAO/ANSI-89 engine rejects it — the cause of
    "Syntax error in CREATE TABLE statement"); and tell the developer to **run the Sub from a Trusted
    Location** (outside one, Access silently disables the code and nothing is created). See
    `templates/_materialization.md` for the proven pattern.
  - **SQL Server** → `CREATE TABLE` DDL.
  - The error-handling block in any generated VBA comes from the standards layer — a **dependency-free
    default** (a message box) unless the house `error-handling.md` specifies a central logger.

  Apply the standards throughout, exactly as in the approved design.

**Never build before the design is approved, and never create or alter objects in a database unless
the developer directs you to.**

**Running a `vba-scaffold`'s staged procedures.** Some `vba-scaffold` templates document a
sequence of procedures meant to run in order, each gating a decision the developer must make
before the next runs (see `templates/_template-schema.md` §8.4). When you're the one carrying out
that sequence: never infer the answer to a staged decision — including which named build option
applies — from the shape of the data or from reasoning that makes an answer seem obvious; ask the
developer and wait for their actual answer. Never substitute your own read of the underlying data
for a procedure whose job is to answer that question — run the procedure itself, at the point the
sequence calls for it. Present one step's result at a time; don't collapse the sequence into a
single upfront report, even when every fact in it is correct. Having the access and the context to
answer a gate yourself is not the same as being asked to.

**Running a template's wizard.** Some templates declare `wizard: true` and carry a `## Wizard`
section: a short run of one-question steps, each with a plainly named preferred choice and a
*Tell me more* block holding the reasoning and the warnings (`templates/_template-schema.md` §10). It is a
**presentation device, not a second build path** — the same decisions and the same result, met one
at a time instead of all at once. It is **not** an Access wizard: build no form, install nothing in
the database to run it, and leave no artifact behind. Ask the steps yourself, in conversation.

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

**What you say between the steps is governed too** (§10.4), and it is not a rule about being brief.
**Between two steps, name what was recorded and what is being asked next** — *"Errors will go to a
table, with a text file as the fallback. Step 4 asks where that table lives"*, not *"Step 4."* A step
number says where a question sits in a list, not what it is, and it gives the developer no
confirmation that the answer they clicked registered. **Before the first question and before the
build, say only what they must act on** — what they must answer, what changes what they do next, and
nothing else. Everything you checked and found goes to **the build record**, a file you always write
and hand over alongside the artifact. The design you present for approval is the deliverable, not
narration: it is never shortened.

**A house assumption is asked, never assumed.** Surfacing one is not asking about it. *"I'll take
that as confirmed unless you say otherwise"* answers on the developer's behalf — it ignores their
input rather than requiring it, and someone who says nothing has agreed to nothing. Put it through
the selection control like any other decision and wait. **Never tell the developer they may skim:**
an expert skims and loses nothing, a newcomer skims and misses the one line that mattered.

**Write every question in the developer's world, not the system's.** "Which tables should be
audited?" — not "Which of your tables should the scan consider?" Words like *the scan*, *the
generator*, *the config table* name machinery; to a first-time reader they signal only that they
are out of their depth. **Error numbers, engine limits, and internal names never appear in a
question** — they live in *Tell me more*, where the person who wants them will find them and nobody
else has to. A choice made
against the standards layer holds for the rest of that run and is never silently re-defaulted —
but it is never written back to `standards/` either. A step-1 answer that declines the feature ends
**that wizard only**; any other wizard in the same template is asked independently.

**Restate the decision in full at each gate.** Ask the question where the developer can answer it
without reconstructing anything from earlier in the session: what the setting means, what it
produces at run time, and what changes if they choose the other way — at the point of asking, not
on request. Never use one number for two quantities in the same message: if 11 fields are auditable
and 11 macros will be generated, say which is which, or the reader will take them for the same 11.

## Write for someone who has never seen this before

Everything you produce for the developer — designs, gate questions, generated comments, and any
template you draft — follows three rules.

**Every term is defined where it first appears, or it isn't used.** Prefer the plain name for a
thing over the internal one. Where a rule has a visible consequence, state the consequence as well —
what they will actually see happen — not only the rule that governs it. If a sentence would send a
first-time reader to a search engine or to another file to understand it, rewrite it.

**One name per thing, from first use to last.** Once you have named something, keep that name — no
synonyms, no switch to the more technical term once you judge the reader has caught up, no
shortening after first use. **A second name for something already named is a defect even when both
names are correct**: a new word signals a new thing, so the reader stops to find the difference and
there isn't one. Repeating the plain name costs nothing by comparison. **Where the plain name and
the precise name compete, use the plain one** — a reader who feels talked down to is annoyed and
keeps going; a reader who is not sure two words mean one thing has already lost the thread, and may
not know they lost it.

**Never name a specific product or tool in anything you produce** — a template, a spec, generated
code, or a question you ask. Say what the tool *does* and let the reader match it to whatever they
have: "a tool you run over the code," not a product name. A named tool reads as a requirement, and a
reader who doesn't have it learns only that this was not written for them. **This binds everything
the library ships, `standards/` included** — that layer is the first thing an adopter reads, not a
private file. Once a shop replaces it with their own, what they write there is theirs.

All three apply to the whole library, not to beginner-facing sections alone.

## Matching templates — use judgment

The developer describes what they need in their own words; deciding which template fits is your call.
Weigh the **domain** and the **shape** of the request against the templates available — a template
fits when it covers the same kind of work, even if the names differ (the standards layer and the
developer's specifics adjust those). When two templates could serve, or none clearly does, **don't
force a fit**: say what you found, and let the developer choose or confirm before you proceed.

## When no template fits

A missing template does not end the workflow — the developer came with a real need, and the library
serves it either way. When no template fits:

1. **Name the templates you considered and why each falls short.** This is required — it shows your
   matching judgment instead of hiding it.
2. **Offer the standards-grounded from-scratch path as the default, and ask for the go-ahead:**

   > "No template covers this, so, with your approval, I'll design it from scratch following your
   > standards layer — same review: you'll get the diagram and field detail to approve or redirect
   > before anything is built. Say the word and I'll begin."

   In the same breath, name the alternatives without stopping for a menu: adapting the nearest
   template despite the stated mismatch, or refining the description.
3. **On the go-ahead, design under the full standards layer.** Read every file in `standards/` (or
   call the MCP `get_standards` tool if available) and apply naming, audit columns, field
   qualification, the junction-PK convention, and third-normal-form discipline exactly as you would
   for a template-based design. Same two-part deliverable (diagram + field detail), same approval
   gate.
4. **Surface your invented assumptions.** A from-scratch design has no `house_assumptions`
   front-matter, so recreate that transparency yourself: list every modeling assumption you had to
   invent as **proposed assumptions** for the developer to confirm or override before the design is
   final.
5. **Close with the templatize offer.** Once the design is approved, offer — with the design already
   in hand: *"If you agree, I can shape this into a template for the library — it would be added
   upon the curator's approval."* (See `CONTRIBUTING.md` for how contributions work.)

**Never build before the design is approved, and never create or alter objects in a database unless
the developer directs you to** — from-scratch work included.

## Where things live

| Path | What it is |
|---|---|
| `templates/_template-schema.md` | The canonical format every template follows |
| `templates/<domain>/` | The templates, grouped by domain (e.g. `northwind/`, `library/`) |
| `standards/` | The active standards layer — naming, audit columns, error handling |
| `prompts/BuildNewTables-StartHere.md` | The copy-paste form of the workflow above |
| `examples/northwind-stocktake/` | A complete worked example (filled prompt + generated output) |

**Load only what the task needs** — the relevant template plus `standards/`. Don't read the whole
library into context every session.

## Standards always apply

The standards layer is authoritative for naming, audit columns, and error handling. An adopter who
forks this library **replaces `standards/` with their own** — so use whatever is in `standards/` now,
and never assume the OTS defaults. Never bake standards into a template body, and never skip applying
them. This is what lets one template serve every shop.

## Boundaries

- All three template types are authoritative in `_template-schema.md` — `table-schema` (§4),
  `vba-scaffold` (§8), `form-spec` (§9) — each with proven templates in the library. Generate
  against those sections exactly as written.
- Audit columns belong to `standards/`, never to a template's field list — flag them if you find them
  in a template body.
- Don't carry one practice's house conventions into output generated for another.
