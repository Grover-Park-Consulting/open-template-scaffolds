---
template: catalog-outcome-first
title: Library Publication Catalog — outcome-first method
domain: library
type: outcome-first
version: 0.1.1
status: stable
implements: catalog-schema
standards_layer:
  - audit-columns
  - naming-conventions
  - error-handling
  - query-style
  - design-principles
house_assumptions:
  - "Noise words — the list of leading words stripped from a title to form its sort title (The, A, An, and
    whatever else a given collection adds) is house-specific — a language, a cataloging convention, or
    a local style choice. A shop cataloging in another language, or with its own list, changes it."
warnings:
  - Choosing the VBA route for Business Rule 2 (see below) means the sort title only stays correct
    for edits made through your entry form. A title changed any other way — directly in the table, by
    an import, by a query — silently keeps its old sort title. This is a disclosed trade-off of that
    route, not a defect; the Data Macro route does not have this gap.
related:
  - "record-finder-scaffold — worth adding once your catalog holds records: gives your entry
    form a way to search and filter by title, creator, or genre instead of scrolling through
    everything."
---

# Library Publication Catalog — outcome-first method

**Status last determined:** 2026-09-26.

**Who reads this.** Everything from *Intent* down to *Standards Layer* is written for the developer
whose database this is. The two sections after that are addressed to the AI assistant building it,
and each says so where it starts.

---

## Intent

**To produce, in an Access database, the results described in *What you end up with* below.** That
section is the specification. It says what the database does once this is built, in words the
developer can check for themselves, and it is followed by the checks that confirm the result arrived
and by the behaviours that must hold however the work was divided up. Those three sections together
are the whole of what this template promises.

**This template realizes two of the five Business Rules the table template states, and leaves the
other three alone.** `catalog-schema` builds the ten tables and states, in its own Business
Rules, what has to be true once the database is in use. Three of those five are already complete the
moment the schema is built — one record per title (Rule 1), no duplicate creator or genre links
(Rule 4), and cascade delete on a publication's own links (Rule 5) are properties of the tables
themselves, enforced by how the grain was chosen, a unique index, and a relationship setting. There is
no route to specify for any of them and nothing this template adds changes them. The other two — the
sort title staying in step with the title (Rule 2) and the volume/set validation (Rule 3) — need
something built on top of the tables to hold, and more than one way exists to build each. This
template states those two as outcomes and leaves the route to whoever builds it, within the choices
named below. It assumes the schema's tables already exist, or are built as part of this same run.

**Nothing else is here.** No procedures, no module names, no code beyond what's needed to reason
through the choices named for Business Rule 2.

---

## What you end up with

### Business Rule 2 — the sort title stays in step with the title

**The sort title — the title with leading noise words like "The", "A", "An" removed — is set
correctly the moment a publication is added, and stays correct whenever the title is later edited
through your entry form.**

**Two routes exist to build this, and which one you pick changes what you actually get — so you're
asked, not defaulted.**

- **The Data Macro route (preferred).** The derivation lives in the table itself, so it fires no
  matter how the title changes — your entry form, a direct edit to the datasheet, an import, a query
  that updates many rows at once. The sort title is always correct, whatever the route.
- **The VBA route.** Simpler to build, and coverage stops at your entry form. A title changed any
  other way — directly in the table, by an import, by a query — silently keeps its old sort title,
  and the catalog quietly sorts that row in the wrong place until someone notices and re-saves it
  through the form. That gap is a known, accepted trade-off of this route, not something to be fixed
  later — choose it only if you're comfortable living with it.

**Why this is offered as a real choice instead of naming one mechanism outright.** Where a platform
genuinely leaves exactly one route to a promise, this library names that route and stops there (see,
for example, how `capital-asset-tracking-outcome-first` names a Data Macro for its
capitalization check). This isn't that case: a sort title is a convenience for browsing and
reporting, not a compliance record, so a lesser but simpler route is a reasonable thing to offer —
provided the person choosing it knows exactly what they're giving up. Naming both mechanisms and
disclosing the difference is what keeps that an informed choice rather than a silent shortcut.

### Business Rule 3 — volume / set validation

**A publication's `PublicationVolume` and `NumberOfVolumes` are never saved in an inconsistent
combination, whatever the route the save was made through** — not through a form, not through a
direct edit to the table, not through an import.

- `NumberOfVolumes` is always present and defaults to 1 when nothing else is specified.
- Where `MultiVolumeSet` is set, `PublicationVolume` must be present and fall between 1 and
  `NumberOfVolumes`, and `NumberOfVolumes` must be at least 2.
- Where `MultiVolumeSet` is not set, `PublicationVolume` must be empty.

A save that violates any of these does not go through, regardless of which route was used to attempt
it.

**No mechanism is named here.** Unlike Business Rule 2, nothing here needs to compare an old value
against a new one or derive and write anything — it only asks whether the row, as it now stands, is
internally consistent. A table-level Validation Rule answers exactly that question, on every route
into the table, which is everything the promise needs; a Data Macro could enforce the same thing and
costs more to build for no additional coverage. Either is a legitimate build.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

**Do these on a copy, and stop if you do not have one.** The checks add, change, and delete records
in your own tables — that is what they test, and there is no version of them that leaves your data
alone. Make the copy once the build is finished, run every check on it, and keep your working file
out of it. Where your database is split into two files, copy both and keep them together.

1. **A new title gets a correct sort title.**
   - Add a new publication titled "The Great Gatsby"
   - Confirm the sort title reads "Great Gatsby" — the leading noise word removed
   - **Where your noise-word list carries house-specific additions** (Step 5 under *Information
     and conditions you need to supply*), repeat this with a title starting with one of those
     additions too — "The Great Gatsby" only exercises the default three words, and a custom
     addition that doesn't work would still pass this check as written otherwise
2. **A title with no leading noise word is stored as itself.**
   - Add a new publication with a title that carries no leading article
   - Confirm the sort title matches the title exactly
3. **Editing the title through your entry form updates the sort title.**
   - Edit an existing publication's title through your entry form
   - Confirm the sort title updates to match the new title
4. **Editing a title directly in the table — only if you chose the Data Macro route.**
   - Edit a publication's title directly in the datasheet, not through the form
   - Confirm the sort title updates too
   - If you chose the VBA route instead, this check is not expected to pass — that's the disclosed
     trade-off of that route, confirmed by *Business Rule 2*'s own text above, not a failure to
     report as one
5. **Editing an unrelated field leaves the sort title untouched.**
   - Edit a field on a publication other than the title
   - Confirm the sort title is exactly what it was before
6. **A multi-volume set without a volume number is refused.**
   - Set `MultiVolumeSet` on a record and leave `PublicationVolume` empty
   - Confirm the save does not succeed
7. **A volume number outside the set's range is refused.**
   - Set `MultiVolumeSet`, `NumberOfVolumes` to 3, and `PublicationVolume` to 5
   - Confirm the save does not succeed
8. **A single-copy record with a volume number set is refused.**
   - Leave `MultiVolumeSet` unset and set `PublicationVolume` to any number
   - Confirm the save does not succeed
9. **A "set" of fewer than two volumes is refused.**
   - Set `MultiVolumeSet` and `NumberOfVolumes` to 1
   - Confirm the save does not succeed
10. **A valid duplicate-copy record succeeds.**
    - Leave `MultiVolumeSet` unset, `PublicationVolume` empty, and `NumberOfVolumes` at 3
    - Confirm the record saves
11. **A valid multi-volume record succeeds.**
    - Set `MultiVolumeSet`, `NumberOfVolumes` to 4, and `PublicationVolume` to 2
    - Confirm the record saves
12. **`NumberOfVolumes` defaults to 1.**
    - Add a new publication without setting `NumberOfVolumes`
    - Confirm it reads 1
13. **Running the build again does not duplicate anything.**
    - Run whatever attaches the Data Macro or Validation Rule a second time
    - Confirm there is still exactly one copy of each — no duplicated Data Macro actions, no second
      Validation Rule
14. **When a build is blocked, it stops rather than leaving tables half-finished.**
    - Leave a table or form open in the database
    - Start the build
    - Confirm it stops, names what is open, and changes nothing

### The same behavior every time, not the same structure

For an AI-assisted template, two builds need not produce the same code. They need to produce a
database that behaves the same way. The following must be true of every build:

- **Which route was chosen for Business Rule 2 is stated plainly in the design and the build record**,
  not left for the developer to discover by testing. They're choosing a coverage tier, and they need
  to know which one they got.
- Where the Data Macro route was chosen, the derivation happens inside the Data Macro itself, not in
  code the macro calls out to — a Data Macro whose only action hands the work to a function elsewhere
  has moved the behaviour out of the database engine and back into whichever route calls that
  function, which reopens exactly the gap the Data Macro route exists to close.
- The volume/set check never writes a value — it only accepts or refuses a save, on every route into
  the table.
- **[your standards]** The error-handling pattern applies to the build-time procedures that attach the
  Data Macro or Validation Rule, not to the refusal itself. When a Validation Rule or a Data Macro's
  `RaiseError` action refuses a save, Access shows its own dialog naming the problem — no VBA runs, no
  `On Error` handler fires, and nothing is logged unless something else is separately watching.
- Running the build again replaces what a previous build attached; it never adds a second copy
  alongside it.

Lines marked **[your standards]** come from your standards layer rather than from this template, and
move with that layer if your shop replaces it. Everything else in `standards/` applies here as it
does to every template.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences, say so while
the design is being worked out, before anything is built. Where you don't choose, the build will
choose based on the rules built into it. The template's promise holds either way.

- **How Business Rule 3 is enforced** — a table-level Validation Rule or a Data Macro. Either
  satisfies the promise; the Validation Rule is the simpler build. (Business Rule 2's mechanism is
  *not* on this list — it's asked directly, under *Information and conditions you need to supply*,
  because the two routes produce different outcomes and the developer needs to choose knowingly.)
- How VBA code, where any is used, divides into procedures, what it's called, and how many there are.
- The wording of everything the developer sees.
- Whether the build reports in message boxes, as returned text, or both.
- How the code is laid out and commented, within whatever your standards already require.
- The order in which you are asked for the things only you can supply.

**If you want something not already in the template, you can have it.** Our promise is specified in
*What you end up with*, together with *The same behavior every time, not the same structure*. The
checks under *How you validate the template's output* are how you confirm you got it in any given
build. Asking for something different is an extension of the basic build, and you are welcome to
experiment — the result is then your responsibility, and the checks above may no longer describe what
you have.

### What the template does not do

- **Fixing the sort title on records already in the table when the build runs.** Both routes cover
  what happens from the build forward; neither retroactively recomputes an existing record's sort
  title. A one-time pass to correct existing rows is yours to run, if you need it, the first time this
  is built into a database that already has data.
- **Catching a title edited through a route the VBA route doesn't cover.** Named plainly under
  *Business Rule 2* and in the warning above — this is the disclosed trade-off of that route, not
  something this template works around.
- **Enforcing title uniqueness.** The schema doesn't ask for it, and this template doesn't add it —
  two publications with the same title are two separate, valid rows.
- **Validating the shape of `ISBN` or any other free-text field.** Only the two rules above are
  enforced; everything else in the schema stays exactly as free-form as the table template left it.

---

## Information and conditions you need to supply

Six things, and nothing here is guessed on your behalf:

1. **Whether you are trying this out on a new database or building it into one you already use.** The
   two behave differently from the first step onward.
2. **Which file holds the tables**, if your database is split.
3. **Whether you have a backup**, if this is a database you already use. If you tell the template
   there is no backup, the build stops rather than continuing.
4. **Which route you want for Business Rule 2 — the Data Macro or VBA.** Read the trade-off under
   *Business Rule 2* above before answering; this isn't a preference with no consequence; it's a
   choice about what you actually get.
5. **Whether your noise-word list is just "The", "A", "An" or carries house-specific additions.** Per
   the declared house assumption above — say now if your collection needs its own list.
6. **Permission to change your tables**, asked immediately before anything is built.

**The accdb must be in a Trusted Location.** Code in an Access file outside one does not run at all,
and Access does not always say so — the failure can show up as the database reporting it cannot find
a procedure that is, in fact, right there. Trust is a setting on each machine, not a property of the
file.

---

## Standards Layer

**Your standards layer decides how this is built. This template decides only what it has to do.**

- **`audit-columns`** — every table this template touches already carries the four standard columns
  per the schema's own standards layer; this template adds nothing new to that set.
- **`naming-conventions`** — what any settings object or new module is called, within whatever this
  layer already requires.
- **`error-handling`** — applies to the build-time procedures, not to a Validation Rule refusal or a
  Data Macro's own `RaiseError` action. See *The same behavior every time* above for what a developer
  actually sees instead.
- **`query-style`** — how any generated code writes and holds its SQL.
- **`design-principles`** — how the work divides into procedures. *Free to choose alternatives* leaves
  that division open on purpose, and this layer is what it is open to.

A shop adopting this library replaces `standards/` with its own. Nothing in this file changes when
they do, which is the point of keeping the two apart.

---

## To the AI assistant building this

**Two sections are the specification, and only those two: *What you end up with* and *The same
behavior every time, not the same structure*, both under that heading above.** Build a system that
satisfies every promise there and passes every entry under *How you validate the template's output*.
Then run every one of those checks yourself, on a copy, and record what each one did — in the build
record, written before you report the build finished, never afterward and never only when asked for
it. The checks bind you twice: the build has to pass them, and you have to run them. How you build is
yours to decide, within *Free to choose alternatives*. Which checks you run is not — all of them,
every build. Every other section in this file — *Intent*, *What the template does not do* — is
context for reading those two. None of it binds on its own, and nothing that binds is stated only
there.

- **Business Rule 2 offers a developer-facing choice between two named mechanisms, and Business
  Rule 3 names none.** For Rule 2, ask which route under *Information and conditions you need to
  supply*, item 4 — never infer it, never default it, and never treat "which is preferred" as
  permission to pick it without asking. For Rule 3, either a table-level Validation Rule or a Data
  Macro satisfies the promise, and that choice sits on the *Free to choose alternatives* list instead.
- **Read `catalog-schema.md`, the table template this realizes, for the fields, the full text
  of all five Business Rules, and the reasoning behind them.** This file restates the outcome of two
  of the five; that file is where the field names, types, and the other three live.
- **Where the Data Macro route is chosen, read `templates/_materialization.md` for the shape of the
  document** — a Data Macro can only be created as an XML document loaded into the table, the shape of
  that document is fixed by the platform, and there is nothing in it for you to decide. Read it for
  the shape only; how the work divides and what things are called stay exactly as declared under
  *Free to choose alternatives*.
- **Read every file in `standards/` and apply it.** Naming, audit columns, error handling, query
  style, and how the work divides into procedures all come from there and never from this file.
- **Ask for the six things under *Information and conditions you need to supply*,** one at a time,
  through the interactive selection control where the answer is a choice and as a question phrased in
  plain language where it is a name. Two of them are gates: a database in real use with no backup
  stops the build, and permission to change the tables is asked immediately before anything is
  changed.
- **After the sixth thing and before you present the design, offer the `Explore options` step**
  (`_template-schema.md` §12.5) over the list under *Free to choose alternatives* — Business Rule 3's
  mechanism and the generic items, nothing more. Business Rule 2's route is not on that list; it was
  already asked directly, and a pick made through `Explore options` is recorded the same way, holds
  for the rest of the run, and is restated in the design you present.
- **Never infer an answer that belongs to the developer** — not from what the database looks like, not
  from reasoning that makes an answer seem obvious. Where a check exists to answer a question, run the
  check at the point the sequence calls for it rather than working the answer out yourself.
- **Surface the house assumption and the warning in the front matter** and get the developer's answer
  on each before building.
- **The build record reports against *How you validate the template's output*, one entry per check,
  each saying what was done and what was observed.** Passed and not passed are the only outcomes,
  including where the first method to run a check hits an obstacle — see `_template-schema.md` §12.2
  for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every entry opens with, and what
  to do before settling for a soft result. **Check 4 is the one exception stated in advance**: where
  the VBA route was chosen, its expected result is stated in the check itself and recording that
  expected, disclosed outcome is what "passed" means for that entry.
- **While the build runs, do not narrate it.** Say once that it has started and what it will produce;
  say anything the developer must act on, as a question; say when it is finished, what was built, and
  where the build record is. Everything else — every procedure written, every check that passed — goes
  to the build record.
- **Once the build is reported finished, and only then, mention each entry under `related` in the
  front matter** (`_template-schema.md` §7.1) — one line per entry, what it is and why. This is not
  part of the build, never a gate, and never read before this point.

## Extra options

*Named optional extensions, none of them filled in for an engagement.*

- **A one-time pass to correct the sort title on records already in the table**, for a database built
  into an existing collection — named above as something this template does not do on its own.
- **Title uniqueness**, if a collection wants to flag or prevent duplicate titles — named above as
  outside this template's promise.
