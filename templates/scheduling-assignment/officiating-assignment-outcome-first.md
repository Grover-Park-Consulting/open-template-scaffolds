---
template: officiating-assignment-outcome-first
title: Officiating Assignment — outcome-first method
domain: scheduling-assignment
type: outcome-first
version: 0.1.0
status: draft
implements: officiating-assignment-schema
standards_layer:
  - audit-columns
  - naming-conventions
  - error-handling
  - query-style
  - design-principles
house_assumptions:
  - "Official age — computed as of today's date, wherever it's shown, not as of any particular
    game date. A league that needs age-as-of-a-game-date for an eligibility rule changes this — the
    schema doesn't ask for that today, so this template doesn't build it."
  - "Game play level — read through its home team specifically, not its away team, matching the
    paired scaffold's own convention. A league whose two teams in a game could genuinely carry
    different play levels needs a different rule than 'read the home team's' — this template assumes
    they always match, per the schema's own declared assumption."
warnings:
  - Choosing the VBA route for Business Rule 3 (see below) means an inactive official can still be
    written into an assignment through any route other than your assignment procedure — a direct edit
    to tblGameOfficial, an import. This is a disclosed trade-off of that route, not a defect; the Data
    Macro route does not have this gap.
  - This template attaches a Data Macro to a live table where the Data Macro route is chosen for
    Business Rule 3. A build against a database in real use is preceded by a backup copy of the file,
    and the developer is asked for one before anything is changed.
related:
  - "app-startup-outcome-first — worth building first, or alongside, if you haven't: it gives your
    application the general folder-ensuring routine that Business Rule 9's photo folder can share,
    instead of a second one built just for photos. Its scaffold sibling, app-startup-scaffold,
    produces the same result from working code if you'd rather build that way."
---

# Officiating Assignment — outcome-first method

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

**This template realizes seven of the nine Business Rules the table template states, and leaves the
other two alone.** `officiating-assignment-schema` builds the ten tables and states, in its own
Business Rules, what has to be true once the database is in use. Two of those nine are already
complete the moment the schema is built — the crew is a junction, not a set of hardcoded columns
(Rule 1), and assignment uniqueness (Rule 2) is enforced by the junction's own unique indexes. There is
no route to specify for either, and nothing this template adds changes them. The other seven — only
active officials may be assigned (Rule 3), a team can't play itself (Rule 4), a game's play level is
derived (Rule 5), the applicable pay rate (Rule 6), game times are ordered (Rule 7), an official's age
is derived (Rule 8), and photos as a file name plus a shared folder setting (Rule 9) — each need
something built on top of the tables to hold, and this template states each of those seven as an
outcome and leaves the route to whoever builds it, within the choices named below. It assumes the
schema's tables already exist, or are built as part of this same run.

**Nothing else is here.** No procedures, no module names, no code beyond what's needed to reason
through the one choice named for Business Rule 3.

---

## What you end up with

### Business Rule 3 — only active officials

**An assignment can never reference an official whose `OfficialIsActive` is not Yes — not through
your assignment procedure, not through a direct edit to `tblGameOfficial`, not through an import.**

**Two routes exist to build this, and which one you pick changes what you actually get — so you're
asked, not defaulted**, the same choice and the same reasoning as the sort-title decision in
`catalog-outcome-first`:

- **The Data Macro route (preferred).** Attached to `tblGameOfficial`, it looks up the referenced
  official's `OfficialIsActive` on every route into the table and refuses the write if that official
  isn't active — the assignment form, a direct edit to the table, an import, all of it.
- **The VBA route.** Simpler to build, and it's what the paired scaffold's `ValidateAssignment`
  already does — coverage stops at whatever calls it. An assignment inserted any other way silently
  goes in against an inactive official, and nothing catches it. That gap is a known, accepted
  trade-off of this route, not something to be fixed later — choose it only if you're comfortable
  living with it, and know that it's the same shape as what already ships in this domain's scaffold.

### Business Rule 4 — a team can't play itself

**A game can never be saved with the same team as both `HomeTeamID` and `AwayTeamID`, whatever the
route the save was made through.**

**No mechanism is named here.** This asks nothing about what changed — only whether the row, as it
now stands, is internally consistent — and a table-level Validation Rule answers exactly that
question on every route into the table. A Data Macro could enforce the same thing and costs more to
build for no additional coverage. Either is a legitimate build.

### Business Rule 5 — a game's play level is derived

**A game's play level is always read through its home team, never stored on the game itself, and it's
always current.** Change a team's play level, and every game that team is the home team for reflects
the new level immediately, with no separate step to bring anything up to date.

**No mechanism is named here.** A saved query, a VBA function, and an Access calculated field all
satisfy this promise equally; nothing about the platform forces one over the others.

### Business Rule 6 — the applicable pay rate

**For a given assignment, the pay rate is the `tblPositionRate` row for the game's play level and the
assignment's position with the latest `EffectiveDate` on or before the game's date.** Where more than
one rate row could apply, the most recent one as of the game date wins — never the newest row by
insertion order, never the highest rate.

**Where no matching rate row exists, nothing is guessed.** The absence is surfaced as a visible gap —
a warning, a blank, something a person notices — never silently treated as a $0 rate that reads as
"this position works for free."

**No mechanism is named here.** This is ordinary lookup logic — find the applicable row, given a date
— and nothing about the platform forces one particular way of writing it.

### Business Rule 7 — game times are ordered

**A game can never be saved with `GameEnd` earlier than or equal to `GameStart`, whatever the route
the save was made through — but only when `GameEnd` is actually present.** A game with no end time
recorded yet saves without complaint.

**No mechanism is named here**, for the same reason as Business Rule 4: a table-level Validation Rule
answers this on every route, and a Data Macro is a legitimate but unnecessary alternative.

### Business Rule 8 — an official's age is derived

**An official's age is always computed from `BirthDate` at the moment it's needed, and never stored
anywhere on `tblOfficial` or written to any other table.**

**No mechanism is named here.** A saved query, a VBA function, and an Access calculated field all
satisfy this promise equally.

### Business Rule 9 — photos are a file name plus one shared folder setting

**Picking a photo for an official copies the chosen file into one shared folder under a controlled
name, and the official's own record stores only that name — never a full path, never the file
itself.** Two officials never collide on the same stored name, and re-picking a photo for the same
official replaces the file cleanly rather than leaving the old one behind, orphaned.

**The folder itself is confirmed before anyone can pick a photo, never silently created somewhere
local.** This is a **shared** folder — `tblOfficial.PhotoFileName` lives in a table every front end
reads, so the file it names has to be reachable from all of them, not just the one that added it.
Creating a local folder instead of confirming the shared one **succeeds with no error**, and the
symptom is worse than a crash: the person who added the photo sees it, nobody else does.

**No mechanism is named for the folder-and-copy pattern itself, because it is not specific to this
domain.** `templates/_materialization.md` → *External file assets* already states it in general terms
— confirm-don't-create for a shared folder, and copy the picked file in under a controlled name before
storing that name — and this template's promise is exactly that pattern applied to
`tblOfficial.PhotoFileName` and the `OfficialPhotoFolder` row of `tblAppSetting`. See *Free to choose
alternatives* for reusing folder-ensuring logic this database may already have.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

**Do these on a copy, and stop if you do not have one.** The checks add, change, and delete records
in your own tables — that is what they test, and there is no version of them that leaves your data
alone. Make the copy once the build is finished, run every check on it, and keep your working file
out of it. Where your database is split into two files, copy both and keep them together.

1. **An assignment against an inactive official is refused through the normal route.**
   - Mark an official inactive
   - Try to assign that official to a game through your assignment procedure
   - Confirm the assignment is refused
2. **An assignment against an inactive official through a direct table edit — only if you chose the
   Data Macro route.**
   - With the same inactive official, add a row to `tblGameOfficial` directly, not through the
     assignment procedure
   - Confirm the write is refused
   - If you chose the VBA route instead, this check is not expected to pass — that's the disclosed
     trade-off of that route, confirmed by *Business Rule 3*'s own text above, not a failure to
     report as one
3. **A game can't be saved against itself.**
   - Try to save a game with the same team as both home and away, once through a form or the table
     directly and once through an append/import query
   - Confirm neither save succeeds
4. **A game with two different teams saves normally.**
   - Save a game with two distinct teams
   - Confirm it saves
5. **A game's play level moves with its home team.**
   - Note a team's play level and a game where that team is the home team
   - Change the team's play level
   - Confirm the game's shown play level reflects the change immediately, with no separate step
6. **The applicable pay rate resolves to the most recent effective row.**
   - Create two `tblPositionRate` rows for the same play level and position, with different
     `EffectiveDate` values, both on or before a game's date
   - Confirm the resolved rate is the one from the row with the later `EffectiveDate`, not the row
     added most recently or the higher of the two rates
7. **A missing rate is surfaced, never guessed.**
   - Create an assignment for a play level and position that has no `tblPositionRate` row at all
   - Confirm the absence is visibly reported — not a $0 rate presented as though it were a real one
8. **Game times out of order are refused.**
   - Try to save a game with `GameEnd` earlier than or equal to `GameStart`, whatever route you use
   - Confirm the save does not succeed
9. **A game with no end time saves without complaint.**
   - Save a game with `GameStart` set and `GameEnd` left empty
   - Confirm it saves
10. **A game with `GameEnd` after `GameStart` saves normally.**
    - Save a game with a valid, later `GameEnd`
    - Confirm it saves
11. **An official's age is correct and not stored.**
    - Compare two officials with different `BirthDate` values
    - Confirm wherever age is shown, it's correct for each, and confirm `tblOfficial` carries no
      separate age column that could drift out of step with `BirthDate`
12. **The shared photo folder is confirmed, never silently created locally.**
    - Point `tblAppSetting.OfficialPhotoFolder` at a shared path that does not currently exist
    - Try to pick a photo for an official
    - Confirm you are told the folder can't be reached — not shown a folder that was quietly created
      on your own machine
13. **Picking a photo copies it into the managed folder under a controlled name.**
    - Point the setting at a folder that exists, and pick a photo for an official
    - Confirm the file appears in that folder under a name that isn't the file's own original name,
      and that the official's record stores only that name
14. **Re-picking a photo replaces it cleanly.**
    - Pick a different photo for the same official
    - Confirm the official's record now points at the new file, and the old one isn't left behind
      under its old name with nothing pointing at it
15. **Changing the folder setting changes where the next photo lands.**
    - Change `tblAppSetting.OfficialPhotoFolder` to a different, valid shared folder
    - Pick a photo for an official
    - Confirm the file lands in the new folder, not the old one
16. **Running the build again does not duplicate anything.**
    - Run whatever attaches the Data Macro(s) or Validation Rules a second time
    - Confirm there is still exactly one copy of each
17. **When a build is blocked, it stops rather than leaving tables half-finished.**
    - Leave a table or form open in the database
    - Start the build
    - Confirm it stops, names what is open, and changes nothing

### The same behavior every time, not the same structure

For an AI-assisted template, two builds need not produce the same code. They need to produce a
database that behaves the same way. The following must be true of every build:

- **Which route was chosen for Business Rule 3 is stated plainly in the design and the build
  record**, not left for the developer to discover by testing. They're choosing a coverage tier, and
  they need to know which one they got.
- Where the Data Macro route was chosen for Rule 3, the active-official check happens inside the Data
  Macro itself, not in code the macro calls out to.
- The pay-rate lookup (Rule 6) reads `tblPositionRate`, never `tblGame` or `tblGameOfficial`, for the
  applicable row — it answers "what does the rate table say applied on this date," not anything about
  what was previously paid.
- The photo-folder check (Rule 9) verifies a shared folder; it never creates one silently in its
  place. Where the folder can't be reached, the person is told, not shown a locally-created
  substitute.
- **[your standards]** The error-handling pattern applies to the build-time procedures, not to a
  Validation Rule refusal or a Data Macro's own `RaiseError` action. When either refuses a save,
  Access shows its own dialog naming the problem — no VBA runs, no `On Error` handler fires, and
  nothing is logged unless something else is separately watching.
- Running the build again replaces what a previous build attached; it never adds a second copy
  alongside it.

Lines marked **[your standards]** come from your standards layer rather than from this template, and
move with that layer if your shop replaces it. Everything else in `standards/` applies here as it does
to every template.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences, say so while
the design is being worked out, before anything is built. Where you don't choose, the build will
choose based on the rules built into it. The template's promise holds either way.

- **How Business Rules 4 and 7 are enforced** — a table-level Validation Rule or a Data Macro, for
  each independently. Either satisfies the promise; the Validation Rule is the simpler build for
  both. (Business Rule 3's mechanism is *not* on this list — it's asked directly, under *Information
  and conditions you need to supply*, because the two routes produce different outcomes and the
  developer needs to choose knowingly.)
- **How Business Rules 5 and 8 are exposed** — a saved query, a VBA function, or an Access calculated
  field, for each independently.
- **Whether Business Rule 9's folder-ensuring logic is its own routine or folds into one this database
  already has.** If you've already built `app-startup` (or anything else that confirms/creates
  folders at open time), extend that routine to cover the photo folder instead of building a second,
  parallel one. If you haven't, this template's own build creates one for the photo folder alone.
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

- **Catching an inactive-official assignment through a route the VBA route doesn't cover.** Named
  plainly under *Business Rule 3* and in the warning above — this is the disclosed trade-off of that
  route, not something this template works around.
- **Age-based eligibility.** Business Rule 8 only derives an official's age; nothing here checks it
  against a play level's minimum or maximum. A league that needs that adds it as its own rule.
- **Correcting a game whose home and away teams carry genuinely different play levels.** The
  house assumption above declares they're expected to match; a league where they legitimately
  wouldn't needs a different rule for Business Rule 5 than "read the home team's."
- **Anything about who's qualified to work which play level or position.** Named in the schema's own
  *Extra Options* as a future extension; this template doesn't model it.
- **A screen for reviewing pay owed, or reconciling it against what was actually paid.** These are
  tables and queries. You can add your own form or report for them.

---

## Information and conditions you need to supply

Five things, and nothing here is guessed on your behalf:

1. **Whether you are trying this out on a new database or building it into one you already use.** The
   two behave differently from the first step onward.
2. **Which file holds the tables**, if your database is split.
3. **Whether you have a backup**, if this is a database you already use, and if you choose the Data
   Macro route for Business Rule 3. If you tell the template there is no backup, the build stops
   rather than continuing.
4. **Which route you want for Business Rule 3 — the Data Macro or VBA.** Read the trade-off under
   *Business Rule 3* above before answering; this isn't a preference with no consequence, it's a
   choice about what you actually get.
5. **Permission to change your tables**, asked immediately before anything is built.

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
- **`query-style`** — how any generated code writes and holds its SQL, notably the pay-rate lookup and
  the derived play-level/age queries.
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

- **Business Rule 3 offers a developer-facing choice between two named mechanisms; Business Rules 4
  and 7 each name none; Business Rules 5 and 8 each name none; Business Rule 9 names a pattern, not a
  domain-specific mechanism.** For Rule 3, ask which route under *Information and conditions you need
  to supply*, item 4 — never infer it, never default it. For Rules 4 and 7, either a table-level
  Validation Rule or a Data Macro satisfies each independently, and that choice sits on the *Free to
  choose alternatives* list. For Rules 5 and 8, a query, a function, or a calculated field all
  satisfy each independently, also on that list.
- **Read `officiating-assignment-schema.md`, the table template this realizes, for the fields,
  the full text of all nine Business Rules, and the reasoning behind them.** This file restates the
  outcome of seven of the nine; that file is where the field names, types, and the other two live.
- **You may read `officiating-assignment-scaffold.md` for one worked decomposition — nothing
  more.** It shows one way to structure `AssignOfficial`, `ValidateAssignment`, and
  `GetApplicablePayRate`, and its `ValidateAssignment` is the VBA route named under *Business Rule 3*
  — read it to see exactly what that route's coverage gap looks like, not as a route this file
  adopts. A different, equally valid decomposition that still satisfies every check here is a
  legitimate build.
- **For Business Rule 9, read `templates/_materialization.md` → *External file assets* for the
  confirm-don't-create and copy-in pattern.** It is written in general terms on purpose — apply it to
  `tblOfficial.PhotoFileName` and `tblAppSetting.OfficialPhotoFolder` specifically. Do not read
  `app-startup-scaffold.md` or `app-startup-outcome-first.md` for this — they are sibling consumers of
  the same pattern, not its source, and reading a sibling domain template here risks importing a
  decomposition this file was never given license to import.
- **Where the Data Macro route is chosen for Business Rule 3, read `templates/_materialization.md` for
  the shape of a Data Macro** — it can only be created as an XML document loaded into the table, the
  shape of that document is fixed by the platform, and there is nothing in it for you to decide.
- **Read every file in `standards/` and apply it.** Naming, audit columns, error handling, query
  style, and how the work divides into procedures all come from there and never from this file.
- **Ask for the five things under *Information and conditions you need to supply*,** one at a time,
  through the interactive selection control where the answer is a choice and as a question phrased in
  plain language where it is a name. Two of them are gates: a database in real use with no backup
  stops the build where the Data Macro route was chosen for Rule 3, and permission to change the
  tables is asked immediately before anything is changed.
- **After the fifth thing and before you present the design, offer the `Explore options` step**
  (`_template-schema.md` §12.5) over the list under *Free to choose alternatives*, and nothing outside
  it. Business Rule 3's route is not on that list; it was already asked directly, and a pick made
  through `Explore options` is recorded the same way, holds for the rest of the run, and is restated
  in the design you present.
- **Never infer an answer that belongs to the developer** — not from what the database looks like, not
  from reasoning that makes an answer seem obvious. Where a check exists to answer a question, run the
  check at the point the sequence calls for it rather than working the answer out yourself.
- **Surface both house assumptions and both warnings in the front matter** and get the developer's
  answer on each before building.
- **The build record reports against *How you validate the template's output*, one entry per check,
  each saying what was done and what was observed.** Passed and not passed are the only outcomes,
  including where the first method to run a check hits an obstacle — see `_template-schema.md` §12.2
  for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every entry opens with, and what
  to do before settling for a soft result. **Check 2 is the one exception stated in advance**: where
  the VBA route was chosen, its expected result is stated in the check itself, and recording that
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

- **Age-based eligibility**, named above as something this template does not do — checking a computed
  age against a play level's minimum or maximum before allowing an assignment.
- **Official qualifications by play level**, already named in the schema's own *Extra Options* — this
  template doesn't model it either.
