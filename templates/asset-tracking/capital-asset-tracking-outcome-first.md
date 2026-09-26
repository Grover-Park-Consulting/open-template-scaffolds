---
template: capital-asset-tracking-outcome-first
title: Capital Asset Tracking — outcome-first method
domain: asset-tracking
type: outcome-first
version: 0.2.1
status: stable
implements: capital-asset-tracking-schema
standards_layer:
  - audit-columns
  - naming-conventions
  - error-handling
  - query-style
  - design-principles
house_assumptions:
  - "Capitalization threshold — held as a single configurable value, not one per asset category.
    A district whose policy sets a different floor for land or land improvements than for equipment
    changes this — see *Information and conditions you need to supply*."
  - "The capitalization threshold value — read from wherever it is stored at the moment each save is
    evaluated, never baked into the check as a literal, so a district that changes its threshold
    later does not need this template rebuilt for the new figure to take effect."
  - "Disposed — resolved by name against tlkpAssetStatus at build time rather than a hardcoded ID,
    because seed order is not guaranteed to put any status at a particular ID, on this build or the
    next."
warnings:
  - This template attaches Data Macros to live tables (Business Rules 1 and 3), which is why backup
    is item 3 under *Information and conditions you need to supply* rather than a separate ask —
    surfacing this warning and asking that item are the same step, not two.
  - A historical or legacy import — assets acquired years ago, under a threshold since raised — is
    refused by the capitalization check exactly as any other insert would be, because the check has
    no way to tell a backdated record from a new purchase that falls short today. This is item 5
    under *Information and conditions you need to supply*; surfacing this warning and asking that
    item are the same step, not two. This template does not supply a way around the check; see
    *What the template does not do*.
related:
  - "stocktake-scan-outcome-first — a similar process for a different purpose: reconciling
    a table of items against barcode scans. The purpose of stocktake-scan is to account for **sale
    inventory** rather than **fixtures and equipment**. The scan-resolution logic differs between the
    two and they solve different problems, so they do not share a scan-resolution mechanism."
---

# Capital Asset Tracking — outcome-first method

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

**This template realizes five of the eight Business Rules the table template states, and leaves the
other three alone.** `capital-asset-tracking-schema` builds the fourteen tables and states, in
its own Business Rules, what has to be true once the database is in use. Three of those eight are
already complete the moment the schema is built — barcode uniqueness (Rule 2), required ownership
fields (Rule 7), and the direct-join reporting shape (Rule 8) are properties of the tables themselves,
enforced by an index, a `Required` flag, or the shape of a foreign key. There is no route to specify
for any of them and nothing this template adds changes them. The other five — the capitalization
threshold (Rule 1), the standing audit trail (Rule 3), computed depreciation (Rule 4), the
required disposal date (Rule 5), and inventory-scan resolution (Rule 6) — need something built on top
of the tables to hold, and more than one way exists to build it. This template states each of those
five as an outcome and leaves the route to whoever builds it. It assumes the schema's tables already
exist, or are built as part of this same run.

**Nothing else is here.** No procedures, no module names, no code beyond the two places the platform
forces a specific mechanism, named below and reasoned through at the point each is named.

---

## What you end up with

### Business Rule 1 — the capitalization threshold

**An asset cannot be saved on file at or under the district's capitalization threshold — not through
a form, not through a direct edit to the table, not through an import, not through a query that
updates a thousand rows at once.** Whatever the route, the same figure is checked and the same result
follows: below or at the threshold, the save does not happen.

**The check applies when the cost itself is what changed, not to every save of a row that happens to
carry a low cost.** This is the distinction that keeps the rule honest over time. A district that
raises its threshold from $400 to $500 does not retroactively touch the desk it bought last year for
$450 — that asset stays on file, exactly as it was capitalized, and a change to its room or its
custodian next month must not be blocked by a cost that was never in question. **Where a check would
re-evaluate a row's cost on every save regardless of whether the cost changed, a threshold increase
would silently start blocking edits to assets that were correctly capitalized under the rules in
force when they were acquired** — the asset's own history and location would become unchangeable
until someone either raised its recorded cost or reclassified it, neither of which reflects anything
that actually happened to the asset. The check exists to catch a cost that should never have been
accepted, not to keep re-litigating one that already was.

**This is why the mechanism is named rather than left open.** A row-level Validation Rule — the
platform's ordinary way to refuse a save that fails a condition — evaluates the row as it stands at
save time; it cannot tell a save that changed `AcquisitionCost` from one that left it untouched,
because a Validation Rule has no access to what the field held a moment before. **A Data Macro can**:
it sees both the old and new value of every field on the row it fires for, on every route into the
table, which is what lets the check ask the only question that actually matters — *did the cost
change, and if so, does the new figure clear the threshold* — instead of the question a Validation
Rule is limited to asking, which would eventually punish every asset the district owns for having
been acquired before the last time the threshold went up. Nothing else on this platform can tell those
two questions apart while still firing on every route into the table.

### Business Rule 3 — the standing audit trail

**Every change to an asset's room, custodian, department, funding source, or status writes a row to
the movement/change history, whoever makes the change and however they make it.** A move recorded by
editing the asset's form, a bulk reassignment run as an update query, someone typing a new room number
directly into the table — all of it lands in the same history, with the prior value, the new value,
the change type, and when it happened.

**This is independent of, and in addition to, the row-level audit columns your standards layer
supplies.** The two answer different questions: the audit columns say who last touched the record and
when; the history says what specifically changed, one row per change, going back to the asset's
acquisition. Keeping both is the normal result, and this template does not ask you to choose between
them.

**The mechanism is named for the same reason it is named in Business Rule 1.** A Data Macro attached
to `tblAsset` fires on every route into the table and can compare old and new values to decide whether
a history row is owed; nothing else on the platform does both. Front-end code fires only through
whatever screen carries it, and the schema's own text already names the consequence: a form, a query,
a second form somebody writes later — any of them can move the asset and leave the trail silently
incomplete. For a district asset register, where the point of the trail is accountability, that gap is
not an acceptable trade against the convenience of writing the logic in VBA instead.

**This one must be an `AfterUpdate` Data Macro, not a `BeforeChange` one — unlike Business Rule 1's
check.** Writing the history row needs `CreateRecord` against `tblAssetHistory`, and `CreateRecord`
is not supported from a `BeforeChange` event; it fails at run time with error 3873
(`templates/_materialization.md`, the Data Macro platform facts). Comparing old and new values to
decide *whether* a row is owed works the same way in `AfterUpdate` as it would in `BeforeChange` —
both see `[Old]` and the current row — so nothing about the comparison logic is lost by the
different event; only the event itself has to change.

### Business Rule 4 — depreciation, computed not stored

**The current book value shown for an asset is always current — it reflects whatever the acquisition
cost, useful life, and salvage value say right now, and it is never a number written to the table at
some earlier point and left to go stale.** Change any of the three inputs and the value shown changes
with them, without anyone running a recalculation step first.

**It is floored at the salvage value and frozen once the useful life has elapsed.** An asset five
years past the end of its useful life shows the same book value it showed on the day the useful life
ran out — it does not keep depreciating into a negative number, and it does not need anyone to notice
and intervene.

**No mechanism is named here.** A saved query, a VBA function called wherever the value is needed, and
an Access calculated field all satisfy this promise; nothing about the platform forces one over the
others, and the choice is yours under *Free to choose alternatives*.

### Business Rule 5 — the required disposal date

**An asset cannot be marked Disposed without a disposal date, whatever the route the change was made
through.** Setting the status field to Disposed and leaving the disposal date empty does not save,
whether that happens on a form, in the table directly, or through a query.

**No mechanism is named here.** Unlike Business Rule 1, this check asks nothing about what changed —
it asks whether the row, as it now stands, is internally consistent. A row-level Validation Rule
answers exactly that question and answers it on every route into the table, which is everything the
promise needs; a Data Macro could enforce the same thing and costs more to build for no additional
coverage. Either is a legitimate build.

### Business Rule 6 — inventory-scan resolution

**A scanned barcode is resolved to an outcome the moment it is recorded, without anyone having to
interpret the scan by hand.** Scan a barcode that matches an asset on file, expected at the site being
counted, and it resolves as Matched. Scan a barcode for an asset whose on-file location belongs to a
different site, and it resolves as Unexpected. Scan the same barcode a second time within the same
session, and the second scan resolves as Duplicate.

**A scanned asset found in a different room than its on-file room is identifiable as misplaced without
changing the asset's own record.** The scan records where it was actually found; the asset's own
`RoomID` is not silently updated by the act of scanning it — reconciling a misplaced asset onto its
new room is a decision for whoever runs the audit, not something a scan does on its own.

**An asset expected at the site and never matched during the session is identifiable by a query against
that session, not by a stored flag anywhere.** Nothing marks an asset "missing" in the table itself;
what makes it reportable as missing is that no Matched scan for it exists in a finished session where
it was expected.

**No mechanism is named here.** This is ordinary procedural logic — resolving a scanned value against
what is on file — and nothing about the platform forces one particular way of triggering or structuring
it. A form's `AfterUpdate` event calling a resolution procedure scan by scan, and a batch procedure run
against a table of raw, unresolved scans, both satisfy every promise above.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

**Do these on a copy, and stop if you do not have one.** The checks add, change, and delete records in
your own tables — that is what they test, and there is no version of them that leaves your data alone.
Make the copy once the build is finished, run every check on it, and keep your working file out of it.
Where your database is split into two files, copy both and keep them together.

1. **A new asset below the threshold is refused.**
   - Add a new asset with an acquisition cost below the district's threshold — do this once through a
     form or the table directly, and once through an append/import query
   - Confirm neither save succeeds
2. **A new asset exactly at the threshold is refused.**
   - Add a new asset with an acquisition cost exactly equal to the threshold
   - Confirm the save does not succeed — the rule is "must exceed," not "must reach"
3. **A new asset above the threshold succeeds.**
   - Add a new asset with an acquisition cost above the threshold
   - Confirm it saves normally
4. **A grandfathered asset can still be edited.**
   - Find or create an asset whose acquisition cost is at or below the district's *current* threshold
     (representing one capitalized before the threshold was last raised)
   - Change a field on it that is not the acquisition cost — its room, say
   - Confirm the save succeeds and is not blocked by the cost check
5. **Lowering an existing asset's cost below the threshold is refused.**
   - Take an asset whose cost is above the threshold and edit its acquisition cost down to at or below
     the threshold
   - Confirm the save does not succeed
6. **A tracked field change is recorded as one history row.**
   - Change one of `RoomID`, `CustodianID`, `DepartmentID`, `FundingSourceID`, or `AssetStatusID` on
     one asset
   - Open the history to confirm it has one new row for that asset naming the change type, the prior
     value, the new value, and when it happened
7. **An untracked field change writes no history row.**
   - Change `AssetNotes` (or another field not named in Business Rule 3) on an asset
   - Confirm no new row appears in the history
8. **Computed book value moves with its inputs.**
   - Note an asset's current book value
   - Change its acquisition cost, useful life, or salvage value
   - Confirm the book value shown changes accordingly, without running any separate recalculation step
9. **Computed book value is floored and freezes.**
   - Take an asset whose useful life has already elapsed
   - Confirm its book value equals its salvage value, not a negative number
   - Wait, or adjust the acquisition date, so more time passes
   - Confirm the book value has not changed further
10. **Disposal without a date is refused.**
    - Set an asset's status to Disposed and leave the disposal date empty
    - Confirm the save does not succeed, whatever the route
11. **A disposed asset keeps its history.**
    - Dispose of an asset that has prior history rows and prior inventory-scan rows
    - Confirm the asset, its history, and its scan history are all still present and readable
12. **A matched scan resolves as Matched.**
    - Start an inventory-audit session for a site
    - Scan the barcode of an asset on file, expected at that site
    - Confirm the scan resolves as Matched
13. **An out-of-site scan resolves as Unexpected.**
    - In the same session, scan the barcode of an asset whose on-file room belongs to a different site
    - Confirm the scan resolves as Unexpected
14. **A repeated scan resolves as Duplicate.**
    - Scan the same barcode a second time in the same session
    - Confirm the second scan resolves as Duplicate and the first is unchanged
15. **A misplaced asset is identifiable without altering its record.**
    - Scan an asset's barcode while recording a room different from its on-file room
    - Confirm the scan records the room it was actually found in
    - Confirm the asset's own `RoomID` is unchanged
16. **A missing asset is identifiable by query, not by a stored flag.**
    - Close a session having left at least one site-expected, Active asset unscanned
    - Confirm nothing on that asset's own record marks it missing
    - Confirm a query against the session identifies it as expected and unmatched
17. **Running the build again does not duplicate anything.**
    - Run whatever attaches the Data Macros and builds the supporting objects a second time
    - Confirm there is still exactly one copy of each — no duplicated Data Macro actions, no second
      settings row, no second copy of any generated procedure
18. **When a build is blocked, it stops rather than leaving tables half-finished.**
    - Leave a table or form open in the database
    - Start the build
    - Confirm it stops, names what is open, and changes nothing

### The same behavior every time, not the same structure

For an AI-assisted template, two builds need not produce the same code. They need to produce a
database that behaves the same way. The following must be true of every build:

- **The capitalization check and the history-writing both happen inside the Data Macro itself, not in
  code the macro calls out to.** A Data Macro whose only action hands the work to a function elsewhere
  has moved the behaviour out of the database engine and into code that a datasheet edit, an update
  query, or any other route can skip. The one exception, as in Business Rule 3's reasoning above, is
  reading the district's current threshold value, which may live in a table the macro looks up.
- **The capitalization check compares old and new `AcquisitionCost`, never the row's current value in
  isolation.** A build that re-checks every save regardless of what changed has reproduced the
  grandfathering fault this template exists to avoid, even though it would pass a check run only on
  new inserts.
- One history row per field that actually changed, for the five tracked fields only. A save that
  leaves all five untouched writes nothing.
- The computed book value is never written to a stored column on `tblAsset`. Wherever it is exposed —
  a query, a function, a calculated field — recomputing it never requires a rebuild.
- A scan's resolution never writes to `tblAsset.RoomID`. The asset's own on-file location changes only
  when someone deliberately reconciles it, never as a side effect of a scan being recorded.
- **[your standards]** Whatever supplies the current capitalization threshold and resolves "Disposed"
  by name is reachable from every file a person edits through, not only the file holding the tables.
  Where the database is split, that means it exists in both the back end and in every copy of the
  front end.
- **[your standards]** Nothing in the capitalization check, the history-writing, or the disposal check
  interrupts a person while they are saving an unrelated, valid change. A message box raised for a
  save that was always going to succeed is an interruption nobody asked for.
- Running the build again replaces what a previous build attached; it never adds a second copy
  alongside it.

Lines marked **[your standards]** come from your standards layer rather than from this template, and
move with that layer if your shop replaces it. Everything else in `standards/` applies here as it does
to every template.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences, say so while the
design is being worked out, before anything is built. Where you don't choose, the build will choose
based on the rules built into it. The template's promise holds either way.

- How VBA code, where any is used, divides into procedures, what they are called, and how many there
  are. They will be functionally equivalent, not necessarily structurally the same. Where the
  capitalization check and the history-writing happen is not open — that is settled above as part of
  the outcome, not the route.
- **How computed book value is exposed** — a saved query, a VBA function, or an Access calculated
  field.
- **How the required-disposal-date rule is enforced** — a table-level Validation Rule or a Before
  Change Data Macro. Either satisfies the promise; the Validation Rule is the simpler build.
- **How a scan is resolved** — as each one is entered, through a form event, or in a batch run against
  a table of unresolved scans. Both produce the same recorded result.
- Names for any settings table or column the build adds, within whatever your naming rules already
  require.
- **Where** the capitalization threshold is stored — the storage location is the free choice here.
  **That it is read live at the moment of each check, never baked in as a literal, is not a choice**
  — that part is the declared house assumption above, restated here only as the constraint on the
  free part, not reopened.
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

- **Accommodating a historical or legacy import below today's threshold.** The capitalization check
  applies to every insert, including a bulk import of assets acquired years ago under a lower
  threshold. A migration like that needs a deliberate way around the check for that one import,
  arranged by whoever builds it — suspending the Data Macro for the duration, most likely. This
  template does not supply that route.
- **Category-specific capitalization floors.** One threshold applies to every asset, per the
  `house_assumptions` above. A district that needs a different floor for land than for equipment
  extends the source the threshold is read from — this template does not model it.
- **Reconciling a misplaced asset onto its newly found room.** A scan records where an asset was
  actually found; moving the asset's own record to match is a decision someone makes, not something
  scanning does automatically.
- **Multi-method depreciation.** Only whatever the schema's seeded `tlkpDepreciationMethod` rows
  support is computed. Adding a second method is an extension, not part of this build.
- **A screen for reviewing history, running an inventory session, or reconciling scans.** These are
  tables and queries. You can add your own forms for them.

---

## Information and conditions you need to supply

You need to provide information about six things; no guesses are made on your behalf.

1. **Whether you are trying this out on a new database or building it into one you already use.** The
   two behave differently from the first step onward.
2. **Which file holds the tables**, if your database is split.
3. **Whether you have a backup**, if this is a database you already use. If you tell the template
   there is no backup, the build stops rather than continuing.
4. **The district's capitalization threshold, and where it should be stored.** A single figure applies
   to every asset unless you tell us otherwise, per the `house_assumptions` above — say so now if your
   district needs a different floor by category, since that changes the design.
5. **Whether you have historical assets to import that fall below today's threshold.** If so, say so
   before the build — see *What the template does not do*. This template does not arrange a way
   around the check for you; you decide how that import happens.
6. **Permission to change your tables**, asked immediately before anything is built.

**The accdb must be in a Trusted Location.** Code in an Access file outside one does not run at all,
and Access does not always say so — the failure can show up as the database reporting it cannot find a
procedure that is, in fact, right there. Trust is a setting on each machine, not a property of the
file.

---

## Standards Layer

**Your standards layer decides how this is built. This template decides only what it has to do.**

- **`audit-columns`** — the four columns recording who created and last changed each row. Every table
  this template touches already carries them per the schema's own standards layer; this template adds
  nothing new to that set. The history table's own change-type, prior-value, and new-value columns are
  not audit columns and are not covered by this file.
- **`naming-conventions`** — what any settings table, its columns, and any new module are called,
  within whatever this layer already requires.
- **`error-handling`** — applies to the build-time procedures (whatever attaches the Data Macros and
  Validation Rules, and any VBA behind Rules 4 and 6), not to the Data Macro and Validation Rule
  refusals themselves. **What those look like instead:** when a Validation Rule or a Data Macro's
  `RaiseError` action refuses a save, Access shows its own dialog box on screen, naming the problem,
  and will not let the record be saved or the cursor move off it until the person fixes the offending
  field or undoes the change — the same dialog Access has always shown for a required field left
  blank; this template just adds a couple more conditions that trigger it. No VBA runs to produce
  that: no procedure is called, no `On Error` handler fires, and nothing is logged anywhere unless
  something else is separately watching. The house error-handling pattern only applies where VBA code
  is actually running, and none is, in that moment.
- **`query-style`** — how the generated code writes and holds its SQL.
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

- **Two mechanisms are named, and only two: a Data Macro for the capitalization check (Business Rule
  1) and a Data Macro for the standing audit trail (Business Rule 3).** Both are named because the
  platform leaves exactly one route that can compare a field's old and new value while firing on every
  route into the table — see the reasoning under each rule above. Nothing else in this file names a
  mechanism. Do not import one for Business Rules 4, 5, or 6 on the assumption that a promise this firm
  elsewhere must mean the same everywhere; it doesn't.
- **Read `capital-asset-tracking-schema.md`, the table template this realizes, for the fields,
  the full text of all eight Business Rules, and the reasoning behind them.** This file restates the
  outcome of five of the eight; that file is where the field names, types, and the other three live.
- **Read `templates/_materialization.md` for the shape of a Data Macro**, the same document the audit
  domain's outcome-first template points to for the same reason: a Data Macro can only be created as
  an XML document loaded into the table, the shape of that document is fixed by the platform, and there
  is nothing in it for you to decide. Read it for the shape of the document only — how the work divides
  and what things are called stay exactly as declared under *Free to choose alternatives*.
- **You may read `audit-logging-lite-scaffold.md` and `audit-logging-lite-outcome-first.md` for one
  worked example of a comparable Data Macro decomposition — nothing more.** Neither is a route this
  file adopts; a different, equally valid decomposition that still satisfies every check here is a
  legitimate build.
- **Read every file in `standards/` and apply it.** Naming, audit columns, error handling, query
  style, and how the work divides into procedures all come from there and never from this file.
- **Ask for the six things under *Information and conditions you need to supply*,** one at a time,
  through the interactive selection control where the answer is a choice and as a question phrased in
  plain language where it is a name. Two of them are gates: a database in real use with no backup
  stops the build, and permission to change the tables is asked immediately before anything is
  changed. **Items 3 and 5 are also this template's two front-matter warnings** — asking those items
  satisfies the instruction elsewhere to surface every `warnings` entry; do not additionally surface
  them as a separate step before the numbered list, which would ask the same thing twice.
- **After the sixth thing and before you present the design, offer the `Explore options` step**
  (`_template-schema.md` §12.5) over the list under *Free to choose alternatives*, and nothing outside
  it. The two named Data Macro mechanisms are not on that list and are never offered an alternative. A
  pick made there is recorded in the build record, holds for the rest of the run, and is restated in
  the design you present.
- **Never infer an answer that belongs to the developer** — not from what the database looks like, not
  from reasoning that makes an answer seem obvious. Where a check exists to answer a question, run the
  check at the point the sequence calls for it rather than working the answer out yourself.
- **Before attaching either Data Macro to a live table, check what has that table open and name it if
  the attempt would be blocked** — `templates/_materialization.md`, "Before altering a table already
  in use, check what has it open." This is what check 18 under *How you validate the template's
  output* is confirming; nothing in this file names a procedure to do it, so the build route you chose
  supplies the check itself, by whatever means it has available.
- **Surface both house assumptions and every warning in the front matter** and get the developer's
  answer on each before building.
- **The build record reports against *How you validate the template's output*, one entry per check,
  each saying what was done and what was observed.** Passed and not passed are the only outcomes,
  including where the first method to run a check hits an obstacle — see `_template-schema.md` §12.2
  for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every entry opens with, and what
  to do before settling for a soft result.
- **While the build runs, do not narrate it.** Say once that it has started and what it will produce;
  say anything the developer must act on, as a question; say when it is finished, what was built, and
  where the build record is. Everything else — every procedure written, every check that passed — goes
  to the build record.
- **Once the build is reported finished, and only then, mention each entry under `related` in the
  front matter** (`_template-schema.md` §7.1) — one line per entry, what it is and why. This is not
  part of the build, never a gate, and never read before this point.

## Extra options

*Named optional extensions, none of them filled in for an engagement.*

- **Category-specific capitalization floors**, named above as something this template does not model.
- **A one-time, checked route for a historical import below today's threshold** — a deliberate,
  logged suspension of the capitalization check for a single, named import run, rather than the
  developer disabling the Data Macro by hand.
- **A form for reconciling a misplaced asset** onto the room a scan actually found it in.
