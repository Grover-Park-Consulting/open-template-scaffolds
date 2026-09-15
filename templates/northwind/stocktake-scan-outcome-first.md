---
template: northwind-stocktake-scan-outcome-first
title: Northwind Scanned Stocktake — outcome-first method
domain: northwind
type: outcome-first
version: 0.4.2
status: draft
extends: Northwind (Access Developer Edition)
requires_tables:
  - Products
requires_fields:
  - Products.ProductID
new_fields:
  - Products.SKUBarCode
  - Products.QuantityInPackage
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
  - design-principles
implements: northwind-stocktake-schema
house_assumptions:
  - "The stored count for a product's count line (CountedQuantity) is kept current as scans are
    added, not recomputed from the scan log at the moment somebody reads it. A reconciled stocktake
    count is a durable audit fact — business decisions are made from it — so it must not shift if
    scan detail is later edited or archived. This mirrors the table template's own declared
    assumption; a build that instead computes the count on demand from the scan log has changed
    that promise."
  - "A count line's flag says only that it needs review, not which direction — shortfall or overage
    — tripped it. That is answered by the sign of the counted-versus-expected variance, read at
    review time; it is not stored a second time on the count line itself. This mirrors the table
    template's own declared assumption; a build that stores the direction on the count line has
    added something this template does not ask for."
warnings:
  - "Two counters can scan the same product, for the first time in a session, at the same moment.
    Both can find no count line for it and both try to create one — the schema's own unique index
    on (session, product) then refuses the second attempt with an engine error. The build must
    turn that refusal into 'use the line the other counter just created', not let it reach the
    counter as a failure. Checking for an existing line first narrows the window; it does not
    close it."
---

# Northwind Scanned Stocktake — outcome-first method

**Who reads this.** Everything from *Intent* down to *Standards Layer* is written for the developer
whose database this is. The section after that is addressed to the AI assistant building it, and
says so where it starts.

---

## Intent

**To produce, in an Access database, the results described in *What you end up with* below.** That
section is the specification: what the database does once this is built, in words the developer can
check for themselves, followed by the checks that confirm the result arrived and the behaviours that
must hold however the work was divided up. Those three sections together are the whole of what this
template promises.

**Nothing else is here.** No procedures, no module names, no code. This template states the result
and leaves the route to whoever builds it. Unlike the audit-logging templates in this library, no
single mechanism is named here, because the platform does not force one: ordinary VBA, run from a
form the counters use, does everything this template asks for. The other version of this template —
the rules-based method, `northwind-stocktake-scan-scaffold` — produces the same result from working
procedure skeletons you fill in. Either one can be built against your own database, and they can be
built one after the other, against separate copies, to compare.

**This template realizes logic the table template leaves open.** `northwind-stocktake-schema` builds
the tables a scanned stocktake needs and states, in its own Business Rules, what has to happen when a
scan comes in — but defers *how* to "the coding section." This is that coding section, stated as an
outcome rather than as code. It assumes those tables already exist.

---

## What you end up with

**Every physical scan is recorded, and nothing scanned is silently dropped.** A counter scans a
barcode; the database resolves it to a product where it can, and holds it for review where it
cannot. Either way, a row exists for that scan.

**A barcode that matches nothing known is held for review, not guessed at and not discarded.** It is
recorded, marked as unmatched, and does not add to any product's counted quantity. Nobody has to
notice it went missing, because it never does.

**Scanning the same item twice, moments apart, is recognized as one item, not two.** Warehouse
scanning produces accidental double-scans — a trigger held a moment too long, a barcode that beeped
twice. When a scan matches one already recorded for the same product in the same session, for the
same quantity, close enough together in time, it is treated as a repeat of the same physical count,
not a second item. **The repeat is still recorded** — it appears in the scan history exactly like any
other scan — but it does not add to the product's counted quantity a second time. How close together
"close enough" means is a number of seconds you supply; see *Information and conditions you need to
supply* below.

**Two items that happen to share a quantity, scanned further apart than that, both count.** The
closeness in time is what tells a repeat apart from a coincidence — two separate boxes of the same
product, scanned minutes apart during ordinary counting, are not duplicates of each other, and
neither is held back.

**A package barcode adds more than one unit; a product barcode adds one.** Where a product is
tracked by the package it ships in as well as by the single unit, scanning the package's own barcode
adds the whole package's quantity to the count in one scan, and scanning a single unit's barcode adds
one.

**The quantity you end up with for a product, in a session, is the sum of what was actually scanned
— minus the repeats.** Not every scan on file; every scan that was not identified as a repeat of an
earlier one.

**Where a product's count strays too far from what the system expected — short or over — it is
flagged for review.** A count can come in low or high, and each direction has its own allowable
tolerance: theft, damage, and misplacement produce small shortfalls that are expected and not
themselves a problem; a receiving error, a return posted wrong, or a miscount produces small
overages the same way. A product can have its own tolerance for either direction; where it doesn't,
the matching database-wide default applies, and a product can override one direction without
overriding the other. A count inside its tolerance passes unremarked, in either direction. One
beyond it is flagged, so somebody looks at it.

**The flag says a count line needs review. It does not say which direction sent it there.** Whether
it was a shortfall or an overage that tripped the flag is not recorded a second time — it is read
off the sign of the variance you already have, the moment somebody reviews it.

**A product the system expected to find none of, and a counter found some of anyway, is always
flagged.** A tolerance stated as a percentage — "flag a shortfall over 5%" — has nothing to be a
percentage *of* when the expected quantity is zero: 5% of zero is zero, so there is no band of
"small and expected" for a percentage tolerance to describe. Rather than pick a stand-in number to
measure against, the check is built so that this case needs no special handling at all — see *To
the AI assistant building this* for the reasoning — and what falls out of it is the plainly correct
answer: any count found where none was expected is notable on its own, not a rounding error near a
threshold, so it is flagged every time.

**A stocktake starts from a baseline, taken once, covering every product it is counting — not
only the products somebody gets round to scanning.** Opening a stocktake writes down what the system
believes is on hand for every product in it, at that one moment, and every variance is measured
against those figures for the rest of the session. Two things follow, and both matter on the floor.
**Stock going on being sold and received while the counting happens does not move the target** — the
figures were taken once and stay put. And **a product nobody scans is still counted**: it ends the
session at zero against whatever the system expected, which is a shortfall of everything, and it is
flagged like any other. A product the system thought you had a hundred of and the counters never
found is the most serious thing a stocktake can turn up, and it is exactly the product that produces
no scans at all.

**Two people can count different products, or the same product, in the same session at the same
time, without stepping on each other's work.** A stocktake with several counters working the floor
at once is the ordinary case this exists for, not an edge case it merely tolerates. Two counters
scanning the same product for the first time in a session, in the same instant, both succeed — their
scans land against the one count line for that product, never two.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

Perform each of these checks against a copy of your database with the tables already built.
**Validate by asking: does this happen or not?**

1. **A barcode that matches a product resolves to that product.**
   - Scan a barcode you know belongs to a product in your catalog
   - Confirm a count line exists for that product in the session, showing that scan's quantity
2. **A barcode that matches nothing is held for review, not lost.**
   - Scan a code that matches no product
   - Confirm it appears in the scan record, marked as unmatched
   - Confirm no product's counted quantity changed because of it
3. **Scanning the same item twice, moments apart, counts once.**
   - Scan a product's barcode
   - Scan the same barcode again immediately, for the same quantity
   - Confirm both scans appear in the scan record
   - Confirm the product's counted quantity reflects only the first
4. **Two separate scans, spaced apart, both count.**
   - Scan a product's barcode
   - Wait past the duplicate-detection window you supplied
   - Scan the same barcode again, for the same quantity
   - Confirm the product's counted quantity now reflects both
5. **A package barcode adds the package quantity.**
   - Scan a product's package barcode, where one exists
   - Confirm the counted quantity increased by that product's package quantity, not by one
6. **A shortfall beyond the allowed tolerance is flagged.**
   - Count a product to fewer than expected, by more than its allowable shortage rate
   - Confirm that product's count line is flagged for review
7. **A shortfall within the allowed tolerance is not flagged.**
   - Count a product to fewer than expected, by less than its allowable shortage rate
   - Confirm that product's count line is not flagged
8. **An overage beyond the allowed tolerance is flagged.**
   - Count a product to more than expected, by more than its allowable overage rate
   - Confirm that product's count line is flagged for review
9. **An overage within the allowed tolerance is not flagged.**
   - Count a product to more than expected, by less than its allowable overage rate
   - Confirm that product's count line is not flagged
10. **The flag does not say which direction, but the variance does.**
    - Flag one count line by shortfall (check 6) and a different one by overage (check 8)
    - Confirm both count lines show the same flag, with nothing on either one distinguishing them
    - Confirm the counted and expected quantities on each still tell you which is which
11. **A count found where none was expected is always flagged.**
    - Find a count line whose `ExpectedQuantity` is zero, or set one up on a product with no
      system on-hand quantity at the moment a session opens
    - Count any nonzero quantity for it
    - Confirm that count line is flagged for review, however small the count
12. **Two counters, one product, the same moment, do not collide.**
    - This one needs two people, or two sessions open at once, counting the same product for the
      first time in the same stocktake within a second or two of each other
    - Confirm both scans succeed
    - Confirm they land against one count line for that product, not two
    - This is the check named in the warning about concurrent counters — it is worth trying
      deliberately, not just trusting the design
    - **Where the tool building this cannot produce two literally simultaneous processes against
      the file** — a single-threaded automation tool, or a machine that opens Access exclusively so
      a second process cannot reach it while a call is in flight — that is not evidence this check
      cannot run. Open a second, independent connection to the same file (a second DAO or ADODB
      connection, not the one the build already holds), and use it to insert the competing row
      between the first connection's own check and its own insert. That forces the engine to raise
      the real collision error rather than asking you to infer one from documentation, and lets you
      confirm the recovery path returns the other connection's row, not a duplicate. See
      `_template-schema.md` §12.2 for the rule this follows
13. **A product nobody scans is still counted, and shows up as missing.**
    - Open a stocktake and count your way through some but not all of the products, exactly as a
      real one would go
    - Pick a product you know the system expected stock of and that nobody scanned
    - Confirm it has a count line in that session anyway, counted zero against what was expected
    - Confirm that line is flagged for review
    - **This check is the one that fails when the baseline was never taken.** A build that creates a
      count line only when a scan arrives passes checks 1 through 12 and fails this one, because the
      product it is asking about is invisible to it

### The same behavior every time, not the same structure

For an AI-assisted template, two builds need not produce the same code. They need to produce a
database that behaves the same way. The following must be true of every build:

- One count line exists per product per session, never more than one — regardless of how many
  counters scan that product or how close together they do it.
- A count line exists for **every product the session covers**, created when the session opens,
  carrying the expected quantity of that moment — not only for the products that were scanned. The
  baseline is taken once and never retaken during the session. A product added to the catalog after
  the session opened is the one case that gets its line later, on first scan.
- A scan is always saved. Matched, unmatched, or a recognized repeat — every one leaves a row behind.
- A recognized repeat is excluded from the counted quantity, never from the scan record. The two
  questions — "was this scanned?" and "does this count?" — have separate, visible answers.
- The counted quantity for a count line is always exactly the sum of that line's scans that were not
  recognized as repeats — kept current as scans arrive, not computed fresh only when somebody asks
  for it (declared in `house_assumptions`).
- Whether a scan is a repeat is judged only against other scans on the same count line — the same
  product, the same session — never across products or across sessions.
- A shortfall check uses the product's own shortage tolerance where one is set, and the matching
  database-wide default otherwise; an overage check does the same against the overage tolerance,
  independently — a product can override one direction without touching the other. Both are
  expressed the same way once found, so the comparison never has to know which source supplied the
  number.
- A product within tolerance, in either direction, is never flagged, and a build does not flag it
  "for visibility" or any reason beyond the rule stated above.
- **The tolerance comparison never divides by the expected quantity, in code or in a query.** A
  count line whose expected quantity is zero is not a special case that needs its own branch — it
  is an ordinary input to a comparison that was never dividing by anything. See *To the AI assistant
  building this* for the reasoning. What that produces is stated as its own behaviour, not derived
  from it: a nonzero count against a zero expected quantity is always flagged, in every build.
- The flag itself never records which direction tripped it. A build that adds a second flag value,
  a direction column, or anything else naming shortfall vs. overage on the count line has added
  something this template does not promise (declared in `house_assumptions`) — the sign of the
  variance is what a reviewer reads instead.
- Two counters racing to be first on the same product never produce two count lines. One creates it;
  the other finds and uses what the first created.
- Nothing here interrupts a counter mid-scan with a message they cannot act on. A repeat is recorded
  quietly; a shortfall flag is something reviewed later, not something that stops a scan from being
  entered.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences, say so while
the design is being worked out, before anything is built. Where you don't choose, the build will
choose based on the rules built into it. The template's promise holds either way.

- How the logic divides into procedures, what they are called, and how many there are. They will be
  functionally equivalent, not necessarily structurally the same.
- Whether the counted quantity is kept current by the same code that records a scan, or by a separate
  mechanism attached to the table itself that runs whenever a scan is added or removed. Both keep the
  quantity current the moment a scan changes; which one runs the update is not fixed here.
- How the race between two counters on the same product is resolved — checking first and catching the
  resulting error, or a different way of making the second counter's attempt land safely. What must
  hold is stated above: one line, however it gets there.
- Which products a stocktake covers, where you want less than all of them. The build takes the
  baseline over every product not marked discontinued unless you say otherwise; what must hold is
  that the baseline is taken over all of them at once, when the session opens.
- Names for anything not already named by the table template.
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

- **Deciding, for a scanned barcode, whether it names a package or a single unit.** That
  determination is not settled by this template or by the table template it realizes — see *Parked /
  future considerations*.
- **A screen for reviewing unmatched scans or flagged count lines.** All of these are recorded.
  Building a screen to work through them is yours to add.
- **Choosing the duplicate-detection window for you.** You supply the number of seconds; see
  *Information and conditions you need to supply*.

---

## Information and conditions you need to supply

You need to provide information about five things; no guesses are made on your behalf.

1. **Whether you are trying this out on a new database or building it into one you already use.**
   The two behave differently from the first step onward, and nothing about your data tells us which
   one this is.
2. **Which file holds the tables.** Where your database is split, that is the back end — the file
   holding data, not the front end holding forms and code. A counter's scanning screen belongs in
   each front end, since it is what each person on the floor runs; the tables it writes to are in the
   one shared back end. A single-file database works the same way; put everything in that one file.
3. **Whether you have a backup**, if this is a database you already use. If you tell the template
   there is no backup, the build stops rather than continuing.
4. **How many seconds apart counts as the same scan.** This is the duplicate-detection window from
   *What you end up with* — two scans of the same product, for the same quantity, closer together
   than this many seconds, are treated as one physical item scanned twice. **120 seconds is a
   reasonable starting point** for a person working through a single aisle at a normal pace; a faster
   or slower counting style may call for a different number. Say a number, or accept 120.
5. **Permission to change your tables**, asked immediately before anything is built.

**The accdb must be in a Trusted Location.** Code in an Access file outside one does not run at all,
and Access does not always say so plainly — the failure can show up as the database reporting it
cannot find a procedure that is, in fact, right there. Trust is a setting on each machine, not a
property of the file.

---

## Standards Layer

**Your standards layer decides how this is built. This template decides only what it has to do.**

- **`error-handling`** — how a generated procedure reports a failure. One constraint on the answer is
  not a preference: nothing in the scanning path interrupts a counter mid-scan. A message box raised
  while somebody is scanning stops them for something they did not ask for and, mid-aisle, cannot act
  on.
- **`query-style`** — how the generated code writes and holds its SQL.
- **`naming-conventions`** — what any new procedures, variables, and settings are called, within
  whatever your standards already require. This template names no procedures on purpose — see *Free
  to choose alternatives*.
- **`design-principles`** — how the work divides into procedures. *Free to choose alternatives*
  leaves that division open on purpose, and this layer is what it is open to.

A shop adopting this library replaces `standards/` with its own. Nothing in this file changes when
they do, which is the point of keeping the two apart.

---

## To the AI assistant building this

**Two sections are the specification, and only those two: *What you end up with* and *The same
behavior every time, not the same structure*, both under that heading above.** Build a system that
satisfies every promise there and passes every entry under *How you validate the template's output*.
Then run every one of those checks yourself, on a copy, and record what each one did — in
`build-record.md`, written before you report the build finished, never afterward and never only
when asked for it. The checks bind you twice: the build has to pass them, and you have to run them. How you build is yours to decide,
within *Free to choose alternatives*. Which checks you run is not — all of them, every build. Every
other section in this file — *Intent*, *What the template does not do*, *Parked / future
considerations* — is context for reading those two. None of it binds on its own, and nothing that
binds is stated only there.

- **This template names no mechanism, unlike the audit templates in this library, because the
  platform leaves more than one workable route here.** Ordinary VBA behind the scanning form,
  a saved query, or a table-attached automatic behavior for the rollup — more than one of these
  satisfies every promise above. Naming one anyway would be route specification, which is exactly
  what *Free to choose alternatives* leaves open. Do not import a mechanism from the paired
  rules-based template on the assumption that a promise this firm must mean one route: it doesn't,
  here.
- **You may read `northwind-stocktake-scan-scaffold.md`, the rules-based method that produces this
  same result, for one worked decomposition — nothing more.** It shows procedure names, a control
  flow, and where the domain logic slots in. None of that is binding here. Copying its shape wholesale
  is a legitimate build; so is a different one that still satisfies every check. That template runs
  the check list below against its own builds and says so where it names this one: the checks belong
  to the result, not to either route.
- **Read `northwind-stocktake-schema.md`, the table template this realizes, for the tables, the
  Business Rules, and the seed values this build reads and writes** — in particular Business Rule 2
  (scan resolution and the duplicate check), Business Rule 3 (the rollup), and Business Rules 7 and 8
  (the variance tolerances, in both directions). This file restates their outcome; that file is where
  the field names and table shapes live.
- **Write the variance check as a comparison of quantities, never as a fraction with
  `ExpectedQuantity` in a denominator — anywhere, in code or in a saved query.** The natural way to
  state a percentage tolerance is `(ExpectedQuantity − CountedQuantity) / ExpectedQuantity > rate`,
  and it divides by zero the moment a count line's `ExpectedQuantity` is zero. Multiplying both
  sides of that comparison by `ExpectedQuantity` — which does not change which side is larger for
  any count line the fraction form could be evaluated on at all — gives
  `ExpectedQuantity − CountedQuantity > rate × ExpectedQuantity`: the same comparison wherever
  `ExpectedQuantity > 0`, with no division anywhere. Business Rule 8 states both directions in this
  form; build to that form, not the fraction form, even though the fraction form is the more natural
  way to *describe* a percentage tolerance in conversation. **This is not a guard clause to add on
  top of the fraction form — it replaces it.** A build that computes the fraction and then checks
  `ExpectedQuantity <> 0` before using it has reintroduced the division the rewritten form exists to
  avoid; it will divide by zero the one time this matters, on the one input path the fraction form
  cannot survive.
- **What this produces at `ExpectedQuantity = 0` is a stated behaviour, not an inference you are
  asked to draw.** Confirm it rather than deriving it: `CountedQuantity` cannot be negative, so the
  shortfall comparison can never hold when `ExpectedQuantity = 0`, and the overage comparison reduces
  to `CountedQuantity > 0`. **Any nonzero count against a zero expected quantity is flagged, every
  time** — this is check 11 under *How you validate the template's output*, and the corresponding
  line under *The same behavior every time, not the same structure*.
- **A negative expected quantity is flagged every time, and that is also stated rather than
  derived.** A host's on-hand calculation can return a negative number where recorded sales have run
  ahead of recorded receipts, so a count line can open with a negative `ExpectedQuantity`. The
  overage comparison then holds for any count at all, including zero, and the line is flagged. Leave
  it that way: the system's own figure was impossible before anybody counted anything, which is worth
  a human looking at. **Do not add a branch for it and do not guard the comparison with a test on
  `ExpectedQuantity`** — that is the division coming back by another route.
- **Read every file in `standards/` and apply it.** Error handling, query style, naming, and how the
  work divides into procedures all come from there and never from this file.
- **Ask for the five things under *Information and conditions you need to supply*,** one at a time,
  through the interactive selection control where the answer is a choice and as a question phrased
  in plain language where it is a number or a name. Two of them are gates: a database in real use
  with no backup stops the build, and permission to change the tables is asked immediately before
  anything is changed.
- **The duplicate-detection window is a value question, not a choice among named options — ask it as
  a question phrased in plain language, with the suggested number stated, and accept whatever the
  developer answers.**
  120 is what the template suggests; it is not a default that applies on its own, and a developer who
  says nothing has not agreed to it. Restate the number they gave (or confirmed) in the design you
  present for approval, in seconds, next to what it does — "two scans of the same product, for the
  same quantity, fewer than N seconds apart, count once" — so approving the design means approving the
  number, not just the feature.
- **After the fifth thing and before you present the design, offer the `Explore options` step**
  (`_template-schema.md` §12.5) over the list under *Free to choose alternatives*, and nothing outside
  it. A pick made there is recorded in the build record, holds for the rest of the run, and is
  restated in the design you present.
- **Never infer an answer that belongs to the developer** — not from what the database looks like,
  not from reasoning that makes an answer seem obvious. Where a check exists to answer a question,
  run the check at the point the sequence calls for it rather than working the answer out yourself.
- **Surface both warnings in the front matter** and get the developer's answer on each before
  building. Neither is a preference to note in passing — the Memo-barcode warning changes what you
  tell the developer about scan-resolution performance on their catalog, and the race-condition
  warning is a requirement you build to, restated under *The same behavior every time, not the same
  structure* as "two counters racing... never produce two count lines."
- **The build record reports against *How you validate the template's output*, one entry per check,
  each saying what was done and what was observed.** Passed and not passed are the only outcomes,
  including where the first method to run a check hits an obstacle — see `_template-schema.md` §12.2
  for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every entry opens with, and what
  to do before settling for a soft result.
- **While the build runs, do not narrate it.** Say once that it has started and what it will produce;
  say anything the developer must act on, as a question; say when it is finished, what was built, and
  where the build record is. Everything else — every procedure written, every check that passed —
  goes to the build record.

## Extra Options

*Named optional extensions, none of them filled in for an engagement; the filled copy is saved to the
developer's own library, not committed here.*

- **Batch / session transaction** — wrap a whole counting session's scans in one transaction (the
  `error-handling.md` transaction guard).
- **Unmatched-scan review queue** — route unmatched scans to a review surface instead of leaving them
  parked in the scan record.

## Parked / future considerations (not in this design)

- **Package-vs-unit disambiguation** — *how* a scan is known to name a package rather than a single
  unit is undecided in the table template this realizes, and stays undecided here; lives in this
  template when it is resolved there.
- **A stocktake that records the scope it was opened for.** You can ask for the baseline to be
  taken over fewer than all the products — one aisle, one category, one supplier — and the build will
  do it. What is missing is the session remembering that afterwards: nothing on the session says
  which products were in scope, so later on a product that was never meant to be counted and a
  product that was in scope and never found look the same. Deciding what a session covers, and
  holding it, is a design this template does not carry.
