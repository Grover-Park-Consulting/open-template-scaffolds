---
step: 03-code
title: "Write and import the code"
platform_facts: [mcp-module-import, mcp-line-numbers, domain-function-transaction, row-lock-errors]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Code

**The tables exist from step 2: do not rebuild them here.** Write and import the procedures the
build needs, divided as the design-principles standard says, within *Free to choose alternatives*.
The rules below bind that code.

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
- **Apply every standard delivered with this template.** Error handling, query style, naming, and how the
  work divides into procedures all come from there and never from this file.
- **A single scan's writes need to be atomic (see the corresponding line under *The same behavior
  every time* and check 14).** This template names no mechanism for it, for the same reason it names
  no mechanism for anything else — see *Intent*. The direct way to get it in Access/DAO is the
  transaction guard `error-handling.md` documents: begin on a `Workspace`, take the `Database` you
  read and write through from that same `Workspace` (`ws.Databases(0)`), and
  wrap one scan's writes and reads in it. **If you use a transaction, a read that needs to see this
  scan's own work is a recordset on the same `Database` object, never a domain function** (fact
  `domain-function-transaction`). **Nothing here asks you to hold a transaction open across more than one scan** — that is
  the Batch / session transaction Extra Option, not the base build, and see that option's note before
  taking it.
