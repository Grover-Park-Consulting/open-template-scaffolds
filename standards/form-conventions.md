# Form Conventions — GPC Default Standards Layer

> **GPC default; fork-and-replace.** Covers form **design** — control prefixes, default control types,
> the standard button set, tab order, sizing — **plus the named reusable patterns** a `form-spec`
> defers to via `standards_layer: [form-conventions]`. Form / subform / record-source *naming* lives in
> `naming-conventions.md`; this file governs *design*. A forked practice swaps this file for its own
> form conventions.

---

## 1. Control-name prefixes

| Prefix | Control | Example |
|---|---|---|
| `txt` | Textbox | `txtPublicationTitle` |
| `cbo` | Combo box | `cboMediaTypeID` |
| `lst` | List box | `lstSelectedItems` |
| `chk` | Check box | `chkMultiVolumeSet` |
| `cmd` | Command button | `cmdSave` |
| `lbl` | Label | `lblFormTitle` |
| `img` | Image | `imgCoverImage` |
| `sfrm` | Subform control | `sfrmPublication_Creator` |
| `grp` | Option group | `grpStatus` |
| `tab` | Tab control | `tabDetails` |

---

## 2. Default control type per data type

| Field / role | Default control |
|---|---|
| Text | Textbox |
| Memo | Multi-line Textbox (grows) |
| Foreign key / lookup | Combo box (row source = the lookup; bound column hidden) |
| Boolean | Check box |
| Date/Time | Textbox + date picker |
| Number / Currency | Textbox |
| Many-to-many relationship | Continuous Subform |

---

## 3. Standard buttons & footer

- The standard action set: **New, Save, Delete, Clear**, plus **record navigation**, in a consistent
  action region (footer or right rail).
- Buttons are captioned with a **mnemonic + a glyph** — e.g. `&Save 💾`, `&New ➕`, `&Delete ❌`. The
  mnemonic is the convention; **the specific glyph is the adopter's choice** (or omit glyphs entirely).
- A quick-add (`➕`) button sits beside a lookup combo where inline add is wanted (see §6).

---

## 4. Tab order

Explicit, in reading order: the record selector first, then detail fields in entry order, then the
action buttons. **Never** left to control-creation order.

---

## 5. Sizing & layout defaults

Label-left / control-right (label-above for Memo), consistent control height, aligned columns. This is
the **default ("ugly but functional") layout** the materialization step emits from a `form-spec`; the
adopter restyles. Positioning and sizing are defaults, not prescriptions.

By default every user-facing data control is generated with its own matching caption label (`lbl…`). A
few controls carry their own caption (a check box) and some are hidden and need none — the developer may
delete a label they don't want, but the build must never silently omit one: better to remove a spare
than to notice one is missing and add it.

---

## 6. Named reusable patterns

A `form-spec` **names** these patterns; the host supplies the implementation (the forms framework).
Documented here so a `form-spec`'s deferral is concrete, not hand-waved. Two tiers:

### Core patterns (generally used)

- **Layered record selector** — *N* filter inputs (e.g. genre, alpha, keyword) rebuild a selector
  combo's row source (dynamic SQL and/or `TempVar`-carried criteria, per `query-style.md`); the
  selector's `AfterUpdate` rewrites the form's record-source `WHERE` (via a SQL-`WHERE` rewriter such as
  `RefreshSQLWhere`). Supports **"All or One"** — an `<All>` entry leaves the form unfiltered.
  **→ tracked as a future `vba-scaffold` (the selector engine is reusable code).**
- **Quick-add lookup** — a `➕` button beside a lookup combo opens the matching manage form in add mode
  (dialog), returns the new ID via a `TempVar`, then requeries + selects it.
- **Validation highlights** — invalid fields are flagged visually on save and cleared on move.

### Optional patterns (engagement-specific — non-mandatory)

- **Audit display** — a per-record audit summary on the form. **Not a universal convention:** it is the
  form-side surfacing of the `audit-columns` standard. Include it where a database carries audit columns
  *and* chooses to show them; omit otherwise. **Fully supported when used** — a `form-spec` that enables
  it names it as an optional feature.
