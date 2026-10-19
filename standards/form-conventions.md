# Form Conventions — OTS Default Standards Layer

**Who reads this:** the AI assistant, applying these rules to what it generates or a shop deciding
what to replace with their own.

If you are building something from a template, you don't need to read this file unless you are curious.
However, we have included clarifying comments to help you interpret what it says, just in case.

**Building from a template?** This file decides how the screens you work on are laid out:
what the boxes and buttons on them are called, which buttons appear, and the order you move through
them in. The names of the screens themselves are decided in
[`naming-conventions.md`](naming-conventions.md).

You do not have to do anything with this information to use the templates. The rules apply to the code you receive
whether or not you read these files.

Using the standards as they are is the normal choice. If you came here because a template asked whether
you want to use these rules as they are or make them your own, you don't answer the question here.

*After you've read about the choice here, go back and answer it where it was asked.*

[`README.md`](README.md) lists all seven files if you want to see the others first.

> **This is the OTS default standards layer.** When you fork the library, replace this file with
> your own house form conventions. You can also add to them here; **do so carefully to avoid
> breaking the template.** It covers form **design** — control prefixes,
> default control types, the standard button set, tab order, sizing — **plus the named reusable
> patterns** a `form-spec` defers to via `standards_layer: [form-conventions]`. Form / subform /
> record-source *naming* lives in `naming-conventions.md`; this file governs *design*.

From here on, this file contains instructions for the AI.

**To the AI generating code:**

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
| Foreign key / lookup | Combo box (row source = a named query over the lookup; bound column hidden) |
| Boolean | Check box |
| Date/Time | Textbox + date picker |
| Number / Currency | Textbox |
| Many-to-many relationship | Continuous Subform |

**A combo box or list box takes its row source from a named query** (`qry[Entity]_cbo` or
`qry[Entity]_lst`, per `naming-conventions.md`) — never a table name and never SQL typed into the
property. The query can then be opened, tested, and changed in one place.

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

**On a continuous subform's detail row, this rule does not mean one label per row.** A continuous
form repeats its detail section once per visible record, so a label placed in the detail section
alongside its control repeats with it — the same caption printed on every row, which is not what
"its own caption label" is asking for. Put the label once, in the **Form Header**, positioned above
its column, the same way a datasheet's own column headings work; the detail section carries the
data control alone, with no attached label of its own.

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
- **All or One selector** — a combo box that lets the person pick one value or everything. Its row
  source is a named query that adds one made-up entry, *All*, to the lookup's rows by `UNION`:

  ```sql
  SELECT TOP 1 0 AS MediaTypeID, "All" AS MediaTypeName FROM MSysObjects
  UNION
  SELECT MediaTypeID, MediaTypeName FROM tlkpMediaType;
  ```

  The *All* entry's key is 0, which no lookup row uses; `MSysObjects` is only a table known to have a
  row. The combo's bound column is hidden, its default value is 0, and picking 0 means no selection:
  whatever the combo drives (a record selection, or a filter on the form) is cleared; any other
  pick applies it. The query's sort order is not part of the pattern; it follows the context. In VBA,
  the 0 is a commented Enum member (`naming-conventions.md` §7).

  **When the pick is applied as the form's own filter, Access saves that filter with the form** and
  re-applies it at the next open, while the combo shows its default. `Form_Open` therefore turns the
  filter off (`Me.FilterOn = False`) and sets the combo to *All*, so the two agree.
- **Quick-add lookup** — a `➕` button beside a lookup combo opens the matching manage form in add mode
  (dialog), returns the new ID via a `TempVar`, then requeries + selects it.
- **Validation highlights** — invalid fields are flagged visually on save and cleared on move.

### Optional patterns (engagement-specific — non-mandatory)

- **Audit display** — a per-record audit summary on the form. **Not a universal convention:** it is the
  form-side surfacing of the `audit-columns` standard. Include it where a database carries audit columns
  *and* chooses to show them; omit otherwise. **Fully supported when used** — a `form-spec` that enables
  it names it as an optional feature.
- **Field help** — a small `?` button beside a field that explains the field's term when pressed. Use
  it where a form's fields carry terms a newcomer will not know. It is made of four parts:
  - **A help table** with one row per term: the term's name (unique), a display order, and its
    explanation, plus the audit columns. The lessons review form calls it `tlkpLessonTerm`.
  - **One routine that opens the help**, `ShowTerm("Term")`, which opens a small read-only form as a
    dialog, filtered to the row whose term name matches.
  - **The buttons**: `cmdHelp_[Term]`, caption `?`, out of the tab order (`TabStop` off), with the
    tip text `What is [Term]?`, and `=ShowTerm("[Term]")` as the click action.
  - **The text ships with the application.** The explanations are written in code and loaded into the
    table by an idempotent routine that creates the table when it is missing, adds a missing row, and
    refreshes a row whose text has changed, so that a new version's wording reaches everyone. The form
    calls it on open (`startup-conventions.md` §2).
- **Locking by record state** — fields lock or unlock as the person moves between records, because
  one attribute of the record decides who owns its text. Situational, so none of the library's
  templates require it. Example, from the lessons review form: a lesson that came with the library has
  its text re-supplied by the library on every update, so `Form_Current` locks every text field of a
  lesson whose origin is Library and leaves its approval editable; a lesson the person added is fully
  editable. The test is a single comparison against the record's own key (an Enum member, per
  `naming-conventions.md` §7), and the same line unlocks the fields again on the next record.
