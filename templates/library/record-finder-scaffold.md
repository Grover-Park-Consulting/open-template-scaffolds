---
template: record-finder-scaffold
title: Record Finder for an Entry Form — VBA Scaffold
domain: library
type: vba-scaffold
version: 0.5.0
status: stable
status_fingerprint: fd6fbf247682
implements: catalog-schema
requires_tables:
  - tblPublication
  - tblPublicationGenre
platform_facts: [form-mapping-rules, form-layout-fidelity]
steps: [01-open-and-confirm-tables, 02-procedures]
standards_layer:
  - design-principles
  - error-handling
  - query-style
  - naming-conventions
  - form-conventions
target_module: modRecordFinder
new_procedures:
  - FinderRowSource
  - WithAllRow
  - JumpFormToRecord
  - RewriteWhere
warnings:
  - The finder itself is a new module, but wiring it up changes the code behind an entry form the
    developer already has. A build against a database in real use is preceded by a backup copy of
    the file, and the developer is asked for one before anything is changed.
---

# Record Finder for an Entry Form — VBA Scaffold

**Status last determined:** 2026-09-26.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Realize the **record finder** the publications entry form defers to code — the pattern
`publication-form` names as a deferred dependency. On a data-entry form over a large
table, this is how a user gets to the *one* record they want to edit: a filter strip narrows a
pick-list dropdown three ways — by **category**, by **first letter (A–Z)**, or by **keyword** — and
choosing an entry **jumps the form to that record**, with a consistent **"show all"** option in every
mode.

This scaffold supplies the **procedure skeletons** — signatures, the row-source builder, the
server-side jump, and the error-handling frame. It does **not** write the domain logic: each
procedure marks where that goes (which columns are searched, the sentinel text, the primary key).
House style (the central error logger, how SQL is written, control names) is deferred to the
standards layer.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here (signatures, the build/jump plumbing, error structure).
- **`[STANDARDS]`** — house style, deferred (`error-handling.md`, `query-style.md`, `naming-conventions.md`, `form-conventions.md`).
- **`[BUSINESS LOGIC]`** — the domain rule you fill in (columns, sentinel text, primary key).

## Prerequisites

| Object | Role |
|---|---|
| `catalog-schema` tables (`tblPublication`, `tblPublicationGenre`) | The finder queries these to build the pick list |
| `publication-form` | The paired form-spec; supplies the finder controls this engine drives (`cboCategory`, the A–Z sub-picker, `txtKeyword` + `cmdKeywordGo`, `cboFinder`) |
| The form's base record source | A saved query (or its SQL) whose `WHERE` `JumpFormToRecord` rewrites to land on one record |
| A sort-title column (e.g. `PubSort`) | The finder list and the jumped form order on it |
| A central error logger | `error-handling.md` |

### Ask before building

**Where this is a database already in use, ask for a backup copy before changing anything.** This
template alters the file it is built into. Two questions, in this order:

1. *"Is this a database you already use, or a new one you're trying this out on?"*
2. Where it is one they already use: *"Make a copy of the file before I start. Say when it's done and
   the build continues; say there's no copy and it stops."*

**Stop and build nothing if they say there is no copy.** A copy made before anything changes is the
only way back.

A new database, or one they are trying this out on, needs no copy — the question ends at step 1.

### Where this module goes in a split database

A **split database** is the normal shape for Access applications, especially those in multi-user
environments: one file holds the tables (the **back end**, usually on a shared network drive — never
on OneDrive, Dropbox, or any other file-syncing cloud folder, which corrupts a shared Access back
end), and each person runs their own copy of a second file holding the forms, reports, and code (the
**front end**), whose tables are *links* pointing at the back end. A single-file database — one
.accdb holding everything — is an acceptable choice for one user, and everything here works there
too; put everything in that one file and ignore the distinction.

| Module | Back end | Front end | Why |
|---|---|---|---|
| `modRecordFinder` (this scaffold) | No | **Yes** | Every one of these procedures serves the form: it reads the form's filter controls, builds a row source for a combo box on that form, and rewrites that form's record source. All of that is front-end work, so the code belongs beside the form. |

`tblPublication` and `tblPublicationGenre` are in the back end; this module reaches them through the
front end's links, which is why the SQL these procedures build matters more than it would locally —
see the two notes below.

**Both filter paths run across the link, so let the back end do the work.** Each rebuild of the
pick list sends a query to the back end and pulls the results back. Filtering in the SQL — which is
what `FinderRowSource` is for — means only the matching rows travel. Pulling the whole table and
filtering it in VBA would work identically on a single-file database and get slower and slower on a
shared one as the catalog grows. This is the practical reason the row source is built as a string
rather than filtered after the fact.

**`JumpFormToRecord` refetches one row on purpose.** Rewriting the record source's `WHERE` means the
back end returns a single record instead of the whole catalog. The alternative — loading everything
and setting the form's `Filter` — is fine on a small local table and increasingly wasteful across a
link. See Parked for the trade-off.

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure, the error-reporting call, and the
  line-number policy come from `error-handling.md`, which ranks two options and says when each
  fits. The `MsgBox` block shown is option 2, which needs nothing installed; a practice with its own
  logger substitutes it at the call site, and may number lines or not. The two pure helpers carry no
  handler by the same standard.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house query
  standard: the `LIKE` keyword search and its term escaping, the `SELECT DISTINCT (RecordID,
  DisplayText, SortIndex)` shape, the `<All ...>` union, and the WHERE-trim in `RewriteWhere`.
- **Naming** — procedure, variable, and parameter names follow `naming-conventions.md`.
- **Form conventions** — the finder control names (`cboCategory`, `cboFinder`, `txtKeyword`,
  `cmdKeywordGo`, the A–Z sub-picker) and the thin-delegate handler rule follow `form-conventions.md`.

## Extra Options

*Named optional extensions, none of them filled in for an engagement; the filled copy is saved to the developer's own
library, not committed here.*

- **"Recently edited" list** — a most-recently-used shortcut alongside the finder, jumping straight
  to a recent record via `JumpFormToRecord`.
- **Multi-column finder** — show author / year beside the title in `cboFinder` for disambiguation.
- **Debounced keyword search** — search as the user pauses typing instead of on a `Go` click.

## Parked / future considerations (not in this design)

- **Server-side vs. client-side jump** — `JumpFormToRecord` refetches a single row server-side (best
  for a large, linked back end). A small, local record set could instead use the form's `Filter` /
  `FilterOn`, keeping the full set loaded; the choice is the adopter's, by data volume.
- **One owner for the base query** — `JumpFormToRecord` assumes a stable base record source. **This
  is not a hypothetical:** in the source database this pattern came from, the same rewrite primitive
  was actually shared by a "recently edited" list and a related form — a live condition, not a future
  one, checked for before this scaffold's own build and flagged where `JumpFormToRecord` is defined
  above. What's still open is the fix a different engagement might prefer: a single owner, or a
  pristine base record source kept untouched, rather than the improve-in-place approach taken there.
