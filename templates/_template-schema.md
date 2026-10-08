---
template: _template-schema
title: Open Template Scaffolds — Canonical Template Format
domain: _meta
type: spec
version: 0.19.0
status: draft
---

# Open Template Scaffolds — Canonical Template Format

**Who reads this:** anyone writing a template, and the AI assistant reading one.

## Contents

- 1. File rules
- 2. Front-matter (YAML)
- 3. Body sections — common core (all template types)
- 4. Body sections — `type: table-schema`
- 5. Field-spec table format (`type: table-schema`)
- 6. Standards Layer boundary
- 7. Extra Options
- 8. `type: vba-scaffold`
- 9. `type: form-spec`
- 10. The OTS Wizard (any template type)
- 11. Minimal skeleton (`type: table-schema`)
- 12. `type: outcome-first`
- 13. Masters and steps (any template type)

**Using a template to build tables, forms, or code in a database?** You do not need to read this
file in order to *use* the template. You only need to read this file if you want to *create* a template.

This is the **format specification** every template file in this library must follow.
It is the contract the template library MCP server keys off: discovery (`list_templates`,
`search_templates`) reads the front-matter; `get_template` and `get_part` deliver a template
in parts, the master first, with the active standards layer and the platform facts each part
needs (§13); `validate` checks a template (or a filled-in copy) against the
rules in this document.

The format specification is meta, not a template itself (`type: spec`, `domain: _meta`) — `validate` skips files
whose `type` is `spec`.

A template is a vetted, standards-baked *starting point*, not a drop-in guarantee. Every
template a developer adopts must be confirmed fit for the developer's intended application
before use.

The library supplies structure and proven decisions for the template itself.

Whether any given template is suitable for a specific engagement is always the adopter's judgment.

This file contains the specifications for the OTS Template format. You may also examine one or more
existing template files to see how they are applied as an example. However, **be careful not to unintentionally change an existing template.**

> **Scope note (build order).** The common core below is proven against real templates:
> `templates/stocktakescan/stocktake-schema.md` (`type: table-schema`),
> `templates/stocktakescan/stocktake-scan-scaffold.md` (`type: vba-scaffold`),
> `templates/library/publication-form.md` (`type: form-spec`), and
> `templates/audit/audit-logging-lite-outcome-first.md` (`type: outcome-first`). **All four type
> sections (§4, §8, §9, §12) are authoritative — the template format is complete**, each proven by
> hand before the template library MCP server's schema-dependent tooling is built against it.

---

## 1. File rules

- **Format:** Markdown (`.md`), UTF-8. No binary, no executable code.
- **Location:** `templates/<domain>/<name>.md`. Files prefixed `_` (e.g. this one) are
  library infrastructure, not domain templates.
- **One template per file.** A template defines one cohesive artifact set for one domain.
- **Single source of truth:** the file *is* the template. CLAUDE.md and the template library
  MCP server are readers of it, never separate copies.
- **Write for a reader who has never seen this system.** Every term is defined where it first
  appears. Templates do not use undefined jargon, coined labels the reader has to look up, or
  concepts assumed from another file. Prefer the plain name for a thing over the internal one.
  Where a rule has a visible consequence, state the consequence as well — what the reader will
  actually see happen — not only the rule that governs it. If a sentence would send a first-time
  reader to a search engine or to another file to understand it, rewrite it. This applies to the
  whole template, not to its introductory sections alone: field tables and Business Rules stay
  precise, but precision and plain language are not in tension. Where a plain paraphrase would
  collide with a term the reader probably already owns, or with a word that already means something
  else here, name both once — plain sentence first, technical term second (§10.5 rule 3).
- **Every action the AI assistant must take is stated where the AI assistant is addressed.** A
  template with more than one reader names each one and marks the switch where it happens, and
  "you" has one referent per section. So an imperative in the developer's section binds the
  developer only. "Perform each of these checks", said to the developer, tells the assistant
  nothing about what it must do, however plainly it is put. Whatever the assistant has to do —
  run every check, run a procedure at the point a sequence calls for it, stop at a gate — is
  stated in the section addressed to the assistant. It may be repeated to the developer, and the
  assistant's section is where it must appear. The test, for each thing the assistant must do:
  which section says so, and whom does that section address? A requirement in the wrong section
  is not a requirement. This holds for every template type; §8.4 and §12.4 are where two of them
  meet it.

---

## 2. Front-matter (YAML)

The front-matter is the machine-readable contract. Keys below marked **required** must be
present on every template; conditional keys are required when their condition holds.

| Key | Req | Type | Notes |
|---|---|---|---|
| `template` | required | string (kebab-case) | Unique slug; matches the filename stem |
| `title` | required | string | Human-readable title |
| `domain` | required | string | Domain folder name (e.g. `stocktakescan`, `sales`, `hr`); `_meta` reserved for infra |
| `type` | required | enum | `table-schema` \| `vba-scaffold` \| `outcome-first` \| `form-spec` \| `spec` |
| `version` | required | semver string | Template version, counted per template and independent of any library version. **Bump it in the same commit as any change to what the template produces** — patch for a correction an adopter needn't act on, minor for anything they would (a new or renamed field, a changed default, an added rule or section), major for a redesign an existing build can't absorb. It is the only thing that distinguishes a copy someone took earlier from the current file; see `CONTRIBUTING.md` → *Versioning a change* for why this is a rule and not a nicety |
| `status` | required | enum | `draft` \| `review` \| `stable` — what moves a template between them is defined in `_maturity-criteria.md`, not by inspection |
| `extends` | conditional | string | Required when the template grafts onto an existing database; names the host (e.g. `Northwind (Access Developer Edition)`) |
| `requires_tables` | conditional | list[string] | Existing tables the template hooks into. Required when `extends` is present |
| `requires_fields` | optional | list[string] | Specific existing fields relied on, as `Table.Field`. **Fields the host must already have** — a template that creates a field on an existing table declares it under `new_fields` instead |
| `new_fields` | optional | list[string] | Fields the template **adds to a table the host already has**, as `Table.Field`. Keeping these out of `requires_fields` matters to both readers: `check_compatibility` stops reporting them as missing pieces the developer has to go and add, and the AI assistant reads them as work the build does rather than a precondition it has to verify |
| `standards_layer` | required | list[enum] | Which standards-layer concerns this template defers; values from §6 |
| `new_tables` | conditional | list[string] | Tables the template defines. Required for `type: table-schema`; must match the `## Entities` headings exactly |
| `implements` | conditional | string | For `type: vba-scaffold` and `form-spec`: the `table-schema` template (by slug) the scaffold realizes / the form edits |
| `target_module` | conditional | string | Required for `type: vba-scaffold`; the module or class the procedures live in |
| `new_procedures` | conditional | list[string] | Required for `type: vba-scaffold`; must match the `## Procedures` `### <name>` headings exactly |
| `record_source` | conditional | string | Required for `type: form-spec`; the form's record source (a query over the edited table) |
| `new_forms` | conditional | list[string] | Required for `type: form-spec`; the forms/subforms defined — each subform appears as a `Subform` control in `## Layout` |
| `seeds` | optional | list[string] | Seed data the template expects, as `Table.RowKey` |
| `house_assumptions` | optional | list[string] | House-particular modeling assumptions deliberately kept in the template body (the "Declared" tier) because they can't be moved to the standards layer or dropped. Each entry is `Target — rationale`, where `Target` names the entity, field, or rule carrying the assumption. Makes embedded house bias machine-visible to adopters and discovery tools. |
| `warnings` | optional | list[string] | **Anything the AI builder must surface *before* building, and act on.** Each entry states the condition and what the developer must confirm or the build must branch on: a platform limit the build cannot get around (e.g. "Data Macros cannot audit Long Text fields — confirm whether any audited table has one"), or what the template does to a database the developer already has (e.g. the backup gate). **These are examples, not a test of what belongs.** Never skip or remove an entry because it does not resemble them. House bias goes in `house_assumptions`, the only exclusion. |
| `platform_facts` | required | list[string] | Ids of the `templates/_materialization.md` sections a build from this template needs (each section carries `<!-- fact: id -->` under its heading). The template library MCP server delivers those sections as parts of the run, so the facts arrive when they are used rather than waiting to be looked up. Name every section whose mechanism the build uses. In a template divided into steps (§13), the master names only the facts the design depends on; a fact marked `route: build` is named by the step that uses it. |
| `steps` | optional | list[string] | The template's step files, in the order they are carried out (§13). Each entry is a file name without `.md`, in the `<template-id>.steps/` folder beside the master. |
| `related` | optional | list[string] | Other templates or standards files worth considering next, once this one is built — never during it (see §7.1). Each entry is `Target — rationale`, where `Target` is a template slug or a `standards/<file>.md` path |

**Rules the `validate` tool enforces on front-matter:**

1. All required keys present and non-empty; `type` and `status` within their enums.
2. `template` is a unique, kebab-case slug that **ends with** the filename stem; a
   domain or pairing prefix may precede it (e.g. file `library/catalog-schema.md` →
   slug `catalog-schema`; a form paired with the catalog → `publication-form`).
3. If `extends` is set, `requires_tables` is non-empty.
4. Every entry in `new_tables` is documented under `## Entities` — either as its own `### <name>`
   heading or as a named row in a grouped lookup sub-table (§4) — and vice versa; the declared and
   documented table sets are identical.
5. `requires_fields` / `new_fields` / `seeds` entries are well-formed `Table.Field` /
   `Table.Field` / `Table.RowKey`.
6. Every `house_assumptions` entry is well-formed (`Target — rationale`), and each `Target`
   resolves to an entity, field, or rule named in this template. (Format check only — `validate`
   cannot judge whether something *should* have been declared; that stays the human review gate.)
7. Every `related` entry is well-formed (`Target — rationale`), and each `Target` resolves —
   either to another template's `template` slug, or to a file that exists under `standards/`.
   (Format and existence only — `validate` cannot judge whether the relationship makes sense;
   that stays the human review gate, same as `house_assumptions`.)

**Scope of `validate` — format, not fitness.** `validate` confirms a template is internally
well-formed: complete front-matter, `new_tables` matching the `### <name>` entity headings, every
`FK → <Table>` resolving to a table named in the template. It does **not** open any host
database. That's a separate `check_compatibility(template, db_path)` tool. And neither check
speaks to fitness: a passing `validate` means well-formed, never *suitable for purpose*. Confirming a
template fits the intended application stays the adopter's responsibility.

### `extends` and the host's own naming convention

**This was an undecided question, not a documented rule, and three templates each answered it
differently on their own** (`audit-logging-lite-scaffold` assumed `tbl…`/`tlkp…` naming that
matches nothing in a non-OTS host; `catalog-schema` declared no host and collided with the
collection it was shaped from; `record-finder-scaffold` named a field its own paired schema renames
away). The decision:

**Grafting onto an existing database (`extends` is set): the host's own naming convention wins over
`standards/naming-conventions.md`, once found — but it has to be found first, and confirmed, never
assumed.** Inspect the host's existing tables before naming a single new object. Where the host's
own convention differs from this file's, ask the developer which one governs the new objects —
state what was found, name the preferred choice as **keep the host's own convention**, and let them
choose. Never silently apply the OTS standard onto a host that has already established its own,
and never silently adopt whatever the host does without asking.

**No host to inspect (a greenfield database, `extends` absent): default to
`standards/naming-conventions.md`, but still ask.** There is no existing convention to defer to, so
the OTS standard is the reasonable preferred choice — but it is still a choice the developer gets
to confirm, not one settled by there being nothing to compare against.

Either way, this is a single wizard-style step, asked once per build, not re-litigated per table or
per template in a multi-template engagement.

---

## 3. Body sections — common core (all template types)

Section headings are canonical: use the exact `##` text below so `validate` and
`get_template` can locate them. Required-for-all sections:

| Section | Purpose |
|---|---|
| `# <title>` | H1 matching `title` |
| `## Intent` | What the template produces and why; the domain framing a reader needs before the detail |
| `## Standards Layer` | What is **omitted** here and supplied by the developer's standards layer (see §6) |
| `## Extra Options` | Engagement-specific stub (see §7) |

Optional-for-all:

| Section | Purpose |
|---|---|
| `## Parked / future considerations` | Named directions explicitly **not** in the current design |

**The backup gate.** A template that alters a database the developer already has asks for a backup
copy before changing anything, and stops if there is none. It goes where that template states its
preconditions — `## Prerequisites`, or *Information and conditions you need to supply* in an
`outcome-first`. **The wording does not vary between templates:** the reason is always that the
template alters the file it is built into, so a reason written per template is a defect, not detail.

---

## 4. Body sections — `type: table-schema`

In addition to the common core, a `table-schema` template **must** contain, in this order:

| Section | Purpose | Required |
|---|---|---|
| `## Prerequisites` | What must be settled before the tables are built: **which file they go in** (see below), and — when the template grafts onto an existing database — the hooks into it, as a table of existing objects the new tables wire into. Required when `extends` is set; otherwise optional, but expected wherever the deployment note applies | conditional |
| `## Entities` | One `### <TableName>` per new table — grain statement, field table (§5), `Indexes:` line. **Trivial, uniform lookups** (`<name>ID` + a descriptor + optional `SortOrder`, nothing more) may instead be **grouped** in a single sub-table of name + seed rows — documentation shorthand only; each row is still its own discrete table (this is *not* a shared/MUCK lookup table). Any lookup carrying extra structure (more fields, an FK, a description) takes its own `### <name>` heading like an entity. | required |
| `## Relationships` | New relationships and hooks into the host schema, as a bulleted list naming parent → child, the join field(s), and cascade behavior | required |
| `## Business Rules` | Numbered list of the logic the generated objects must honor (grain constraints, rollups, derivations, deferred-logic notes) | required |
| `## Validating the build` | The structural checklist (§4.2) confirming the tables, keys, indexes, relationships, and seed rows exist and behave as this template specifies | required |

**The deployment note.** Every `table-schema` states, in `## Prerequisites`, which file its tables
are built into. This library's templates are oriented toward a **split database** — the normal shape
for Access applications, especially those in multi-user environments: one file holds the tables (the
**back end**), and each person runs their own copy of a second file holding the forms, reports, and
code (the **front end**), whose tables are links to the back end. Tables go in the back end. A
single-file database — one .accdb holding everything — is an acceptable choice for one user and
works the same way, so the note always says so rather than implying that splitting is required.

Say it even when it seems obvious: a schema designed as though there were only ever one file can be
built as one, but going the other way later costs a rebuild. Templates that carry a behavior
sensitive to the split — a data macro, a VBA function a macro calls, an external file folder — say so
where that behavior is defined, not only here.

**A schema that names no mechanism for a Business Rule is choosing to leave enforcement to the
paired `outcome-first` stage, not omitting it.** Two schemas in this library make different choices
here, and both are right for their own case: one names a Data Macro directly, in its own text, for a
comparably-shaped rule; another states an equivalent rule as a field-table constraint and says
nothing about how it is enforced, leaving that to the outcome-first template built against it. Read a
given schema's own `## Standards Layer` and `## Business Rules` to see which way it chose — don't
infer one schema's choice from another's, and don't read a schema's silence on mechanism as a gap to
fill. The same in-flight choice X11 names for a paired scaffold's field naming: the library states
the decision once, in the template that owns it, and a developer or a later template reads it there
rather than assuming it travels.

### 4.1 `validate` rules for `table-schema`

1. Every table documented under `## Entities` — whether as a `### <name>` heading or as a named row
   in a grouped lookup sub-table — appears in front-matter `new_tables`, and vice versa.
2. Every field table conforms to §5 (columns and type vocabulary).
3. Every `FK → <Table>` named in a field table resolves to either another entity in this
   template, a `requires_tables` entry, or another `new_tables` entry.
4. Every table named in `## Relationships` is an entity, a lookup, or a `requires_tables` entry.
5. Audit columns (the house set — e.g. `CreatedDate`, `CreatedBy`, `ModifiedDate`, `ModifiedBy` under
   the OTS default) do **not** appear in field tables — they belong to the standards layer (§6) and
   are flagged if present.
6. `## Validating the build` is present and non-empty. `validate` confirms the section exists — it
   cannot judge whether the checks it lists are the right ones; that stays the human review gate.
7. Every bullet in `## Relationships` states its cascade behavior explicitly — the text contains
   "cascade" or "no cascade"/"restrict". A relationship whose behavior isn't stated either way is
   flagged, not silently assumed either direction.
8. Every `(Business Rule N)` citation elsewhere in the template resolves to an actual numbered item
   in `## Business Rules`. Catches a stale reference after a rule is renumbered, reworded away, or
   removed — the citation is text, so nothing else would catch this drift.
9. Every front-matter `seeds` entry is described with its actual value(s) somewhere in
   `## Entities`, not merely named in the list. A seed that is declared but never specified gives a
   builder nothing to insert.

### 4.2 The `table-schema` checklist — structural, not business logic

**Why `table-schema` needs its own checklist rather than deferring to its paired template's.** A
`vba-scaffold` or `form-spec` defers its own checklist to the paired `outcome-first` template's —
they promise the same result by a different route, so one checklist serves both. A `table-schema`
has no such sibling to defer to: it is usually built *before* its pair exists, and what it promises
is different in kind — not a business outcome, but that the tables, keys, indexes, relationships,
and seed rows exist and behave the way this template specifies. **Confirmed necessary, not merely
tidy:** a `table-schema` build once shipped with a leftover `DefaultValue` blocking every insert into
a new table — reported passing, because nothing in the template asked anyone to try one. The
equivalent structural defect, in a template with a checklist, was caught by an ordinary insert.

**Every `table-schema` template's `## Validating the build` states its own checks, instantiating
this baseline against its own tables, fields, relationships, and seed rows** — not a generic list
copied unchanged, and not reinvented from scratch per template:

1. **Every new table accepts an ordinary insert** — the minimum a table must do, and the specific
   gap that went unnoticed without this rule. Run one plain insert per table, through whatever the
   template's own build route produces (a DAO `Sub`, or the AI assistant inserting through the tool
   it built the table with), not a hand-crafted edge case.
2. **Every declared `Required` field refuses a missing value, and every declared `AllowZeroLength`
   field accepts an empty string** where the template says it should (see
   `_materialization.md`'s sink-field rule and its log/audit-domain exception).
3. **Every unique index refuses the duplicate it names.** A junction's compound unique index, a
   lookup's name uniqueness — whatever `## Entities` declares as unique, tested with an attempt that
   should fail.
4. **Every relationship in `## Relationships` behaves as declared on the delete side** — a cascade
   deletes its children, a non-cascade refuses the parent's own delete while children exist.
5. **Every seed row named in front-matter `seeds` is present, exactly as specified**, after the
   build — not merely that the table exists.
6. **Audit columns stamp correctly**, where the template attaches `standards/audit-columns.md`'s
   mechanism: `CreatedDate`/`CreatedBy` on insert, `ModifiedDate`/`ModifiedBy` on update, `Created*`
   left frozen on a later update.
7. **Every relationship also behaves as declared on the insert side** — a child row whose FK value
   doesn't exist in the parent is refused. Check 4 tests deleting the parent; this tests the other
   direction, and neither stands in for the other.
8. **A value too long for a `Text(n)` field is refused, not silently truncated** — except where a
   field's own documentation names truncation as the intended behavior (`ErrorDescription` in
   `error-logging-schema.md` is the one declared exception in the library today). Silent truncation
   passing as success is the same failure shape X17/X18 name for Data Macros, one layer down, at the
   field type itself. **Test it on every route the build itself writes by, and only those:** a bound
   form or a recordset refuses an over-long value, while a SQL `INSERT` stores the first n characters
   without raising an error, so where the build inserts by SQL it checks each value against the
   field's own definition before writing. **Where the build copies existing rows into a table it
   created, compare each source field's definition with its target's before copying;** a source
   wider than its target is settled with the developer, never copied and cut short. A SQL `INSERT`
   the build does not make, such as a query or import the developer writes later, is not tested:
   what the engine does with it is a platform fact (`_materialization.md`, "A SQL `INSERT` shortens
   an over-long text value without an error"), told to the developer in the design.
9. **`Description` actually landed on the field, not only in the template's prose.** The classic
   order-of-operations defect `_materialization.md` documents — `Description` set *after*
   `TableDefs.Append`, never before, or error 3219 — is cheap to check and has already bitten this
   library once for exactly this reason.
10. **The build `Sub` survives being run a second time** — either it re-runs cleanly, or it fails
    naming what's already there, never a bare "duplicate object" surprise. The structural sibling of
    the "running the build again duplicates nothing" check several `vba-scaffold` checklists already
    run.

Report against this list in the build record exactly as `_template-schema.md` §12.2 states for
every checklist in the library: one entry per check, a literal `Result: PASSED` or
`Result: NOT PASSED`, passed and not passed the only outcomes.

---

## 5. Field-spec table format (`type: table-schema`)

Each entity's fields are a Markdown table with exactly these columns:

```
| Field | Type | Key / Req | Purpose & rules |
```

- **Field** — backtick-wrapped field name. OTS field-qualification rules apply (no bare
  reserved/ambiguous nouns: `Status` → `<Entity>StatusID`, `Notes` → `<Entity>Notes`).
- **Type** — from the Access type vocabulary: `AutoNumber`, `Long`, `Integer`, `Byte`,
  `Single`, `Double`, `Currency`, `Text(n)`, `Memo`, `Date/Time`, `Boolean`. `GUID` is
  permitted only when documented as a deliberate choice (the stocktake template strips
  Dataverse GUID keys in favor of `AutoNumber` — treat GUID as a smell to justify).
- **Key / Req** — one or more of: `PK`, `FK → <Table>`, `PK + FK → <Table>` (shared key),
  `Required`, `Nullable`.
- **Purpose & rules** — one line: what the field is for and any field-level rule.

Each entity also carries an `Indexes:` line naming the PK, any unique index (with its
columns), and FK indexes. Derived values that are computed rather than stored are noted
explicitly as "Derived (not stored): …".

---

## 6. Standards Layer boundary

The `## Standards Layer` section lists what the template **deliberately omits** so that the
same template produces house-conforming output for any practice. Front-matter
`standards_layer` enumerates which of these apply. Recognized values:

| Value | What it covers |
|---|---|
| `audit-columns` | The house who-and-when columns (`CreatedDate` / `CreatedBy` / `ModifiedDate` / `ModifiedBy` under the OTS default) on new tables. A host database may already use different names for the same thing — check it before naming new ones, per `standards/audit-columns.md`'s own Notes section. Never named in the template body |
| `naming-conventions` | Table/field prefix policy (e.g. Northwind no-prefix vs the OTS `tbl`/`tlkp`). The template states which house style it follows; a different practice builds the same entities under its own conventions without editing the template |
| `error-handling` | The house `errHandler` / global-error pattern for any VBA generated alongside |
| `query-style` | How VBA and saved queries write and run SQL — where SQL lives, aliasing/qualification, formatting, and safe criteria. Applies to any generated code that touches data (notably `vba-scaffold`) |
| `form-conventions` | Form **design** defaults (control prefixes, control types, buttons, tab order, sizing) + the named reusable form patterns (selector, quick-add, validation highlights; audit display optional). Used by `form-spec` |
| `design-principles` | The reasoning behind the specific rules — one-job-per-procedure, separation of concerns, encapsulation, cohesion/coupling, DRY, strong contracts — that any generated VBA is shaped by |
| `startup-conventions` | How a generated Access application initializes on open — the `AutoExec` → `Startup()` convention, the idempotent `EnsureAppFolders()` slot, and reliable external-file-asset folders. Used by a `form-spec` that materializes a full application |

A template **describes the boundary**; it does not embed the standards. Where a house-specific
*modeling* assumption cannot be cleanly separated, resolve it by the lowest tier that fits: drop
it from the published template (Private), park it in `## Extra Options` (Optional), or — if it is
load-bearing — keep it and declare it in the `house_assumptions:` front-matter list (Declared), so
it is machine-visible rather than buried in prose.

---

## 7. Extra Options

Every template ends with a `## Extra Options` section: a **stub in the base library**,
listing named, optional extensions a developer fills per client engagement. The filled-in
copy is saved to the developer's own library — never committed back here. Extra Options are
how a template absorbs natural depth without bloating the core (e.g. the stocktake template
parks cloud/mobile migration and category-level shrinkage here).

### 7.1 The `related` front-matter key

Moved to `templates/_method.md`, method `related-after-finish`.

## 8. `type: vba-scaffold`

A `vba-scaffold` template defines a set of **procedure skeletons** that realize logic a paired
`table-schema` template defers to code. It provides structure — signatures, recordset plumbing,
control flow, and the error-handling frame — with the **domain logic marked but not written** and
the **house style deferred** to the standards layer. The defining idea is a three-way split:

- **`[SCAFFOLD]`** — structure the template provides.
- **`[STANDARDS]`** — house style, deferred (error-handling, query-style, naming).
- **`[BUSINESS LOGIC]`** — the domain rule, filled per engagement, sourced from the paired
  table template's numbered Business Rules.

### 8.1 Front-matter (in addition to the common keys in §2)

| Key | Req | Notes |
|---|---|---|
| `implements` | optional | The `table-schema` template (by `template` slug) whose Business Rules this scaffold realizes. Present when the scaffold is paired with a schema. |
| `target_module` | required | The module or class the procedures live in (e.g. `modStockTakeScan`). |
| `new_procedures` | required | The procedures the template defines; must match the `### <Procedure>` headings under `## Procedures` exactly. |

`requires_tables` (the tables the code runs against) and `standards_layer` (which **must** include
`error-handling`, and typically `query-style` and `naming-conventions`) carry their §2 meanings.

### 8.2 Body sections

In addition to the common core (§3), a `vba-scaffold` template **must** contain, in this order:

| Section | Purpose | Required |
|---|---|---|
| `## Prerequisites` | The objects the code runs against — the paired schema's tables, host fields, and the central error logger. Required when `requires_tables` or `implements` is set | conditional |
| `## Procedures` | One `### <ProcedureName>` per procedure (matching `new_procedures`), each with its scope + signature and an annotated `vba` code block (§8.3) | required |

A `vba-scaffold` has **no** `## Relationships` or `## Business Rules` — those live in the paired
`table-schema` template; the scaffold *cites* its Business Rule numbers in `[BUSINESS LOGIC]`
markers rather than restating them.

### 8.3 Procedure entry format

Each `### <ProcedureName>` heading is followed by the procedure's scope and signature and a single
fenced `vba` block. The block is a **complete, compilable skeleton** carrying three kinds of comment
annotation:

- `' [SCAFFOLD] ...` — structure provided by the template.
- `' [STANDARDS — <file>] ...` — a point deferred to a standards-layer file.
- `' [BUSINESS LOGIC #n] ...` — a domain rule to fill in, citing the paired template's Business
  Rule number(s) where applicable. Insertion points use the `>>> ... <<<` marker.

**Conventions:**

- **No line numbers.** Scaffolds never hard-code line numbers; numbering is house-specific and
  deferred to `error-handling.md` (which may number via `Erl`, or not at all).
- **The `errHandler` block is shown once and referenced.** Because the errHandler block's form in
  `error-handling.md` is identical in every procedure, show it in full in the first procedure and
  reference it (`standard errHandler block`) thereafter.
- **Scope is explicit.** Each procedure is `Public` or `Private` as its usage requires; a sub
  performs an action, a function returns a value.

### 8.4 Staged execution and facilitation

Moved to `templates/_method.md`, method `staged-procedures`.

### 8.5 `validate` rules for `vba-scaffold`

1. `target_module` is present and non-empty.
2. Every entry in `new_procedures` has a matching `### <ProcedureName>` heading under
   `## Procedures`, and vice versa — the declared and documented procedure sets are identical.
3. Each `### <ProcedureName>` is followed by at least one fenced `vba` block.
4. `standards_layer` includes `error-handling`, and every value is recognized (§6).
5. If `implements` is set, it is a well-formed template slug. *(Format only — `validate` does not
   open the named template or check that cited Business Rule numbers exist; that is the human
   review gate.)*

### 8.6 Code that finds its own targets

Some generated procedures are handed a list of objects to work on. Others **discover** their
targets — asking the database which tables have a certain property, then acting on whatever comes
back. A procedure that discovers its own targets **and then changes or deletes something** follows
the rules below. They exist because discovery returns things nobody chose: the database's own
objects, and objects belonging to a developer who never opted into any of this.

**1. Names are a floor, never the mechanism.** Do not decide what is safe to change by reading an
object's name. A shop's tables may be `tblCompany`, `Company` or `CompanyT`, and their own
housekeeping tables may be called anything at all — a template cannot predict any of it, and a
template that guesses gets it wrong silently. Gate destructive action on one of two things
instead: **a list the developer confirmed**, or **a test that the artifact is one this template created**.
The second is usually available and usually better. Generated artifacts can be made to
carry a recognizable mark, and a mark survives every naming convention.

**2. Two names are off limits regardless, and they are off limits for different reasons.**

| Prefix | Whose it is | Rule |
|---|---|---|
| `MSys` | Access's own | **Read it freely when the information is needed** — which tables carry a Data Macro, what objects exist, and the like. **Never insert, update, or delete anything in it**, and never let a discovery routine make one a target of a change or delete action. No opt-in, no exception to the write prohibition. |
| `USys` | The developer's own hidden tables | Left alone **unless the developer opted that object in themselves**, by naming it in whatever list or configuration table the template uses for scope. |

The `USys` rule is not the same rule with a softer edge. Creating a `USys` table is a deliberate
act by someone who has taken responsibility for managing part of the database themselves; the
template's business is to leave that alone until invited. `MSys` is not the developer's to opt in
with in the first place — reading it for information is always allowed, but nothing about that
information ever earns it a place on a list of things to change.

**3. The name check has to come before the ownership test**, because a discovery routine must never
add an `MSys…` object to a change-or-delete list regardless of what the ownership test would say
about it — the name alone disqualifies it, before any other reasoning runs. Order the guards
accordingly.

**4. Back up before removing, and keep the backup when you decline to remove.** A procedure that
declines to touch something should still leave the developer a record of what it found.

## 9. `type: form-spec`

A `form-spec` template defines a **default, functional form layout** that edits a paired
`table-schema` and realizes its UI-level behaviors. It captures the controls, their arrangement (by
region and order), and the features the form must support — and **stops at function, not polish**: a
working, unstyled default ("ugly but correct" is a pass), aesthetics left to the adopter. It is the
most standards-dependent type: house design defaults and a reusable forms framework are deferred to
the standards layer, which the template **names, not redefines**.

Three layers, kept distinct:

- **`[LAYOUT]`** — controls + default arrangement (the template).
- **`[STANDARDS]`** — house design defaults (`form-conventions.md`) + the named forms framework.
- **`[BUSINESS LOGIC]`** — UI behaviors realizing the paired table-schema's Business Rules.

### 9.1 Front-matter (in addition to the common keys in §2)

| Key | Req | Notes |
|---|---|---|
| `implements` | optional | The `table-schema` the form edits (shared with `vba-scaffold`). |
| `record_source` | required | The form's record source (a query over the edited table). |
| `new_forms` | required | The forms/subforms the template defines; each subform appears as a `Subform` control in `## Layout`, and the main form is the one Layout describes. |

`standards_layer` must include `form-conventions`, and typically `naming-conventions`.

### 9.2 Body sections

In addition to the common core (§3), a `form-spec` template **must** contain, in this order:

| Section | Purpose | Required |
|---|---|---|
| `## Prerequisites` | The paired schema, the record source, the standards/framework depended on | required |
| `## Layout` | Named **regions**, each with an ordered **control inventory** (§9.3) | required |
| `## Features` | The functions supported, as named behaviors — citing the schema's Business Rule numbers for UI behaviors, and naming deferred framework patterns | required |
| `## Materialization` | How the spec becomes a real form (§9.4) | required |

A `form-spec` has no field-spec or procedure sections; its content is the control inventory + the
feature list.

### 9.3 Layout / control-inventory format

Layout is described **structurally, never by pixel**. Each region is a heading or labeled group; under
it, a control inventory table with exactly these columns:

```
| Control | Type | Bound to | Notes |
```

- **Control** — the control name (per `form-conventions` prefixes: `txt`/`cbo`/`chk`/`cmd`/`sfrm`).
- **Type** — Textbox, Combo, Checkbox, Subform, Button, Label, Image, …
- **Bound to** — the field/table the control binds to, or `—` for unbound/framework controls.
- **Notes** — one line: lookup target, multi-line, quick-add, a Parked UI behavior, etc.

Arrangement = region + row order; exact positioning and sizing default in the materialization step, not
in the spec. Hidden/internal controls (PK, sort key, image-link) are listed and marked.

### 9.4 Materialization

A form-spec materializes by **building the form live through an Access MCP server's form-creation and
control-creation tools**, with a default stacked layout and the code-behind wired to the named
framework helpers (and any paired `vba-scaffold`). The markdown is the source of truth.

The same design also expresses as **importable Access form text** (`SaveAsText`/`LoadFromText`), and
the markdown → Access-text mapping is proven by hand before a generator is built. **That text is not
a second build route.** Where no Access MCP server is connected the run ends at the approved design;
importable text is not generated for the developer to import as though it were a finished form,
because nothing will have opened it to confirm it loads. If they ask for it knowing that, it goes to
them headed `UNVERIFIED`. See `_materialization.md` for the full mapping rules and a hand-validated
fragment.

### 9.5 `validate` rules for `form-spec`

1. `record_source` is present and non-empty.
2. Every subform in `new_forms` appears as a `Subform` control in `## Layout`; the main form is the
   one Layout describes.
3. `## Layout` and `## Features` are present; `## Layout` has at least one control-inventory table
   conforming to §9.3.
4. `standards_layer` includes `form-conventions`, and every value is recognized (§6).
5. If `implements` is set, it is a well-formed template slug. *(Format only — `validate` does not open
   the named schema or confirm cited Business Rule numbers.)*

---

## 10. The OTS Wizard (any template type)

**The OTS Wizard is the process by which a build moves from the developer's first request to an
artifact that meets the request, converting what is unknown into what is known one decision at a time.** A
request arrives underdetermined. The template carries decisions already made, the standards layer
carries more, and what is left over is whatever only the developer can settle. The wizard is how
that remainder gets settled, before anything is built.

It works by interaction rather than by formula. No rule here computes an answer: the AI assistant
asks, the developer decides, and the build proceeds on what they settle between them. That
exchange is the point, and it is the thing a fixed procedure cannot reproduce.

**It is implemented as a short run of one-question steps**, each naming a **preferred choice**
(§10.7, deliberately not called a "default") and carrying an explanation that stays closed until
the reader opens it. That shape was chosen because it is the most familiar one available, not
because the process requires it. The unknowns, and the answers that resolve them, are the same
whether the developer meets them one at a time or all at once.

**Access wizards are familiar to the developers this library serves, and that is why this process
is called the OTS Wizard.**

### 10.1 What marks a template whose questions are written out

The `## Wizard` section (§10.2), and nothing else. There is no front-matter key for it: a reader
who has the template already has the section in front of them, so a flag announcing it would
arrive at the same moment and say nothing the section does not.

Where that section is present, the template's `warnings` and `house_assumptions` (§2) are
surfaced **at the step each one belongs to**, rather than all at once before the first question.
That is the one thing its presence changes outside the steps themselves.

### 10.2 The `## Wizard` section

A wizard template carries a `## Wizard` section immediately **before** the first section that
produces the artifact — `## Entities` for a `table-schema`, `## Procedures` for a `vba-scaffold`,
`## Layout` for a `form-spec`, or any declarations preamble that precedes one of those. The
developer decides, then sees what their decisions produce. Inside it, one `### Step <n> —
<question>` heading per step, in the order they are asked, each laid out like this:

```markdown
### Step 2 — Where should errors be recorded?

**Ask:** Where should errors be recorded?

| Option | Short description |
|---|---|
| `A table in this database` | Errors go into a table you can open, sort, and filter. |
| `A text file` | Errors are appended as lines of text to an external file. |

**Preferred:** `A table in this database` — this template's own; the standards layer does not
speak to this choice.

**Skip when:** Step 1 was answered "No".

<details>
<summary>Tell me more about where errors are recorded</summary>

One or two facts that might tip the choice, plus any warning that belongs to this decision.

</details>
```

The `<details>` block is how the *file* stores the explanation — collapsed on GitHub, so a person
reading the template sees the same shape the wizard has. **It is not how the explanation reaches
the developer at run time.** There is no collapsed block in a conversation, and simulating one by
writing "say *tell me more* if you want…" is exactly the failure §10.3 forbids: it turns a click
into typing. At run time *Tell me more* is an option in the selection control, and this block is
what that option shows.

Option rows list only the **substantive** answers. `Tell me more about <topic>` and
`Go back to the previous question` are added by the mechanism on every step (§10.3) and are never
written into the table.

### 10.3 How a step is asked — the mechanism

Moved to `templates/_method.md`, method `wizard`.

### 10.4 What the AI assistant says outside a step

Moved to `templates/_method.md`, method `quiet-build`.

### 10.5 Rules

Moved to `templates/_method.md`, method `wizard`.

### 10.6 The entry question

Moved to `templates/_method.md`, method `wizard`.

### 10.7 "Preferred choice", not "default" — and why the word matters

Moved to `templates/_method.md`, method `wizard`.

## 11. Minimal skeleton (`type: table-schema`)

```markdown
---
template: <domain>-<name>-schema
title: <Human Title>
domain: <domain>
type: table-schema
version: 0.1.0
status: draft
extends: <Host DB>            # if grafting onto an existing database
requires_tables: [<Existing>] # if extends is set
standards_layer: [audit-columns, naming-conventions, error-handling]
new_tables: [<TableA>, <TableB>]
---

# <Human Title>

## Intent
<what this produces and why; domain framing>

## Prerequisites
| Existing object | Used as | Notes |
|---|---|---|
| `<Existing>.<PK>` | <role> | <hook note> |

## Entities

### <TableA>
Grain: <one row per …>

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `<TableA>ID` | AutoNumber | PK | Surrogate key |

Indexes: PK on `<TableA>ID`.

## Relationships
- `<Parent> (1) → (∞) <Child>` on `<Field>` — cascade behavior

## Business Rules
1. <rule>

## Standards Layer
- **Audit columns** — supplied by the host audit convention.
- **Naming conventions** — <house style this template follows>.
- **Error handling** — house pattern for any VBA generated alongside.

## Extra Options
*Named optional extensions, none of them filled in for an engagement.*
- <named optional extension>
```

## 12. `type: outcome-first`

An **outcome-first** template states the finished condition a build has to reach and says nothing
about the route to it. Where a `vba-scaffold` (§8) ships working procedures with the house-specific
parts marked for substitution, an outcome-first template ships no code at all. It specifies what the
database must do once the build is finished, how the developer confirms it, and what is theirs to
decide — and leaves the decomposition, the naming and the code itself to whoever builds it under
the standards layer.

The two are not a draft and a finished version of each other. **A domain may carry one of each**,
built from the same paired `table-schema` and run against the same database, so a shop can see
what the two methods produce for one requirement. `templates/audit/` carries the proven pair:
`audit-logging-lite-scaffold.md` (§8) and `audit-logging-lite-outcome-first.md` (§12).

An outcome-first template hands the route to whoever builds it, so it assumes a builder that can be
trusted with one. Where the assistant in use cannot be, the `vba-scaffold` in the same domain is the
path, and that is the second reason a domain carries both.

### 12.1 Front matter

The common core of §2 applies unchanged. Beyond it:

| Key | Required? | Notes |
|---|---|---|
| `type` | required | `outcome-first` |
| `implements` | conditional | the `table-schema` template (by slug) whose tables the build realizes |
| `standards_layer` | required | must include `error-handling` — the build generates code, so a house pattern for reporting a failure always applies |

**`target_module` and `new_procedures` must be absent.** They name a module and a list of
procedures, which is route specification. Declaring either is the one way a template of this type
stops being one, and `validate` reports it.

### 12.2 Body sections

In addition to the common core (§3), an outcome-first template **must** contain:

| Section | Holds |
|---|---|
| `## What you end up with` | The specification: what the database does once this is built, in terms the developer can check without reading any code. Carries the checks that confirm it and the behaviours that must hold however the work was divided up. Its checklist subsection opens with a line telling the developer to confirm every check ran, not only the ones the AI mentions running, pointing to `README.md`, "Check that every validation check ran." |
| `## Information and conditions you need to supply` | Everything only the developer can answer, and which of those are gates that stop a build. |
| `## To the AI assistant building this` | The build instruction, addressed to the assistant and marked as such where it starts. |

There is **no `## Procedures` section**, and `validate` rejects one.

**This section restates the build record's timing rule in its own words — it does not rely on
`_materialization.md` alone to carry it.** Per §12.4, a requirement stated outside the sections this
heading names is context, not specification, so `_materialization.md`'s *"write it before you say the
build is finished, not when you are asked for it"* does not reach the build unless this section says
so too. Tie the restatement to the checks this section already requires the builder to run: the build
record is where each result is recorded, it exists before the build is reported finished, and it is
never written later, only when asked for.

The build-record checklist rule moved to `templates/_method.md`, method `checklist-rule`.

### 12.3 `validate` rules for `outcome-first`

| Rule | Check |
|---|---|
| `OF` | each of the three required sections above is present |
| `OF1` | `standards_layer` includes `error-handling` |
| `OF2` | `implements`, where set, is a well-formed slug |
| `OF3` | no `target_module`, no `new_procedures`, and no `## Procedures` section |

### 12.4 The specification is only what the closing instruction names

An outcome-first template hands its builder no route, so **every statement that binds has to sit where
the builder is told the binding statements are.** `## To the AI assistant building this` names the
sections that make up the specification. A requirement stated anywhere else — in `## Intent`, in a
platform-facts or background section, in a list of what the template does not do — is read as context
and does not reach the build.

This is not a limit on what a builder can read. It is a limit on what it can infer. A person assigns
weight to a sentence by working out why its section is there; a builder assigns weight by designation.
So a load-bearing sentence under a heading like "Facts about the platform" is invisible as a
requirement however plainly it is written, and no amount of rewriting it helps.

**This has happened.** The audit outcome-first template stated its central mechanism in a facts
section — three headings away from the promise it served, and outside the sections its own closing
instruction named as the specification. A build substituted a different mechanism, and passed every
check the template listed while doing it.

Two authoring tests, both answerable yes or no:

1. **Is every binding sentence inside a section the closing instruction names as the specification?**
   If a sentence would make a build wrong by its absence, it belongs in one of those sections. Others
   may restate it; none may be its only home.
2. **For each promise, what is the cheapest build that passes the sentence as written?** If a mechanism
   you did not intend passes it, the sentence is a summary of the requirement rather than the
   requirement itself. Strengthen it until the unintended build fails — by naming the condition that
   tells the two apart, which is still an outcome, and not by naming the route.

A third test, common to every template type, is in §1: for each thing the builder must do, which
section says so, and whom does that section address?

**A mechanism may be named in the specification only where the platform leaves exactly one.** Where two
would both satisfy the promise, name neither: that is route, and it is the builder's. Where only one
can satisfy it, the mechanism is part of the outcome, and saying so is not route specification. Give
the reason in the same breath, so the constraint travels with what makes it necessary — a constraint
whose reason is somewhere else is the defect above in a new place.

A section that does not bind says so in one line where it starts, and says where the binding statements
are instead.

**What arrives with the template binds too, and the closing instruction says so.** The template
library MCP server delivers the method and the platform facts a build needs alongside the template body, outside the
sections the closing instruction names. By the rule above they would be read as context. So every
`## To the AI assistant building this` carries, right after it names the specification sections, this
sentence: *"The method and platform facts delivered with this template bind this build as fully as the
sections named above."* A template body then restates a delivered fact only where it applies it, and
names the fact by its id.

### 12.5 The `Explore options` step

Moved to `templates/_method.md`, method `explore-options`.

---

## 13. Masters and steps (any template type)

**Why templates are divided.** An AI client shows a tool's answer in full only up to a size; above
it, the answer is saved to a file and the assistant sees a pointer instead. One client measured on
2026-10-08 did this above 50,000 characters. An assistant that has to go and read a file has been
handed a pointer, not the knowledge, and may never read it. So the template library MCP server
delivers a template in **parts**, one per call, each under `RESPONSE_LIMIT` (45,000 characters,
leaving headroom for clients not measured), and each at the point in the run where it is used.

**The parts, in order:**

1. `master`: the template file itself, from `get_template`.
2. `standards-gate`: the standards gate (`templates/_standards-gate.md`), which every run asks
   before any other question. A from-scratch run gets it from `get_standards`.
3. `design-standards`: the standards the design follows.
4. `design-facts`: the platform facts the design depends on.
5. Any step marked `route: both` (below).
6. `build-method`: how a build is conducted. The design route stops before this part.
7. `build-standards`: the standards the code follows (`error-handling`, `query-style`,
   `startup-conventions`).
8. Each remaining step, in order. A template not yet divided gets one `build-facts` part here instead.

Every answer names the next call and when to make it (`next`), and carries `served` (part, version,
hash), so a build record shows every part fetched and any part that was not. A standards or facts
part too large for one answer is split by the server into numbered parts; a step is not split by the
server, because its author knows where the seams are.

**A divided template is a master plus a folder of step files.** The master stays at its usual path
and keeps every section §3 and its type require for the design: its intent, prerequisites, house
assumptions and warnings, and what the developer must supply. Step files live in
`<template-id>.steps/` beside it, and the master lists them under `steps`, in order. Like a main
procedure calling its subprocedures, the master holds the order; each step holds the work.

**A step file:**

```markdown
---
step: 02-rules-module
title: "Write and import the rules module, modTimeOffRules"
platform_facts: [vba-import-xml-entities, mcp-module-import, mcp-line-numbers]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures
...
```

- `step` matches the file name; `title` is quoted, since titles often hold a colon.
- `platform_facts` names the facts this step uses. A fact may be named by the master and by a step,
  and by more than one step: it is delivered each time, at each point it is used.
- `route: both` marks a step the design is drafted against (a checklist the design must pass, the
  wizard's design-time questions). It arrives with the design parts, on both routes. Such steps are
  listed before every build step.
- `when:` overrides the default timing text for the step's part.
- A step opens with an `##` heading or a reader line; one that opens on a `###` heading would be read
  as part of the previous step's section. Procedures in a step sit under a `## Procedures` heading
  (`## Procedures (continued)` after the first); `validate` reads the master and its steps together
  for the §8 and §12 rules.
- The wizard method arrives with whichever part holds the template's `## Wizard` heading, once.

**Dividing an existing template moves its text; it does not reword it.** Only the version, the
`platform_facts` line, `steps`, and any lead a step needs to say what it is are new.

**Rules `validate` enforces:** every listed step has a file whose front matter parses and names it,
with a `title` and a reader line (ST1, ST4), and no step file is unlisted (ST1); every step's facts
resolve (ST2); no fact marked `route: build` sits in a divided template's master (ST3); `route: both`
steps come first (ST5); and every part the run can fetch is within `RESPONSE_LIMIT`, measured as the
server sends it (SZ1). While the library is being divided, `steps` is optional (ST0 off).
