---
template: _template-schema
title: Open Template Scaffolds — Canonical Template Format
domain: _meta
type: spec
version: 0.8.0
status: draft
---

# Open Template Scaffolds — Canonical Template Format

**Who reads this:** anyone writing a template, and the AI assistant reading one.

**Using a template to build tables, forms, or code in a database?** You do not need to read this
file in order to use the template.

This is the **format specification** every template file in this library must follow.
It is the contract the template library MCP server keys off: discovery (`list_templates`,
`search_templates`) reads the front-matter; `get_template` composes the body with the
active standards layer; `validate` checks a template (or a filled-in copy) against the
rules in this document.

It is meta, not a template itself (`type: spec`, `domain: _meta`) — `validate` skips files
whose `type` is `spec`.

A template is a vetted, standards-baked *starting point*, not a drop-in guarantee. Every
template a developer adopts must be confirmed fit for the developer's intended application
before use — the library supplies structure and proven decisions; suitability for a specific
engagement is always the adopter's judgment.

> **Scope note (build order).** The common core below is proven against real templates:
> `templates/northwind/stocktake-schema.md` (`type: table-schema`),
> `templates/northwind/stocktake-scan-scaffold.md` (`type: vba-scaffold`), and
> `templates/library/publication-form.md` (`type: form-spec`). **All three type sections (§4, §8,
> §9) are authoritative — the template format is complete**, each proven by hand before the
> template library MCP server's schema-dependent tooling is built against it.

---

## 1. File rules

- **Format:** Markdown (`.md`), UTF-8. No binary, no executable code.
- **Location:** `templates/<domain>/<name>.md`. Files prefixed `_` (e.g. this one) are
  library infrastructure, not domain templates.
- **One template per file.** A template defines one cohesive artifact set for one domain.
- **Single source of truth:** the file *is* the template. CLAUDE.md and the template library
  MCP server are readers of it, never separate copies.
- **Write for a reader who has never seen this system.** Every term is defined where it first
  appears, or it isn't used — no undefined jargon, no coined label the reader has to look up, no
  concept assumed from another file. Prefer the plain name for a thing over the internal one.
  Where a rule has a visible consequence, state the consequence as well — what the reader will
  actually see happen — not only the rule that governs it. If a sentence would send a first-time
  reader to a search engine or to another file to understand it, rewrite it. This applies to the
  whole template, not to its introductory sections alone: field tables and Business Rules stay
  precise, but precision and plain language are not in tension. Where a plain paraphrase would
  collide with a term the reader probably already owns, or with a word that already means something
  else here, name both once — plain sentence first, technical term second (§10.5 rule 3).

---

## 2. Front-matter (YAML)

The front-matter is the machine-readable contract. Keys below marked **required** must be
present on every template; conditional keys are required when their condition holds.

| Key | Req | Type | Notes |
|---|---|---|---|
| `template` | required | string (kebab-case) | Unique slug; matches the filename stem |
| `title` | required | string | Human-readable title |
| `domain` | required | string | Domain folder name (e.g. `northwind`, `sales`, `hr`); `_meta` reserved for infra |
| `type` | required | enum | `table-schema` \| `vba-scaffold` \| `form-spec` \| `spec` |
| `version` | required | semver string | Template version, counted per template and independent of any library version. **Bump it in the same commit as any change to what the template produces** — patch for a correction an adopter needn't act on, minor for anything they would (a new or renamed field, a changed default, an added rule or section), major for a redesign an existing build can't absorb. It is the only thing that distinguishes a copy someone took earlier from the current file; see `CONTRIBUTING.md` → *Versioning a change* for why this is a rule and not a nicety |
| `status` | required | enum | `draft` \| `review` \| `stable` |
| `extends` | conditional | string | Required when the template grafts onto an existing database; names the host (e.g. `Northwind (Access Developer Edition)`) |
| `requires_tables` | conditional | list[string] | Existing tables the template hooks into. Required when `extends` is present |
| `requires_fields` | optional | list[string] | Specific existing fields relied on, as `Table.Field` |
| `standards_layer` | required | list[enum] | Which standards-layer concerns this template defers; values from §6 |
| `new_tables` | conditional | list[string] | Tables the template defines. Required for `type: table-schema`; must match the `## Entities` headings exactly |
| `implements` | conditional | string | For `type: vba-scaffold` and `form-spec`: the `table-schema` template (by slug) the scaffold realizes / the form edits |
| `target_module` | conditional | string | Required for `type: vba-scaffold`; the module or class the procedures live in |
| `new_procedures` | conditional | list[string] | Required for `type: vba-scaffold`; must match the `## Procedures` `### <name>` headings exactly |
| `record_source` | conditional | string | Required for `type: form-spec`; the form's record source (a query over the edited table) |
| `new_forms` | conditional | list[string] | Required for `type: form-spec`; the forms/subforms defined — each subform appears as a `Subform` control in `## Layout` |
| `seeds` | optional | list[string] | Seed data the template expects, as `Table.RowKey` |
| `house_assumptions` | optional | list[string] | House-particular modeling assumptions deliberately kept in the template body (the "Declared" tier) because they can't be moved to the standards layer or dropped. Each entry is `Target — rationale`, where `Target` names the entity, field, or rule carrying the assumption. Makes embedded house bias machine-visible to adopters and discovery tools. |
| `warnings` | optional | list[string] | Hard **platform caveats** the AI builder must surface to the developer *before* building — engine limits, not house bias (e.g. "Data Macros cannot audit Long Text fields — confirm whether any audited table has one"). Surfaced alongside `house_assumptions` in the review step; each entry states the limit and what the developer must confirm or the build must branch on. |
| `wizard` | optional | boolean | `true` when the template's set-up decisions are presented as an OTS wizard (§10). A wizard template surfaces its `warnings` and `house_assumptions` **at the step each one belongs to** rather than all at once before the first question |

**Rules the `validate` tool enforces on front-matter:**

1. All required keys present and non-empty; `type` and `status` within their enums.
2. `template` is a unique, kebab-case slug that **ends with** the filename stem; a
   domain or pairing prefix may precede it (e.g. file `library/catalog-schema.md` →
   slug `library-catalog-schema`; a form paired with the catalog → `library-catalog-publication-form`).
3. If `extends` is set, `requires_tables` is non-empty.
4. Every entry in `new_tables` is documented under `## Entities` — either as its own `### <name>`
   heading or as a named row in a grouped lookup sub-table (§4) — and vice versa; the declared and
   documented table sets are identical.
5. `requires_fields` / `seeds` entries are well-formed `Table.Field` / `Table.RowKey`.
6. Every `house_assumptions` entry is well-formed (`Target — rationale`), and each `Target`
   resolves to an entity, field, or rule named in this template. (Format check only — `validate`
   cannot judge whether something *should* have been declared; that stays the human review gate.)

**Scope of `validate` — format, not fitness.** `validate` confirms a template is internally
well-formed: complete front-matter, `new_tables` matching the `### <name>` entity headings, every
`FK → <Table>` resolving to a table named in the template. It does **not** open any host
database — that's a separate `check_compatibility(template, db_path)` tool. And neither check
speaks to fitness: a passing `validate` means well-formed, never *suitable*. Confirming a
template fits the intended application stays the adopter's responsibility.

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

---

## 4. Body sections — `type: table-schema`

In addition to the common core, a `table-schema` template **must** contain, in this order:

| Section | Purpose | Required |
|---|---|---|
| `## Prerequisites` | What must be settled before the tables are built: **which file they go in** (see below), and — when the template grafts onto an existing database — the hooks into it, as a table of existing objects the new tables wire into. Required when `extends` is set; otherwise optional, but expected wherever the deployment note applies | conditional |
| `## Entities` | One `### <TableName>` per new table — grain statement, field table (§5), `Indexes:` line. **Trivial, uniform lookups** (`<name>ID` + a descriptor + optional `SortOrder`, nothing more) may instead be **grouped** in a single sub-table of name + seed rows — documentation shorthand only; each row is still its own discrete table (this is *not* a shared/MUCK lookup table). Any lookup carrying extra structure (more fields, an FK, a description) takes its own `### <name>` heading like an entity. | required |
| `## Relationships` | New relationships and hooks into the host schema, as a bulleted list naming parent → child, the join field(s), and cascade behavior | required |
| `## Business Rules` | Numbered list of the logic the generated objects must honor (grain constraints, rollups, derivations, deferred-logic notes) | required |

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

### 4.1 `validate` rules for `table-schema`

1. Every table documented under `## Entities` — whether as a `### <name>` heading or as a named row
   in a grouped lookup sub-table — appears in front-matter `new_tables`, and vice versa.
2. Every field table conforms to §5 (columns and type vocabulary).
3. Every `FK → <Table>` named in a field table resolves to either another entity in this
   template, a `requires_tables` entry, or another `new_tables` entry.
4. Every table named in `## Relationships` is an entity, a lookup, or a `requires_tables` entry.
5. Audit columns (`AddedBy`, `AddedOn`, `ModifiedBy`, `ModifiedOn`) do **not** appear in field
   tables — they belong to the standards layer (§6) and are flagged if present.

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
| `audit-columns` | `AddedBy` / `AddedOn` / `ModifiedBy` / `ModifiedOn` on new tables, supplied by the host's audit convention — never in the template body |
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

---

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
- **The `errHandler` block is shown once and referenced.** Because the VBE-reflection form in
  `error-handling.md` is identical in every procedure, show it in full in the first procedure and
  reference it (`standard errHandler block`) thereafter.
- **Scope is explicit.** Each procedure is `Public` or `Private` as its usage requires; a sub
  performs an action, a function returns a value.

### 8.4 Staged execution and facilitation

Some `vba-scaffold` templates document procedures meant to run in a specific order, where each
step gates a decision the developer must make before the next one runs — picking among named
build options, reviewing a generated list before the next procedure acts on it, and the like.
When a template documents this kind of sequence, it must also state, alongside the sequence, a
facilitation rule for any assistant carrying out the steps on the developer's behalf:

- **Never infer the answer to a staged decision.** Not from the shape of the data, not from
  domain reasoning that makes an answer seem obvious — ask the developer, and wait for their
  actual answer, even when it looks predictable.
- **Never substitute your own analysis for a procedure whose job is to answer the question.** If
  the sequence includes a check or scan procedure, run *that procedure*, at the point the
  sequence calls for it — don't read the underlying data directly and report a conclusion in its
  place.
- **Present one step at a time.** Don't collapse a staged sequence into a single upfront report,
  even where every fact in it turns out correct — the sequence exists so the developer reviews
  and approves each gate, not just the end state.
- **Restate the decision in full at the gate.** Ask where the developer can answer without
  reconstructing anything from earlier in the session: what the setting means, what it produces
  at run time, and what changes if they choose the other way — at the point of asking, not on
  request. And never use one number for two quantities in the same message: if 11 fields are
  auditable and 11 macros will be generated, say which is which, because a reader will otherwise
  take them for the same 11.

This is in addition to — not a substitute for — a project's own standing rule that no edit happens
without explicit approval. It addresses a different failure mode: an assistant that has enough
context and initiative to *answer* a gate the developer was meant to answer, even where it never
touches a file. (See the `## Wizard` section of `templates/audit/audit-logging-lite-scaffold.md`
for a worked example of this note in place, next to the steps it governs. §10 is how a template
presents such a sequence to the developer; this section is the rule the AI assistant follows while
running it.)

### 8.5 `validate` rules for `vba-scaffold`

1. `target_module` is present and non-empty.
2. Every entry in `new_procedures` has a matching `### <ProcedureName>` heading under
   `## Procedures`, and vice versa — the declared and documented procedure sets are identical.
3. Each `### <ProcedureName>` is followed by at least one fenced `vba` block.
4. `standards_layer` includes `error-handling`, and every value is recognized (§6).
5. If `implements` is set, it is a well-formed template slug. *(Format only — `validate` does not
   open the named template or check that cited Business Rule numbers exist; that is the human
   review gate.)*

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

A form-spec materializes as **importable Access form text** (`SaveAsText`/`LoadFromText`) with a
default stacked layout and the code-behind wired to the named framework helpers (and any paired
`vba-scaffold`). The markdown → Access-text mapping is proven by hand before a generator is built; the
alternative path builds the form live through an Access MCP server's form-creation and
control-creation tools. The
markdown is the source of truth; the Access text is a generated target. See `_materialization.md` for
the full mapping rules and a hand-validated fragment.

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

## 10. Wizard steps (any template type)

A template may present its set-up decisions as an **OTS wizard**: a short run of one-question
steps, asked one at a time, each naming a **preferred choice** (§10.7 — deliberately not called a
"default") and carrying an explanation that stays closed until the reader opens it.

**It is a presentation device, not a second build path.** The decisions and the artifact they
produce are exactly what they would have been without it. What changes is the order the developer
meets things in: one question at a time, with the reasoning and the warnings available on request
instead of fired at them before they have chosen anything.

**It is not an Access wizard.** Nothing is installed in the developer's database to run it, and no
form is left behind. The AI assistant asks the questions in conversation; this file is where the
questions, the options, and the explanations are written down. The resemblance to an Access wizard
is one of shape only.

### 10.1 Front-matter

`wizard: true` (§2) marks a template whose set-up runs this way.

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
| `A table in this database` | Errors go to a table you can open, sort, and filter. |
| `A text file` | Errors are appended as lines of text to a file. |

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

**Every step is put to the developer through the interactive selection control** — the one that
renders each option as something they click. **A step rendered as prose, a markdown table, or a
list the developer has to answer by typing is a failed step**, however good its content: it asks
them to compose an answer where they were promised a choice.

What this file holds is the **source** for that control, not the thing shown. The `**Ask:**` line
becomes the question, each option row becomes a clickable option with its short description
underneath, and the `**Preferred:**` line is stated with them.

Four rules follow from it:

- **One step per ask.** Never two steps in one control, even where the second seems to follow.
- **`Tell me more about <topic>` is always the last option**, on every step. There is no other way
  for the developer to reach it — they must be able to click it. Choosing it shows the
  *Tell me more* text and then **asks the same step again, unchanged**, so the explanation costs
  them nothing but a click.
- **`Go back to the previous question` is an option on every step after the first**, wherever there
  is room for it alongside the substantive options and *Tell me more*.
- **The control takes at most four options.** A step therefore carries **at most three substantive
  answers plus *Tell me more***. A decision with more than three natural answers is either two
  decisions, or has two answers that should be one — resolve it in the template. Never resolve it
  by dropping *Tell me more*, and never by silently cutting an option (rule 8).

### 10.4 What the AI assistant says outside a step

§10.3 governs the step itself. Everything else said during a wizard — before the first question,
between two steps, before the build begins, and while it runs — has no specified shape, and
unspecified space is where ordinary explaining habits reassert themselves. Four rules govern it.
**The first two fix opposite problems, and neither is a rule about being brief.**

Nothing here is written in a template file. These are run-time rules: what the AI assistant says as the
wizard runs, composed in the conversation and never authored anywhere.

**1. Between two steps, name what was recorded and what is being asked next.**

One line for each, in the developer's words:

> *"Errors will go to a table, with a text file as the fallback. Step 4 asks where that table lives."*

Not:

> *"Step 4."*

A step number says where a question sits in a list, not what it is. A developer four questions in has
no other confirmation that the answer they clicked registered. **This rule makes what is said between
steps longer, not shorter** — that is what it is for. Where the wizard branches, or a step is asked
twice, this is where that is said.

**2. Before the first question, and before the build, say only what the developer must act on.**

These are the two moments with the most to report — the template that was matched and why, the
build-wide warnings, the house assumptions, what was found on opening the files — and the least use
for it. At the first, the developer has chosen nothing yet. At the second, they have chosen
everything and are waiting. Three things are said at these moments and nothing else:

- Anything they must answer or confirm — asked as a question, never stated in prose (rule 3).
- Anything that changes what they do next.
- The disclosure line below, before the first question only.

Everything else — what was checked, what was found, what it meant — goes to **the build record**.

**This rule does not govern the design presented for approval.** The diagram and field detail are the
deliverable the whole workflow exists to produce; they are not narration, and they are not shortened.
The rule governs the prose around them.

**The build record is always written**, and delivered as a file alongside the artifact. Without it
this rule deletes the detail rather than routing it, and the disclosure line promises something the
format does not keep. **`templates/_materialization.md`, "The build record", defines what it is
called, where it goes, and what belongs in it.**

**3. A house assumption is asked, never assumed.**

A template's `house_assumptions` entries are surfaced before the first question. **Surfacing is not
asking.** *"I'll take that as confirmed unless you say otherwise"* states the assumption and then
answers it on the developer's behalf — it **ignores their input rather than requiring it**, and the
developer who says nothing has not agreed to anything. It is put through the selection control like
any other decision, and the build waits for the answer.

This is the §10.7 trap one level up, and worse: a preferred choice at least appears in a question the
developer is looking at. A self-confirmed assumption appears in prose they were free to skim — and an
expert skims and loses nothing, while a newcomer skims and misses the one line that mattered.

**4. While the build runs, do not narrate it.**

The questions are over and the developer is waiting for a result. Everything happening now is work
they already approved, so a running commentary on it reports progress to nobody: they cannot act on
it, cannot verify it, and cannot tell from it whether anything is going wrong. Rule 2 covers the
moment before the build; this covers the build itself, which is longer and where the habit is
strongest. Three things are said between the last question and the finished artifact:

- **That it has started**, once, and what it will produce. Silence for several minutes is its own
  failure — this is the line that prevents it.
- **Anything that needs the developer to act** — a failure they have to clear, or something the
  build hit that no question covered. Asked as a question, never narrated past.
- **That it is finished**: what was built, and where the build record is.

Every object created, every procedure run, every check that passed, every step that went exactly as
expected: all of it goes to the build record. **A build that goes to plan produces nothing between
its first message and its last.**

**Progress commentary has a real audience, and it is not this one.** Someone developing or trialling
a template does want to watch each step land — they are reading for the template's behaviour, not
for their own database, and that is a different reader (see the three readers in `CLAUDE.md`). They
will say so. Absent that, the developer wants their tables, not a transcript of them being made.

**The disclosure line**, said once, immediately before the entry question (§10.6):

> *"While I build your \<artifact\>, I keep notes as I go — what I checked, what worked, and anything
> that surprised me. When it's finished you get them as a file alongside the \<artifact\> itself: a
> record of how it was built, not just the thing. Anything you need to decide is in a question I ask
> you. Nothing you have to act on will be buried in what I say in between."*

**Do not tell the developer they may skim.** Knowing which paragraph is safe to skip is what
experience buys: an expert skims and loses nothing, a newcomer skims and misses the one line that
mattered. The disclosure line gives an anchor instead — everything you must act on is in a question.

### 10.5 Rules

1. **One decision per step.** A step that asks two things is two steps.
2. **The `Ask:` line is one short question, in the developer's words.** No clause explaining why it
   is being asked, no naming of the machinery behind it. "Which tables should be audited?" — not
   "Which of your tables should the scan consider for auditing? This is the one boundary decided in
   code, and everything finer-grained is a switch you flip in a table afterwards." Words like *the
   scan*, *the generator*, *the config table*, *the boundary* mean nothing to someone meeting this
   for the first time; what they convey is that they are out of their depth, and the likeliest
   response is to stop using it. **If the question needs a second sentence, that sentence belongs in
   *Tell me more*.**
3. **One name per thing, from the first step to the last.** Once something has been named — a file,
   a folder, a setting, a table, a step — it keeps that name in every question, every option, and
   every *Tell me more*. No synonyms, no switch to the more technical term later, no shortening
   after first use. **A second name for something already named is a defect even when both names
   are correct**: a new word signals a new thing, so the reader stops to work out what the
   difference is and finds none. That pause costs more than the repetition would have. **Where the
   plain name and the precise name compete, use the plain one** — a reader who feels talked down to
   is annoyed and keeps going; a reader who is not sure two words mean one thing has already lost
   the thread, and may not know they lost it. If the precise name is genuinely needed, it replaces
   the plain one from first use. **The exception is a term the developer probably already owns** —
   *referential integrity*, *cascade delete* — **or a plain word already taken by something else in
   this material**, where the paraphrase misdirects rather than merely under-informs. There, name
   both once, plain sentence first and the technical term marked as such: *"nothing in the database
   enforces that reference — the database term for this is referential integrity."* After the
   pairing the plain name carries on alone; pairing is a definition given once, and a synonym
   appearing later is still a defect. Where the plain word is taken, the pairing belongs **in the
   question**, not in *Tell me more* — a reader who does not open *Tell me more* has already taken
   the wrong meaning.
4. **A short description says what the option *is*, in one line — never why it is better.** No
   bolding, no ordering by preference, no "recommended". Every comparison lives in *Tell me more*.
   A description may carry a consequence the developer needs *at the moment of choosing* ("any Data
   Macros those tables already have are replaced"), but never the reasoning behind it.
5. **Error numbers, engine limits, version caveats, and internal names never appear outside *Tell
   me more*.** Someone who meets "error 3870", "VBE reflection", or "`Application.LoadFromText`" in
   a question they are being asked to answer learns one thing: this was not written for them. Put
   it one click away, where the person who wants it will find it and nobody else has to.
6. **Every step names a preferred choice — never a "default".** See §10.7. The `**Preferred:**`
   line is the only signal a reader gets about which option the library would point at first, and
   it is enough: no bolding, no "(Recommended)", no argument. Where the standards layer answers the
   question, the preferred choice is that answer; where the standards layer is silent, it is the
   template's own and the line says so. It may follow an earlier answer, in which case the line says
   which step it follows.

   **Say where it came from in plain words — never as a file name or a section number.** Rule 5
   forbids an internal name in a question, and a `Preferred:` line is part of the question. So write
   *"the naming style these templates follow"*, not `standards/naming-conventions.md` §1.1. **A line
   that cites a file forces whoever reads it aloud to invent a paraphrase**, and the paraphrase is
   then unreviewed: one such line produced *"from **your** naming conventions"* in a live run —
   claiming the developer had authored a file they had never seen. Give the spoken wording in the
   template and there is nothing to invent.

   **Avoid the possessive entirely.** *"Your standards"* is wrong for anyone who has not adopted a
   layer; *"the house standard"* assumes a house the reader may not have; *"the template's"* is
   inaccurate, since the template follows the layer rather than defining it. *"The standards these
   templates follow"* claims nothing about whose they are.
7. **A confirmation step has no preferred choice.** Where a step asks the developer to attest to
   something rather than to prefer something — that they have a backup, that a list the build will
   act on is correct — write `**Preferred:** none` and say why: nothing the library picks can stand
   in for the developer's own word.
8. **Options are re-ranked, never removed.** A choice the library ranks last is still offered, in
   the same plain form as the others.
9. ***Tell me more* stays closed until asked for** and gives one or two facts that might tip the
   choice — drawn from the standards files and the template's own description, not restated from
   them, and not exhaustive.
10. **Warnings live at the step they belong to.** A front-matter `warnings` entry that governs one
    decision is surfaced inside that step's *Tell me more*; one that governs the whole build is
    surfaced before step 1. This is the point of the format: the warnings are not less visible, they
    are visible where they are actionable.
11. **A choice made against the standards layer holds for that run** — carried forward to every
    later step, never quietly reverted, and never written back to the standards files. The next run
    starts from the standards again. Flexibility within limits.
12. **Going back is always available.** Every step after the first offers it, and the developer may
    name any earlier step at any time. **Changing an answer discards every answer after it** and the
    wizard resumes forward from the changed step — so a revised decision can never leave a stale one
    standing behind it.
13. **A wizard of more than three steps opens with the entry question** (§10.6), which is where the
    developer chooses whether to answer every step or have the preferred choices used. It is never
    an option inside Step 1.
14. **Ending early ends one wizard, not the run.** Where a step-1 answer declines the whole feature,
    that wizard stops; any other wizard in the same template is asked independently.
15. **§8.4's facilitation rules apply in full.** Never infer the answer to a step, present one step
    at a time, and restate the decision at the gate so it can be answered without reconstructing
    anything from earlier in the session.

### 10.6 The entry question

**A wizard of more than three steps opens with one question before Step 1**, asked through the same
selection control as every other step:

> **Ask:** This takes *n* questions. Do you want to answer them, or shall I just build it?

| Option | Short description |
|---|---|
| `Ask me the questions` | Go through them one at a time. |
| `Just build it` | I use the preferred choice at each step, and only stop where a step needs something from you. |

**Preferred:** `Ask me the questions`.

It exists because an instruction to proceed — "find a template and run it" — is not permission to
put seven questions in front of someone. The entry question costs them one, and it is the only
place the wizard interposes itself between the instruction and the build.

Six rules govern it:

- **Asked once, before Step 1, and never again.** It is not an option inside Step 1, and no later
  step re-opens it.
- **`Just build it` cannot skip a confirmation step** (rule 7). A step with no preferred choice has
  nothing to fall back on, and passing one silently would answer for the developer on exactly the
  questions they were meant to answer. Say up front how many of those remain.
- **State the preferred choices before acting on them** — the answer being used at each skipped
  step, in a short list. `Just build it` authorizes known answers; it is not consent to be
  surprised.
- **A preferred choice that contradicts what the developer asked for is not a preferred choice on
  that run — ask the step, and say why you are asking it.** Preferred choices are written into a
  template before anyone has said what they want, so a request can arrive that one of them directly
  contradicts. A developer who asks for the feature on the database they already have has ruled out
  the step whose preferred choice builds a set of sample tables to try it on; using it anyway is
  precisely the surprise the rule above forbids. This applies to one step at a time — the rest of
  `Just build it` stands.
- **Ask it even when the developer sounded impatient.** Especially then: an imperative instruction
  is what this question is for, and answering it takes one click.
- ***n* is this template's own count** — the steps in its `## Wizard` section that apply to this
  run, a number the developer could arrive at from the file. A run can turn up questions no
  template carries: a build route where a connected tool offers one, a file that has to be made
  writable first, a step re-asked under the rule above. Don't fold those into *n* and don't try to
  predict them. Ask each where it arises and say it is one more than the number given at the start.

### 10.7 "Preferred choice", not "default" — and why the word matters

**A wizard step names a *preferred choice*. This library does not use the word "default" for it,
anywhere, deliberately.**

"Default" carries two meanings and nothing in the word says which is meant:

- **the choice we would point at first** — a recommendation, which still has to be offered; and
- **what happens when nobody chooses** — a fallback that fires on its own.

Written into a file that an AI reads and acts on, the second meaning wins. A step labelled
`Default:` reads as standing permission to skip the question, and the question stops being asked.
That is not a hypothetical: it is how `standards/error-handling.md` came to say *"emit option 3…
and say so"* and how a build came to pick its own error-handling option and announce the result to
a developer who had never been asked.

**The preferred choice becomes the answer in exactly two situations, and both are an act by the
developer:**

1. They decline to choose — "you pick", "whatever you think".
2. They ask to get on with it — the entry question's *"just build it"* answer (§10.6).

**It never becomes the answer on the AI assistant's initiative.** No amount of obviousness, data
shape, or convenience converts a preferred choice into a decision nobody made.

> **A note for anyone writing a template.** This distinction does not arise when you write code:
> a default parameter value simply *is* the fallback, and no reader expects otherwise. It arises
> the moment your reader is an agent that will act on what you wrote. Words that are precise in a
> function signature turn ambiguous in an instruction, and the ambiguity resolves toward *action*,
> because acting is what the reader is there to do. When in doubt, name the act you want and the
> act you don't.

`validate` does not check §10 at all — the format is proven by hand first, exactly as the three
template types were (see the scope note at the top of this file). `templates/errors/error-logging-scaffold.md`
is the worked example.

**Two different gaps sit inside that, and only one of them closes.** §10.2 and §10.5 describe things
that are in the file — a `### Step n —` heading, an `**Ask:**` line, a `**Preferred:**` line, a
two-column option table, a `<details>` block per step — and a checker could assert every one of them.
**§10.4 cannot be checked here at any point**, because nothing it governs is in a file: the line
between two steps is composed in the conversation, the disclosure line is spoken, and the build
record is written into the adopter's own folder. A green `validate` run says nothing about §10.4
either way, and would look identical if the rules were never followed. Until a checker exists, the
only thing enforcing §10 is the AI assistant reading it and a person noticing afterwards.

---

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
*Empty in the base template. Filled per client engagement.*
- <named optional extension>
```
