---
template: _template-schema
title: Open Template Scaffolds — Canonical Template Format
domain: _meta
type: spec
version: 0.1.0
status: draft
---

# Open Template Scaffolds — Canonical Template Format

This is the **format specification** every template file in this library must follow.
It is the contract the reference MCP server keys off: discovery (`list_templates`,
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
> §9) are authoritative — the template format is complete**, each proven by hand before the MCP's
> schema-dependent tooling is built against it.

---

## 1. File rules

- **Format:** Markdown (`.md`), UTF-8. No binary, no executable code.
- **Location:** `templates/<domain>/<name>.md`. Files prefixed `_` (e.g. this one) are
  library infrastructure, not domain templates.
- **One template per file.** A template defines one cohesive artifact set for one domain.
- **Single source of truth:** the file *is* the template. CLAUDE.md and the MCP are readers
  of it, never separate copies.

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
| `version` | required | semver string | Template version, independent of library version |
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
| `## Supplementals` | Engagement-specific stub (see §7) |

Optional-for-all:

| Section | Purpose |
|---|---|
| `## Parked / future considerations` | Named directions explicitly **not** in the current design |

---

## 4. Body sections — `type: table-schema`

In addition to the common core, a `table-schema` template **must** contain, in this order:

| Section | Purpose | Required |
|---|---|---|
| `## Prerequisites` | The hooks into the existing schema — a table of existing objects the new tables wire into. Required when `extends` is set | conditional |
| `## Entities` | One `### <TableName>` per new table — grain statement, field table (§5), `Indexes:` line. **Trivial, uniform lookups** (`<name>ID` + a descriptor + optional `SortOrder`, nothing more) may instead be **grouped** in a single sub-table of name + seed rows — documentation shorthand only; each row is still its own discrete table (this is *not* a shared/MUCK lookup table). Any lookup carrying extra structure (more fields, an FK, a description) takes its own `### <name>` heading like an entity. | required |
| `## Relationships` | New relationships and hooks into the host schema, as a bulleted list naming parent → child, the join field(s), and cascade behavior | required |
| `## Business Rules` | Numbered list of the logic the generated objects must honor (grain constraints, rollups, derivations, deferred-logic notes) | required |

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

- **Field** — backtick-wrapped field name. GPC field-qualification rules apply (no bare
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
| `naming-conventions` | Table/field prefix policy (e.g. Northwind no-prefix vs GPC `tbl`/`tlkp`). The template states which house style it follows; a different practice regenerates the same entities under its own conventions without editing the template |
| `error-handling` | The house `errHandler` / global-error pattern for any VBA generated alongside |
| `query-style` | How VBA and saved queries write and run SQL — where SQL lives, aliasing/qualification, formatting, and safe criteria. Applies to any generated code that touches data (notably `vba-scaffold`) |
| `form-conventions` | Form **design** defaults (control prefixes, control types, buttons, tab order, sizing) + the named reusable form patterns (selector, quick-add, validation highlights; audit display optional). Used by `form-spec` |

A template **describes the boundary**; it does not embed the standards. Where a house-specific
*modeling* assumption cannot be cleanly separated, resolve it by the lowest tier that fits: drop
it from the published template (Private), park it in `## Supplementals` (Optional), or — if it is
load-bearing — keep it and declare it in the `house_assumptions:` front-matter list (Declared), so
it is machine-visible rather than buried in prose.

---

## 7. Supplementals

Every template ends with a `## Supplementals` section: a **stub in the base library**,
listing named, optional extensions a developer fills per client engagement. The filled-in
copy is saved to the developer's own library — never committed back here. Supplementals are
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

### 8.4 `validate` rules for `vba-scaffold`

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
alternative path builds the form live via the Access MCP `create_form`/`create_control` tools. The
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

## 10. Minimal skeleton (`type: table-schema`)

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

## Supplementals
*Empty in the base template. Filled per client engagement.*
- <named optional extension>
```
