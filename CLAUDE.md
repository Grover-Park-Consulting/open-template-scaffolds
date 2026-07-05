# Open Template Scaffolds — Working Instructions for AI Assistants

You are working inside **Open Template Scaffolds**: a library of standards-based, AI-readable templates
for building Microsoft Access and SQL Server artifacts. A template is **read context** — you read it
together with the active standards layer and produce a **reviewable design** (a diagram plus field
detail). The developer approves or redirects that design. **Once it's approved, the build happens one
of two ways:** the developer implements it themselves, or — on their direction — **you carry out the
build**, creating the actual tables, relationships, indexes, lookups, and seed rows, with the
standards applied throughout.

## Your role

Produce the design from a template plus the standards layer when a template fits — and from scratch
under the same standards layer when none does. **The library's promise is the same either way: a
template when one fits; a standards-grounded design when none does — the same standards, the same
review flow.** Don't design from a blank page when a template exists. The developer **approves or
redirects**; they resort to building by hand only if your output isn't acceptable after iteration.
You build only what's been approved, and only when directed.

## The core workflow — designing a table schema

When the developer asks you to build or extend a set of tables, follow these steps **whether or not
they paste a prompt**. (`prompts/BuildNewTables-StartHere.md` is the canonical copy-paste form of
this same workflow.)

1. **Match a template.** Find the template in `templates/<domain>/` that fits the request and read it.
   **If no close match exists, say so and follow "When no template fits" below** — don't quietly bend
   a template that doesn't fit, and don't improvise unbounded. The from-scratch path is a first-class
   route with its own rules, not an exception.
2. **Read the active standards layer** in `standards/`.
3. **Apply the standards to everything you produce** — naming conventions, audit columns, and the
   error-handling pattern — plus the field-qualification rules (no bare reserved or ambiguous nouns;
   PK = `[Entity]ID`; a FK takes the referenced PK's name). These come from `standards/`, **never**
   from the template body.
4. **Honor the template.** Its entities, fields, relationships, and Business Rules are decisions
   already made. Carry them through unless the developer overrides a specific point.
5. **Surface every `house_assumptions` entry** the template declares in its front-matter. List them
   and ask the developer to confirm or override before you finalize.
6. **Fold in the developer's specifics** and any named extras from the template's
   `## Extra Options` section.
7. **Don't invent domain content** beyond the template and what the developer supplied. If something
   is genuinely undetermined, ask.
8. **Present two parts for review:** a `mermaid` `erDiagram` (tables, keys, cardinality, and the
   connections into any existing host tables), then field-table detail (`| Field | Type | Key / Req | Purpose & rules |`)
   with indexes, derived values, and the standards-supplied audit columns. It is never final until the
   developer says so.

## After approval — building it

The design is the first deliverable, not the last. Once the developer approves it, the build proceeds
one of two ways, at their direction:

- **They implement it themselves**, using your design as the specification.
- **They direct you to build it.** First **ask which platform the tables are for**, then generate the
  matching artifact (keys, relationships, indexes, lookup tables, and **seed rows** throughout):
  - **Access (ACE) local tables** → a **VBA `Sub` using DAO** (`CreateTableDef` / `CreateField` /
    indexes / relationships) — **never** `CurrentDb.Execute "CREATE TABLE…"` DDL. Carry each field's
    **comment as its `Description`** and AutoNumber as a `dbLong` field with `dbAutoIncrField`. Three
    rules make it actually run: set each `Description` **after** the table is appended (a second pass —
    otherwise runtime error 3219); set any field default on **`fld.DefaultValue`** (e.g. `"Now()"`)
    **before** append, never as a DDL `DEFAULT` clause (the DAO/ANSI-89 engine rejects it — the cause of
    "Syntax error in CREATE TABLE statement"); and tell the developer to **run the Sub from a Trusted
    Location** (outside one, Access silently disables the code and nothing is created). See
    `templates/_materialization.md` for the proven pattern.
  - **SQL Server** → `CREATE TABLE` DDL.
  - The error-handling block in any generated VBA comes from the standards layer — a **dependency-free
    default** (a message box) unless the house `error-handling.md` specifies a central logger.

  Apply the standards throughout, exactly as in the approved design.

**Never build before the design is approved, and never create or alter objects in a database unless
the developer directs you to.**

## Matching templates — use judgment

The developer describes what they need in their own words; deciding which template fits is your call.
Weigh the **domain** and the **shape** of the request against the templates available — a template
fits when it covers the same kind of work, even if the names differ (the standards layer and the
developer's specifics adjust those). When two templates could serve, or none clearly does, **don't
force a fit**: say what you found, and let the developer choose or confirm before you proceed.

## When no template fits

A missing template does not end the workflow — the developer came with a real need, and the library
serves it either way. When no template fits:

1. **Name the templates you considered and why each falls short.** This is required — it shows your
   matching judgment instead of hiding it.
2. **Offer the standards-grounded from-scratch path as the default, and ask for the go-ahead:**

   > "No template covers this, so, with your approval, I'll design it from scratch following your
   > standards layer — same review: you'll get the diagram and field detail to approve or redirect
   > before anything is built. Say the word and I'll begin."

   In the same breath, name the alternatives without stopping for a menu: adapting the nearest
   template despite the stated mismatch, or refining the description.
3. **On the go-ahead, design under the full standards layer.** Read every file in `standards/` (or
   call the MCP `get_standards` tool if available) and apply naming, audit columns, field
   qualification, the junction-PK convention, and third-normal-form discipline exactly as you would
   for a template-based design. Same two-part deliverable (diagram + field detail), same approval
   gate.
4. **Surface your invented assumptions.** A from-scratch design has no `house_assumptions`
   front-matter, so recreate that transparency yourself: list every modeling assumption you had to
   invent as **proposed assumptions** for the developer to confirm or override before the design is
   final.
5. **Close with the templatize offer.** Once the design is approved, offer — with the design already
   in hand: *"If you agree, I can shape this into a template for the library — it would be added
   upon the curator's approval."* (See `CONTRIBUTING.md` for how contributions work.)

**Never build before the design is approved, and never create or alter objects in a database unless
the developer directs you to** — from-scratch work included.

## Where things live

| Path | What it is |
|---|---|
| `templates/_template-schema.md` | The canonical format every template follows |
| `templates/<domain>/` | The templates, grouped by domain (e.g. `northwind/`, `library/`) |
| `standards/` | The active standards layer — naming, audit columns, error handling |
| `prompts/BuildNewTables-StartHere.md` | The copy-paste form of the workflow above |
| `examples/northwind-stocktake/` | A complete worked example (filled prompt + generated output) |

**Load only what the task needs** — the relevant template plus `standards/`. Don't read the whole
library into context every session.

## Standards always apply

The standards layer is authoritative for naming, audit columns, and error handling. An adopter who
forks this library **replaces `standards/` with their own** — so use whatever is in `standards/` now,
and never assume the GPC defaults. Never bake standards into a template body, and never skip applying
them. This is what lets one template serve every shop.

## Boundaries

- All three template types are authoritative in `_template-schema.md` — `table-schema` (§4),
  `vba-scaffold` (§8), `form-spec` (§9) — each with proven templates in the library. Generate
  against those sections exactly as written.
- Audit columns belong to `standards/`, never to a template's field list — flag them if you find them
  in a template body.
- Don't carry one practice's house conventions into output generated for another.
