# Consistency & Conformity Review — Open Template Scaffolds

**Who reads this:** George (and any future AI assistant). This is a cross-cutting review of the
whole library, not a build record for one template — hence it sits at the `build-records/` root
rather than in a single template's folder.

**Author:** the AI assistant (deepseek-v4-pro) running in this workspace.
**Date/time:** 2026-09-26 06:11
**Scope:** all 21 domain template files across 7 domains, the `standards/` layer, the template
schema, and the prompts. `.claude/pilot/` is excluded as instructed (planning documents replaced by
actual components).

---

## Conforms (what is solid)

- **Validator is green and meaningful** — `python mcp-server/run_validate.py` reports 21/21 templates
  well-formed. The tool genuinely enforces structure: entity/procedure/table set matching, FK
  resolution, field-type vocabulary, required sections, cascade statements, and
  `new_tables` ↔ `## Entities` identity.
- **Standards layer is coherent and self-consistent** — all 7 standards files share one header shape;
  each names its reader, defines its `standards_layer` value, and states the fork-and-replace
  boundary identically.
- **The four template types are uniformly authoritative** — `_template-schema.md` §4/§8/§9/§12 are
  complete, each with a proven shipped template and per-type `validate` rules.
- **Cross-references resolve** — `implements`, `related`, `requires_tables`, and `new_fields` targets
  all point at real slugs/tables/fields.
- **Core conventions hold everywhere** — every table-schema carries field tables in the exact
  `| Field | Type | Key / Req | Purpose & rules |` format; audit columns never leak into template
  bodies; naming rules are consistently deferred to the standards layer, not embedded.

## Inconsistent / non-conforming

### Naming drift (the prior "second review" #1 finding is still open)

- **Slug ↔ filename ↔ folder misalignment in 4 of 7 domains.** `asset-tracking` folder holds slugs
  `school-district-asset-tracking-*` in files `asset-tracking-*.md`; `scheduling-assignment` holds
  `sports-officiating-assignment-*` in `officiating-assignment-*.md`; `stocktakescan` holds
  `northwind-stocktake-*` in `stocktake-*.md`; `library` drops the prefix in some filenames
  (`catalog-schema.md` → slug `library-catalog-schema`). The `domain` key (folder name) and the
  slug's domain word do not match.
- **Title convention broken in two schema files** — `error-logging-schema` is titled "Error Log
  Table" (no "– Table Schema" suffix) and `asset-tracking-schema` is titled "School District Capital
  Asset Tracking"; the other four schemas use the "– Table Schema" suffix.
- **One scaffold title adds a scope word** — `officiating-assignment-scaffold` says "Assignment & Pay
  VBA Scaffold" while its siblings say only "Assignment".

### Front-matter format drift

- **`build_paths:` is an undocumented key.** It appears on `audit-logging-lite-schema` and
  `audit-logging-lite-scaffold` but is absent from the §2 front-matter key table. The Path A/Path B
  mechanism is proven and shipped but not folded back into the schema.
- **`implements` is under-documented in §2** — the key table lists it for `vba-scaffold`/`form-spec`
  only, though §12.1 also uses it for `outcome-first`.
- **`house_assumptions` entries are unevenly formed.** §2 requires `Target — rationale`; several
  entries lack a leading Target (e.g. `app-startup-outcome-first`'s first entry begins "The locations
  of any folders…"), and separators vary between ` — ` (em dash) and ` - ` (hyphen).
- **YAML quoting style is inconsistent** — some `house_assumptions`/`warnings`/`related` entries are
  quoted strings, others unquoted multiline blocks.
- **`standards_layer` is written two ways** — inline `[...]` vs block `-` lists — and the set of
  standards declared per type varies without an evident rule (`catalog-schema` declares only
  `audit-columns, naming-conventions`; `asset-tracking-schema` adds `error-handling, query-style`;
  `error-logging-schema` includes `design-principles` but omits `error-handling`).

### The standards gate is still pilot-scoped

- `_standards-gate.md` opens "**Pilot scope.** Only the two audit-logging templates run this gate
  today," and `audit-logging-lite-scaffold.md` still calls itself "the gate's pilot." The library has
  grown to 21 templates across 7 domains, but the gate's reach has not been generalized or
  re-declared. The prior "standards gate needs a definitive scope" finding stands unchanged.

### Template families are uneven (prior #2 finding, still open)

- Only `scheduling-assignment` ships the full set (schema + outcome-first + scaffold + form).
  `library` has schema + outcome-first + form + a finder-scaffold but no catalog scaffold;
  `app-startup` has scaffold + outcome-first but no schema; `stocktakescan`, `audit`, and `errors`
  have schema + outcome-first + scaffold but no form; `asset-tracking` has schema + outcome-first
  only. No matrix labels partial families.

### Status metadata is uninformative

- All 21 templates carry `status: draft` despite several being heavily trialled live
  (`audit-logging-lite-scaffold` at v0.18.0, `stocktake-scan-scaffold` at v0.10.0). None have advanced
  to `review`/`stable`, so maturity is invisible in metadata.

### Prompt coverage is uneven and inconsistently named

- 16 prompts, but no scaffold prompt for `asset-tracking` (only outcome-first), and prompt names use
  different casing/words than slugs (`StocktakeScan` vs `northwind-stocktake`, `OfficiatingAssignment`
  vs `sports-officiating`).

## Bottom line

- **Strong:** architecture, standards layer, validator, and per-type conformity.
- **Weak:** the exact items the prior "second review" flagged — naming normalization, a coverage
  matrix, uneven family completeness, an un-scoped standards gate, and unlabeled legacy vs current —
  are all still present. No structural redesign is needed; the remaining work is governance and
  naming consistency, plus folding the proven `build_paths`/Path-A-B mechanism back into the schema.
