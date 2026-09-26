---
template: _coverage
title: Template Family Coverage Matrix
domain: _meta
type: spec
version: 0.1.0
status: draft
---

# Template Family Coverage Matrix

**Who reads this:** the AI assistant, when working out what the library already carries for a
developer's request; and George, as a factual inventory of the seven domains.

Which of the four template types each domain has today. A filled cell names the template of that
type; a dash is a type the domain does not have. A template's type is read from its front-matter
`type:`, not from its filename.

| Domain | table-schema | outcome-first | vba-scaffold | form-spec |
|---|---|---|---|---|
| `app-startup` | — | `app-startup-outcome-first` | `app-startup-scaffold` | — |
| `asset-tracking` | `capital-asset-tracking-schema` | `capital-asset-tracking-outcome-first` | — | — |
| `audit` | `audit-logging-lite-schema` | `audit-logging-lite-outcome-first` | `audit-logging-lite-scaffold` | — |
| `errors` | `error-logging-schema` | `error-logging-outcome-first` | `error-logging-scaffold` | — |
| `library` | `catalog-schema` | `catalog-outcome-first` | `record-finder-scaffold` | `publication-form` |
| `scheduling-assignment` | `officiating-assignment-schema` | `officiating-assignment-outcome-first` | `officiating-assignment-scaffold` | `officiating-assignment-form` |
| `stocktakescan` | `stocktake-schema` | `stocktake-scan-outcome-first` | `stocktake-scan-scaffold` | — |
