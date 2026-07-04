---
template: fm6-bad-house-assumption
title: Broken Fixture — FM6 (house_assumptions entry without rationale)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
house_assumptions:
  - TblWidget is modeled this way
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — FM6 (house_assumptions entry without rationale)

## Intent

Deliberately broken test fixture: the `house_assumptions` entry lacks the
"Target — rationale" form (no separator, no rationale).

## Entities

### TblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
