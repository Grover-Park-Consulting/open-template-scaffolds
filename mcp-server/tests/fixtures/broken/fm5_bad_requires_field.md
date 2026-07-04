---
template: fm5-bad-requires-field
title: Broken Fixture — FM5 (requires_fields entry without Table.Field form)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
requires_fields:
  - ProductsProductID
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — FM5 (requires_fields entry without Table.Field form)

## Intent

Deliberately broken test fixture: a `requires_fields` entry has no dot separator.

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
