---
template: fm3-extends-no-requires
title: Broken Fixture — FM3 (extends without requires_tables)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
extends: Northwind (Access Developer Edition)
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — FM3 (extends without requires_tables)

## Intent

Deliberately broken test fixture: `extends` names a host database but `requires_tables`
is empty. (Also triggers TS's Prerequisites-section rule; the test asserts FM3 is present.)

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
