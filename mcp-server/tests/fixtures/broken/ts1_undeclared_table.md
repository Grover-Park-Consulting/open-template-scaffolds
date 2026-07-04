---
template: ts1-undeclared-table
title: Broken Fixture — TS1 (documented entity not declared in new_tables)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — TS1 (documented entity not declared in new_tables)

## Intent

Deliberately broken test fixture: `### TblGadget` is documented under Entities
but absent from `new_tables`.

## Entities

### TblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key |

### TblGadget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `GadgetID` | AutoNumber | PK | Surrogate key |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
