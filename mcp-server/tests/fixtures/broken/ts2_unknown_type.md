---
template: ts2-unknown-type
title: Broken Fixture — TS2 (field with unknown data type)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — TS2 (field with unknown data type)

## Intent

Deliberately broken test fixture: `WidgetName` is typed `Varchar(50)`, which is not
an Access scalar type.

## Entities

### TblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key |
| `WidgetName` | Varchar(50) | Required | Not an Access type |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
