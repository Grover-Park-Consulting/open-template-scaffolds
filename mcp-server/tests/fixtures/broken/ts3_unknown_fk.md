---
template: ts3-unknown-fk
title: Broken Fixture — TS3 (FK target resolves to no known table)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — TS3 (FK target resolves to no known table)

## Intent

Deliberately broken test fixture: a field declares `FK → NoSuchTable`, which is neither
declared in `new_tables` nor listed in `requires_tables`.

## Entities

### TblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key |
| `ParentID` | Long | FK → NoSuchTable, Required | Dangling reference |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
