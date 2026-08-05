---
template: ts6-unqualified-field
title: Broken Fixture — TS6 (unqualified field name)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
standards_layer:
  - naming-conventions
new_tables:
  - tblWidget
---

# Broken Fixture — TS6 (unqualified field name)

## Intent

Deliberately broken test fixture: `Status` is one bare noun, and `TypeID` is a bare
noun wearing an `ID` suffix that does not qualify it. `WidgetID` and `WidgetStatus`
are correct and must not be reported.

## Entities

### tblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key — passes: `Widget` names the table |
| `WidgetStatus` | Text(30) | Required | Passes: qualified with the entity |
| `Status` | Text(30) | Required | Fails: one unqualified word |
| `TypeID` | Long | Nullable | Fails: `Type` is unqualified, and `ID` does not qualify it |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
