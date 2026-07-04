---
template: ts5-audit-column
title: Broken Fixture — TS5 (audit column in a template field table)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — TS5 (audit column in a template field table)

## Intent

Deliberately broken test fixture: `AddedBy` appears in the field table. Audit columns
belong to the standards layer, never to a template body.

## Entities

### TblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key |
| `AddedBy` | Text(50) | Required | Belongs to the standards layer, not here |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
