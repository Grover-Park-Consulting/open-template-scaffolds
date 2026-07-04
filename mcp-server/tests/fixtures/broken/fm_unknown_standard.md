---
template: fm-unknown-standard
title: Broken Fixture — FM (unrecognized standards_layer value)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
standards_layer:
  - no-such-standard
new_tables:
  - TblWidget
---

# Broken Fixture — FM (unrecognized standards_layer value)

## Intent

Deliberately broken test fixture: `standards_layer` names a value outside the
registered set.

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

- **Unknown** — fixture only.

## Extra Options

*Empty — fixture stub.*
