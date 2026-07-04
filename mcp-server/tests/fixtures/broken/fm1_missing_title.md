---
template: fm1-missing-title
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — FM1 (missing required front-matter key `title`)

## Intent

Deliberately broken test fixture: the required `title` front-matter key is absent.
Every other rule should pass, so validate_template reports FM1 and only FM1-family errors.

## Entities

### TblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key |
| `WidgetName` | Text(50) | Required | Display name |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
