---
template: fm7-unknown-fact
title: Broken Fixture — FM7 (platform_facts id with no section)
domain: fixtures
type: table-schema
version: 0.0.1
status: draft
platform_facts: [no-such-fact]
standards_layer:
  - naming-conventions
new_tables:
  - TblWidget
---

# Broken Fixture — FM7 (platform_facts id with no section)

## Intent

Deliberately broken test fixture: `platform_facts` names an id that no
`<!-- fact: id -->` marker in `_materialization.md` carries.

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
