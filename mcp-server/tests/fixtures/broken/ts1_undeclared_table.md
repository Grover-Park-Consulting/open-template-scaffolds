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

Deliberately broken test fixture: `### TblGadgetPart` is documented under Entities
but absent from `new_tables`. Its key is a compound name so that this fixture
violates one rule only — a single-word key on an undeclared table would also be
reported as unqualified, which is a different rule's business.

## Entities

### TblWidget

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `WidgetID` | AutoNumber | PK | Surrogate key |

### TblGadgetPart

| Field | Type | Key / Req | Purpose & rules |
|---|---|---|---|
| `GadgetPartID` | AutoNumber | PK | Surrogate key |

## Relationships

- None.

## Business Rules

1. None — fixture only.

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
