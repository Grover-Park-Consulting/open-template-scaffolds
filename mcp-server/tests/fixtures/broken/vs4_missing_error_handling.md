---
template: vs4-missing-error-handling
title: Broken Fixture — VS4 (vba-scaffold without error-handling standard)
domain: fixtures
type: vba-scaffold
version: 0.0.1
status: draft
target_module: modFixture
new_procedures:
  - DoFixtureThing
standards_layer:
  - naming-conventions
---

# Broken Fixture — VS4 (vba-scaffold without error-handling standard)

## Intent

Deliberately broken test fixture: a `vba-scaffold` whose `standards_layer` omits
`error-handling`, which the type requires.

## Procedures

### DoFixtureThing

```vba
Public Sub DoFixtureThing()
    ' [SCAFFOLD] fixture body — never executed
End Sub
```

## Standards Layer

- **Naming conventions** — per `naming-conventions.md`.

## Extra Options

*Empty — fixture stub.*
