# Standards Layer — GPC Defaults

This folder holds the **GPC default standards layer**: the house rules a template defers to, so the
same template produces house-conforming output for any practice. A template names what it defers in
its `standards_layer` front-matter; when the generator builds artifacts, it reads the matching
file(s) here and applies them.

| File | `standards_layer` value | Covers |
|---|---|---|
| [`naming-conventions.md`](naming-conventions.md) | `naming-conventions` | Object, table, and field/column naming + the qualified-field rule (Access + SQL Server) |
| [`audit-columns.md`](audit-columns.md) | `audit-columns` | The `CreatedDate` / `CreatedBy` / `ModifiedDate` / `ModifiedBy` / `AccessTS` set and how it's maintained |
| [`error-handling.md`](error-handling.md) | `error-handling` | The VBA `errHandler` / central-logging pattern for any code generated alongside |
| [`query-style.md`](query-style.md) | `query-style` | How VBA and saved queries write and run SQL — where SQL lives, aliasing/qualification, formatting, safe criteria |
| [`form-conventions.md`](form-conventions.md) | `form-conventions` | Form design defaults (control prefixes, control types, buttons, tab order, sizing) + named form patterns (selector, quick-add, validation highlights; audit display optional) |

## Forking

Replace these files with your own house rules. **Nothing in a template body depends on the specific
names or patterns here** — that's the standards/template split. Swap in your standards, regenerate,
and the templates produce artifacts in your conventions without any edit to the templates themselves.
