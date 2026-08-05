# Standards Layer — OTS Defaults

**Who reads this:** the AI assistant, applying these rules to what it generates — and a shop deciding what to replace with their own.

**Using a template?** These rules reach you in the code you receive. You do not need to read this file in order to use the template.

This folder holds the **OTS default standards layer**: the house rules a template defers to, so the
same template produces house-conforming output for any practice. A template names what it defers in
its `standards_layer` front-matter; when the generator builds artifacts, it reads the matching
file(s) here and applies them.

| File | `standards_layer` value | Covers |
|---|---|---|
| [`naming-conventions.md`](naming-conventions.md) | `naming-conventions` | Object, table, and field/column naming + the qualified-field rule (Access + SQL Server) |
| [`audit-columns.md`](audit-columns.md) | `audit-columns` | The `CreatedDate` / `CreatedBy` / `ModifiedDate` / `ModifiedBy` / `AccessTS` set and how it's maintained |
| [`error-handling.md`](error-handling.md) | `error-handling` | The VBA `errHandler` / central-logging pattern for any code generated alongside |
| [`query-style.md`](query-style.md) | `query-style` | How VBA and saved queries write and run SQL — where SQL lives, aliasing/qualification, formatting, safe criteria |
| [`design-principles.md`](design-principles.md) | `design-principles` | The reasoning behind the specific rules — one-job-per-procedure, separation of concerns, encapsulation, cohesion/coupling, DRY, strong contracts; what generated VBA is shaped by |
| [`form-conventions.md`](form-conventions.md) | `form-conventions` | Form design defaults (control prefixes, control types, buttons, tab order, sizing) + named form patterns (selector, quick-add, validation highlights; audit display optional) |
| [`startup-conventions.md`](startup-conventions.md) | `startup-conventions` | How a generated Access application initializes on open — the `AutoExec` → `Startup()` convention, the idempotent `EnsureAppFolders()` slot, and reliable external-file-asset folders |

## Using your own standards

You can replace the standards files with your own rules. Put your own rules in place, and the next
design built from a template follows them, with no edit to the template itself. Tables and code you
have already built are not changed by this — the new rules apply to what you build from then on.

However, the templates do depend on these 7 standards being available and consistent with the way
they are written. If you edit or replace one or more of these standards files, be sure to preserve
their shape.

When you run any of the templates in the library, you'll be asked how to use the standards.
