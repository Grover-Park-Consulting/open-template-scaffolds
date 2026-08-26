# Standards Layer — OTS Defaults

**Who reads this:** the AI assistant, applying these rules to what it generates or a shop deciding
what to replace with their own.

If you are building something from a template, you don't need to read this file unless you are curious.
However, we have included clarifying comments to help you interpret what it says, just in case.

**Building from a template?** These seven files decide the part of how your database is
built that the template itself does not settle. That includes what things are called, what extra columns every
table carries, what happens when something goes wrong, and how the code is laid out. The table below
says which file covers which, one line each.

You do not have to do anything with this information to use the templates. The rules apply to the code you receive
whether or not you read these files.

Using the standards as they are is the normal choice. If you came here because a template asked whether
you want to use these rules as they are or make them your own, you don't answer the question here.

*After you've read about the choice here, go back and answer it where it was asked.*

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

You can replace any or all of the standards files with your own rules. If you put your own rules in place, the next
design built from a template follows them, *with no edit to the template itself.* Tables and code you
have already built are not changed by this; the new rules apply to what you build from then on.

## Standards are required for OTS templates to run

While you can use your own standards, the OTS template library depends on these seven standards being
available. Moreover, your standards files must be internally consistent with the way OTS standards
files are written. If you edit or replace one or more of these standards files, be sure to preserve their shape, i.e. the same sections with similar rules throughout.

When you run any of the templates in the library, you'll be asked how to use the standards.
