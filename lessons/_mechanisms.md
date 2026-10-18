# Mechanisms

**Who reads this:** the AI assistant, which tags and looks up lessons by these names; the developer, who approves new rows.

A **mechanism** is the part of Access or the build process a lesson is about. Lessons and templates use the canonical name only. A lookup matches aliases and never shows them as the name. **`always: yes`** means the lessons for that mechanism are delivered with every template.

| Canonical name | Description | Aliases | always |
|---|---|---|---|
| field-properties | Settings on a column that decide what the table accepts: Required, AllowZeroLength, defaults, validation rules. | column settings | no |
| schema-migration | Changing a table that already holds data: renames, widening, clearing values. | table migration | no |
| derived-values | Values computed from other columns and stored. | calculated values, rollups | no |
| error-handling | The handler pattern inside procedures, including line numbers and what a handler does with an error. | error handler, error logging | yes |
| vba-generation | Writing, importing, and running VBA from outside the Access window. | COM automation, import route | yes |
| build-verification | Running checks against a built artifact and trusting what they report. | checks, testing a build | yes |
| split-database | Front end and back end in separate files, links between them. | linked tables, back end | no |
| host-adoption | Fitting a template onto a database that already exists. | existing database | no |
| ace-sql | What the Access database engine's SQL accepts and refuses. | Access SQL, Jet SQL | no |
| form-building | Creating forms and controls by tool. | form creation | no |
| data-macro | Table-level macros written as generated XML. | data macros, AXL | no |

To add a row: the AI assistant checks this table first and reuses a row when one fits. When none fits, it proposes a row (name, description, why no existing row covers it) and the developer approves it with the lessons.
