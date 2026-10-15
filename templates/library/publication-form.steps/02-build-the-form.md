---
step: 02-build-the-form
title: "Build the form"
platform_facts: [form-live-build, form-hand-validation, mcp-module-import, mcp-run-procedure, mcp-line-numbers, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

**The tables are confirmed from step 1.** Build the forms and the record-source query from *Layout*
and *Features* in the master, as the facts below say. Change nothing before the developer has given
permission.

## Materialization

The form-spec materializes by building the form live through an Access MCP server's form-creation and
control-creation tools, with a default stacked layout and the code-behind wired in — event handlers
calling the framework helpers (and any paired `vba-scaffold` procedures). The markdown → Access-text
mapping is **defined and hand-validated** as part of this template type's proof; the actual generator
is built in the template library MCP server phase (B3). Where no Access MCP server is connected the
run ends at the approved design, and no importable form text is generated for the developer to import
as though it were a finished form — see `_materialization.md`.
