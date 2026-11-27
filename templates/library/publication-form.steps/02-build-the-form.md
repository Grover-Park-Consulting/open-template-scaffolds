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

On the build route, build the forms live through an Access MCP server's form-creation and
control-creation tools: a default stacked layout, code-behind wired to the framework helpers and the
companion `record-finder-scaffold`. On the handover route, write them as importable form text by the
mapping rules, marked as proposed code (method `handover-code`). The markdown is the source of truth.
