---
step: 02-code
title: "Write and import the code"
platform_facts: [open-existing-startup, mcp-module-import, mcp-line-numbers, app-startup-autoexec]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Code

Open the database the developer supplied, as the facts below say. Write and import the procedures the
build needs, divided as the design-principles standard says, within *Free to choose alternatives*.
Change nothing before the developer has given permission.

**An existing `AutoExec` stays as it is** (method `backup-question`). Read it in your copy. Where it
does open-time work, write that work into `Startup()` as its first lines, commented out, under a
comment naming them a proposed translation of what the developer's `AutoExec` does today. Say in that
comment that they either switch the lines on and make `RunCode Startup()` the `AutoExec`'s only
action, or leave them off and add `RunCode Startup()` above the `AutoExec`'s own actions, and that
doing both runs that work twice. Never switch the lines on yourself.

