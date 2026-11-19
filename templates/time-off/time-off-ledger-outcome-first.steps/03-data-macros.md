---
step: 03-data-macros
title: Build the Data Macros the chosen routes need
platform_facts: [data-macro-rules, vba-import-xml-entities]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Data Macros

**The tables exist from step 2: do not rebuild them here.** Build the Data Macros the routes the
developer chose for Business Rule 1 and the posting checks call for. Read any Data Macro a table
already carries before writing to it, and merge. Confirm they exist before you fetch step 4. A Data
Macro needs only the name of a function it calls, not its code, so write the call now and leave the
function to step 4.

**From the master's instructions to the AI assistant, for this step.** They bind as the master's do.

- **Read the existing Data Macro on `tblTimeOffEntry` before writing anything to it** (fact
  `data-macro-rules`). The audit stamping lives in that document. Merge the new logic into it, and
  where the two do not obviously share a branch, ask before writing.
