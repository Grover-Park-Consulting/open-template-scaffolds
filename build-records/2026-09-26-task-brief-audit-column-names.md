# Task Brief: Stop Hardcoding a Source Database's Audit-Column Names

**Who reads this:** the AI assistant picking up this task (currently DeepSeek, via OpenCode).
Re-confirm you've read `FREEZE.md`. This task does not require touching any of the 18 frozen
files — `templates/_template-schema.md` has already been corrected by Claude Code as part of the
same finding this task fixes; you don't need to touch it again.

## The problem

George caught this from memory, and it checked out: two templates hardcode one specific source
database's audit-column names (`AddedBy`/`AddedOn`/`ModifiedBy`/`ModifiedOn`) as if that were the
house standard, instead of presenting the actual house standard and instructing the developer to
check their own target database first. The house standard, per `standards/audit-columns.md`, is
`CreatedDate` / `CreatedBy` / `ModifiedDate` / `ModifiedBy` (/ `AccessTS` on SQL Server).

**Why this is wrong, not just inconsistent:** a template that bakes in one source's column names
teaches every future build to expect that one source's convention, even for a developer whose
database has never seen Northwind or the asset-tracking contributor's schema. `standards/audit-columns.md`
already says the right thing in its own Notes section — check the target database's existing
convention before naming new audit columns, and fall back to the house names only when no existing
convention is present. The two templates below don't follow their own standards layer's rule.

**The model to imitate:** `templates/audit/audit-logging-lite-scaffold.md`, around line 948, gets
this right. It names `AddedOn`/`AddedBy`/`ModifiedOn`/`ModifiedBy` explicitly as *"one form you will
meet, and there are others"* and instructs finding out how the host fills those columns before
proceeding. Match that shape, not the wording — the fix isn't to copy that file's sentences, it's to
apply the same principle: present the house default, and instruct checking the actual target
database rather than asserting one fixed set of names.

## The two files to fix

### `templates/stocktakescan/stocktake-schema.md`

- **Line ~368** (checklist item 6): currently asserts `AddedBy`/`AddedOn`/`ModifiedBy`/`ModifiedOn`
  "per the Northwind data-macro pattern" as what the check validates. Rewrite so the check validates
  against whatever the active standards layer specifies — don't name a fixed set of columns in a
  validating-build checklist item at all; that's exactly what the standards layer exists to supply.
- **"Standards Layer" section, "Audit columns" bullet** (~line 383): currently states the specific
  Northwind names as the contributed convention. Rewrite to state the house default
  (`CreatedDate`/`CreatedBy`/`ModifiedDate`/`ModifiedBy`) and note, as an example, that a
  Northwind-derived host may already use a different convention — check the target database first,
  per `standards/audit-columns.md`'s own Notes section — rather than asserting the Northwind names
  as this template's contributed standard.

### `templates/asset-tracking/capital-asset-tracking-schema.md`

Same defect, same fix, different source database:

- **Checklist item 6** (~line 310): same rewrite — validate against the active standards layer, not
  a named fixed set.
- **"Standards Layer" section, "Audit columns" bullet** (~line 321): same rewrite — state the house
  default, note that the contributing host's own database may have used something else, instruct
  checking the actual target database.

## What NOT to change

- `templates/audit/audit-logging-lite-scaffold.md` — already correct, it's your reference, not a
  file to edit.
- The naming-conventions bullets in either file (the `tbl`/`tlkp` vs. no-prefix discussion) — that's
  a separate, already-correct pattern (a template stating which house style *it* follows is fine;
  the audit-column problem was asserting a fixed name set as if it were universal, which is a
  different kind of claim). Leave naming-conventions content alone unless you find the identical
  audit-column-style defect in it, in which case name it in your build record rather than fixing it
  silently — that would be outside this task's stated scope.
- Anything in `examples/northwind-stocktake/` — that folder is a worked example of a real Northwind
  build, where Northwind-specific names are correct, not a defect. It's also frozen. Do not touch it.

## Verification

- `python mcp-server/run_validate.py` → confirm still 21/21.
- Re-read both edited sections after writing them and confirm neither still names a fixed audit-
  column set as the expected/validated one — the whole point is that the template no longer asserts
  one specific database's names as universal.

## When you're done

Build record as before, in `build-records/`. State plainly whether you touched any of the 18 frozen
files (should be no) and quote the before/after of each of the four edited spots so the fix is easy
to check without opening four diffs.
