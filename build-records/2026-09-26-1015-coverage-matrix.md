# Build Record: Template Family Coverage Matrix

**Who reads this:** George (and any future AI assistant working in this library). The third
implementation task — a new `templates/COVERAGE.md`, per the brief at
`build-records/2026-09-26-task-brief-family-coverage-matrix.md`.

**Author:** the AI assistant (deepseek-v4-pro), via OpenCode.
**Date/time:** 2026-09-26 10:15.
**Task:** a factual inventory of which of the four template types each of the 7 domains has.

**Frozen files (`FREEZE.md`, the 18): no.** I touched none of them in this task. The one frozen file
in the working tree's diff, `templates/_template-schema.md`, is from my earlier tasks and was
covered by the prior review — it is not from this one.

---

## What I created

One new file, `templates/COVERAGE.md`. A filled cell names the template of that type; a dash is a
type the domain does not have. Types are read from each template's front-matter `type:`, not its
filename.

| Domain | table-schema | outcome-first | vba-scaffold | form-spec |
|---|---|---|---|---|
| `app-startup` | — | `app-startup-outcome-first` | `app-startup-scaffold` | — |
| `asset-tracking` | `capital-asset-tracking-schema` | `capital-asset-tracking-outcome-first` | — | — |
| `audit` | `audit-logging-lite-schema` | `audit-logging-lite-outcome-first` | `audit-logging-lite-scaffold` | — |
| `errors` | `error-logging-schema` | `error-logging-outcome-first` | `error-logging-scaffold` | — |
| `library` | `catalog-schema` | `catalog-outcome-first` | `record-finder-scaffold` | `publication-form` |
| `scheduling-assignment` | `officiating-assignment-schema` | `officiating-assignment-outcome-first` | `officiating-assignment-scaffold` | `officiating-assignment-form` |
| `stocktakescan` | `stocktake-schema` | `stocktake-scan-outcome-first` | `stocktake-scan-scaffold` | — |

## Two things about the file you should know (neither was in the brief)

**1. It carries a `type: spec` front-matter block.** Without it, the validator would pick
`templates/COVERAGE.md` up as a 22nd template and fail it: the `iter_templates` scan skips only
`_`-prefixed files and `README.md`, and `validate_library` skips `type: spec`. The spec type is what
makes the brief's "still 21/21" true, and it matches the library's existing infrastructure files
(`_template-schema.md` is the same shape). I made the call to satisfy the stated verification; it is
flagged here rather than left silent.

**2. It still appears in `list_templates`.** Confirmed by calling the tool: `list_templates` filters
on the `_` prefix, not on `type: spec`, so `_coverage` shows up there as an entry with
`domain: _meta`, `type: spec`. Renaming the file to `templates/_COVERAGE.md` would remove it from
every tool — the underscore is the library's infrastructure convention, used by all three existing
infra files — but that departs from the filename in the brief, so I left it as `COVERAGE.md` for you
to decide.

## A factual correction to the brief's premise

The brief (and my own earlier review) says only `scheduling-assignment` ships a full set. The matrix
shows **`library` also has all four types**: `catalog-schema` (table-schema), `catalog-outcome-first`
(outcome-first), `record-finder-scaffold` (vba-scaffold), `publication-form` (form-spec). If
`record-finder-scaffold` is being read as a generic mechanism rather than a library-catalog artifact,
that would give a count of one; by front-matter type and domain, the count is two.

## Linking (the brief asked me to say, not act)

I think the matrix belongs where an assistant will see it. `CLAUDE.md`'s "Where things live" table
is the natural home and is not frozen. `README.md` is frozen, so a link there would need a lift. I
added no link; your call.

## Verification

- `python mcp-server/run_validate.py` → **21/21 well-formed**.
- `python -m unittest discover -s mcp-server/tests` → **31 OK, 1 skipped**.
- `git status`: the only file this task added is `templates/COVERAGE.md`. The working tree also
  carries my two earlier uncommitted tasks (naming-drift, front-matter-drift); no existing file was
  modified by this task.
