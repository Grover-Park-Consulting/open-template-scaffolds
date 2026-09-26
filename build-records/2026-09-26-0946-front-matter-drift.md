# Build Record: Front-Matter Format Drift Cleanup

**Who reads this:** George (and any future AI assistant working in this library). A record of the
second implementation task assigned after the naming-drift pass, per the brief at
`build-records/2026-09-26-task-brief-frontmatter-drift.md`.

**Author:** the AI assistant (deepseek-v4-pro), via OpenCode.
**Date/time:** 2026-09-26 09:46.
**Task:** fix the "Front-matter format drift" section of
`build-records/2026-09-26-0611-consistency-conformity-review.md` — the three items named in the
brief, and only those three.

---

## What I changed

### Item 1 — `house_assumptions` "Target — rationale" shape

Ten entries across seven files lacked the leading `Target —`. I added a Target to each (prepended
the name, then an em dash). Targets chosen to be real names already in each template's own body —
this turned out to be a hard constraint, see "What the validator taught me" below.

| File | Target added |
|---|---|
| `app-startup-outcome-first` | `Shared-folder locations` |
| `capital-asset-tracking-outcome-first` (1st) | `Capitalization threshold` |
| `error-logging-outcome-first` | `Error log` |
| `catalog-outcome-first` | `Noise words` |
| `officiating-assignment-outcome-first` (1st) | `Official age` |
| `officiating-assignment-outcome-first` (2nd) | `Game play level` |
| `stocktake-scan-outcome-first` (1st) | `CountedQuantity` |
| `stocktake-scan-outcome-first` (2nd) | `RemediationStatus` |
| `stocktake-schema` (2nd) | `RemediationStatus` |
| `stocktake-schema` (3rd) | `ProductVarianceAllowance` |

Each change was the minimal rewrite: the leading noun phrase became the Target, the rest of the
sentence became the rationale. No rationale text was otherwise reworded.

### Item 2 — separator (em dash vs hyphen)

**No work.** Every `Target — rationale` separator across all templates was already an em dash. The
"hyphen" I reported in the original review was a misread — the terminal renders the em dash (`—`)
as `-`. Confirmed by a code-point check; there were zero hyphen separators to fix.

### Item 3 — YAML quoting

Standardized **per-key majority**, as George ruled:

- `house_assumptions` → **quoted** (converted 15 plain entries to double-quoted).
- `warnings` → **plain** (converted 6 quoted entries to unquoted).
- `related` → **quoted** (already uniform; no change).

One plain entry (`audit-logging-lite-scaffold`, "Identity source") carried embedded double quotes
around "Before you write the modules". Wrapping it in quotes would have broken the YAML, so I
converted those inner quotes to single quotes as part of the same pass.

## What the validator taught me (and why two Targets changed)

The `validate` tool's FM6 rule checks that a `house_assumptions` Target's **first word appears
literally in the template body**. Two of my original Targets failed this and were corrected:

- `error-logging-outcome-first`: I first used `tblErrorLog` (mirroring its paired schema). The
  outcome-first body never names `tblErrorLog` — it says "the log" / "log table". Changed to
  `Error log`.
- `stocktake-scan-outcome-first`: I first used `StockTakeCount.CountedQuantity` (mirroring its
  paired schema). The outcome-first body never names the table. Changed to `CountedQuantity`.

The lesson: a `house_assumptions` Target must name something the template's own body actually
names — not the paired schema's name.

## Out of scope (unchanged, per the brief)

- The undocumented `build_paths:` key and the `implements` under-documentation in §2 (both need a
  frozen-file change; left for George).
- Standards-gate scope, template-family matrix, `status: draft` metadata, prompt naming/coverage.
- `templates/errors/error-logging-scaffold.md` is under `FREEZE.md` and was left untouched — its
  own `warnings` (plain) and `related` (quoted) already conform, so no change was needed there.

## Verification

- `python mcp-server/run_validate.py` → **21/21 well-formed**.
- `python -m unittest discover -s mcp-server/tests` → **31 OK, 1 skipped**.
- Sweep confirms: all `house_assumptions` and `related` entries are quoted, all `warnings` plain,
  and every `house_assumptions` entry has a `Target —` (em-dash) separator. No dangling
  references to the old form.

## Files changed (this task)

`app-startup-outcome-first`, `app-startup-scaffold`, `capital-asset-tracking-outcome-first`,
`audit-logging-lite-outcome-first`, `audit-logging-lite-scaffold`, `audit-logging-lite-schema`,
`error-logging-outcome-first`, `catalog-outcome-first`, `catalog-schema`,
`officiating-assignment-outcome-first`, `officiating-assignment-scaffold`,
`officiating-assignment-schema`, `stocktake-scan-outcome-first`, `stocktake-schema` — 14 template
files.
