# House Comparison — AccessCodeLib

**Who reads this:** a contributor, or a shop comparing its own house rules with this library's. It
records how one outside body of work was weighed against the OTS standards layer and George
Hepworth's GPC house standards, and what was decided about each point.

This file is a record, not a rule. Templates and the AI assistant never read it during a run; the
rules themselves live in `standards/` (OTS) and in the GPC house standards. Where a row says
*Adopted*, the adopted rule is in the file the row names, credited in its code.

## The source

| | |
|---|---|
| Name | AccessCodeLib |
| Where | access-codelib.net; reviewed from a downloaded copy of the repository's `master` branch, 2026-11-18 |
| What it is | A library of reusable VBA modules and classes for Microsoft Access, with recipes, tests and a module catalog. Code, not a standards document: its conventions were read from its contributor guide, its recipes and its code |
| Licence | BSD 3-Clause License, copyright 2010 access-codelib.net |
| Reviewed by | George Hepworth, with an AI assistant, 2026-11-18 |
| Permission to publish adapted work | Confirmed in writing by the source's author, 2026-11-18: use the code, or parts of it, for this library |

## Rulings

Four rulings are used, and each row carries exactly one:

- **Adopted** — the method is now part of OTS and GPC, credited to the source in the code.
- **House difference** — both ways work; each house keeps its own. Recorded so the difference is
  visible, not because either is wrong.
- **Conflict** — the source's practice is one the OTS standards rule out, for the reason given.
- **Parked** — worth a closer look later; no decision yet.

| Topic | OTS | GPC | AccessCodeLib | Ruling | Reason |
|---|---|---|---|---|---|
| Values in a statement built in code | `query-style.md` §4: `PARAMETERS` on a temporary QueryDef | `VBA-Standards.md` §8 | `data/dao/DaoHandler.cls`, `ExecuteParamSql` | Adopted | Values never enter the SQL text, and the engine checks their types |
| Number and date literals in built SQL | `query-style.md` §4: `Str$`, escaped ISO date formats | `VBA-Standards.md` §8 | `data/SqlTools.cls`, `NumberToSqlText`, `DateToSqlText` | Adopted | Plain concatenation follows the machine's regional settings. Tested by the source's contributors on other regional settings, not on George's machine |
| Date ranges | `query-style.md` §4: on or after the first day, before the day after the last | `VBA-Standards.md` §8 | `data/SqlTools.cls`, `BuildCriteria` with `SQL_Add_WildCardSuffix` | Adopted | `BETWEEN` misses every row after midnight on the last day |
| An empty search value | `query-style.md` §5: adds no condition | `VBA-Standards.md` §8 | `data/SqlTools.cls`, `data/FilterStringBuilder.cls` | Adopted | A comparison with Null matches no row |
| Re-raising an error to the caller | `error-handling.md`: copy `Err` first, name the procedure in the source | `VBA-Standards.md` §6 | `data/dao/DaoHandler.cls` handlers; `_codelib/recipes/dao-data-access.md` | Adopted | Tested 2026-11-18: releasing a class in cleanup can clear `Err`, and the old shape then raised error 5 in place of the real error |
| Recording the path an error took | `templates/errors/error-logging-schema.md`: `ErrorSource` column | — | `base/modErrorHandler.bas`, `->` source delimiter | Adopted | The logging procedure's own `Erl` reports only the line of its call |
| String comparison in modules | `Option Compare Database` | `Option Compare Database` (§2) | `Option Compare Text` (`CONTRIBUTING.md`) | House difference | Both are valid; this one changes how `=` and `InStr` compare text, so a module moved between houses must be checked |
| Variable names | Type prefix (`naming-conventions.md` §7) | Type prefix (§3) | No prefix (`RecordsAffected`, `Criteria`) | House difference | A naming choice |
| Private class fields | `m` prefix (`mForm`) | `m` prefix (§3) | `m_` prefix (`m_TransIsStarted`) | House difference | A naming choice |
| Error handler labels | `errHandler:`, `Cleanup:` | Same (§6) | `HandleErr:`, `ExitHere:` | House difference | A naming choice |
| Line numbers and `Erl` | Numbered where the handler reads `Erl` | Numbered (§5) | None; `Erl` is not used anywhere in the library | House difference | OTS and GPC use the failing line in the log; the source logs without it |
| Hidden system table prefix | `USys` | `USys` (§3) | `usys_` (`usys_DbmsConnection`) | House difference | A naming choice |
| Module header | No header required | GPC attribution header (§2) | Natural Docs comments and a `<codelib>` block listing file, dependencies, licence, tests | House difference | The source's header feeds its own import tool; see Parked for provenance |
| What the person sees on an error | Settled once, in the logger (`error-handling.md`) | `GlblErrMsg` (§6) | A mode chosen per call or as a default: raise, message box, ignore, plus file | House difference | Same idea of one central decision; the ignore mode is a conflict, below |
| Ignoring an error | Not permitted | Not permitted | `aclibErrIgnore` mode in `base/modErrorHandler.bas` | Conflict | An ignored error leaves no record and lets the caller carry on as if the step worked |
| Naming the failing procedure | `MODULE_NAME` / `PROC_NAME` constants | Same (§6) | A literal passed to `HandleError` (`HandleError(Err.Number, "ExampleProc", …)`) | Conflict | A literal drifts when the procedure is renamed |
| Auditing data changes | Data Macros on the table (`audit-columns.md`) | Same | `form/data/AuditTrail.cls`, form events writing through ADODB | Conflict | Many shops use form events, but they record only changes made through that form; a query, an import or another form writes unrecorded |
| Task-oriented recipes | — | — | `_codelib/recipes/`: Keywords, Import, Usage, Notes, See also | Parked | A possible shape for the lesson store or OTS Pro; a later side project |
| Transaction wrapper class | `booInTrans` flag (`error-handling.md`, *Transaction guard*) | Same (§6) | `data/dao/DaoTransactionHandler.cls`: rolls back when the object is destroyed | Parked | A safety net worth reviewing later |
| Provenance header in generated modules | — | — | `<codelib>` block | Parked | Idea for OTS Pro: name the template, version and fingerprint a module came from |
