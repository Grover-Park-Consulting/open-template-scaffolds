---
template: sports-officiating-assignment-scaffold
title: Sports Officiating Assignment — Assignment & Pay VBA Scaffold
domain: sports
type: vba-scaffold
version: 0.3.0
status: draft
implements: sports-officiating-assignment-schema
requires_tables:
  - tblGame
  - tblOfficial
  - tblGameOfficial
  - tlkpOfficialPosition
  - tblPositionRate
  - tblTeam
  - tlkpPlayLevel
  - tblAppSetting
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
target_module: modOfficiatingAssignment
new_procedures:
  - AssignOfficial
  - ValidateAssignment
  - GetApplicablePayRate
  - GetAppSetting
---

# Sports Officiating Assignment — Assignment & Pay VBA Scaffold

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Realize the assignment and pay logic the officiating **table** template
(`sports-officiating-assignment-schema`) defers to code: assigning an official to a game
position with friendly validation, resolving the effective-dated pay rate, and reading
application settings. As with every scaffold, this supplies **procedure skeletons** —
signatures, recordset plumbing, control flow, error structure — with the domain logic marked
against the table template's numbered **Business Rules**, and house style deferred to the
standards layer.

A query-style note worth keeping: the source database this domain was shaped from stored the
crew as two hardcoded columns and needed a `UNION ALL` query to un-pivot them into
one-row-per-assignment. Against the junction, that whole query collapses to
`SELECT ... FROM tblGameOfficial` with joins — the before/after argument for the
normalization, written per `query-style.md`.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here (signature, plumbing, control flow, error structure).
- **`[STANDARDS]`** — house style, deferred to the standards layer (`error-handling.md`, `query-style.md`, `naming-conventions.md`).
- **`[BUSINESS LOGIC]`** — the domain rule you fill in, sourced from the table template's Business Rules.

## Prerequisites

| Object | Role |
|---|---|
| `sports-officiating-assignment-schema` tables | The scaffold runs against the tables that template creates (`tblGame`/`tblOfficial`/`tblGameOfficial`, the lookups, `tblPositionRate`, `tblAppSetting`) |
| `tblAppSetting.OfficialPhotoFolder` seed row | Read by `GetAppSetting` for the photo feature (schema Business Rule 9) |
| A central error logger | `error-handling.md` |

### Where this module goes in a split database

A **split database** is the normal shape for Access applications, especially those in multi-user
environments: one file holds the tables (the **back end**, usually on a shared network drive — never
on OneDrive, Dropbox, or any other file-syncing cloud folder, which corrupts a shared Access back
end), and each person runs their own copy of a second file holding the forms, reports, and code (the
**front end**), whose tables are *links* pointing at the back end. A single-file database — one
.accdb holding everything — is an acceptable choice for one user, and everything here works there
too; put everything in that one file and ignore the distinction.

| Module | Back end | Front end | Why |
|---|---|---|---|
| `modOfficiatingAssignment` (this scaffold) | No | **Yes** | Every procedure here is called by the assignment form — it validates what the user picked, resolves a rate to display, and reads a setting the form needs. That is front-end work, so the code belongs beside the form. |

**Assigning officials is a shared job.** An assignor at the desk and a second person working from
home are both making assignments into the one back end, and that changes two things in this module:

1. **`ValidateAssignment` cannot be relied on alone.** It checks that the position is free, then
   `AssignOfficial` inserts. Between those two moments somebody else can insert the same position
   for the same game — both people's checks passed, and the junction's unique index on
   (`GameID`, `OfficialPositionID`) refuses the second insert with an engine error. Checking first does not
   remove the race; it makes it rare and keeps the *ordinary* refusal friendly, which is its real
   job. `AssignOfficial` must still trap the duplicate-key error and turn it into the same plain
   message, or the second assignor sees a raw Access error instead of "that position is already
   filled."
2. **`GetAppSetting` reads a shared table.** `tblAppSetting` lives in the back end, so a setting
   changed once reaches everyone — which is exactly why the photo folder is a setting (schema
   Business Rule 9) rather than a constant compiled into each front end.

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. **Every procedure ends
with the same `errHandler` block** — shown in full in `AssignOfficial` and referenced
thereafter, because the VBE-reflection form (`error-handling.md`) is *identical* in every
procedure by design. Line numbers are deliberately absent (house-specific; see
`error-handling.md`).

### AssignOfficial — `Public Function` → `Long` (entry point)

```vba
Public Function AssignOfficial(ByVal lGameID As Long, _
                               ByVal lOfficialID As Long, _
                               ByVal lOfficialPositionID As Long) As Long
    ' [SCAFFOLD] Assign one official to one position on one game.
    '            Returns the new GameOfficialID, or 0 when the assignment is refused.
    Dim sRefusalReason As String

    On Error GoTo errHandler

    If Not ValidateAssignment(lGameID, lOfficialID, lOfficialPositionID, sRefusalReason) Then
        ' [SCAFFOLD] friendly refusal before the junction's unique indexes can bark
        MsgBox sRefusalReason, vbExclamation
        GoTo Cleanup
    End If

    ' [BUSINESS LOGIC #1,#2] insert the tblGameOfficial row (GameID, OfficialID, OfficialPositionID)
    ' >>> insert per query-style.md; set AssignOfficial = new GameOfficialID <<<

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] error reporting comes from the standards layer.
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```

### ValidateAssignment — `Private Function` → `Boolean`

```vba
Private Function ValidateAssignment(ByVal lGameID As Long, _
                                    ByVal lOfficialID As Long, _
                                    ByVal lOfficialPositionID As Long, _
                                    ByRef sRefusalReason As String) As Boolean
    ' [SCAFFOLD] All-or-nothing pre-checks for one assignment; on the first failure,
    '            set sRefusalReason and exit False. True = safe to insert.
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #2] the position is not already filled for this game
    '                     (unique on GameID + OfficialPositionID, checked kindly here first)
    ' >>> existence query per query-style.md; on hit: sRefusalReason, GoTo Cleanup <<<

    ' [BUSINESS LOGIC #2] the official is not already working this game in another position
    '                     (unique on GameID + OfficialID)
    ' >>> existence query per query-style.md; on hit: sRefusalReason, GoTo Cleanup <<<

    ' [BUSINESS LOGIC #3] the official is active (tblOfficial.OfficialIsActive = True)
    ' >>> lookup per query-style.md; on inactive: sRefusalReason, GoTo Cleanup <<<

    ValidateAssignment = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    Resume Cleanup
End Function
```

### GetApplicablePayRate — `Private Function` → `Currency`

```vba
Private Function GetApplicablePayRate(ByVal lPlayLevelID As Long, _
                                      ByVal lOfficialPositionID As Long, _
                                      ByVal dtGameDate As Date) As Currency
    ' [SCAFFOLD] Resolve the effective-dated rate for (level, position) as of a game date.
    '            Returns 0 when no rate row applies — caller surfaces a warning, never guesses.
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #6] the tblPositionRate row for (lPlayLevelID, lOfficialPositionID) with the
    '                     latest EffectiveDate on or before dtGameDate
    ' >>> TOP 1 ... ORDER BY EffectiveDate DESC query per query-style.md <<<
    sSql = vbNullString

    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then GetApplicablePayRate = rs!PayRate

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Function
```

*(The game's play level is derived through its home team — schema Business Rule 5 — so a
caller resolves `lPlayLevelID` from `tblGame → tblTeam.PlayLevelID` before calling.)*

### GetAppSetting — `Public Function` → `String`

```vba
Public Function GetAppSetting(ByVal sSettingName As String) As String
    ' [SCAFFOLD] Read one named setting from tblAppSetting; empty string when absent.
    '            Single purpose: settings live in data, never hardcoded (Business Rule 9's
    '            photo folder is the first consumer; every later setting reuses this).
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [SCAFFOLD] look up SettingValue by SettingName (unique)
    ' >>> lookup query per query-style.md <<<
    sSql = vbNullString

    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then GetAppSetting = Nz(rs!SettingValue, vbNullString)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block
    Resume Cleanup
End Function
```

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure plus the error-reporting call and
  line-number policy come from `error-handling.md`, which ranks three options and says when each
  fits. The `MsgBox` block shown is option 3, which needs nothing installed; a practice with its
  own logger substitutes it at the call site, and may number lines or not.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house
  query standard (aliasing, where querydefs live, formatting, safe criteria).
- **Naming** — procedure, variable, and parameter names follow `naming-conventions.md`.

## Extra Options

*Named optional extensions, none of them filled in for an engagement; the filled copy is saved to the
developer's own library, not committed here.*

- **RemoveAssignment / ReassignOfficial** — the un-assign and swap counterparts of
  `AssignOfficial`, with the same friendly validation.
- **Crew-completeness check** — a per-game function reporting unfilled positions against the
  engagement's required crew.
- **Compensation rollup** — per-official or per-game pay summaries built on
  `GetApplicablePayRate` (pairs with the schema's compensation-ledger Extra Option).

## Parked / future considerations (not in this design)

- **Availability / conflict checking** — refusing assignments that overlap an official's other
  games by date/time; needs an engagement's travel-time and doubleheader policies first.
- **Rate-snapshot on assignment** — if pay must survive later rate edits, the assignment row
  stores the resolved rate at assignment time (the schema's compensation-ledger Extra Option
  carries the shape).
