---
template: sports-officiating-assignment-scaffold
title: Sports Officiating Assignment — Assignment & Pay VBA Scaffold
domain: scheduling-assignment
type: vba-scaffold
version: 0.5.0
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
  - EnsureGameValidationRule
  - GameLevelID
  - OfficialAge
  - EnsurePhotoFolder
  - SetOfficialPhoto
warnings:
  - "This build's active-official check (Business Rule 3, inside ValidateAssignment) only covers
    assignments made through AssignOfficial. An assignment inserted directly into tblGameOfficial, or
    by an import, is not checked. This is the VBA route named in
    sports-officiating-assignment-outcome-first — that template's Data Macro route closes this gap,
    at greater build cost. Choosing this scaffold is choosing this trade-off."
---

# Sports Officiating Assignment — Assignment & Pay VBA Scaffold

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not have to read it to use the template.

## Intent

Realize seven of the officiating **table** template's (`sports-officiating-assignment-schema`)
nine Business Rules in code: assigning an official to a game position with friendly validation
and an active-official check (Rules 2, 3), resolving the effective-dated pay rate (Rule 6),
keeping a game internally consistent — two different teams, an end time after the start
(Rules 4, 7) — deriving a game's play level and an official's age rather than storing either
(Rules 5, 8), and handling a photo as a file name plus one shared, confirmed folder (Rule 9).
The other two — the crew being a junction, and assignment uniqueness — are already complete the
moment the table template's tables are built; nothing here touches them. As with every scaffold,
this supplies **procedure skeletons** — signatures, recordset plumbing, control flow, error
structure — with the domain logic marked against the table template's numbered **Business
Rules**, and house style deferred to the standards layer.

**The other version of this template — the outcome-first method,
`sports-officiating-assignment-outcome-first` — produces the same result from a specification
rather than working code, and offers a genuine choice for Rule 3** (a Data Macro that closes the
coverage gap named in the warning above, or this scaffold's own VBA route). Either can be built
against your own database, and they can be built one after the other, against separate copies, to
compare.

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
| `tblAppSetting.OfficialPhotoFolder` seed row | Read by `GetAppSetting`, `EnsurePhotoFolder`, and `SetOfficialPhoto` (Business Rule 9). **In a split database this must be an absolute shared path**, e.g. `\\server\share\OfficialPhotos\` — the schema's own seeded value is relative, which only suits a single-file database. |
| A photo picker | The screen that calls `SetOfficialPhoto` with the file the person chose — a form concern, deferred to a `form-spec` template. |
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

## Validating the build

**To the AI assistant.** This template and `sports-officiating-assignment-outcome-first` reach the
same result by different routes, and each says so where it names the other. That result is what gets
checked, not which route built it, so there is one checklist for both rather than two.

**Run every entry under `sports-officiating-assignment-outcome-first.md`'s "How you validate the
template's output" against this build, on a copy, exactly as that template requires.** Do this
whether you generated the code yourself or handed the developer the files to import — the checks read
the database this build produced, not the procedures that produced it. Report against that same
numbered list in the build record: one entry per check, what was done and what was observed. Passed
and not passed are the only outcomes, including where the first method to run a check hits an
obstacle — see `_template-schema.md` §12.2 for the full rule, the `Result: PASSED` / `Result: NOT
PASSED` line every entry opens with, and what to do before settling for a soft result.

**Check 2 is the one entry that turns on which route built the database, and that list already says
so.** It tests that an inactive official is refused on a direct edit to `tblGameOfficial`, which only
the Data Macro route reaches. This template is the VBA route named under that file's *Business Rule
3*: `ValidateAssignment` covers whatever calls it and nothing else. Run the check anyway, record what
you observed, name the route this build took, and report the entry the way check 2's own text
directs — it is a disclosed trade-off of this route, not something this build got wrong.

**Do not devise your own list in place of that one.** Reading the procedures below and working out a
plausible set of checks from them is easy to do and produces a list that tests what this code does.
The list above tests what the developer was promised, which is a different thing and the only one of
the two that is validation. So a build has passed when it has passed those checks, and a report
saying validation passed means those checks and no others — name the list you ran, so the developer
can see which one it was.

Three things follow from this being procedure skeletons rather than an open route.

- **Run `EnsureGameValidationRule` before checks 3, 4, 8, 9 and 10.** All five test a game-level rule
  "whatever route you use," one of them through an append or import query — and nothing in VBA
  reaches a save made that way. The table-level Validation Rule does, which is what that procedure
  puts in place. It is a one-time setup call rather than part of the assignment path, so a build that
  never runs it fails all five at once with the wrong cause attached to each.

- **Drive checks 1 and 2 through `AssignOfficial`, not by inserting rows into `tblGameOfficial`.** An
  insert made directly bypasses `ValidateAssignment`, which on this route is the whole of what those
  two checks examine.

- **Compile the host's VBA project before running any of them, and record that you did.** The checks
  read a database, and code that will not compile never reaches it — so an uncompiled build fails
  every check on the list at once, with the wrong cause attached to each of them. A compile catches
  that in seconds.

## Procedures

Each procedure shows its scope, signature, and an annotated skeleton. **Every procedure ends
with the same `errHandler` block** — shown in full in `AssignOfficial` and referenced
thereafter, because the errHandler block's form (`error-handling.md`) is *identical* in every
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

*(The game's play level is derived through its home team — schema Business Rule 5 —
so `GetApplicablePayRate`'s `lPlayLevelID` argument is resolved by calling `GameLevelID`,
below, rather than read from anywhere stored.)*

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

### EnsureGameValidationRule — `Public Sub`

```vba
Public Sub EnsureGameValidationRule()
    ' [SCAFFOLD] Attach the row-level Validation Rule enforcing Business Rules 4 and 7 to
    '            tblGame. Run once, after the schema's tables are built; safe to run again -
    '            it overwrites the same text rather than layering a second copy.
    Dim tdf As DAO.TableDef

    On Error GoTo errHandler

    Set tdf = CurrentDb.TableDefs("tblGame")

    ' [BUSINESS LOGIC #4,#7] a team can't play itself; when GameEnd is present, it's later
    '                        than GameStart
    tdf.ValidationRule = "[HomeTeamID]<>[AwayTeamID] And " & _
                          "([GameEnd] Is Null Or [GameEnd]>[GameStart])"
    tdf.ValidationText = "A game needs two different teams, and an end time (if given) " & _
                          "after the start time."

Cleanup:
    On Error Resume Next
    Set tdf = Nothing
    Exit Sub

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Sub
```

### GameLevelID — `Public Function` → `Long`

```vba
Public Function GameLevelID(ByVal lGameID As Long) As Long
    ' [SCAFFOLD] A game's play level, read through its home team - never stored on the game
    '            itself (schema Business Rule 5). Returns 0 when the game or its home team
    '            can't be found.
    Dim db   As DAO.Database
    Dim rs   As DAO.Recordset
    Dim sSql As String

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #5] tblGame -> tblTeam (on HomeTeamID) -> PlayLevelID
    ' >>> join query per query-style.md <<<
    sSql = vbNullString

    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then GameLevelID = Nz(rs!PlayLevelID, 0)

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```

### OfficialAge — `Public Function` → `Variant`

```vba
Public Function OfficialAge(ByVal lOfficialID As Long) As Variant
    ' [SCAFFOLD] An official's age, computed from BirthDate as of today - never stored
    '            anywhere (schema Business Rule 8). Returns Null when BirthDate is unknown,
    '            which is why this is a Variant rather than an Integer.
    Dim db      As DAO.Database
    Dim rs      As DAO.Recordset
    Dim sSql    As String
    Dim dtBirth As Variant

    On Error GoTo errHandler
    Set db = CurrentDb

    ' [BUSINESS LOGIC #8] look up tblOfficial.BirthDate for lOfficialID
    ' >>> lookup query per query-style.md <<<
    sSql = vbNullString

    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then
        dtBirth = rs!BirthDate
        If Not IsNull(dtBirth) Then
            ' Subtract a year when this year's birthday hasn't happened yet.
            OfficialAge = DateDiff("yyyy", dtBirth, Date) - _
                IIf(Format(Date, "mmdd") < Format(dtBirth, "mmdd"), 1, 0)
        End If
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```

### EnsurePhotoFolder — `Public Function` → `Boolean`

```vba
Public Function EnsurePhotoFolder() As Boolean
    ' [SCAFFOLD] Confirm the shared folder named in tblAppSetting.OfficialPhotoFolder exists -
    '            never create one locally in its place (schema Business Rule 9; see
    '            templates/_materialization.md -> External file assets, "shared content
    '            folders: VERIFY, never create"). Returns False, with a message naming the
    '            folder, when it can't be reached.
    Dim sFolder As String

    On Error GoTo errHandler

    sFolder = GetAppSetting("OfficialPhotoFolder")

    If Len(sFolder) = 0 Then
        MsgBox "No photo folder is configured (OfficialPhotoFolder is blank).", vbExclamation
        GoTo Cleanup
    End If

    If Len(Dir$(sFolder, vbDirectory)) = 0 Then
        MsgBox "The shared photo folder cannot be reached:" & vbCrLf & sFolder, vbExclamation
        GoTo Cleanup
    End If

    EnsurePhotoFolder = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
End Function
```

### SetOfficialPhoto — `Public Function` → `Boolean`

```vba
Public Function SetOfficialPhoto(ByVal lOfficialID As Long, ByVal sSourcePath As String) As Boolean
    ' [SCAFFOLD] Copy a picked photo into the shared folder under a controlled name, and
    '            record that name on the official - never the source path, never the file
    '            itself (schema Business Rule 9). The controlled name is always the same for
    '            a given official, so a re-pick overwrites cleanly instead of leaving the old
    '            file behind, orphaned.
    Dim sFolder As String
    Dim sExt    As String
    Dim sName   As String
    Dim lDot    As Long

    On Error GoTo errHandler

    If Not EnsurePhotoFolder() Then GoTo Cleanup

    sFolder = GetAppSetting("OfficialPhotoFolder")
    If Right$(sFolder, 1) <> "\" Then sFolder = sFolder & "\"

    lDot = InStrRev(sSourcePath, ".")
    sExt = IIf(lDot > 0, Mid$(sSourcePath, lDot), vbNullString)
    sName = "Official_" & lOfficialID & sExt

    FileCopy sSourcePath, sFolder & sName

    ' [BUSINESS LOGIC #9] UPDATE tblOfficial SET PhotoFileName = sName WHERE OfficialID = lOfficialID
    ' >>> update per query-style.md <<<

    SetOfficialPhoto = True

Cleanup:
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] standard errHandler block (see AssignOfficial)
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Cleanup
    Resume
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
