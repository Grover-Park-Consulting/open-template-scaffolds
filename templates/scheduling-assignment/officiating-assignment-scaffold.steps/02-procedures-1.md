---
step: 02-procedures-1
title: "The procedures, part 1 of 2"
platform_facts: [mcp-module-import, mcp-line-numbers, recordset-append-crash, sql-insert-truncation, dao-table-build, vbe-line-continuation, vba-import-xml-entities]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

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

### ValidateAssignment — `Public Function` → `Boolean`

```vba
Public Function ValidateAssignment(ByVal lGameID As Long, _
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

### GetApplicablePayRate — `Public Function` → `Variant`

```vba
Public Function GetApplicablePayRate(ByVal lPlayLevelID As Long, _
                                     ByVal lOfficialPositionID As Long, _
                                     ByVal dtGameDate As Date) As Variant
    ' [SCAFFOLD] Resolve the effective-dated rate for (level, position) as of a game date.
    '            Returns Null when no rate row applies — caller surfaces a warning, never
    '            guesses a rate. The return type is Variant, not Currency, because VBA's own
    '            Currency data type (unlike an Access table field of that type, which is
    '            nullable) is a fixed-point numeric type and cannot hold Null - assigning Null
    '            to it raises runtime error 94. A $0 rate and "no rate exists" must stay
    '            distinguishable, matching what the paired outcome-first template's query
    '            returns for the same case (Business Rule 6, "never silently a $0 rate").
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
    If rs.EOF Then
        GetApplicablePayRate = Null
    Else
        GetApplicablePayRate = rs!PayRate
    End If

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

**Every caller must guard for `Null` before doing arithmetic that feeds a typed variable.**
Storing the return value directly — into another `Variant`, or into a control's bound property —
needs no guard; `Null` displays as blank and propagates cleanly. But summing or otherwise combining
it into a strictly-typed accumulator (`Currency`, `Long`, `Double`) raises runtime error 94 the
moment one game has no applicable rate. Wrap it first: `curTotal = curTotal + Nz(GetApplicablePayRate(...), 0)`
— and decide deliberately what the zero in that `Nz` means for whatever is being totaled, since it
is exactly the silent-$0 conflation Business Rule 6 exists to prevent. This is the one hazard the
**Compensation rollup** Extra Option below will hit on its first missing-rate row if built without
this guard.

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
    ' [SCAFFOLD] Setting tdf.ValidationRule needs tblGame to itself, the same as any DAO
    '            table alter - a bound form, an open datasheet, or a query holding it open
    '            all refuse this with error 3420, which names nothing about what is open.
    '            Check first and name it, rather than let the engine's own uninformative
    '            error be the first thing the developer sees.
    Dim tdf   As DAO.TableDef
    Dim sOpen As String

    On Error GoTo errHandler

    sOpen = ListOpenObjects()
    If Len(sOpen) > 0 Then
        MsgBox "Stopped. Something in this database is still open:" & vbCrLf & vbCrLf & _
               sOpen & vbCrLf & "Close it and run this again.", vbExclamation
        Exit Sub
    End If

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

### ListOpenObjects — `Public Function` → `String`

**Names everything open in this database, and returns an empty string when nothing is.**
`EnsureGameValidationRule` needs `tblGame` to itself to alter it, and Access refuses that while
anything is using the table — a bound form being the usual cause, and one that never names the
table it uses. This turns that prerequisite into something the code can check and report by name,
rather than something a developer meets as error 3420 with no idea what to close.

```vba
Public Function ListOpenObjects() As String
    ' [SCAFFOLD] Recognition only. Returns "" when nothing in this database is open, and
    '            otherwise one indented line per open object for a caller to print.
    '            Access's own tables are skipped: AllTables lists them, they are never this
    '            system's business, and an open one is not something a developer would be
    '            asked to close.
    '            No errHandler and therefore no line numbers: an object that cannot be asked
    '            is reported as not open, and the caller stops on whatever else it finds.
    Dim obj   As Object
    Dim sList As String

    sList = ""

    On Error Resume Next

    For Each obj In CurrentProject.AllForms
        If obj.IsLoaded Then sList = sList & "  Form: " & obj.Name & vbCrLf
    Next obj

    For Each obj In CurrentProject.AllReports
        If obj.IsLoaded Then sList = sList & "  Report: " & obj.Name & vbCrLf
    Next obj

    For Each obj In CurrentData.AllTables
        If Left$(obj.Name, 4) <> "MSys" Then
            If SysCmd(acSysCmdGetObjectState, acTable, obj.Name) <> 0 Then
                sList = sList & "  Table: " & obj.Name & vbCrLf
            End If
        End If
    Next obj

    For Each obj In CurrentData.AllQueries
        If SysCmd(acSysCmdGetObjectState, acQuery, obj.Name) <> 0 Then
            sList = sList & "  Query: " & obj.Name & vbCrLf
        End If
    Next obj

    On Error GoTo 0

    ListOpenObjects = sList
End Function
```
