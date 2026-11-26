---
step: 00-procedures
title: "The procedures this template writes"
platform_facts: []
route: both
---

**Who reads this:** the AI assistant, drafting the design.

## The procedures

These are this template's procedures as its code steps write them. Draft the design spec from them:
give each procedure the declaration below and say what it does. Draft no signature or behaviour of
your own. A change to either is the developer's decision: ask it before you present the design.

### AssignOfficial

```vba
Public Function AssignOfficial(ByVal lGameID As Long, _
                               ByVal lOfficialID As Long, _
                               ByVal lOfficialPositionID As Long) As Long
```

Assign one official to one position on one game. Returns the new GameOfficialID, or 0 when the assignment is refused.

### ValidateAssignment

```vba
Public Function ValidateAssignment(ByVal lGameID As Long, _
                                    ByVal lOfficialID As Long, _
                                    ByVal lOfficialPositionID As Long, _
                                    ByRef sRefusalReason As String) As Boolean
```

All-or-nothing pre-checks for one assignment; on the first failure, set sRefusalReason and exit False. True = safe to insert.

### GetApplicablePayRate

```vba
Public Function GetApplicablePayRate(ByVal lPlayLevelID As Long, _
                                     ByVal lOfficialPositionID As Long, _
                                     ByVal dtGameDate As Date) As Variant
```

Resolve the effective-dated rate for (level, position) as of a game date. Returns Null when no rate row applies — caller surfaces a warning, never guesses a rate. The return type is Variant, not Currency, because VBA's own Currency data type (unlike an Access table field of that type, which is nullable) is a fixed-point numeric type and cannot hold Null - assigning Null to it raises runtime error 94. A $0 rate and "no rate exists" must stay distinguishable, matching what the paired outcome-first template's query returns for the same case (Business Rule 6, "never silently a $0 rate").

### GetAppSetting

```vba
Public Function GetAppSetting(ByVal sSettingName As String) As String
```

Read one named setting from tblAppSetting; empty string when absent. Single purpose: settings live in data, never hardcoded (Business Rule 9's photo folder is the first consumer; every later setting reuses this).

### EnsureGameValidationRule

```vba
Public Sub EnsureGameValidationRule()
```

Attach the row-level Validation Rule enforcing Business Rules 4 and 7 to tblGame. Run once, after the schema's tables are built; safe to run again - it overwrites the same text rather than layering a second copy.

### GameLevelID

```vba
Public Function GameLevelID(ByVal lGameID As Long) As Long
```

A game's play level, read through its home team - never stored on the game itself (schema Business Rule 5). Returns 0 when the game or its home team can't be found.

### OfficialAge

```vba
Public Function OfficialAge(ByVal lOfficialID As Long) As Variant
```

An official's age, computed from BirthDate as of today - never stored anywhere (schema Business Rule 8). Returns Null when BirthDate is unknown, which is why this is a Variant rather than an Integer.

### EnsurePhotoFolder

```vba
Public Function EnsurePhotoFolder() As Boolean
```

Confirm the shared folder named in tblAppSetting.OfficialPhotoFolder exists - never create one locally in its place (schema Business Rule 9; see templates/_materialization.md -> External file assets, "shared content folders: VERIFY, never create"). Returns False, with a message naming the folder, when it can't be reached.

### SetOfficialPhoto

```vba
Public Function SetOfficialPhoto(ByVal lOfficialID As Long, ByVal sSourcePath As String) As Boolean
```

Copy a picked photo into the shared folder under a controlled name, and record that name on the official - never the source path, never the file itself (schema Business Rule 9). The controlled name carries the picked file's own extension, which can change between picks for the same official (a .jpg replaced by a .png) - so the new name is not always the old name. Before writing the new file, this reads the official's CURRENT PhotoFileName - the full stored name, extension included - and deletes exactly that file if it differs from the name about to be written. Deleting by the recorded full name, not by guessing the old extension matches the new one, is what keeps this accurate across an extension change.

### ListOpenObjects

```vba
Public Function ListOpenObjects() As String
```

Recognition only. Returns "" when nothing in this database is open, and otherwise one indented line per open object for a caller to print. Access's own tables are skipped: AllTables lists them, they are never this system's business, and an open one is not something a developer would be asked to close. No errHandler and therefore no line numbers: an object that cannot be asked is reported as not open, and the caller stops on whatever else it finds.
