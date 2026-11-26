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

### LogError

```vba
Public Sub LogError(ByVal strModule As String, _
                    ByVal strProcedure As String, _
                    ByVal lngLine As Long)
```

Called from an errHandler block that is already dealing with a failure. Nothing in here may raise an error of its own.

### WriteErrorToTable

```vba
Private Function WriteErrorToTable(ByVal lngErrNumber As Long, _
                                   ByVal strErrDescription As String, _
                                   ByVal strErrSource As String, _
                                   ByVal strModule As String, _
                                   ByVal strProcedure As String, _
                                   ByVal lngLine As Long, _
                                   ByRef lngLogID As Long) As Boolean
```

Guarded throughout; reports success as a Boolean instead of raising.

### WriteErrorToFile

```vba
Private Function WriteErrorToFile(ByVal lngErrNumber As Long, _
                                  ByVal strErrDescription As String, _
                                  ByVal strErrSource As String, _
                                  ByVal strModule As String, _
                                  ByVal strProcedure As String, _
                                  ByVal lngLine As Long) As Boolean
```

Guarded throughout; reports success as a Boolean instead of raising.

### GetErrorLogFilePath

```vba
Private Function GetErrorLogFilePath() As String
```

Guarded: a bad path returns the fallback, never an error.

### ShowErrorToUser

```vba
Private Sub ShowErrorToUser(ByVal lngErrNumber As Long, _
                            ByVal strErrDescription As String, _
                            ByVal strModule As String, _
                            ByVal strProcedure As String, _
                            ByVal lngLine As Long, _
                            ByVal lngLogID As Long, _
                            ByVal booRecorded As Boolean)
```

Guarded: a failure to show the message must not replace the failure we were called about.

### CreateErrorLogTable

```vba
Public Function CreateErrorLogTable(Optional booSilent As Boolean = False) As String
```

Creates tblErrorLog (schema template entity). Idempotent: an existing table is reported and skipped, so this is safe to re-run. booSilent:=True suppresses the message box and returns the same text.

### SetFieldDescription

```vba
Private Sub SetFieldDescription(db As DAO.Database, _
                                ByVal strField As String, _
                                ByVal strDescription As String)
```

Appending a property that already exists raises 3367, so look for it first and set it where it is. Any other error is real and reaches the caller's handler: a field that cannot be described is a build that went wrong.

### LinkErrorLogTable

```vba
Public Function LinkErrorLogTable(Optional ByVal strBackEndPath As String = "", _
                                  Optional booSilent As Boolean = False) As String
```

Links tblErrorLog from the back end into this front end. Needed only where Step 4 put the log table in the back end. Idempotent: a table that is already here is reported and left alone, so this is safe to re-run. booSilent:=True suppresses the message box, returning the text.
