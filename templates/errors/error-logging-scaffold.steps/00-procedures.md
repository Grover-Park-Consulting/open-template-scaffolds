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
Public Sub LogError(ByVal sModule As String, _
                    ByVal sProcedure As String, _
                    ByVal lLine As Long)
```

Called from an errHandler block that is already dealing with a failure. Nothing in here may raise an error of its own.

### WriteErrorToTable

```vba
Private Function WriteErrorToTable(ByVal lErrNumber As Long, _
                                   ByVal sErrDescription As String, _
                                   ByVal sErrSource As String, _
                                   ByVal sModule As String, _
                                   ByVal sProcedure As String, _
                                   ByVal lLine As Long, _
                                   ByRef lLogID As Long) As Boolean
```

Guarded throughout; reports success as a Boolean instead of raising.

### WriteErrorToFile

```vba
Private Function WriteErrorToFile(ByVal lErrNumber As Long, _
                                  ByVal sErrDescription As String, _
                                  ByVal sErrSource As String, _
                                  ByVal sModule As String, _
                                  ByVal sProcedure As String, _
                                  ByVal lLine As Long) As Boolean
```

Guarded throughout; reports success as a Boolean instead of raising.

### GetErrorLogFilePath

```vba
Private Function GetErrorLogFilePath() As String
```

Guarded: a bad path returns the fallback, never an error.

### ShowErrorToUser

```vba
Private Sub ShowErrorToUser(ByVal lErrNumber As Long, _
                            ByVal sErrDescription As String, _
                            ByVal sModule As String, _
                            ByVal sProcedure As String, _
                            ByVal lLine As Long, _
                            ByVal lLogID As Long, _
                            ByVal bRecorded As Boolean)
```

Guarded: a failure to show the message must not replace the failure we were called about.

### CreateErrorLogTable

```vba
Public Function CreateErrorLogTable(Optional bSilent As Boolean = False) As String
```

Creates tblErrorLog (schema template entity). Idempotent: an existing table is reported and skipped, so this is safe to re-run. bSilent:=True suppresses the message box and returns the same text.

### SetFieldDescription

```vba
Private Sub SetFieldDescription(db As DAO.Database, _
                                ByVal sField As String, _
                                ByVal sDescription As String)
```

Guarded because a documentation comment is never worth failing a build for, and because appending a property that already exists raises 3367.

### LinkErrorLogTable

```vba
Public Function LinkErrorLogTable(Optional ByVal sBackEndPath As String = "", _
                                  Optional bSilent As Boolean = False) As String
```

Links tblErrorLog from the back end into this front end. Needed only where Step 4 put the log table in the back end. Idempotent: a table that is already here is reported and left alone, so this is safe to re-run. bSilent:=True suppresses the message box, returning the text.
