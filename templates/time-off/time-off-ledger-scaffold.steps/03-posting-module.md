---
step: 03-posting-module
title: Write and import the posting module, modTimeOffPosting
platform_facts: [mcp-module-import, mcp-line-numbers, sql-insert-truncation, recordset-append-crash, domain-function-transaction]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures: the posting module

The two error shapes, and *the standard reporting handler*, are described at the start of step 2;
the handler is shown in full in `PostOnePeriod` below.

**The module header, `modTimeOffPosting`**

```vba
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modTimeOffPosting"
Private Const NOTE_MAX    As Long = 255     ' the width of TimeOffEntryNote
```

### ExpectedRefusal — `Private Function` → `Boolean`

```vba
Private Function ExpectedRefusal(ByVal lErr As Long) As Boolean
    ' [SCAFFOLD] Errors that mean "the table said no", not "something broke": the Data Macro's own
    '            refusal (3939) and the engine's refusals of a duplicate, a missing parent, a bad
    '            value, or a row held by someone else. A refusal is reported to the person as a
    '            sentence and is not logged. Anything else is a failure nobody planned for and is
    '            logged. Never mistake the second kind for the first.
    Select Case lErr
        Case 3939, 3022, 3200, 3201, 3314, 3315, 3316, 3317, 3163, 3218, 3260, 3188
            ExpectedRefusal = True
    End Select
End Function
```

### AppendEntry — `Private Function` → `Long`

```vba
Private Function AppendEntry(ByVal db As DAO.Database, ByVal lEmp As Long, ByVal lCat As Long, _
                             ByVal sReason As String, ByVal dDate As Date, ByVal dHours As Double, _
                             ByVal vPeriod As Variant, ByVal vCorrects As Variant, _
                             ByVal vReplaces As Variant, ByVal sNote As String) As Long
    ' [SCAFFOLD] Appends one entry and returns its TimeOffEntryID. Raises whatever the table or its
    '            Data Macro raises; the caller decides what that means.
    '            Written by a parameterised SQL INSERT on db, never a recordset append: a recordset
    '            append from code crashed Access while the Data Macro's function read through
    '            CurrentDb (_materialization.md, "A recordset append from code can crash Access...").
    '            Parameters carry every value, so an apostrophe in a note needs no quoting.
    Dim qd As DAO.QueryDef

    ' >>> If sNote is longer than the size of TimeOffEntryNote read from db.TableDefs, raise 3163 (the
    '     engine's own "field is too small") with a sentence naming the limit, so ExpectedRefusal
    '     treats it as the refusal a recordset would have given. CreateQueryDef("", a PARAMETERS ...
    '     INSERT INTO tblTimeOffEntry statement) on db; set every parameter, Null for the optional
    '     ones not given; Execute dbFailOnError; read the new key with SELECT @@IDENTITY on the same
    '     db; close the QueryDef <<<
End Function
```

### PostOnePeriod — `Private Function` → `Long`

```vba
Private Function PostOnePeriod(ByVal lEmp As Long, ByVal lCat As Long, ByVal dPeriod As Date, _
                               ByVal bOnPurpose As Boolean, ByVal sNote As String, _
                               ByRef sMsg As String) As Long
    ' [SCAFFOLD] Posts the hours earned for one period. Returns 1 = posted, 2 = already there,
    '            3 = could not be posted (sMsg says why). Business Rule 4.
    '            The employee's lock is taken before the "already there?" look and held until the
    '            entry is written, so two runs reaching one period take turns and the second finds
    '            the first one's entry. Neither reports an error.
    Const PROC_NAME As String = "PostOnePeriod"
    Dim ws       As DAO.Workspace
    Dim db       As DAO.Database
    Dim bInTrans As Boolean
    Dim dHours   As Double
    Dim sWhy     As String
    Dim lResult  As Long

    On Error GoTo errHandler
    ' [STANDARDS — error-handling.md, "Transaction guard"] db from the same Workspace as BeginTrans.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)
    ws.BeginTrans
    bInTrans = True
    Set gdbTxn = db                    ' the checks the macro runs read through this transaction
    LockEmployee db, lEmp

    ' [BUSINESS LOGIC #4] Unless bOnPurpose, where CountEarnedForPeriod is above zero: lResult = 2 and
    '            sMsg says the employee already has hours for the period, and that a cancelled entry
    '            counts. bOnPurpose skips this look; the table's own check still refuses a second
    '            ACTIVE entry.
    ' [BUSINESS LOGIC #3] Otherwise dHours = AccrualHours(...). Where it is 0: lResult = 3 and sMsg
    '            says why, and that nothing was posted for it.
    ' [BUSINESS LOGIC #4] Otherwise AppendEntry an Earned entry (EntryDate and AccrualPeriodStart both
    '            the period's first day); lResult = 1; sMsg says how many hours, for whom, which
    '            period, and the new entry's number.
    ' >>> the three branches <<<

    ws.CommitTrans
    bInTrans = False
    PostOnePeriod = lResult

Cleanup:
    Set gdbTxn = Nothing
    Set db = Nothing
    Set ws = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] reporting shape, plus the transaction guard's rollback. A
    '            refusal the table made becomes a sentence; anything else is logged and the person
    '            is told only that something unexpected happened and was recorded.
    If ExpectedRefusal(Err.Number) Then
        sMsg = "Not posted: " & Err.Description
        If bInTrans Then ws.Rollback
        bInTrans = False
    Else
        If bInTrans Then ws.Rollback
        bInTrans = False
        LogError MODULE_NAME, PROC_NAME, Erl
        sMsg = "Not posted: an unexpected problem occurred and was recorded."
    End If
    PostOnePeriod = 3
    Resume Cleanup
    Resume
End Function
```

### PostEarnedForPeriod — `Public Function` → `String`

```vba
Public Function PostEarnedForPeriod(ByVal lEmp As Long, ByVal lCat As Long, ByVal dPeriodStart As Date, _
                                    Optional ByVal sNote As String = "") As String
    ' [SCAFFOLD] Posts the hours earned for one period; a period that already has an Earned entry,
    '            cancelled or not, is left alone. A note longer than NOTE_MAX is refused here, with
    '            a sentence, before anything is written. Calls PostOnePeriod with bOnPurpose False;
    '            where it answers 2 the sentence begins "Nothing posted:". No dialog.
    ' >>> body <<<
End Function
```

### PostEarnedOnPurpose — `Public Function` → `String`

```vba
Public Function PostEarnedOnPurpose(ByVal lEmp As Long, ByVal lCat As Long, ByVal dPeriodStart As Date, _
                                    Optional ByVal sNote As String = "") As String
    ' [SCAFFOLD] How a person gives hours that a Correction took away (Business Rule 4). Calls
    '            PostOnePeriod with bOnPurpose True. The table still refuses a second active entry.
    ' >>> body, as PostEarnedForPeriod <<<
End Function
```

### PostDuePeriods — `Public Function` → `String`

```vba
Public Function PostDuePeriods(Optional ByVal lEmp As Long = 0) As String
    ' [SCAFFOLD] Posts every accrual period that is due, for one employee (lEmp) or all of them, and
    '            answers with one sentence: how many it posted, how many were already there, how many
    '            it could not post, and the first it could not post with its reason. A period it
    '            could not post is never skipped without being counted. Reporting shape.
    Const PROC_NAME As String = "PostDuePeriods"
    Dim db     As DAO.Database
    Dim rsE    As DAO.Recordset
    Dim rsC    As DAO.Recordset
    Dim lPosted As Long, lThere As Long, lFailed As Long
    Dim sFirst As String

    On Error GoTo errHandler
    Set db = CurrentDb         ' [SCAFFOLD] Reads only the employee and category lists. Each period is
                               '            posted by PostOnePeriod in a transaction of its own, so
                               '            nothing here needs a transaction's view.

    ' [BUSINESS LOGIC #4] For each employee (or the one asked for) and each category, walk the periods
    '            from HireDate in blocks of AccrualPeriodMonths, computing each period's start from
    '            HireDate by whole blocks (never by adding to the previous start, which drifts at
    '            month ends). A period is DUE when it has started (on or before today), is on or
    '            after the employee's AccrualStartDate (HireDate where that is empty), and starts no
    '            later than the InactiveDate where there is one. Call PostOnePeriod for each due
    '            period and count the three outcomes. Put an upper limit on the walk so a bad
    '            AccrualPeriodMonths cannot loop forever.
    ' >>> the two loops, per query-style.md <<<

Cleanup:
    On Error Resume Next
    If Not rsC Is Nothing Then rsC.Close
    If Not rsE Is Nothing Then rsE.Close
    Set rsC = Nothing: Set rsE = Nothing: Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] reporting shape. Say how many periods were posted before it
    '            stopped, because those are in the ledger and cannot be taken back.
    LogError MODULE_NAME, PROC_NAME, Erl
    PostDuePeriods = "Not finished: an unexpected problem occurred and was recorded. Posted " & _
                     lPosted & " period(s) before it."
    Resume Cleanup
    Resume
End Function
```

### PostTimeTaken — `Public Function` → `String`

```vba
Public Function PostTimeTaken(ByVal lEmp As Long, ByVal lCat As Long, ByVal dDate As Date, _
                              ByVal dHours As Double) As String
    ' [SCAFFOLD] Records time taken for one day. dHours is the amount taken, positive; it is stored
    '            negative. The checks (balance, one kind of time off per day, inactive date) run in
    '            the Data Macro, not here. One transaction on the default Workspace; gdbTxn set after
    '            BeginTrans; LockEmployee first; AppendEntry a Taken entry; CommitTrans; the answer
    '            names the hours, category, employee, day, entry number and the balance now.
    '            Hours of zero or less are refused with a sentence before anything is written.
    Const PROC_NAME As String = "PostTimeTaken"
    ' >>> as PostOnePeriod's frame: BeginTrans, gdbTxn, LockEmployee, write, CommitTrans <<<
    ' errHandler: the standard reporting handler (see PostOnePeriod). Cleanup clears gdbTxn.
End Function
```

### CancelTimeOffEntry — `Public Function` → `String`

```vba
Public Function CancelTimeOffEntry(ByVal lEntryID As Long, ByVal sNote As String) As String
    ' [SCAFFOLD] Cancels one entry with a Correction of equal and opposite hours. The entry cancelled
    '            is left exactly as it was. A Correction of hours earned needs a note, and the
    '            Data Macro refuses one without.
    Const PROC_NAME As String = "CancelTimeOffEntry"
    ' [BUSINESS LOGIC #7] Read the entry's employee, category, date and hours first (a snapshot
    '            recordset, before the transaction); where it does not exist, say so.
    ' >>> BeginTrans, gdbTxn, LockEmployee, AppendEntry a Correction on the entry's own date with
    '     the opposite hours, naming the entry; CommitTrans; the answer gives the new balance <<<
    ' [SCAFFOLD] The table's unique index on the corrected entry refuses a second cancellation
    '            (error 3022). That is an expected refusal: the answer says the entry has already
    '            been cancelled. Give the recovery its own statement; never re-run a variable that
    '            still holds the failed one.
    ' errHandler: the standard reporting handler.
End Function
```

### ReplaceTimeOffEntry — `Public Function` → `String`

```vba
Public Function ReplaceTimeOffEntry(ByVal lEntryID As Long, ByVal dNewHours As Double, _
                                    Optional ByVal dNewDate As Date = 0, _
                                    Optional ByVal sNote As String = "") As String
    ' [SCAFFOLD] Replaces a wrong Earned or Taken entry: the corrected entry first, naming the one it
    '            replaces, then the Correction that cancels the wrong one, the two together or not
    '            at all (Business Rule 10). dNewHours is positive. An Earned replacement covers the
    '            same period; a Taken replacement may move to dNewDate. The note is required where
    '            hours earned are replaced.
    Const PROC_NAME As String = "ReplaceTimeOffEntry"
    Dim ws       As DAO.Workspace
    Dim db       As DAO.Database
    Dim bInTrans As Boolean

    On Error GoTo errHandler
    ' [BUSINESS LOGIC #10] Before the transaction: read the entry (snapshot); refuse, with a
    '            sentence, one that does not exist, is not Earned or Taken, or (Earned) comes with no
    '            note, or has corrected hours of zero or less.
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.Databases(0)
    ws.BeginTrans
    bInTrans = True
    Set gdbTxn = db
    LockEmployee db, lEmp               ' held across BOTH writes, not released between them
    gbReplacing = True                  ' tells TimeOffEntryCheck a replacement is allowed

    ' [BUSINESS LOGIC #10] AppendEntry the corrected entry (naming lEntryID in the replaced column).
    CheckHook "betweenHalves"           ' [SCAFFOLD] a check copy places a competing row here
    ' [BUSINESS LOGIC #10,#7] AppendEntry the Correction that cancels lEntryID: on the old entry's
    '            date, with the old hours reversed, and a note (the person's, or one saying which
    '            entry replaced it).
    ' [SCAFFOLD] The Data Macro judges the Correction here, and must SEE the corrected entry written
    '            a moment ago. It reads through gdbTxn, which is why gdbTxn is set above.
    ws.CommitTrans
    bInTrans = False

Cleanup:
    On Error Resume Next
    gbReplacing = False                 ' cleared on every path, including failure
    Set gdbTxn = Nothing
    Set db = Nothing: Set ws = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] the standard reporting handler, plus: a refusal at either
    '            write rolls back BOTH, so the corrected entry is not left behind, and the answer
    '            says the ledger holds what it held before. A refusal for time taken that the
    '            corrected hours no longer cover is the table's own sentence, passed through.
    If bInTrans Then ws.Rollback
    bInTrans = False
    ' >>> ExpectedRefusal branch as PostOnePeriod; else LogError MODULE_NAME, PROC_NAME, Erl <<<
    Resume Cleanup
    Resume
End Function
```
