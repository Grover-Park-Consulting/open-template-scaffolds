---
template: time-off-ledger-scaffold
title: Time Off Ledger — VBA Scaffold
domain: time-off
type: vba-scaffold
version: 0.2.1
status: draft
implements: time-off-ledger-schema
requires_tables:
  - tlkpTimeOffCategory
  - tlkpEntryReason
  - tblEmployee
  - tblAccrualSchedule
  - tblTimeOffEntry
standards_layer:
  - error-handling
  - query-style
  - naming-conventions
  - design-principles
target_module: modTimeOffRules, modTimeOffPosting
new_procedures:
  - CheckHook
  - ReadDb
  - ScalarValue
  - LockEmployee
  - TimeOffBalanceOn
  - TimeOffBalance
  - CompletedYears
  - AccrualHours
  - CountEarnedForPeriod
  - TimeOffEntryCheck
  - ExpectedRefusal
  - AppendEntry
  - PostOnePeriod
  - PostEarnedForPeriod
  - PostEarnedOnPurpose
  - PostDuePeriods
  - PostTimeTaken
  - CancelTimeOffEntry
  - ReplaceTimeOffEntry
platform_facts: [dao-table-build, data-macro-rules, domain-function-transaction, row-lock-errors, sql-insert-truncation, recordset-append-crash, open-existing-startup, mcp-module-import, mcp-run-procedure, mcp-line-numbers, mcp-file-release, app-startup-autoexec]
warnings:
  - These procedures write real rows into the ledger, and the ledger is never edited or deleted once
    written, so a run that goes wrong cannot be undone by deleting what it wrote. A build against a
    database in real use is preceded by a backup copy of the file, and the developer is asked for one
    before anything is changed.
  - The function the Data Macro calls, TimeOffEntryCheck, and everything it calls must exist in the
    back end and in every front end. In a split database, two front ends holding different copies
    hold two people to different rules. Nothing keeps the copies in step, so a changed function has
    to be imported everywhere.
  - The checks, the lock and the replacement all depend on one rule about where a read comes from. A
    read that must see work a transaction has just done is made on the Database that transaction
    was begun on, never through a domain function. A build that breaks this compiles, runs, and
    returns a wrong answer without raising anything. See the transaction guard in error-handling.md.
  - CheckHook is empty in the delivered database and only a check copy fills it in. Delivering it
    filled in would run test code on every posting.
related:
  - "time-off-ledger-outcome-first — the same ledger, produced by an open route in place of these skeletons. The promise is identical and each template says so. It carries the checks this one is validated against."
  - "error-logging-scaffold — builds the shared logger these procedures call when an unexpected error occurs. Build it first if the database has none."
---

# Time Off Ledger — VBA Scaffold

**Status last determined:** 2026-10-05.

**Who reads this:** the AI assistant, building this alongside the developer who asked for it.

**If that developer is you:** this file holds the decisions already made on your behalf. You do not
have to read it to use the template.

## Intent

Realize the logic that `time-off-ledger-schema` defers to code: the checks a new entry must pass
(Business Rules 4 to 10), the balance (Rule 2), hours earned by length of service (Rule 3), and the
six ways entries get posted. This scaffold supplies the **procedure skeletons**: signatures,
recordset plumbing, control flow, the transaction and lock structure, and the error-handling frame.
It does not write the domain logic. Each procedure marks where that goes, sourced from the table
template's numbered Business Rules. House style is deferred to the standards layer.

Three layers, kept distinct throughout:

- **`[SCAFFOLD]`** — structure provided here.
- **`[STANDARDS]`** — house style, deferred to the standards layer.
- **`[BUSINESS LOGIC]`** — the domain rule you fill in, sourced from the table template's Business
  Rules.

**What this scaffold does not build.** The tables (`time-off-ledger-schema`), the Data Macro that
keeps a posted entry from changing and calls `TimeOffEntryCheck` (Business Rule 1 and the checks;
`time-off-ledger-outcome-first` states the routes and `templates/_materialization.md` has the
shape), and the open-time macro for posting. This is the VBA those pieces call.

## Prerequisites

| Object | Role |
|---|---|
| `time-off-ledger-schema` tables | The five tables the code runs against |
| A Data Macro on `tblTimeOffEntry` calling `TimeOffEntryCheck` | Where the checks run, on every route into the table |
| A central error logger | `error-handling.md`; `error-logging-scaffold` builds one |

### Ask before building

**Where this is a database already in use, ask for a backup copy before changing anything.** Two
questions, in this order:

1. *"Is this a database you already use, or a new one you're trying this out on?"*
2. Where it is one they already use: *"Make a copy of the file before I start. Say when it's done and
   the build continues; say there's no copy and it stops."*

**Stop and build nothing if they say there is no copy.** A new database needs no copy.

**One choice belongs to the developer and is never inferred:** when earned hours are posted. It is
item 7 under *Information and conditions you need to supply* in `time-off-ledger-outcome-first`.
Ask it as that template says.

### Where the code goes: it depends on how the database is built

**A single-file database.** Everything goes in that one file. There is nothing more to decide.

**A split database** (one file holding the tables, the back end, and a copy of a second file for
each person, the front end). Where the code goes follows from two facts:

- The Data Macro on `tblTimeOffEntry` calls `TimeOffEntryCheck`, and a Data Macro can call only a
  function that exists in the file where the change is made. So the rules code (the procedures up to
  and including `TimeOffEntryCheck`) must be in the back end.
- The posting procedures are called from whatever each person uses to post, so they go in each front
  end. They call the rules code, so each front end holds the rules code as well.

So the rules code is in the back end and in every front end, and the posting code is in the front
ends. Nothing keeps the copies in step: where the rules code changes, import it everywhere, or two
people are held to different rules.

## Validating the build

**To the AI assistant.** This template and `time-off-ledger-outcome-first` promise the identical
result by different routes. That promise is what is checked, so there is one checklist for both.

**Run every entry under that template's "How you validate the template's output" against this build,
on a copy, as it requires, and report against the same numbered list in the build record.** Do not
devise a list from reading the procedures below: that tests what this code does, not what the
developer was promised. Name the list you ran.

Seven things follow from this being skeletons rather than an open route.

- **Drive every check through the six posting procedures**, never by inserting into
  `tblTimeOffEntry` by hand, except where a check says to make a direct insert. A direct insert
  bypasses the lock and the transaction, which is most of what is being checked.
- **Compile the VBA project before running any check, and record that you did.** Code that will not
  compile fails every check at once with the wrong cause attached to each.
- **Leave `CheckHook` empty in the delivered database.** The checks that need a competing writer
  (outcome-first checks 8, 10, 17, 31 and 32) are run on a check copy whose `CheckHook` is filled in
  to place the competing row at the point the check names: `afterCheck:Earned`, `afterCheck:Taken`,
  `afterCheck:Correction` after a check has passed, and `betweenHalves` between the two writes of a
  replacement. Record what the copy changed.
- **The competing writer must be a second process you started and identified.** Observed in the
  2026-10-05 trial: a second connection opened inside the same process is not independent, because
  `CurrentDb` belongs to the default workspace the first writer is using, and a file that an Access
  MCP server holds open exclusively cannot be opened by a second process at all. Close that session,
  start both processes yourself on a check copy, and identify each by process ID, as
  `CLAUDE.md` requires. Where you cannot, record the check as not tested.
- **Run the same checks once with `LockEmployee` emptied**, on a copy, for the checks that force a
  collision, so that you know each can fail. Record that it did.
- **Check 30 forces a failure between the two writes of a replacement.** Force it from `CheckHook`
  at `betweenHalves`, in `modTimeOffPosting`, never inside `TimeOffEntryCheck`. An error raised
  there reaches the engine as an unhandled error and can stop the run with a dialog.
- **Checks 2 and 23 carry their own conditions** (the route chosen for Business Rule 1, and whether
  hours are posted by the developer, on open, or on a schedule). State them in the entry.

## Procedures

Each procedure shows its scope, signature and an annotated skeleton. Line numbers are deliberately
absent (house-specific; see `error-handling.md`). Two error shapes appear, and each procedure uses
exactly one:

- **Reporting shape**, for a procedure that is the end of the line: it rolls back where it began a
  transaction, logs, and returns a sentence. Shown in full in `PostOnePeriod` and referenced
  thereafter as *the standard reporting handler*.
- **Propagating shape**, for a procedure whose caller depends on the outcome: local cleanup, then
  `Err.Raise Err.Number, Err.Source, Err.Description`, no report. `AccrualHours` and
  `TimeOffEntryCheck` take it, because the posting code and the Data Macro both act on their answer.

**The module header, `modTimeOffRules`**

```vba
Option Compare Database
Option Explicit

Private Const MODULE_NAME As String = "modTimeOffRules"

' [SCAFFOLD] While ReplaceTimeOffEntry writes its two entries these hold the transaction's Database
'            and say so, so the checks the Data Macro runs can see the corrected entry that has not
'            been saved yet. Every posting procedure sets gdbTxn after BeginTrans and clears it in
'            Cleanup, including when the posting fails. gbReplacing is set only by ReplaceTimeOffEntry.
Public gdbTxn      As DAO.Database
Public gbReplacing As Boolean
```

### CheckHook — `Public Sub`

```vba
Public Sub CheckHook(ByVal sPoint As String)
    ' [SCAFFOLD] Empty on purpose. The delivered database keeps it empty. A check copy fills it in to
    '            place a competing row at the moment a check needs one. Called after every posting
    '            check has passed and before the write (TimeOffEntryCheck), and between the two
    '            writes of a replacement (ReplaceTimeOffEntry). A competing row placed any earlier is
    '            read by the first writer, and the check proves nothing.
End Sub
```

### ReadDb — `Private Function` → `DAO.Database`

```vba
Private Function ReadDb() As DAO.Database
    ' [SCAFFOLD] The Database every read in this module goes through: the transaction's own while a
    '            posting has one open, otherwise the current one. A domain function would not see the
    '            transaction's uncommitted rows.
    If gdbTxn Is Nothing Then
        Set ReadDb = CurrentDb
    Else
        Set ReadDb = gdbTxn
    End If
End Function
```

### ScalarValue — `Private Function` → `Variant`

```vba
Private Function ScalarValue(ByVal db As DAO.Database, ByVal sSql As String) As Variant
    ' [SCAFFOLD] First column of the first row, read on the Database it is handed. Null where there is
    '            no row. Never a domain function. Used by every read below that wants one value.
    Dim rs As DAO.Recordset
    ScalarValue = Null
    Set rs = db.OpenRecordset(sSql, dbOpenSnapshot)
    If Not rs.EOF Then ScalarValue = rs.Fields(0).Value
    rs.Close
    Set rs = Nothing
End Function
```

### LockEmployee — `Public Sub`

```vba
Public Sub LockEmployee(ByVal db As DAO.Database, ByVal lEmp As Long)
    ' [SCAFFOLD] The write lock that makes two simultaneous entries for one employee take turns.
    '            Taken before the checks read anything and held until the transaction (or, for a
    '            single statement, the statement) finishes. This is what satisfies Business Rules
    '            4, 5, 9 and 10 against two people at once; checking first and then writing narrows
    '            the window and does not close it.
    '            It writes a column back to itself. The employee's Before Change Data Macro stamps
    '            only where a column really changed, so this stamps nothing.
    ' [BUSINESS LOGIC #4,#5,#9,#10] >>> UPDATE tblEmployee SET <a nullable column> = itself WHERE
    '            EmployeeID = lEmp, executed on db with dbFailOnError, per query-style.md <<<
End Sub
```

### TimeOffBalanceOn — `Public Function` → `Double`

```vba
Public Function TimeOffBalanceOn(ByVal db As DAO.Database, ByVal lEmp As Long, _
                                 ByVal lCat As Long) As Double
    ' [BUSINESS LOGIC #2] The balance: the sum of TimeOffHours over the employee's entries in the
    '            category, 0 where there are none. Nothing stores it. Entries dated in the future
    '            count (the outcome-first template's house assumption).
    ' [SCAFFOLD] Read with ScalarValue on db, never DSum.
    ' >>> SELECT Sum(...) per query-style.md; Nz the result <<<
End Function
```

### TimeOffBalance — `Public Function` → `Double`

```vba
Public Function TimeOffBalance(ByVal lEmp As Long, ByVal lCat As Long) As Double
    ' [SCAFFOLD] The balance for a caller that has no Database of its own: reads through ReadDb, so
    '            it sees a transaction's own rows while one is open.
    TimeOffBalance = TimeOffBalanceOn(ReadDb(), lEmp, lCat)
End Function
```

### CompletedYears — `Public Function` → `Long`

```vba
Public Function CompletedYears(ByVal dHire As Date, ByVal dOn As Date) As Long
    ' [BUSINESS LOGIC #3] Whole years between dHire and dOn, not counting a year that has not yet
    '            been completed on dOn's month and day, and never below 0.
    ' >>> year difference less one where the anniversary falls after dOn <<<
End Function
```

### AccrualHours — `Public Function` → `Double`

```vba
Public Function AccrualHours(ByVal lEmp As Long, ByVal lCat As Long, ByVal dPeriodStart As Date, _
                             ByRef sWhy As String) As Double
    ' [SCAFFOLD] Hours earned for one period; 0 with a reason in sWhy where none can be worked out.
    '            Nothing is ever posted as zero hours (Business Rule 3). Propagating shape.
    Const PROC_NAME As String = "AccrualHours"
    Dim db     As DAO.Database
    Dim vHire  As Variant
    Dim vHours As Variant
    Dim lYears As Long

    On Error GoTo errHandler
    sWhy = ""
    Set db = ReadDb()

    ' [BUSINESS LOGIC #3] vHire = the employee's HireDate. Where there is none, sWhy says there is no
    '            such employee and the procedure ends.
    ' >>> read via ScalarValue, per query-style.md <<<
    lYears = CompletedYears(CDate(vHire), dPeriodStart)

    ' [BUSINESS LOGIC #3] vHours = HoursPerPeriod of the schedule row for the category with the
    '            greatest MinYearsOfService not above lYears. Read from tblAccrualSchedule; no figure
    '            is typed into code. Where no row applies, sWhy names the category and the completed
    '            years, and the result stays 0.
    ' >>> SELECT TOP 1 ... ORDER BY MinYearsOfService DESC, per query-style.md <<<

Cleanup:
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] propagating shape: local cleanup only, then pass it on.
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
```

### CountEarnedForPeriod — `Public Function` → `Long`

```vba
Public Function CountEarnedForPeriod(ByVal db As DAO.Database, ByVal lEmp As Long, _
                                     ByVal lCat As Long, ByVal dPeriod As Date) As Long
    ' [BUSINESS LOGIC #4] How many Earned entries the employee has for the category and period,
    '            cancelled or not. A period that already has one counts as posted when periods that
    '            are due are posted, so a cancellation made to take hours away is not quietly undone
    '            by the next run.
    ' >>> SELECT Count(*) via ScalarValue on db, per query-style.md <<<
End Function
```

### TimeOffEntryCheck — `Public Function` → `String`

```vba
Public Function TimeOffEntryCheck(ByVal vEmp As Variant, ByVal vCat As Variant, ByVal vReason As Variant, _
                                  ByVal vDate As Variant, ByVal vHours As Variant, ByVal vPeriod As Variant, _
                                  ByVal vCorrects As Variant, ByVal vReplaces As Variant, _
                                  ByVal vNote As Variant) As String
    ' [SCAFFOLD] Called by the Before Change Data Macro on tblTimeOffEntry for every new entry.
    '            Returns "" where the entry may be posted, otherwise ONE sentence saying what is wrong
    '            and what to do next: the hours available, the date already taken, the entry already
    '            cancelled. Raises nothing for a refusal. The macro raises the sentence, so the person
    '            sees this wording whoever posted the entry.
    '            No message box: someone is saving a row.
    '            Parameters are Variants because the macro passes Null for a column that is empty.
    Const PROC_NAME As String = "TimeOffEntryCheck"
    Dim db     As DAO.Database
    Dim rs     As DAO.Recordset
    Dim sMsg   As String
    Dim sReason As String
    Dim sActive As String
    Dim dEffect As Double

    On Error GoTo errHandler
    If IsNull(vEmp) Or IsNull(vCat) Or IsNull(vReason) Or IsNull(vDate) Or IsNull(vHours) Then GoTo Cleanup
    Set db = ReadDb()

    ' [SCAFFOLD] The lock comes before the first read, so two entries for one employee take turns.
    LockEmployee db, CLng(vEmp)
    ' >>> sReason = the reason's name, read via ScalarValue on db <<<

    ' The order below is the order the build was proven in. Each refusal sets sMsg and goes to
    ' Cleanup; the first one that applies wins.

    ' [BUSINESS LOGIC #6] Dated after the employee's InactiveDate: refused, naming the date.

    ' [BUSINESS LOGIC #8,#7,#10] What each reason must and must not carry: Earned has positive hours
    '            and an accrual period and no corrected entry; Taken has negative hours and no
    '            accrual period and no corrected entry; Correction names the entry it cancels, has no
    '            accrual period and no replaced entry. Any other reason is refused.

    ' [BUSINESS LOGIC #7] A Correction: the entry named exists; belongs to the same employee and
    '            category; is not itself a Correction; is not already cancelled; carries hours equal
    '            and opposite; and, where it cancels hours earned, carries a note.

    ' [BUSINESS LOGIC #10] An entry that names one it replaces: refused unless gbReplacing is True
    '            (only ReplaceTimeOffEntry sets it, because it writes the cancellation in the same
    '            transaction). Where it is True: the entry named exists, is the same employee and
    '            category and reason, is still active, has not been replaced, and is not a
    '            Correction; an Earned replacement covers the same period.

    ' [SCAFFOLD] The entries that count against the one-per-period and one-per-day rules are the
    '            ACTIVE ones, meaning no Correction cancels them, and the entry being replaced is
    '            left out so it does not count against its own replacement. Build that as one
    '            condition fragment (sActive) and add it to both counts below.
    ' >>> sActive = " AND e.TimeOffEntryID NOT IN (...Corrects...)" [+ " AND e.TimeOffEntryID <> " & replaced id] <<<

    ' [BUSINESS LOGIC #4] An Earned entry where an active Earned entry already exists for the
    '            employee, category and period: refused, saying how to change it (replace it).
    ' [BUSINESS LOGIC #9] A Taken entry where an active Taken entry already exists for the employee
    '            and day, whatever the category: refused, saying how to change it.

    ' [BUSINESS LOGIC #5] dEffect = the entry's own hours, except that a replacement counts for its
    '            hours less the hours of the entry it replaces. Where dEffect is negative and the
    '            balance plus dEffect is below zero: refused, with wording by case (a Correction of
    '            hours earned; a replacement of hours earned that leaves time taken uncovered; time
    '            taken beyond what is available, naming the hours available). Read the balance with
    '            TimeOffBalanceOn on db. Exactly zero is accepted.

    ' [SCAFFOLD] Every rule passed. A check copy places a competing row here.
    CheckHook "afterCheck:" & sReason

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    TimeOffEntryCheck = sMsg
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] propagating shape. An unexpected error here refuses the save:
    '            the macro raises it and Access shows its own message. Nothing is logged unless the
    '            build makes it so; the build record says which.
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
```

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

## Standards Layer

- **Error handling** — the `errHandler`/`Cleanup` structure, both shapes, the transaction guard and the
  line-number policy come from `error-handling.md`. A function a Data Macro calls reports a refusal
  to the macro and does not open a message box of its own. The posting procedures use the shared
  logger, never a message box, because nothing in a posting may interrupt a person with a dialog.
- **Query style** — every `>>> ... per query-style.md <<<` marker is SQL written to the house query
  standard.
- **Naming** — procedure, variable and parameter names follow `naming-conventions.md`.
- **Design principles** — how these procedures divide is already settled here; a build that merges or
  splits them keeps every rule above.

**To the AI assistant: once the build is reported finished, and only then, mention each entry under
`related` in the front matter** (`_template-schema.md` §7.1), one line per entry, what it is and why.
This is not part of the build, never a gate, and never read before this point.

## Extra Options

*Named optional extensions, none of them filled in for an engagement.*

- **Role-based accrual, going below zero, a stored balance, date-range entries, holidays, carry-over
  and expiry**, all named in the schema's own *Extra Options*. This scaffold builds none of them.
  Each changes which procedures read or write what, and the outcome-first checks describe the base
  build only. Re-check every read against the transaction guard if one is taken.
- **A schedule that posts by itself.** A Windows scheduled task opens the database, calls
  `PostDuePeriods`, and closes it. This scaffold builds what the task runs; the task is the
  developer's.

## Parked / future considerations (not in this design)

- **A procedure that finds the open-time macro and adds the posting to it** lives in
  `app-startup-scaffold` and `standards/startup-conventions.md`, not here.
