---
step: 02-rules-module
title: Write and import the rules module, modTimeOffRules
platform_facts: [vba-import-xml-entities, mcp-module-import, mcp-line-numbers, domain-function-transaction, row-lock-errors]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

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
