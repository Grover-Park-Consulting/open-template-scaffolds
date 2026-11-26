---
step: 02-rules-module
title: Write and import the rules module, modTimeOffRules
platform_facts: [vba-import-xml-entities, mcp-module-import, mcp-line-numbers, domain-function-transaction, row-lock-errors]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Procedures

Each procedure shows its scope, signature and an annotated skeleton. Line numbers are deliberately
absent (house-specific; see `error-handling.md`). Each handler is the one `error-handling.md`'s
table gives for the procedure's role, and its first comment names it. Three shapes appear:

- **Reporting shape**, for an entry point: it rolls back where it began a transaction, tells the
  person a refusal's sentence or logs a fault, and returns a sentence. Shown in full in
  `PostOnePeriod` and referenced thereafter as *the standard reporting handler*.
- **Propagating shape**, for a step whose caller depends on the outcome: copy `Err` into locals,
  local cleanup, then `Err.Raise` from the locals with `ModuleName.ProcName->` in front of the
  source, no report. `AccrualHours` takes it, because the posting code acts on its answer.
- **Refusing shape**, for `TimeOffEntryCheck`, which the Data Macro calls: it never raises and never
  shows anything. It returns `""` or a sentence, and on an unexpected error it logs quietly and
  returns a sentence carrying the reference, so the save is refused either way.

**The module header, `modTimeOffRules`**

```vba
Option Compare Database
Option Explicit

Private Const ModuleName As String = "modTimeOffRules"

' [SCAFFOLD] While ReplaceTimeOffEntry writes its two entries these hold the transaction's Database
'            and say so, so the checks the Data Macro runs can see the corrected entry that has not
'            been saved yet. Every posting procedure sets dbTxn after BeginTrans and clears it in
'            Cleanup, including when the posting fails. booReplacing is set only by ReplaceTimeOffEntry.
Public dbTxn      As DAO.Database
Public booReplacing As Boolean
```

### CheckHook — `Public Sub`

```vba
Public Sub CheckHook(ByVal strPoint As String)
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
    If dbTxn Is Nothing Then
        Set ReadDb = CurrentDb
    Else
        Set ReadDb = dbTxn
    End If
End Function
```

### ScalarValue — `Private Function` → `Variant`

```vba
Private Function ScalarValue(ByVal db As DAO.Database, ByVal strSQL As String) As Variant
    ' [SCAFFOLD] First column of the first row, read on the Database it is handed. Null where there is
    '            no row. Never a domain function. Used by every read below that wants one value.
    Dim rst As DAO.Recordset
    ScalarValue = Null
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rst.EOF Then ScalarValue = rst.Fields(0).Value
    rst.Close
    Set rst = Nothing
End Function
```

### LockEmployee — `Public Sub`

```vba
Public Sub LockEmployee(ByVal db As DAO.Database, ByVal lngEmp As Long)
    ' [SCAFFOLD] The write lock that makes two simultaneous entries for one employee take turns.
    '            Taken before the checks read anything and held until the transaction (or, for a
    '            single statement, the statement) finishes. This is what satisfies Business Rules
    '            4, 5, 9 and 10 against two people at once; checking first and then writing narrows
    '            the window and does not close it.
    '            It writes a column back to itself. The employee's Before Change Data Macro stamps
    '            only where a column really changed, so this stamps nothing.
    ' [BUSINESS LOGIC #4,#5,#9,#10] >>> UPDATE tblEmployee SET <a nullable column> = itself WHERE
    '            EmployeeID = lngEmp, executed on db with dbFailOnError, per query-style.md <<<
End Sub
```

### TimeOffBalanceOn — `Public Function` → `Double`

```vba
Public Function TimeOffBalanceOn(ByVal db As DAO.Database, ByVal lngEmp As Long, _
                                 ByVal lngCat As Long) As Double
    ' [BUSINESS LOGIC #2] The balance: the sum of TimeOffHours over the employee's entries in the
    '            category, 0 where there are none. Nothing stores it. Entries dated in the future
    '            count (the outcome-first template's house assumption).
    ' [SCAFFOLD] Read with ScalarValue on db, never DSum.
    ' >>> SELECT Sum(...) per query-style.md; Nz the result <<<
End Function
```

### TimeOffBalance — `Public Function` → `Double`

```vba
Public Function TimeOffBalance(ByVal lngEmp As Long, ByVal lngCat As Long) As Double
    ' [SCAFFOLD] The balance for a caller that has no Database of its own: reads through ReadDb, so
    '            it sees a transaction's own rows while one is open.
    TimeOffBalance = TimeOffBalanceOn(ReadDb(), lngEmp, lngCat)
End Function
```

### CompletedYears — `Public Function` → `Long`

```vba
Public Function CompletedYears(ByVal dtHire As Date, ByVal dtOn As Date) As Long
    ' [BUSINESS LOGIC #3] Whole years between dtHire and dtOn, not counting a year that has not yet
    '            been completed on dtOn's month and day, and never below 0.
    ' >>> year difference less one where the anniversary falls after dtOn <<<
End Function
```

### AccrualHours — `Public Function` → `Double`

```vba
Public Function AccrualHours(ByVal lngEmp As Long, ByVal lngCat As Long, ByVal dtPeriodStart As Date, _
                             ByRef strWhy As String) As Double
    ' [SCAFFOLD] Hours earned for one period; 0 with a reason in strWhy where none can be worked out.
    '            Nothing is ever posted as zero hours (Business Rule 3). Propagating shape.
    Const ProcName As String = "AccrualHours"
    Dim db     As DAO.Database
    Dim vHire  As Variant
    Dim vHours As Variant
    Dim lngYears As Long
    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    On Error GoTo errHandler
    strWhy = ""
    Set db = ReadDb()

    ' [BUSINESS LOGIC #3] vHire = the employee's HireDate. Where there is none, strWhy says there is no
    '            such employee and the procedure ends.
    ' >>> read via ScalarValue, per query-style.md <<<
    lngYears = CompletedYears(CDate(vHire), dtPeriodStart)

    ' [BUSINESS LOGIC #3] vHours = HoursPerPeriod of the schedule row for the category with the
    '            greatest MinYearsOfService not above lngYears. Read from tblAccrualSchedule; no figure
    '            is typed into code. Where no row applies, strWhy names the category and the completed
    '            years, and the result stays 0.
    ' >>> SELECT TOP 1 ... ORDER BY MinYearsOfService DESC, per query-style.md <<<

Cleanup:
    Set db = Nothing
    Exit Function

errHandler:
    ' [STANDARDS — error-handling.md] propagating shape: copy Err, local cleanup only, then pass it on
    '            with this procedure named in the source.
    '            Adapted from AccessCodeLib (access-codelib.net), data/dao/DaoHandler.cls, BSD 3-Clause License.
    lngErrNumber = Err.Number
    strErrSource = Err.Source
    strErrDescription = Err.Description
    Set db = Nothing
    Err.Raise lngErrNumber, ModuleName & "." & ProcName & "->" & strErrSource, strErrDescription
End Function
```

### CountEarnedForPeriod — `Public Function` → `Long`

```vba
Public Function CountEarnedForPeriod(ByVal db As DAO.Database, ByVal lngEmp As Long, _
                                     ByVal lngCat As Long, ByVal dtPeriod As Date) As Long
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
    Const ProcName As String = "TimeOffEntryCheck"
    Dim db     As DAO.Database
    Dim rst     As DAO.Recordset
    Dim strMsg   As String
    Dim strReason As String
    Dim strActive As String
    Dim dblEffect As Double
    Dim lngErrNumber      As Long
    Dim strErrSource      As String
    Dim strErrDescription As String

    Dim lngRef As Long
    On Error GoTo errHandler
    If IsNull(vEmp) Or IsNull(vCat) Or IsNull(vReason) Or IsNull(vDate) Or IsNull(vHours) Then GoTo Cleanup
    Set db = ReadDb()

    ' [SCAFFOLD] The lock comes before the first read, so two entries for one employee take turns.
    LockEmployee db, CLng(vEmp)
    ' >>> strReason = the reason's name, read via ScalarValue on db <<<

    ' The order below is the order the build was proven in. Each refusal sets strMsg and goes to
    ' Cleanup; the first one that applies wins.

    ' [BUSINESS LOGIC #6] Dated after the employee's InactiveDate: refused, naming the date.

    ' [BUSINESS LOGIC #8,#7,#10] What each reason must and must not carry: Earned has positive hours
    '            and an accrual period and no corrected entry; Taken has negative hours and no
    '            accrual period and no corrected entry; Correction names the entry it cancels, has no
    '            accrual period and no replaced entry. Any other reason is refused.

    ' [BUSINESS LOGIC #7] A Correction: the entry named exists; belongs to the same employee and
    '            category; is not itself a Correction; is not already cancelled; carries hours equal
    '            and opposite; and, where it cancels hours earned, carries a note.

    ' [BUSINESS LOGIC #10] An entry that names one it replaces: refused unless booReplacing is True
    '            (only ReplaceTimeOffEntry sets it, because it writes the cancellation in the same
    '            transaction). Where it is True: the entry named exists, is the same employee and
    '            category and reason, is still active, has not been replaced, and is not a
    '            Correction; an Earned replacement covers the same period.

    ' [SCAFFOLD] The entries that count against the one-per-period and one-per-day rules are the
    '            ACTIVE ones, meaning no Correction cancels them, and the entry being replaced is
    '            left out so it does not count against its own replacement. Build that as one
    '            condition fragment (strActive) and add it to both counts below.
    ' >>> strActive = " AND e.TimeOffEntryID NOT IN (...Corrects...)" [+ " AND e.TimeOffEntryID <> " & replaced id] <<<

    ' [BUSINESS LOGIC #4] An Earned entry where an active Earned entry already exists for the
    '            employee, category and period: refused, saying how to change it (replace it).
    ' [BUSINESS LOGIC #9] A Taken entry where an active Taken entry already exists for the employee
    '            and day, whatever the category: refused, saying how to change it.

    ' [BUSINESS LOGIC #5] dblEffect = the entry's own hours, except that a replacement counts for its
    '            hours less the hours of the entry it replaces. Where dblEffect is negative and the
    '            balance plus dblEffect is below zero: refused, with wording by case (a Correction of
    '            hours earned; a replacement of hours earned that leaves time taken uncovered; time
    '            taken beyond what is available, naming the hours available). Read the balance with
    '            TimeOffBalanceOn on db. Exactly zero is accepted.

    ' [SCAFFOLD] Every rule passed. A check copy places a competing row here.
    CheckHook "afterCheck:" & strReason

Cleanup:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Set db = Nothing
    TimeOffEntryCheck = strMsg
    Exit Function

errHandler:
    ' [STANDARDS - error-handling.md] refusing shape: a Data Macro calls this, so it never
    '            raises and never shows anything. An unexpected error refuses the save with a
    '            sentence carrying the log reference; the macro raises it, so the person sees it in
    '            Access's own dialog, or the posting procedure receives it as 3939 and tells it.
    '            Option 2 has no log: the sentence carries Err.Description instead.
    lngRef = LogError(ModuleName, ProcName, Erl, booShow:=False)
    TimeOffEntryCheck = "Not saved: the entry could not be checked. If it keeps happening, " & _
                        "quote reference " & lngRef & "."
    Resume Cleanup
    Resume
End Function
```
