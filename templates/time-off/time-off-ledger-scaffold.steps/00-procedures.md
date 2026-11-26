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

### CheckHook

```vba
Public Sub CheckHook(ByVal strPoint As String)
```

Empty on purpose. The delivered database keeps it empty. A check copy fills it in to place a competing row at the moment a check needs one. Called after every posting check has passed and before the write (TimeOffEntryCheck), and between the two writes of a replacement (ReplaceTimeOffEntry). A competing row placed any earlier is read by the first writer, and the check proves nothing.

### ReadDb

```vba
Private Function ReadDb() As DAO.Database
```

The Database every read in this module goes through: the transaction's own while a posting has one open, otherwise the current one. A domain function would not see the transaction's uncommitted rows.

### ScalarValue

```vba
Private Function ScalarValue(ByVal db As DAO.Database, ByVal strSQL As String) As Variant
```

First column of the first row, read on the Database it is handed. Null where there is no row. Never a domain function. Used by every read in this module that wants one value.

### LockEmployee

```vba
Public Sub LockEmployee(ByVal db As DAO.Database, ByVal lngEmp As Long)
```

The write lock that makes two simultaneous entries for one employee take turns. Taken before the checks read anything and held until the transaction (or, for a single statement, the statement) finishes. This is what satisfies Business Rules 4, 5, 9 and 10 against two people at once; checking first and then writing narrows the window and does not close it. It writes a column back to itself. The employee's Before Change Data Macro stamps only where a column really changed, so this stamps nothing.

### TimeOffBalanceOn

```vba
Public Function TimeOffBalanceOn(ByVal db As DAO.Database, ByVal lngEmp As Long, _
                                 ByVal lngCat As Long) As Double
```

Read with ScalarValue on db, never DSum.

### TimeOffBalance

```vba
Public Function TimeOffBalance(ByVal lngEmp As Long, ByVal lngCat As Long) As Double
```

The balance for a caller that has no Database of its own: reads through ReadDb, so it sees a transaction's own rows while one is open.

### CompletedYears

```vba
Public Function CompletedYears(ByVal dtHire As Date, ByVal dtOn As Date) As Long
```

Whole years between dtHire and dtOn, not counting a year that has not yet been completed on dtOn's month and day, and never below 0.

### AccrualHours

```vba
Public Function AccrualHours(ByVal lngEmp As Long, ByVal lngCat As Long, ByVal dtPeriodStart As Date, _
                             ByRef strWhy As String) As Double
```

Hours earned for one period; 0 with a reason in strWhy where none can be worked out. Nothing is ever posted as zero hours (Business Rule 3). Propagating shape.

### CountEarnedForPeriod

```vba
Public Function CountEarnedForPeriod(ByVal db As DAO.Database, ByVal lngEmp As Long, _
                                     ByVal lngCat As Long, ByVal dtPeriod As Date) As Long
```

How many Earned entries the employee has for the category and period, cancelled or not. A period that already has one counts as posted when periods that are due are posted, so a cancellation made to take hours away is not quietly undone by the next run.

### TimeOffEntryCheck

```vba
Public Function TimeOffEntryCheck(ByVal vEmp As Variant, ByVal vCat As Variant, ByVal vReason As Variant, _
                                  ByVal vDate As Variant, ByVal vHours As Variant, ByVal vPeriod As Variant, _
                                  ByVal vCorrects As Variant, ByVal vReplaces As Variant, _
                                  ByVal vNote As Variant) As String
```

Called by the Before Change Data Macro on tblTimeOffEntry for every new entry. Returns "" where the entry may be posted, otherwise ONE sentence saying what is wrong and what to do next: the hours available, the date already taken, the entry already cancelled. Raises nothing for a refusal. The macro raises the sentence, so the person sees this wording whoever posted the entry. No message box: someone is saving a row. Parameters are Variants because the macro passes Null for a column that is empty.

### ExpectedRefusal

```vba
Private Function ExpectedRefusal(ByVal lngErr As Long) As Boolean
```

Errors that mean "the table said no", not "something broke": the Data Macro's own refusal (3939) and the engine's refusals of a duplicate, a missing parent, a bad value, or a row held by someone else. A refusal is reported to the person as a sentence and is not logged. Anything else is a failure nobody planned for and is logged. Never mistake the second kind for the first.

### AppendEntry

```vba
Private Function AppendEntry(ByVal db As DAO.Database, ByVal lngEmp As Long, ByVal lngCat As Long, _
                             ByVal strReason As String, ByVal dtDate As Date, ByVal dblHours As Double, _
                             ByVal vPeriod As Variant, ByVal vCorrects As Variant, _
                             ByVal vReplaces As Variant, ByVal strNote As String) As Long
```

Appends one entry and returns its TimeOffEntryID. Raises whatever the table or its Data Macro raises; the caller decides what that means. Written by a parameterised SQL INSERT on db, never a recordset append: a recordset append from code crashed Access while the Data Macro's function read through CurrentDb (_materialization.md, "A recordset append from code can crash Access..."). Parameters carry every value, so an apostrophe in a note needs no quoting.

### PostOnePeriod

```vba
Private Function PostOnePeriod(ByVal lngEmp As Long, ByVal lngCat As Long, ByVal dtPeriod As Date, _
                               ByVal booOnPurpose As Boolean, ByVal strNote As String, _
                               ByRef strMsg As String) As Long
```

Posts the hours earned for one period. Returns 1 = posted, 2 = already there, 3 = could not be posted (strMsg says why). Business Rule 4. The employee's lock is taken before the "already there?" look and held until the entry is written, so two runs reaching one period take turns and the second finds the first one's entry. Neither reports an error.

### PostEarnedForPeriod

```vba
Public Function PostEarnedForPeriod(ByVal lngEmp As Long, ByVal lngCat As Long, ByVal dtPeriodStart As Date, _
                                    Optional ByVal strNote As String = "") As String
```

Posts the hours earned for one period; a period that already has an Earned entry, cancelled or not, is left alone. A note longer than NOTE_MAX is refused here, with a sentence, before anything is written. Calls PostOnePeriod with booOnPurpose False; where it answers 2 the sentence begins "Nothing posted:". No dialog.

### PostEarnedOnPurpose

```vba
Public Function PostEarnedOnPurpose(ByVal lngEmp As Long, ByVal lngCat As Long, ByVal dtPeriodStart As Date, _
                                    Optional ByVal strNote As String = "") As String
```

How a person gives hours that a Correction took away (Business Rule 4). Calls PostOnePeriod with booOnPurpose True. The table still refuses a second active entry.

### PostDuePeriods

```vba
Public Function PostDuePeriods(Optional ByVal lngEmp As Long = 0) As String
```

Posts every accrual period that is due, for one employee (lngEmp) or all of them, and answers with one sentence: how many it posted, how many were already there, how many it could not post, and the first it could not post with its reason. A period it could not post is never skipped without being counted. Reporting shape.

### PostTimeTaken

```vba
Public Function PostTimeTaken(ByVal lngEmp As Long, ByVal lngCat As Long, ByVal dtDate As Date, _
                              ByVal dblHours As Double) As String
```

Records time taken for one day. dblHours is the amount taken, positive; it is stored negative. The checks (balance, one kind of time off per day, inactive date) run in the Data Macro, not here. One transaction on the default Workspace; dbTxn set after BeginTrans; LockEmployee first; AppendEntry a Taken entry; CommitTrans; the answer names the hours, category, employee, day, entry number and the balance now. Hours of zero or less are refused with a sentence before anything is written.

### CancelTimeOffEntry

```vba
Public Function CancelTimeOffEntry(ByVal lngEntryID As Long, ByVal strNote As String) As String
```

Cancels one entry with a Correction of equal and opposite hours. The entry cancelled is left exactly as it was. A Correction of hours earned needs a note, and the Data Macro refuses one without.

### ReplaceTimeOffEntry

```vba
Public Function ReplaceTimeOffEntry(ByVal lngEntryID As Long, ByVal dblNewHours As Double, _
                                    Optional ByVal dtNewDate As Date = 0, _
                                    Optional ByVal strNote As String = "") As String
```

Replaces a wrong Earned or Taken entry: the corrected entry first, naming the one it replaces, then the Correction that cancels the wrong one, the two together or not at all (Business Rule 10). dblNewHours is positive. An Earned replacement covers the same period; a Taken replacement may move to dtNewDate. The note is required where hours earned are replaced.
