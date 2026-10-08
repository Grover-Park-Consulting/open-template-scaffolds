---
step: 01-checks
title: How you validate the template's output
platform_facts: [row-lock-errors, domain-function-transaction]
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

**Do these on a copy, and stop if you do not have one.** The checks add, change and try to delete
records in your own tables, which is what they test, and there is no version of them that leaves
your data alone. Make the copy once the build is finished, run every check on it, and keep your
working file out of it. Where your database is split into two files, copy both and keep them
together.

1. **A posted entry cannot be changed through an ordinary route.**
   - Post an entry, then try to change its hours or date the way anyone ordinarily would: an edit
     in the table, a query that updates it, a form of your own
   - Confirm each change is refused
2. **A posted entry cannot be changed or deleted directly in the table: only if you chose the Data
   Macro route for Business Rule 1.**
   - Open the table directly, change a posted entry's hours, and separately try to delete one
   - Confirm both are refused
   - If you chose the VBA route instead, this check is not expected to pass. That is the disclosed
     trade-off of that route, confirmed by *Business Rule 1*'s own text above, not a failure to
     report as one
3. **The balance is the sum of the entries, and nothing stores it.**
   - Post an Earned entry of 8 hours and a Taken entry of 3 hours
   - Confirm the balance reads 5
   - Confirm no table has a balance column
4. **A Taken entry dated in the future already reduces the balance.**
   - Post a Taken entry dated next month
   - Confirm the balance reflects it now
5. **Hours earned follow length of service.**
   - Set up a schedule with a band at 0 years and another at 5 years, with different hours
   - Post accruals for an employee with two completed years and one with six
   - Confirm each was paid from the right band
6. **A missing schedule row is reported, never posted as zero.**
   - Post an accrual for a category with no schedule row at all
   - Confirm nothing was posted and the gap is visibly reported
7. **Posting the same period twice leaves one active entry.**
   - Post an accrual for an employee, category and period, then post it again, and then try to add a
     second Earned entry for that period by a route that does not go through the posting (a direct
     insert into the table)
   - Confirm exactly one active Earned entry exists, and that the direct insert was refused with a
     sentence saying the period already has one
8. **Two posting runs at the same moment leave one entry.**
   - **Pick an employee and period with no Earned entry yet**
   - Have two runs reach it at the same moment. Where that cannot be done with two people or two
     sessions, open a second, independent connection to the file and use it to insert the competing
     Earned entry between the first run's own check and its own write. The competing entry has to
     land after the first run has made its check; before that, the first run simply finds the entry
     already there, which any build handles and which proves nothing
   - Confirm exactly one active entry exists and the posting run finishes without an error shown to
     anyone. The competing writer ends one of three ways, and each is acceptable: it is a posting
     run that finds the entry already there; it is a posting run held back until the first
     finishes, which then finds it; or it is an entry inserted directly, which the table refuses
     with a plain sentence. Record which of the three you saw
9. **Time taken within the balance is accepted; time taken beyond it is refused.**
   - With 8 hours, post a Taken entry of 8 hours, then separately one of 8.25
   - Confirm the first is accepted, leaving 0, and the second is refused with the available hours
     named. Repeat the refused attempt directly in the table if you chose the Data Macro route for
     the posting checks
   - Then post an Earned entry, spend all of it, and try to cancel the Earned entry with a
     Correction and a note. Confirm that is refused, and accepted once the time taken against it has been
     cancelled
10. **Two entries that fit alone but not together never both go in.**
    - With 8 hours, have two entries of 8 hours each reach the table at the same moment, using a
      second independent connection as in check 8 where needed. Place the competing entry after the
      first entry's balance has been read. Before that, the balance already includes it, so any
      build refuses the first entry correctly, with or without protection against two at once, and
      the check proves nothing
    - Confirm one is posted, the other is refused, and the balance is 0, not negative
11. **An entry after the `InactiveDate` is refused.**
    - Give an employee an `InactiveDate`, then try an entry dated the day after, and one dated on
      that date
    - Confirm the first is refused and the second is accepted
12. **A correction cancels one entry and restores the balance.**
    - Post a Taken entry, then a Correction naming it with equal and opposite hours
    - Confirm the balance is back where it was and the original entry is unchanged
13. **A bad correction is refused.**
    - Try a Correction that names nothing, one naming an entry that does not exist, one naming
      another employee's entry, one naming an entry in a different category, one naming another
      Correction, one with the wrong hours, one naming an entry already cancelled, and one that
      cancels hours earned and carries no note
    - Confirm each is refused
14. **Wrong signs, zero hours and part quarter hours are refused.**
    - Try an Earned entry with negative hours, a Taken entry with positive hours, any entry with
      zero hours, and an entry of 4.1 hours
    - Confirm each is refused, and that an entry of 4.25 hours is accepted
15. **Two categories cannot share a day.**
    - Post a Taken entry for one category on a date, then try Taken for a different category on the
      same date, and then the same category again on that date
    - Confirm both are refused
16. **A cancelled day can be booked again, and other days are unaffected.**
    - Cancel the Taken entry from check 15 with a Correction, then book that date again
    - Confirm it is accepted, and that an entry for the next day was never in question
17. **Two entries for one day at the same moment never both go in.**
    - Using a second independent connection as in check 8, make two Taken entries for the same
      employee and date reach the table together
    - Confirm one is posted and the other refused
18. **The audit columns still stamp.**
    - Insert an entry
    - Confirm `CreatedDate` and `CreatedBy` are filled, as before this template attached anything
19. **Running the build again does not duplicate anything.**
    - Run whatever attaches the Data Macro and the functions a second time
    - Confirm there is still exactly one copy of each
20. **When a build is blocked, it stops rather than leaving things half-finished.**
    - Leave a table or form open in the database
    - Start the build
    - Confirm it stops, names what is open, and changes nothing
21. **A missed period is caught up, and posting again adds nothing.**
    - For an employee with an annual category, post nothing for the last two periods that have
      started, then post every period that is due
    - Confirm both missed periods were posted and none was posted twice
    - Post every period that is due again and confirm it posted nothing and reported the periods
      as already there
    - Cancel one posted period with a Correction and a note, post every period that is due, and
      confirm that period was **not** posted again
22. **Counting starts where the accrual start date says.**
    - Give an employee a `HireDate` several years back and an `AccrualStartDate` of this year
    - Post every period that is due
    - Confirm no period before the `AccrualStartDate` was posted. Then clear the `AccrualStartDate`
      on a second employee with the same `HireDate` and confirm periods from the `HireDate` onward
      were posted
23. **A period that cannot be posted is reported, and any trigger you chose fires.**
    - Add a category with no schedule row and post every period that is due
    - Confirm the answer counts the periods it could not post and names the first with its reason
    - Where you chose to have hours posted when the database opens or on a schedule, make a period
      due, then open the database (or let the scheduled run happen) as a person would, and confirm
      the period was posted and that opening it a second time posted nothing more. If you chose to
      run it yourself, this part does not apply
24. **A wrong accrual is replaced by a right one when the balance allows.**
    - Post an accrual of 96 hours and take 40 against it. Replace the 96 with 64
    - Confirm the corrected 64 and the cancellation of the 96 are both in the ledger, the balance
      went from 56 to 24, the entry replaced is unchanged, and the new entry names it
    - Confirm that during the replacement neither the two entries for one period nor the cancellation
      was refused for that reason
25. **A replacement that would leave time taken uncovered is refused as a whole.**
    - With 96 earned and 80 taken, try to replace the 96 with 64
    - Confirm the whole replacement is refused, the person is told that 16 hours of time taken are
      not covered, and afterwards the ledger holds exactly the entries it held before: no corrected
      entry was left behind and the balance is unchanged
26. **A wrong accrual is replaced by a larger one.**
    - Replace a 64-hour accrual that has 40 hours taken against it with one of 96
    - Confirm the balance went from 24 to 56 and both entries are in the ledger
27. **Time taken on the wrong day is replaced, even when it used the whole balance.**
    - With 8 hours earned, take 8 on the wrong day, then replace it with 8 on the right day
    - Confirm it is accepted, the balance is 0, and the day it was taken on is the new one. A day
      that is already booked, or the wrong entry's own day, is not in the way
    - Then replace a Taken entry of 4 hours with one of 8 hours on the same day, with enough
      balance, and confirm it is accepted; with too little, confirm it is refused as a whole
28. **Hours earned are taken away only with a note, and only when the balance allows.**
    - With 96 earned and nothing taken, cancel the 96 with a Correction and a note
    - Confirm it is accepted and the balance is 0. Then post every period that is due and confirm the
      period was not posted again
    - With 96 earned and 80 taken, try the same and confirm it is refused. Try it with no note and
      confirm that is refused too
29. **A bad replacement is refused.**
    - Try replacements that each break one requirement: one that names an entry that does not
      exist, another employee's entry, an entry in a different category, a Correction, an entry already
      cancelled, an entry already replaced, an Earned replacement for a different period, a Taken
      replacement of an Earned entry (different reason), a Taken replacement that books a day
      already booked by another entry, and a replacement inserted directly into the table on its own
    - Confirm each is refused with a sentence, and that none left an entry behind
30. **A replacement that fails partway leaves nothing behind.**
    - Cause a failure after the corrected entry has been written and before the cancellation, other
      than the balance refusal in check 25. For example, make the cancellation fail with an error of
      its own. Force the failure in a procedure your own posting code calls between the two halves,
      not inside the function a Data Macro calls: an error raised there reaches the engine as an
      unhandled error, and in a development session it opens a dialog that stops the run
    - Confirm no corrected entry was left behind, the entry replaced is still active, and the
      balance is what it was
31. **Nothing gets between the two halves of a replacement.**
    - Give an employee 100 hours earned and 10 taken, then replace the 100 earned with 64. Using a
      second independent connection as in check 8, insert a Taken entry of 60 hours for the same
      employee after the replacement has made its last check on the cancellation and before it
      finishes. Size it so that lack of time off cannot be the reason it is refused: 60 hours fits
      the 90 available before the replacement but not the 54 left after it. A larger entry is
      refused for lack of time off whatever else happens, so any build passes and the check proves
      nothing
    - Confirm the competing entry is refused or held back until the replacement has finished, and
      that the balance never went below zero
    - If you run this once with the protection against two at once switched off, expect the
      competing entry to be refused anyway: the replacement's own uncommitted rows are locked by
      the database engine until it finishes. The check then shows that nothing gets between the two
      halves, not which of the two protections stopped it. Say so in the build record
32. **An entry is replaced once, and cancelled once, even at the same moment.**
    - Using a second independent connection, post two replacements of the same entry, and two
      cancellations of the same entry, so that each pair reaches the table together
    - Confirm one of each pair is accepted and the other refused
