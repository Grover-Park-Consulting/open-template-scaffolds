---
step: 01-checks
title: "How you validate the template's output"
platform_facts: []
route: both
when: before you draft the design; every check is run when the build is finished
---

**Who reads this:** the developer whose database this is. The AI assistant designs to pass every
check below and runs every one of them once the build is finished.

### How you validate the template's output

**Before you trust this build: confirm every one of these ran, not just the ones the AI mentions
running.** See `README.md`, "Check that every validation check ran."

Do these on a copy, not on the live application. You need your front end, its data file, and
somewhere else to move the data file to. **Validate by asking: does this happen or not?**

1. **It opens.**
   - Open the front end the way anyone would
   - Confirm your startup form appears and shows data
2. **It notices the data file is gone.**
   - Close the front end
   - Move the data file to another folder, or rename it
   - Open the front end
   - Confirm it tells you the data file cannot be reached, names the place it looked, and offers to
     let you find it
3. **It reconnects to the file you choose.**
   - Point it at the data file in its new place
   - Confirm your startup form opens and shows data
   - Open the linked-table manager on the External Data ribbon and confirm every linked table now
     names the new location — this is what shows the links themselves moved, rather than the
     application working around them
4. **It stays reconnected.**
   - Close the front end and open it again
   - Confirm it goes straight to your startup form without asking — the new location was written
     into the links for good, not held for one session
5. **It refuses the wrong file.**
   - Close the front end and move the data file away again
   - Open the front end and, when it asks, choose a different Access file — a brand-new empty one is
     easiest
   - Confirm it refuses that file, tells you nothing was changed, and your startup form does not open
   - Put the data file back where it was in check 3 and open the front end again
   - Confirm it works — that is what "nothing was changed" meant
6. **It stops when nobody chooses a file.**
   - Move the data file away once more
   - Open the front end and cancel instead of choosing
   - Confirm it tells you the application cannot continue, and your startup form does not open
7. **A record survives the data file being unusable.**
   - Leave the data file where the application cannot reach it, or replace it with a file of the same
     name that is not a database at all
   - Open the front end
   - Confirm you are told there is a problem
   - If your standards keep a written record of errors, confirm that record is somewhere you can
     still read it — a file beside the front end, or on screen — and not inside the data file, which
     is the thing that is not working
8. **The folders.**
   - Delete the working folder the application creates beside the front end
   - Open the front end and confirm the folder is back
   - Rename the shared folder the application points at
   - Open the front end again and confirm it tells you that folder cannot be reached, and names it
   - Put the folder back
9. **It reconnects on its own when it already knows where the file went.**
   - Starting from a working application: close it, move the data file somewhere new, and then
     change the place the application remembers — the row in its own settings table, or the small
     file beside it — to the new location by hand
   - Open the front end
   - Confirm it goes straight to your startup form, asking nothing and saying nothing — this is the
     reconnection nobody ever sees happen
   - If your build keeps the remembered location in the code instead of in a table or a file, this
     check cannot be done without editing code, and that half of the promise goes unchecked
10. **A data file behind a password.**
    - Only if yours has one
    - Repeat checks 2 and 3 with it
    - Confirm it reconnects, and at no point is anyone asked for the password
