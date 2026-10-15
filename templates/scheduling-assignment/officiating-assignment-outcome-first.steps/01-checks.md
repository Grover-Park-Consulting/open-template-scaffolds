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

**Do these on a copy, and stop if you do not have one.** The checks add, change, and delete records
in your own tables — that is what they test, and there is no version of them that leaves your data
alone. Make the copy once the build is finished, run every check on it, and keep your working file
out of it. Where your database is split into two files, copy both and keep them together.

1. **An assignment against an inactive official is refused through the normal route.**
   - Mark an official inactive
   - Try to assign that official to a game through your assignment procedure
   - Confirm the assignment is refused
2. **An assignment against an inactive official through a direct table edit — only if you chose the
   Data Macro route.**
   - With the same inactive official, add a row to `tblGameOfficial` directly, not through the
     assignment procedure
   - Confirm the write is refused
   - If you chose the VBA route instead, this check is not expected to pass — that's the disclosed
     trade-off of that route, confirmed by *Business Rule 3*'s own text above, not a failure to
     report as one
3. **A game can't be saved against itself.**
   - Try to save a game with the same team as both home and away, once through a form or the table
     directly and once through an append/import query
   - Confirm neither save succeeds
4. **A game with two different teams saves normally.**
   - Save a game with two distinct teams
   - Confirm it saves
5. **A game's play level moves with its home team.**
   - Note a team's play level and a game where that team is the home team
   - Change the team's play level
   - Confirm the game's shown play level reflects the change immediately, with no separate step
6. **The applicable pay rate resolves to the most recent effective row.**
   - Create two `tblPositionRate` rows for the same play level and position, with different
     `EffectiveDate` values, both on or before a game's date
   - Confirm the resolved rate is the one from the row with the later `EffectiveDate`, not the row
     added most recently or the higher of the two rates
7. **A missing rate is surfaced, never guessed.**
   - Create an assignment for a play level and position that has no `tblPositionRate` row at all
   - Confirm the absence is visibly reported — not a $0 rate presented as though it were a real one
8. **Game times out of order are refused.**
   - Try to save a game with `GameEnd` earlier than or equal to `GameStart`, whatever route you use
   - Confirm the save does not succeed
9. **A game with no end time saves without complaint.**
   - Save a game with `GameStart` set and `GameEnd` left empty
   - Confirm it saves
10. **A game with `GameEnd` after `GameStart` saves normally.**
    - Save a game with a valid, later `GameEnd`
    - Confirm it saves
11. **An official's age is correct and not stored.**
    - Compare two officials with different `BirthDate` values
    - Confirm wherever age is shown, it's correct for each, and confirm `tblOfficial` carries no
      separate age column that could drift out of step with `BirthDate`
12. **The shared photo folder is confirmed, never silently created locally.**
    - Point `tblAppSetting.OfficialPhotoFolder` at a shared path that does not currently exist
    - Try to pick a photo for an official
    - Confirm you are told the folder can't be reached — not shown a folder that was quietly created
      on your own machine
13. **Picking a photo copies it into the managed folder under a controlled name.**
    - Point the setting at a folder that exists, and pick a photo for an official
    - Confirm the file appears in that folder under a name that isn't the file's own original name,
      and that the official's record stores only that name
14. **Re-picking a photo replaces it cleanly.**
    - Pick a different photo for the same official
    - Confirm the official's record now points at the new file, and the old one isn't left behind
      under its old name with nothing pointing at it
15. **Changing the folder setting changes where the next photo lands.**
    - Change `tblAppSetting.OfficialPhotoFolder` to a different, valid shared folder
    - Pick a photo for an official
    - Confirm the file lands in the new folder, not the old one
16. **Running the build again does not duplicate anything.**
    - Run whatever attaches the Data Macro(s) or Validation Rules a second time
    - Confirm there is still exactly one copy of each
17. **When a build is blocked, it stops rather than leaving tables half-finished.**
    - Leave a table or form open in the database
    - Start the build
    - Confirm it stops, names what is open, and changes nothing
