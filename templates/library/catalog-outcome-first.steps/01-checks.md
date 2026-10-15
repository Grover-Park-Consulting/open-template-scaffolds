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

1. **A new title gets a correct sort title.**
   - Add a new publication titled "The Great Gatsby"
   - Confirm the sort title reads "Great Gatsby" — the leading noise word removed
   - **Where your noise-word list carries house-specific additions** (Step 5 under *Information
     and conditions you need to supply*), repeat this with a title starting with one of those
     additions too — "The Great Gatsby" only exercises the default three words, and a custom
     addition that doesn't work would still pass this check as written otherwise
2. **A title with no leading noise word is stored as itself.**
   - Add a new publication with a title that carries no leading article
   - Confirm the sort title matches the title exactly
3. **Editing the title through your entry form updates the sort title.**
   - Edit an existing publication's title through your entry form
   - Confirm the sort title updates to match the new title
4. **Editing a title directly in the table — only if you chose the Data Macro route.**
   - Edit a publication's title directly in the datasheet, not through the form
   - Confirm the sort title updates too
   - If you chose the VBA route instead, this check is not expected to pass — that's the disclosed
     trade-off of that route, confirmed by *Business Rule 2*'s own text above, not a failure to
     report as one
5. **Editing an unrelated field leaves the sort title untouched.**
   - Edit a field on a publication other than the title
   - Confirm the sort title is exactly what it was before
6. **A multi-volume set without a volume number is refused.**
   - Set `MultiVolumeSet` on a record and leave `PublicationVolume` empty
   - Confirm the save does not succeed
7. **A volume number outside the set's range is refused.**
   - Set `MultiVolumeSet`, `NumberOfVolumes` to 3, and `PublicationVolume` to 5
   - Confirm the save does not succeed
8. **A single-copy record with a volume number set is refused.**
   - Leave `MultiVolumeSet` unset and set `PublicationVolume` to any number
   - Confirm the save does not succeed
9. **A "set" of fewer than two volumes is refused.**
   - Set `MultiVolumeSet` and `NumberOfVolumes` to 1
   - Confirm the save does not succeed
10. **A valid duplicate-copy record succeeds.**
    - Leave `MultiVolumeSet` unset, `PublicationVolume` empty, and `NumberOfVolumes` at 3
    - Confirm the record saves
11. **A valid multi-volume record succeeds.**
    - Set `MultiVolumeSet`, `NumberOfVolumes` to 4, and `PublicationVolume` to 2
    - Confirm the record saves
12. **`NumberOfVolumes` defaults to 1.**
    - Add a new publication without setting `NumberOfVolumes`
    - Confirm it reads 1
13. **Running the build again does not duplicate anything.**
    - Run whatever attaches the Data Macro or Validation Rule a second time
    - Confirm there is still exactly one copy of each — no duplicated Data Macro actions, no second
      Validation Rule
14. **When a build is blocked, it stops rather than leaving tables half-finished.**
    - Leave a table or form open in the database
    - Start the build
    - Confirm it stops, names what is open, and changes nothing
