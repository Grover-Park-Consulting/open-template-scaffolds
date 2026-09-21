---
template: app-startup-outcome-first
title: Application Startup and Back-End Relinking — outcome-first method
domain: app-startup
type: outcome-first
version: 0.2.0
status: draft
standards_layer:
  - design-principles
  - error-handling
  - naming-conventions
  - query-style
  - startup-conventions
house_assumptions:
  - The locations of any folders everybody shares are settings kept in the data file. That is why the
    folders are dealt with after the data connection and not before. A practice that keeps folder
    locations somewhere else — beside the front end, or in the code — changes that ordering.
warnings:
  - This template changes the front end people already open. It adds code that runs the moment the
    file opens, and it rewrites where the table links point. A build against an application in real
    use is preceded by a backup copy of the front end, and the developer is asked for one before
    anything is changed. The data file is read and never written to, so the front end is the file
    that needs copying.
related:
  - "error-logging-outcome-first — worth adding once app-startup is built: your application now has
    a place it starts from, and this template gives the errors it runs into somewhere to go instead
    of nowhere."
  - "standards/error-handling.md — worth a second look once you're building error logging. This
    template's own entry adds a constraint that would not be obvious from the standards file alone:
    whatever logs errors here must never write to the back end, because that is exactly the
    situation this template exists to handle."
---

# Application Startup and Back-End Relinking — outcome-first method

**Who reads this.** Everything from *Intent* down to *Standards Layer* is written for the developer
whose application this is. The section after that is addressed to the AI assistant building it, and
says so where it starts.

---

## Intent

**To produce, in an Access application, the results described in *What you end up with* below.** That
section is the specification: what the application does once this is built, in words a person using
it can check for themselves, followed by the checks that confirm the result arrived and the
behaviours that must hold however the work was divided up. Those together are the whole of what this
template promises.

**Nothing else is here.** No procedures, no module names, no code. This template states the result
and leaves the route to whoever builds it. No single mechanism is named, because the platform does
not force one: ordinary VBA, called from `AutoExec`, satisfies everything this template asks for. The
other version of this template — the rules-based method, `app-startup-scaffold` — produces the same
result from working procedure skeletons you fill in and import. Either one can be built against your
own application, and they can be built one after the other, against separate copies, to compare.

---

## What you end up with

Your **front end** — the file each person opens, holding the forms, reports and code — opens by
itself and finds its **back end**, the data file holding the tables the front end is linked to. An
application split into those two files is called a **split database**, and that is the shape this
template is for. Where everything lives in one file instead, the reconnecting half of this has
nothing to look for and passes straight through; the folder half below still runs.

If the data file has moved, the application either reconnects to it without saying anything, or asks
the person where it went and reconnects to the file they choose — refusing a file that does not hold
this application's data. Where the data file is kept behind a database password, you are asked for
that password once, while this is being set up. Afterwards nobody is asked again — the links
themselves carry what is needed. Where it cannot reconnect at all, it says so in words the person can
act on, naming the place it looked, instead of opening normally and then failing on the first screen
that shows data.

Once the data can be reached, and before any screen opens, the folders the application keeps files in
are dealt with: a folder everybody shares is confirmed to be there, and reported if it is not; a
folder belonging to one person alone is created if it is missing. Then your startup form opens.

**Nothing has to be set up by hand to get it started.** The first time the application opens with
links that work, it records where they point, and that is what it falls back on later. You never type
a path in to begin with.

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

### The same behavior every time, not the same structure

For an AI-assisted template, two builds need not produce the same code. They need to produce an
application that behaves the same way — the same things happening, in the same situations, for the
people using it. The work may be divided into different procedures, given different names, and
commented differently; some builds will come out very alike and others not. None of that is a fault,
and **comparing one build against another is not the test.** What follows must be true of every
build — two builds that agree on all of it have both succeeded, however alike or unlike they read.

- **[your standards]** The front end remembers where its data file is, somewhere that still works
  when the data file cannot be reached. A setting kept inside the data file cannot tell you where the
  data file is.
- **[your standards]** The data connection is checked before anything that touches data, and before
  any folder is checked — the shared folders' locations are themselves settings kept in the data
  file, so they cannot be read until the data file can be.
- **[your standards]** A file the person chooses is tested before any linked table is repointed at
  it, and refused if it fails the test.
- **[your standards]** The remembered location is written only after the application has actually
  reached the data file at that place. A remembered wrong place is worse than none, because it is
  tried first on every future open.
- **[your standards]** Nothing in this part of the application writes its record of an error into the
  data file. Everything here can be running at a moment when the data file is exactly what cannot be
  reached, so a record written there is lost precisely when it is needed.
- **[your standards]** Nothing here is harmed by running again: it runs on every open, and the same
  routine that creates a folder the first time confirms it every time after. The technical term for
  this is that the startup is **idempotent**.
- **[your standards]** A folder everybody shares is confirmed, never created.
- **[your standards]** When the data file could not be reached, the application does not carry on to
  a screen that shows data.
- Links that do not point at an Access file — a link to a database server, an **ODBC link** in the
  usual term, or a link to a spreadsheet or text file — are left exactly as they are.
- Testing whether the data file can be reached never locks it against anyone else. Nobody is shut out
  of the application because somebody else's front end is checking.
- Whether the data file is kept behind a database password is asked before anything is built, and a
  build told that it has one does not go on without the password.

Lines marked **[your standards]** come from the startup rules in your standards layer
(`standards/startup-conventions.md`) rather than from this template, and move with that layer if your
shop replaces it. The three unmarked lines are this template's own. Everything else in `standards/`
applies here as it does to every template.

### Free to choose alternatives

The template does not decide any of the following. If you have specific preferences about any of
them, say so while the design is being worked out, before anything is built. Where you don't choose,
the build will choose based on the rules built into it. The template's promise holds either way.

- How the work is divided into procedures, what those procedures are called, and which of them the
  rest of the application can call.
- **Where the remembered location is kept** — a small table inside the front end, or a text file
  beside it — provided it can still be read when the data file cannot be. **The paired scaffold's
  answer is a table called `USysLocalSetting`; that's one worked example, not a requirement.**
  Writing it into the code is a third choice and costs two things: the application can no longer
  update it when somebody points it at a new file, and check 9 above cannot be performed.
- The wording of everything the person sees, provided the message reporting that the data file cannot
  be reached names the place it looked.
- Whether those words arrive in a message box or on a form of your own.
- Whether a start that cannot go on leaves the application sitting open or closes it.
- Whether a shared folder that cannot be reached stops the start or only warns.
- How the code is laid out and commented, within whatever your standards already require.
- The order in which you are asked for the things only you can supply.

**If you want something not already in the template, you can have it.** Our promise is specified in
*What you end up with*, together with *The same behavior every time, not the same structure*. The
checks under *How you validate the template's output* are how you confirm you got it in any given
build. Asking for something different is an extension of the basic build, and you are welcome to
experiment — the result is then your responsibility, and the checks above may no longer describe what
you have.

### What the template does not do

Some of this is left out by design. Some of it the template cannot do — the check it makes does not
reach that far, which is a limit rather than a decision. Either way it comes to you the same: nothing
here does it, and nothing here tells you it isn't being done.

- **Checking that the front end and the data file are the same version of the application.** They can
  be reconnected to each other and still disagree about what the tables look like.
- **Telling this data file apart from an old copy or a backup of it.** The test is that the chosen
  file contains a table this application's data file must contain — and last year's backup contains
  it too. Somebody who picks a backup gets a working application running on stale data, and nothing
  here will say so. A version check, named below, is what would close this.
- **Checking that the chosen file holds everything the application needs.** One named table is what
  it looks for.
- **Front ends linked to more than one Access file** — a second data file, an archive, a shared
  lookup file. It looks at the first linked table it finds and treats that one file's location as the
  answer for all of them.
- **Noticing that only some of the links reach the data file.** Because it looks at the first linked
  table it finds, a front end left half-reconnected by an earlier failure passes the check on the next
  open.
- **Putting back a reconnection that stopped part way through.** If some links moved to the new file
  and then one failed, the ones that moved stay moved. You are told the reconnection failed; you are
  not told that part of it succeeded.
- **Links to a database server rather than to an Access file.** Those fail for reasons of their own —
  a server name, a driver, a sign-in — and repointing one at a file path would make matters worse.
- **A data file whose database password is changed later.** The links carry the password as it was
  when they were made. Once it no longer matches, the data file reads as one that cannot be reached,
  and choosing that same file when asked gets it refused as not this application's data file. Putting
  that right means relinking the front end yourself; it is not something this sets up.
- **Adding links for tables the front end does not already link to.** If the data file has gained a
  table since the front end was built, nothing here notices.
- **Searching likely folders for the data file** instead of asking. A search that finds the wrong copy
  is worse than a question that gets the right answer.
- **Telling apart the different ways of not reaching the data file.** Missing, damaged, or present but
  not permitted all produce the same report: that it could not be reached, and where it looked.
- **Fixing everybody's front end at once.** Each copy remembers its own answer, so each person is
  asked once. That is usually right — two people can legitimately reach the same file by different
  routes — but one person's answer does not help anybody else.
- **Stopping somebody opening the front end in a way that skips the check entirely.** *Close the
  Shift-key bypass*, under *Extra Options* below, is one way to close that specific route.
- **Controlling what the data file itself does when somebody opens it directly.**

---

## Information and conditions you need to supply

Six things, and nothing here can be guessed — the sixth applies only to some applications:

1. **A backup copy of your front end**, if this is an application people are using. What gets built
   here runs the moment the file opens, and it rewrites where your table links point, so a build that
   goes wrong goes wrong at the one moment you need the file to work. A copy of the front end, made
   before anything is changed, is what you go back to. Your data file is only read, never written to,
   so the front end is the file that needs copying. **If you tell the template there is no backup,
   the build stops rather than continuing.**
2. **Your startup form** — the switchboard, menu, or home form the application opens once everything
   checks out.
3. **A table your data file must contain.** Naming one is how a chosen file is tested: a file that
   does not have it is not this application's data file. Pick one that no other database of yours
   would have.
4. **What to call the application** in the messages people see.
5. **Your folders, if the application uses any** — which of them everybody shares, and which belong
   to one person. Say where the shared ones are, and where that location is recorded.
6. **Whether your data file is kept behind a database password, and the password if it is.** Access
   has two different things called a password, and only one of them matters here. *User-level
   security* is the older mechanism, and it applies only to the `.mdb` file format — rare to meet
   today. What's meant here is a **database password**, which either the front end or the back end,
   or both, can carry independently. **What matters for this template is specifically the back end's
   password**, because reaching the back end to confirm it's the right file is what needs it. If it
   has one and you would rather not supply it, the build stops — nothing here goes after a data file
   whose password you have not given. One thing worth knowing, briefly: a link to a password-protected
   back end carries that password in plain text, readable by anyone who can open the front end — that
   is Access's own doing, and it rarely matters for what this template is used for.

**And one thing to arrange rather than supply.** The application has to sit in a folder Access trusts.
Code in a file anywhere else does not run at all, and Access does not say so — it reports that it
cannot find the procedure, which sends people looking at the code. Trust is a setting on each machine,
not a property of the file, so it is arranged once per machine and applies to everyone running a copy.
**Nothing here works without it, and nothing here can arrange it for you.**

### Facts about the platform

**Nothing in this section is part of the specification.** It is here so a build knows what Access
will and will not do, and so a developer reading it can see why the design looks as it does. Anything
that binds is stated in *What you end up with*, in full, and a requirement that appears to live only
here is a defect in this template.

- A front end whose links to tables are broken still opens normally. Nothing complains until
  something touches data, which is usually a form. With no check when the file opens, the first sign
  of trouble is an error on somebody's screen.
- A link records the data file's path as it was at the moment the link was made, and nothing
  automatically keeps that path up to date afterwards.
- A link to a table in another Access file records that file's path as `;DATABASE=<path>`. A link to
  a database server, or to a spreadsheet or a text file, records something else entirely, which is how
  the two are told apart.
- A link to an Access file **that has a database password** records something different again:
  `MS Access;PWD=<the password>;DATABASE=<path>`. It does not begin with `;DATABASE=`, so anything
  looking for that beginning does not find the link at all.
- Changing that recorded text does not reconnect anything by itself. Refreshing the link is what
  re-establishes the connection — `RefreshLink`, in the technical name — and without it the link
  holds new text and still reaches nothing.
- A file chooser hands back whatever the person picked and can tell you nothing about it. Anything
  that matters about the chosen file has to be established by opening it and looking.
- Opening the data file to test it **without asking for sole use** leaves everyone else able to open
  it. A test that opens it exclusively locks out every other person for as long as the test runs.
- A path on a server that is not answering can take many seconds to fail rather than failing at once.
  While that is happening the application looks like it has hung.
- An `AutoExec` macro runs when a **person** opens an Access file. In normal use the data file is
  opened by the database engine on behalf of a front end, not by a person, so an `AutoExec` placed in
  the data file never runs as part of the application starting up.
- **Holding Shift while opening a file tells Access to skip the `AutoExec` macro and the startup
  settings altogether.** That is a legitimate recovery path — it is how a developer gets back into a
  front end whose startup is broken — and it is also a way past every check this template builds. **It
  can be closed**, per file, by setting the database property `AllowBypassKey` to `False` — see *Close
  the Shift-key bypass* under *Extra Options*, below. Closing it removes the same recovery path it
  removes for anyone trying to get around the checks, so it is a deliberate trade-off, not a pure
  hardening step.
- Access has a startup setting of its own — *Display Form*, under Options → Current Database, and
  `StartUpForm` where the file's own properties are listed — naming a form to open when the file
  opens. It is separate from anything this template builds, and a form opened that way is not covered
  by the checks above, so leave that setting empty and let the code open your startup form.
- `RunCode`, the macro action that calls VBA, can call a `Function` and never a `Sub`. That is why the
  entry point returns a value even where nothing reads it.
- `MkDir` creates one folder level per call. A folder two levels deep needs two calls.
- A folder path that does not start from a drive letter or a server name is worked out from whatever
  folder the session happens to be pointing at, which is not necessarily the one holding the front
  end.
- Creating a folder on the local machine where a shared one was intended **succeeds**, silently. No
  error is raised, the person who put a file there sees it, and nobody else does.
- An object whose name begins with `USys` is hidden from the navigation pane. Nothing is set on the
  object itself to do it — the name alone is what hides it — so the table is not missing, and *Show
  System Objects* in the navigation options reveals it.
- Code in a file that is not in a trusted location does not run, and the symptom does not mention
  trust: Access reports that it cannot find the procedure. A trusted parent folder is not always
  enough — a file under the temporary folder stays untrusted even where the folder above it is
  trusted.
- A data file that several people share must not sit in a folder that syncs to the cloud. Those
  folders copy a whole file once it stops changing, which is not how several people writing to one
  database at once works, and the file is corrupted rather than merged.
- Code that depends on an object library which is not installed on the machine the front end lands on
  does not merely fail at that line — the whole project refuses to compile, so nothing in it runs at
  all. Anything used at startup either sticks to what is always present or asks for the object by
  name at the moment it is needed.

---

## Standards Layer

**Your standards layer decides how this is built. This template decides only what it has to do.**

- **`startup-conventions`** — the `AutoExec` → `Startup()` convention, `Startup()` being a `Public
  Function` because `RunCode` cannot call a `Sub`, the idempotence rule, the relink convention, and
  the two kinds of folder. The lines marked **[your standards]** under *The same behavior every time,
  not the same structure* are outcomes that layer requires; they are stated here so the developer can
  see the whole bar in one place, and they move with the layer if a shop replaces it.
- **`error-handling`** — how a generated procedure reports a failure. One constraint on the answer is
  not a preference and is stated as an invariant above: whatever reports errors here must not write
  into the data file.
- **`query-style`** — how the generated code writes and holds its SQL.
- **`naming-conventions`** — procedure, variable and table names, within whatever your standards
  already require.
- **`design-principles`** — how the work divides into procedures. *Free to choose alternatives* leaves
  that division open on purpose, and this layer is what it is open to.

A shop adopting this library replaces `standards/` with its own. Nothing in this file changes when
they do, which is the point of keeping the two apart.

---

## To the AI assistant building this

**Two sections are the specification, and only those two: *What you end up with* and *The same
behavior every time, not the same structure*, both under that heading above.** Build an application
that satisfies every promise there and passes every entry under *How you validate the template's
output*. Then run every one of those checks yourself, on a copy, and record what each one did — in
`build-record.md`, written before you report the build finished, never afterward and never only when
asked for it. The checks bind you twice: the build has to pass them, and you have to run them. How you
build is yours to decide, within *Free to choose alternatives*. Which checks you run is not — all of
them, every build. Every other section in this file — *Intent*, *Facts about the platform*, *What the
template does not do* — is context for reading those two. None of it binds on its own, and nothing
that binds is stated only there.

- **This template names no mechanism, because the platform leaves more than one workable route.**
  Ordinary VBA called from `AutoExec` satisfies every promise above. Do not import a mechanism from
  another template in this library on the assumption that a promise this firm must mean one route: it
  doesn't, here.
- **You may read `app-startup-scaffold.md`, the rules-based method that produces this same result, for
  one worked decomposition — nothing more.** It shows procedure names, a control flow, and where the
  storage and the probes sit. None of that is binding here. Copying its shape wholesale is a
  legitimate build; so is a different one that still satisfies every check.
- **This template names no procedures, no module, and no table on purpose.** How the work divides,
  what it is called, and where the remembered location lives are all declared free above. Do not
  import a decomposition from anywhere else, and do not treat the count of things in this file as a
  count of procedures to write.
- **Read every file in `standards/` and apply it.** Naming, error handling, query style, and the
  error-handling frame all come from there and never from this file.
- **Ask for the six things under *Information and conditions you need to supply*,** one at a time,
  through the interactive selection control where the answer is a choice and as a plain question where
  it is a name. Two of them are gates. The first is asked before anything else in the list: an
  application in real use with no backup copy of the front end stops the build. The sixth: if the
  developer says the data file has a database password and then declines to supply it, stop and build
  nothing.
- **After the sixth thing and before you present the design, offer the `Explore options` step**
  (`_template-schema.md` §12.5) over the list under *Free to choose alternatives*, and nothing outside
  it. A pick made there is recorded in the build record, holds for the rest of the run, and is
  restated in the design you present.
- **Never infer an answer that belongs to the developer** — not from what the application looks like,
  not from reasoning that makes an answer seem obvious. Where a check exists to answer a question, run
  the check at the point the sequence calls for it rather than working the answer out yourself.
- **Surface the house assumption in the front matter** and get the developer's answer before building.
- **The build record reports against *How you validate the template's output*, one entry per check,
  each saying what was done and what was observed.** Passed and not passed are the only outcomes,
  including where the first method to run a check hits an obstacle — see `_template-schema.md` §12.2
  for the full rule, the `Result: PASSED` / `Result: NOT PASSED` line every entry opens with, and what
  to do before settling for a soft result.
- **While the build runs, do not narrate it.** Say once that it has started and what it will produce;
  say anything the developer must act on, as a question; say when it is finished, what was built, and
  where the build record is. Everything else — every procedure written, every check that passed — goes
  to the build record.
- **Once the build is reported finished, and only then, mention each entry under `related` in the
  front matter** (`_template-schema.md` §7.1) — one line per entry, what it is and why. This is not
  part of the build, never a gate, and never read before this point.

## Extra Options

*Named optional extensions, none of them filled in for an engagement.*

- **Close the Shift-key bypass**, per `AllowBypassKey`, a database property. Closing it removes the
  same recovery path it removes for anyone trying to get around this template's checks, so it is a
  deliberate trade-off rather than a pure hardening step, and it isn't part of the base promise.
  `app-startup-scaffold`'s *Extra Options* carries one accepted way to do it, adapted for the same
  reason a route decision is: to `error-handling.md`. Run it once, by hand, as a deployment step —
  never from inside `Startup()`. **Closing it can be undone, but not from inside the file it was
  closed on.** Setting the property back needs code to run, and a front end whose startup is broken
  is one you cannot get code to run in, which is the situation you would be undoing it for. The route
  back is a second Access file, or a small script, that opens the closed file through the database
  engine and sets the property to `True` from outside it. Tools that do exactly this exist; this
  library does not supply one. The other route back is the backup copy you were asked for at the
  start, which still has the bypass open.
- **An explainer form with a *Try again* button**, so a failed start can be retried without closing
  and reopening the application. (Whether the words arrive on a form or in a message box is already
  free; the retry is the extra.)
- **A "where is my data?" menu item** — the same reconnection offered on a button, so a move can be
  handled deliberately rather than waiting for the next failure.
- **A version check**, comparing something recorded in the front end against something recorded in
  the data file. It can only run after the data connection succeeds, so it goes after everything
  above. It is also what would close the two exclusions about backups and stale data, named under
  *What the template does not do*.
