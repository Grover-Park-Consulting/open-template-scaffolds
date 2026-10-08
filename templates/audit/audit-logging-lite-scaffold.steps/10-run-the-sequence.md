---
step: 10-run-the-sequence
title: "Run the setup sequence, with wizard Steps 6 to 9"
platform_facts: [mcp-run-procedure, table-in-use, mcp-file-release]
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Running the setup

### Before you run the generator

**Close every object in this database — tables, forms, reports and queries — and leave them closed
until the whole run is finished.** Opening something to look at it between the numbered steps is
enough to stop the next one. Each step checks before it starts and tells you what is open, so
nothing is changed while you sort it out.
`Four_GenerateAllAuditDataMacros` opens each table in design view to attach its macros, and it
cannot do that while anything is using the table. This is true whether or not another copy of the
database is open, and it is the commonest reason a run only half works.

Close every other copy of the database as well — in a split design, the back end and each front
end.

Anything left open stops the run on that one table; the other tables still finish, so what you get
is some tables generated and others not. Re-running is safe (generation replaces rather than
accumulates), so the fix is to close everything and run it again.

## What the developer runs, and in what order

**To the AI assistant.** This is the canonical sequence, and it is the source for the `runbook.md`
that `_materialization.md` requires whenever the developer is left procedures to run. **Write that
runbook whenever the developer runs any of this themselves** — which this staged sequence has them
do by design.
Naming the procedures
`One_`, `Two_`, `Three_` does not tell anyone what the first two are, that the config table has to
be reviewed in between, or that three of the twenty procedures are the only ones they ever call
directly. Turn the sequence below into their runbook in plain words; do not paste this section into
it.

**Only four procedures are ever run by hand** — `One_CheckAuditReadiness`,
`Two_CreateAuditTables`, `Three_PopulateConfigTable`, `Four_GenerateAllAuditDataMacros` — plus
`Zero_CreateSampleTables` on Path A, and the verification helpers afterwards. Everything else in
these modules is called by those. **They are numbered in the order they are run**, which is the
only thing the number means.

Each is a `Function` returning a report, so it is run from the Immediate window with a leading `?`
to print what it returns: `?Two_CreateAuditTables()`.

| # | What they do | Path |
|---|---|---|
| 1 | Set `AUDIT_SCOPE_MODE` (and `AUDIT_SCOPE_LIST`, where the answer was a list of names) at the top of `modAddDataMacros` to Step 4's answer, import the four modules, put each in the right file (see the split-database table above), and compile. | Both |
| 2 | `?Zero_CreateSampleTables()` — builds three made-up tables to try the system on. | A only |
| 3 | `?One_CheckAuditReadiness()` — reports which tables cannot be audited as they stand, before anything at all is created. | Both; required on B |
| 4 | `?Two_CreateAuditTables()` — creates `tblAuditLog`, `tblLongTextBackup`, `tblAuditLogConfig`. | Both |
| 5 | `?Three_PopulateConfigTable()` on Path A — every field starts switched **on**. `?Three_PopulateConfigTable(False)` on Path B — every field starts switched **off**. | Both |
| 6 | **Open `tblAuditLogConfig` and set the `IsAuditable` switches.** This is where the audit net is actually drawn, and nothing else does it for them. | Both |
| 7 | **Close every table, form, report and query in the database, and leave them closed until step 9 is finished.** | Both |
| 8 | `?Four_GenerateAllAuditDataMacros()` — attaches the macros and reports per table. | Both |
| 9 | Make one real edit to an audited table and open `tblAuditLog`. | Both |

**Step 3 comes before anything is created, and that is the point of it.** It reads table
definitions and changes nothing, so the developer learns which of their tables cannot be audited
while there is still nothing to undo. Running it later — after the three system tables exist and
after the config table has been filled in — tells them the same thing too late to act on it
cheaply.

**Step 6 is a step, not a note.** It is the one place the developer has to make decisions in data
rather than answer a question, and a runbook that folds it into step 5 produces a build that audits
everything or nothing.

**Step 7 is where a run goes wrong.** Say what happens if they skip it: the tables that were open
keep the macros they already had, the report names them, and running step 8 again after closing
everything fixes it.

`BackupAndRemoveAllDataMacros`, `DumpTableMacros` and `ListMacroEvents` are optional tools, not
steps — say what each is for and that a normal run never calls them.

### Step 6 — Check the tables first?

**Ask:** Should I read your table definitions and report any table this system can't track as it
stands?

| Option | Short description |
|---|---|
| `Yes, check them first` | I look at every table and tell you what I find. Nothing is changed. |
| `No, skip the check` | Go straight on. |

**Preferred:** `Yes, check them first` — this template's own.

**Skip when:** the tables were created by this template's own setup step, in this run. Those tables
were built to the shape this system needs, so the check has nothing left to find. Any other table
has never been looked at. This is always all-or-nothing, never a mix: Step 4 locks the demo's scope
to exactly the tables the setup step just created, so a run never both creates some of what it
audits and adopts the rest in the same pass.

<details>
<summary>Tell me more about the check</summary>

It reads your table definitions and reports whether each one will work. It changes nothing at all,
so there is no risk in running it, and it can be run again at any time. The procedure behind it is
`One_CheckAuditReadiness`.

**What it is looking for:** every table needs **one single number field as its primary key, set to
auto-number**. Most tables you designed yourself already look like this. Older or inherited tables
sometimes don't — a table with no primary key set, one that uses two or more fields together as its
key, or one keyed on a text code will not work with this system as it stands.

A table designed before this system existed is where such a key turns up, which is why the check is
here.

If a table isn't ready you have two ways out: fix that table's primary key, or leave the table out
by switching its fields off. Adapting the template to a different key design is possible, but it is
your adaptation, not something this template supports.

</details>

### Step 7 — Are these the long-text fields?

**Ask:** These are the long-text fields I found. Is that right?

| Option | Short description |
|---|---|
| `Yes, that's right` | Tables with one of these get some extra handling. |
| `No — let me look first` | Nothing happens until you say so. |

**Preferred:** none. It is your database, and the build acts on this list.

**Skip when:** never — but where no long-text field is found anywhere, this is an empty list to
confirm rather than a decision.

<details>
<summary>Tell me more about Long Text fields</summary>

**A Data Macro cannot read or write a Long Text field at all.** That is a limit of the Access
engine, not a choice this template made, and it cannot be worked around inside the macro.

So a table carrying one takes a different route: a Before Change and a Before Delete macro call a
VBA function (`BackupLongTextFieldsDM`) that copies the old value into a staging table first, and
the audit macro reads it back from there. A table with no Long Text field needs only the three
simpler After macros.

That is why the list matters: it decides which macros each table gets. Getting it wrong doesn't
produce an error — it produces an audit trail that quietly records nothing for that field.

The function this depends on lives in `modAuditLongText`, which must exist in the back end **and in
every front end**, for the reason given before step 1.

</details>

### Step 8 — Have you set the tracking switches the way you want them?

**Ask:** Have you set the tracking switches the way you want them?

| Option | Short description |
|---|---|
| `Yes, they're how I want them` | Carry on. |
| `Not yet — I'll set them now` | I close the database so you can open `tblAuditLogConfig` and set them, then wait until you say you're done. |

**Preferred:** none. Only you know whether the list says what you meant.

**Skip when:** never. This is the review the whole design is built around.

**To the AI assistant: if you are building the database yourself rather than handing over a script,
you are holding the file open and the developer cannot open it.** Close it, say that you have, and
wait for them. Then reopen it to carry on. This applies at every step that asks the developer to go
and look at a table, not only this one — an instruction to open something they are locked out of
reads as the wizard being broken.

<details>
<summary>Tell me more about the switches</summary>

This is where you actually decide what gets tracked, and you decide it as data rather than in code
— one line per field, with a Yes/No switch. Open `tblAuditLogConfig`, sort by table name, and set
them.

It is worth spending a minute on, because the next step acts on exactly what is in that table. It
is also the cheapest thing here to change your mind about later: flip a switch, run the last step
again, and the macros are rebuilt.

The rows for the three system tables are switched off and must stay off — auditing the audit trail
would loop.

</details>

### Step 9 — Ready to switch auditing on?

**Ask:** Close every table, form, report and query in the database first. Anything left open stops
this on that table. Ready to switch auditing on?

| Option | Short description |
|---|---|
| `Yes, switch it on` | Attaches the tracking to your tables and reports what happened to each one. Any Data Macros they already carry are replaced. |
| `Not yet` | Nothing is changed. |

**Preferred:** none. This is the step that changes your tables.

**Skip when:** never.

<details>
<summary>Tell me more about what this step does</summary>

For each table still switched on, the generator writes the macro definitions out as XML and loads
them onto the table. It has to be done this way: **DAO cannot create Data Macros**, and writing
UTF-16 XML and loading it with `Application.LoadFromText` is the only build path there is.

`LoadFromText` replaces a table's entire macro set. It never merges. This is why the house
audit-column stamping and the change auditing are generated together, in one Before Change
macro, so that one macro carries both and neither can replace the other. But any *other* Data Macro
a table carries,
***business logic of your own written for unrelated reasons is replaced too.*** The generator backs a
table's existing macros up to `DataMacroBackups\` before replacing them and names the affected
tables in its report; putting that logic back is your call.

**And if your tables already record who created a record and who changed it, the macro doing that is
replaced too — by this system's version, which may not behave exactly like yours.** Both fill the
same columns, so the difference is easy to miss. Yours might fill the changed-by pair when a record
is first created, or put a value back if somebody clears it. This one fills the created pair when the
record is made and the changed pair only on a later change, and it never puts anything back. Your
original is in `DataMacroBackups\` if you want to compare the two, or fold something back in.

**It is safe to run again.** Generation replaces rather than accumulates, so a run that stopped
part-way is fixed by closing everything and running it again.

**The one thing that stops this is an object left open**, and it is common enough to expect. In
practice it is usually the very first table this hits, not a later one — whatever tool is running
this build often holds the database open on its own account, which is the case people miss, so the
Ask above already tells you to close everything first rather than let this surface as the failure.
Attaching the macros needs each table to itself, and it cannot have that while a form, report, query
or the table itself is open on it. Where it does turn up partway through — some tables done, one
newly opened after the run started — Access says so at the time, one message per table it could not
do. The tables it could not do keep the macros they already had; the rest are finished normally; and
the report at the end says how many were built and how many failed. Close everything and run it
again.

**A table with auditing switched off still gets a stamping macro.** That is not an oversight: the
house audit columns are `Required`, and a table with no macro at all has no way to fill them, so it
would refuse every insert.

</details>
