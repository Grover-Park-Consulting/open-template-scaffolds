---
step: 01-wizard-questions
title: "The wizard: the questions asked before anything is built"
platform_facts: [target-file]
route: both
---

**Who reads this:** the AI assistant, carrying out this step of the build.

## Wizard

Nine questions, asked one at a time, preceded by the entry question
(`templates/_template-schema.md` §10.6) that asks whether you want to answer them at all. The
standards gate comes from its own part: run it before any of these. It is not one of the nine.

**One of the nine — Step 3 — is only asked where you are adding auditing to a database you already
use**, so the try-it-out build asks eight and the other asks all nine.

**If the entry question is answered `Just build it`:** Steps 1, 4, 5 and 6 use their preferred
choices, and the rest are still asked — they have no preferred choice, because each needs something
only you can supply: which file the tables are in, whether you have a backup, whether a list is
right, whether the switches say what you meant, and permission to change your tables. On the
try-it-out demo that is Steps 2, 7, 8 and 9 — four questions instead of eight. On a database you
already use it is those four and Step 3 as well — five instead of nine, because only you can say
whether you have a copy to go back to. State the preferred choices being used before acting on them.

These questions are how the build gets from what you asked for to something that fits your own
database: the parts only you can settle, asked one at a time, before anything is built. Nothing is
installed to run them and no form is built; the AI assistant asks them in conversation. See
`templates/_template-schema.md` §10.

> **If an AI assistant is running this for someone:** ask each step and wait for the answer. Never
> infer one — not from what the database looks like, not from reasoning that makes an answer seem
> obvious ("it already has real data, so it must be Path B"). Never work out a check procedure's
> answer yourself by reading the tables: run the procedure at the step that calls for it, show what
> it said, and stop there. Don't collapse the sequence into a single upfront report, even where
> every fact in it turns out correct. Offer to go back at every step after the first, and when an
> earlier answer changes, discard the answers after it and resume forward from there.

**Before step 1**, three things that apply to the whole build whatever the answers are, and that
you can act on right now:

- **Close every copy of the database** — the back end and every front end, for the whole run and
  not only for the last step. The step that attaches the macros opens each table in design view,
  and a table held open by another Access instance stops the run part-way, leaving some tables
  done and some not. Re-running is safe, so the fix for a partial run is to close everything and
  repeat.
- **This module runs in the same file as the tables it audits** — the back end of a split design.
  Data Macros attach to tables in the file the tables actually live in.
- **`modAuditLongText` goes in every front end as well**, because it holds `AuditUser()`, which the
  generated macros call on every audited table. Without it there, a front end cannot insert a row
  into a table whose who-and-when columns this system fills. If
  this build uses a function your own database already has for that instead, the same requirement
  applies to whichever module holds that one.

Nothing else is said here. Every other warning this template carries is raised at the step where
you can do something about it.

### Step 1 — Which database are you building this into?

**Ask:** Are you trying this out on made-up tables, or adding it to a database you already use?

| Option | Short description |
|---|---|
| `A try-it-out demo` | Three made-up tables are created and audited. Nothing you already have is touched. |
| `A database you already use` | Auditing is switched on for your own tables. No tables are created for you. |

**Preferred:** `A try-it-out demo` — this template's own. The standards layer does not speak to
this choice.

**Skip when:** never. This is the first question after the entry question.

<details>
<summary>Tell me more about the two builds</summary>

**The try-it-out demo** creates three made-up tables — a client list, a support ticket list, and a
short pick-list of ticket priorities — and switches auditing on for them, so you can watch the
audit trail work before you touch anything real. Good for a first look, a demo, or learning what
this system does. Nothing in your own database is affected, because your own tables are not
involved at all.

**A database you already use** works directly against the tables you have. It does not create the
made-up tables. Because it changes something real, it is much less forgiving: Data Macros get
attached to your live tables, and that is not a step to redo casually.

Both run the same numbered procedures. What differs is `Zero_CreateSampleTables` (demo only), the
starting point for the tracking flags, and two safety steps that only the second path needs.

</details>

### Step 2 — Which file holds the tables?

**Ask:** Which file holds the tables you want audited?

| Option | Short description |
|---|---|
| *one row per database file found* | The tables live in this one. |
| `They're all in one file` | Everything is in a single file, so there is nowhere else for them to be. |

**Preferred:** none. Only you know which file is which, and working it out from a file name is
exactly the kind of guess this wizard is built to avoid.

**Skip when:** never.

**To the AI assistant: the option list holds three answers at most, and `They're all in one file`
is one of them — so two database files fill it exactly.** An ordinary folder holds more than two:
the file with the real tables, a copy of the file people run, and a backup of either. Narrow the
list before offering it rather than truncating it afterwards. A file holding only links to tables
elsewhere cannot be the answer, and neither can one with no tables in it at all, so neither is
offered. Drop `They're all in one file` as soon as a second database file has been found — the
folder has already answered it. If more than three candidates still survive, offer three and say in
the question itself that a different file can be named instead. There is no room on this step for
`Go back to the previous question`; that is expected here, not an omission.

<details>
<summary>Tell me more about why the file matters</summary>

The tracking gets attached to the tables themselves, so it has to be built in the file the tables
actually live in. Build it in the wrong one and nothing errors — you simply get a database with
some code in it and no tracking anywhere.

Most Access applications with more than one user are split across two files: one holds the tables
(usually on a shared drive), and each person runs their own copy of a second file holding the
forms, reports and code. The second file doesn't really contain the tables — it contains links
pointing at them. Tracking has to go where the real tables are.

If the two files are named alike, or one is a copy, open each and look: the one with real tables in
it rather than linked ones is the answer.

</details>

### Step 3 — Have you made a backup copy of this database?

**Ask:** Have you made a backup copy of this database?

| Option | Short description |
|---|---|
| `Yes, I have a copy I can go back to` | Carry on to the next question. |
| `No — stop so I can make one` | Everything stops here. Nothing has been changed. |

**Preferred:** none. Only you can say whether you have one.

**Skip when:** Step 1 chose the try-it-out demo — a demo touches nothing you would want back.

<details>
<summary>Tell me more about why a copy matters here</summary>

This build attaches Data Macros directly to your live tables, and `Application.LoadFromText`
**replaces a table's entire macro set** — it never merges. Any other Data Macro a table already
carries, written by you for reasons unrelated to auditing, is replaced along with everything else.

Existing macros are backed up to a file before they are replaced, and the last step tells you which
tables that happened to. But nothing puts that logic back afterward: re-adding it is your call,
from the backup.

Copying the database file is the same precaution you would take before any change you cannot easily
undo, and it is the only one that covers everything at once.

</details>

### Step 4 — Which tables should be audited?

**Ask:** Which tables should be audited?

| Option | Short description |
|---|---|
| `The tables you propose` | I look at every table and propose the ones that hold your data; you confirm or change the list. |
| `Every table in this file` | All of them, apart from the three kinds that are never included. |
| `A different set — I'll tell you which` | You give me the table names, and only those are used. |

**Preferred:** `The tables you propose` — it decides by what each table holds, which a table's name
only hints at.

**Skip when:** Step 1 chose the try-it-out demo. The demo's scope is fixed to the three tables it
creates — asking here would offer `Every table in this file` or another list, either of which could
reach real tables sharing the file the demo was built into, and Step 3's backup gate was skipped on
the promise that a demo run touches nothing you already have. Set `AUDIT_SCOPE_MODE = "List"` and
`AUDIT_SCOPE_LIST = "tblClient;tlkpTicketPriority;tblSupportTicket"` for the demo without asking,
and `AUDIT_STAMP_LIST` to the same three: step 0 gives them their who-and-when columns, and nothing
else fills them.

**To the AI assistant: three answers plus `Tell me more` fills the control, so there is no room on
this step for `Go back to the previous question`.** That is expected here rather than an omission,
the same as on Step 2.

**On `The tables you propose`, decide each table by what it holds, never by its name alone.** Read
every local table in your copy: its fields, its relationships, and whether people add, change or
delete its records as part of the application's work. A table that holds the application's data is
proposed; one that supports the application's administration (settings, configuration, staging,
working tables, logs) is not. A name's prefix or suffix is a clue, never the test.

**Show every local table in the file, each marked *proposed*, *not proposed* or *unsure*, with one
line saying why.** The three kinds never included (linked, `MSys`, `USys`) appear too, with that
reason. Every other reason names something you saw besides the name: a field, a relationship, a
form, the row count.

Clues, none decisive:
- **Data:** related to other tables by keys, or edited through a form, a lookup list included.
  Stamping columns already present are a strong clue.
- **Administration:** one row, or name-and-value pairs, with no relationships. Rows written only by
  code, carrying their own date and person: a log. Emptied and refilled by an import or a procedure:
  staging.
- **The name:** a prefix or suffix points one way.

**Where the clues disagree, mark the table *unsure*, with the reason.** The developer decides it.
Ask through the selection control: `Use this list`, or `Change it` (they say what to add or remove).
Ask again after every change until they choose `Use this list`.

Whichever answer comes back, set `AUDIT_SCOPE_MODE` in the module-level declarations before the
module is imported — `"List"` for the first and third answers, `"All"` for the second — and set
`AUDIT_SCOPE_LIST` to the confirmed or given names. Nothing at run time changes either value.

<details>
<summary>Tell me more about which tables get audited</summary>

This answer is the **only** one in the whole system that ends up written into code. Every
finer-grained
choice — which tables, which individual fields — is a flag you set in the config table afterward,
as data. That is deliberate: changing your mind about a field should not mean editing VBA.

**Three kinds of table are left out whatever you choose here**, because including them either cannot
work or is not this system's business:

- **Access's own system tables**, whose names begin `MSys`. Never touched.
- **Your own hidden tables**, whose names begin `USys`. Yours to manage, so one is included only
  where you name it yourself in the third answer.
- **Linked tables** — tables that live in another file and only appear in this one. Access cannot
  attach the tracking to a linked table at all: it has to be built in the file where the table
  really lives, which is what Step 2 was asking. Name one in the third answer and it is left out,
  and the report says which, rather than passing over it quietly.

`Every table in this file` includes working and staging tables, if your file has any. The first
answer leaves them out, because it looks at what each table holds rather than what it is called;
you see the list and confirm it before anything is built.

</details>

### Step 5 — Start by tracking everything, or nothing?

**Ask:** Should every field start out tracked, or should none of them?

| Option | Short description |
|---|---|
| `Track everything to start` | Everything is switched on, and you switch off what you don't want. |
| `Track nothing to start` | Everything is switched off, and you switch on what you do want. |

**Preferred:** follows Step 1 — `Track everything to start` on the demo, `Track nothing to start`
on a database you already use. This template's own.

**Skip when:** never.

<details>
<summary>Tell me more about the starting point</summary>

Either way you get one line per field and you decide the rest by flipping switches, so neither
answer locks anything in. Under the covers this sets a single argument on
`Three_PopulateConfigTable` — nothing after it starts everything on, `False` starts everything off.

**Starting on** suits the demo, where there are nine or ten fields and you want to see the trail
working immediately.

**Starting off** is the safer footing on tables this system was not designed around, where "track
everything" can sweep in more than you meant — long free-text notes, columns another process
rewrites constantly, fields you would rather not have a second copy of.

Some rows are switched **off no matter which you choose**: the three system tables (auditing the
audit trail would loop), each table's own primary key (its value is already on every log row), and
the house audit columns and other always-changing system fields. Nothing is hidden — those rows
are written and shown, just switched off.

**What you will see, so it does not look wrong:** most of the rows belong to the three system
tables, and on a small database that can be well over half of them. Sort by table name and review
only the tables you recognise as your own.

</details>
