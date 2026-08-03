# The Standards Gate

**Who reads this:** the AI assistant, running the gate with the developer who asked for a build.

**If that developer is you:** this file holds the questions you will be asked and why. You do not
have to read it — every decision in it reaches you as a question.

**Pilot scope.** Only `audit/audit-logging-lite-scaffold.md` runs this gate today. Nothing else
points at it, and nothing else should until the pilot has been tried and judged.

---

## 1. What the gate is

**The standards layer governs everything a template produces — and until the gate existed, a
developer could run a template end to end, successfully, and never learn the layer was there.** The
gate is the one question that fixes that: whose rules govern this build, asked where the developer
can answer it.

It is asked through the same interactive selection control as every wizard step, under the same
rules (`_template-schema.md` §10.3, §10.4, §10.5). It is **its own wizard**, separate from the
template's — §10.5 rule 14 already allows more than one, and the gate's outcome outlives the build
in a way a template's own answers do not.

**The result is settled once and reused.** A developer who settles it while building audit logging
is not asked again from scratch on their next build; they are asked whether the settled answer still
holds (§4).

### 1.1 Where it sits in a run

```
1. Template matched, offered, accepted.
2. The disclosure line (§10.4).
3. THE STANDARDS GATE.            ← this file
4. House assumptions asked; build-wide warnings surfaced.
5. The template's entry question (§10.6).
6. The template's own wizard steps.
```

**The disclosure line comes before the gate, not before the template's entry question.** §10.4
places it "immediately before the entry question" because that used to be the first question of a
run. The gate is now first, and the line's promise — *anything you need to decide is in a question I
ask you* — has to be made before the first question of any kind or it is made too late.

### 1.2 The gate asks no entry question of its own

§10.6 asks for one at more than three steps. The gate is **one step** until the developer opens the
walkthrough, and `Use them as they are` already does exactly what `Just build it` does: it takes the
preferred choice everywhere and stops. A second version of the same question is the ceremony §10.6
exists to prevent. Do not ask one.

### 1.3 Writing to the standards files is what this gate is for

**§10.5 rule 11 stands, and it does not reach the gate.** Rule 11 governs a choice made *inside a
template's wizard*: a step that offers a message box where the standards layer calls for a central
logger is exactly its case — the developer's choice holds for that build, and the standards file is
left untouched. That is right, and nothing here changes it.

The gate is a different act. It is not a build decision that departs from the layer for one run; it
is the developer setting the layer. Two of its three ownership paths change `standards/` on purpose
and permanently, because they said to. Make the change, and record it (§7).

---

## 2. Step G1 — the gate, on a first run

> **Step source. Ask this through the selection control.**

```markdown
### Step G1 — Use the rules these templates follow?

**Ask:** These templates follow a written set of rules about how things are named, how errors are
handled, and how changes are recorded. Use them as they are?

| Option | Short description |
|---|---|
| `Use them as they are` | The build follows the rules that came with the library. |
| `Let me look at them first` | We go through the rules one file at a time and you decide each one. |

**Preferred:** `Use them as they are` — the library's own; these are the rules the templates were
written against.

**Skip when:** a previous run settled this. Ask Step G1-again instead.

<details>
<summary>Tell me more about these rules</summary>

The rules live in seven short files that came with the library. They are the reason a table comes
out named one way rather than another, and the reason every table gets the same five columns
recording who changed a record and when. The templates were written against them, so changing a rule
changes what every template builds — not just this one.

You can use them as they are today and change them later; you will be asked again the next time you
run a template. If you go through the rules now, you settle it for this build and every build after
it.

</details>
```

**`Use them as they are` ends the gate in one click.** Go straight on to the house assumptions. That
is the shipped path and it has to cost nothing.

---

## 3. The walkthrough — on `Let me look at them first`

Seven files, one question each, in this order. **The five that shape the build in front of the
developer come first**, so they meet each one while the reason for it is visible; the last two are
part of the same set and are settled now so they are not asked again.

| # | File | Shapes an audit-logging build? |
|---|---|---|
| 1 | `naming-conventions` | Yes — decides which of their tables the build looks at |
| 2 | `audit-columns` | Yes — the set the build writes to |
| 3 | `error-handling` | Yes — appears in every procedure built |
| 4 | `query-style` | Yes |
| 5 | `design-principles` | Yes |
| 6 | `form-conventions` | No |
| 7 | `startup-conventions` | No |

**Never send the developer off to read a standards file on their own.** Each step below carries the
plain-words account of what that file covers; that account is what the walkthrough *is*. A developer
told to go and read seven markdown files has not been walked through anything.

`Go back to the previous question` is available on every step here after the first.

### Step G2.1 — The naming rules

```markdown
**Ask:** Use the naming rules as they are, or make them your own?

| Option | Short description |
|---|---|
| `Use these as they are` | The build follows the naming rules that came with the library. |
| `Make this one your own` | You choose how: change these rules, or use rules you already have. |

**Preferred:** `Use these as they are` — the library's own; these are the rules the templates were
written against.

<details>
<summary>Tell me more about the naming rules</summary>

They cover names for tables, fields, queries, forms, and the code that goes with them. Two parts are
important in what gets built today: table names start with a short prefix (`tbl` for a table of your
own data, `tlkp` for a short list of choices), and no field is left with a bare everyday word like
`Date` or `Name` as its name. The database already uses those words for things of its own, and it
has no way to tell that you meant your field — it takes the word for the thing it already knows, and
what comes back is not your data.

The table prefix decides which of your tables this build looks at. If your tables are named a
different way, change or replace this file, or point me at your own. Otherwise the build may find
none of them.

</details>
```

### Step G2.2 — The rules about tracking columns

```markdown
**Ask:** Use the rules about tracking columns as they are, or make them your own?

| Option | Short description |
|---|---|
| `Use these as they are` | The build follows the tracking-column rules that came with the library. |
| `Make this one your own` | You choose how: change these rules, or use rules you already have. |

**Preferred:** `Use these as they are` — the library's own; these are the rules the templates were
written against.

<details>
<summary>Tell me more about the tracking columns</summary>

Five columns are added to the end of every table: who created the record and when, who last changed
it and when, and one the database maintains itself. They go at the end, in that order, on every
table.

They matter more here than in most builds. The audit system this template installs writes to those
five columns every time anyone changes a record, so their names are built into the code you get. If
you change the names here, the code changes with them.

</details>
```

### Step G2.3 — The rules about what happens when something goes wrong

```markdown
**Ask:** Use the rules about what happens when something goes wrong as they are, or make them your
own?

| Option | Short description |
|---|---|
| `Use these as they are` | The build follows the rules that came with the library. |
| `Make this one your own` | You choose how: change these rules, or use rules you already have. |

**Preferred:** `Use these as they are` — the library's own; these are the rules the templates were
written against.

<details>
<summary>Tell me more about what happens when something goes wrong</summary>

When the code hits a problem while it is running, two things can happen: a message appears on screen
so the person knows, and the problem is written down somewhere you can look at afterwards. The rules
that came with the library do the first and depend on nothing else, so the code works in any
database as it stands.

If your shop already writes problems down in one place, this is the file to change — and every
procedure I build for you today will report problems your way instead.

</details>
```

### Step G2.4 — The rules about reading and writing data

```markdown
**Ask:** Use the rules about reading and writing data as they are, or make them your own?

| Option | Short description |
|---|---|
| `Use these as they are` | The build follows the rules that came with the library. |
| `Make this one your own` | You choose how: change these rules, or use rules you already have. |

**Preferred:** `Use these as they are` — the library's own; these are the rules the templates were
written against.

<details>
<summary>Tell me more about reading and writing data</summary>

The instructions that fetch and change data can sit inside the code, or in saved, named queries you
can open and look at. This file decides which, and how those instructions are written — including
how a name with an apostrophe in it, or a date, is handled so it does not break the instruction it
sits in.

</details>
```

### Step G2.5 — The rules about how the code is organized

```markdown
**Ask:** Use the rules about how the code is organized as they are, or make them your own?

| Option | Short description |
|---|---|
| `Use these as they are` | The build follows the rules that came with the library. |
| `Make this one your own` | You choose how: change these rules, or use rules you already have. |

**Preferred:** `Use these as they are` — the library's own; these are the rules the templates were
written against.

<details>
<summary>Tell me more about how the code is organized</summary>

This is the reasoning the other rules are built on: how much one procedure is allowed to do before it
becomes two, and how much any one procedure is allowed to know about the rest.

It is why what you get today is a set of small, separately named procedures — one that sets the
tables up, one that scans them, one that switches auditing on — rather than a single long one. You
can run them one at a time and see what each did.

</details>
```

### Step G2.6 — The rules about data-entry screens

```markdown
**Ask:** Use the rules about data-entry screens as they are, or make them your own?

| Option | Short description |
|---|---|
| `Use these as they are` | The build follows the rules that came with the library. |
| `Make this one your own` | You choose how: change these rules, or use rules you already have. |

**Preferred:** `Use these as they are` — the library's own; these are the rules the templates were
written against.

<details>
<summary>Tell me more about data-entry screens</summary>

How a screen is laid out — what the boxes and buttons on it are called, which buttons appear, and
what order the cursor moves in when someone presses Tab.

Nothing being built today is a screen, so this changes nothing you will see. It is here because
settling it now means you are not asked again the next time you build something that does have one.

</details>
```

### Step G2.7 — The rules about what happens when the database opens

```markdown
**Ask:** Use the rules about what happens when the database opens as they are, or make them your
own?

| Option | Short description |
|---|---|
| `Use these as they are` | The build follows the rules that came with the library. |
| `Make this one your own` | You choose how: change these rules, or use rules you already have. |

**Preferred:** `Use these as they are` — the library's own; these are the rules the templates were
written against.

<details>
<summary>Tell me more about what happens when the database opens</summary>

Which code runs first when someone opens the database, and what it puts in place before anyone can
do anything — checking that the folders it needs exist, for instance.

Nothing being built today runs at open, so this changes nothing you will see. It is here because
settling it now means you are not asked again the next time you build something that does.

</details>
```

---

## 4. Step G2.n-b — how to make one file your own

Asked **only** when a step in §3 was answered `Make this one your own`, and asked about that one
file. The wording below is the same for every file bar its name.

```markdown
**Ask:** How do you want to set the naming rules?

| Option | Short description |
|---|---|
| `Change them with me` | We go through them together now and change what you want changed. |
| `Replace them with your file` | You give me your own file and it takes the place of this one. |
| `Point me at yours` | Your rules stay where they are and I read them from there. |

**Preferred:** `Change them with me` — the library's own; it is the one way that cannot break the
shape the templates rely on.

<details>
<summary>Tell me more about the three ways</summary>

Changing them together keeps everything else in the file intact, which matters because the templates
expect these rules to be laid out a particular way. Before the first change, I'll put a copy of the
rules as they came with the library aside, so you can go back to them at any time.

Giving me your own file, or pointing me at one, replaces the whole thing. That is the right move if
you already have written rules. I check what you give me against what the templates expect and tell
you about anything missing before we go on — you decide what to do about it.

</details>
```

**This step has no `Go back to the previous question` slot** — three answers plus *Tell me more*
fills the control, which §10.3 allows. If the developer says in words that they want to go back, go
back.

**All three paths keep the file's shape.** Filenames never change, and sections keep their order and
their numbering — every template depends on that. Only what the sections *say* varies.

---

## 5. The backup — an action, not a question

**Copy `standards/` to `standards-original/` the moment the first `Make this one your own` is
answered** — before any file is changed, and never later than that. Do not ask the developer to do
it and do not ask permission: it is the thing that makes every other path safe.

**`standards-original/` holds the files exactly as they shipped, and it is inviolable.** Written
once; never written again, by any path, for any reason. If it already exists, leave it alone — it is
the one thing in the library that cannot be regenerated from anything the developer has, and
overwriting it with already-edited files would quietly destroy the only way back.

Saving further versions is a separate matter and a separate folder. Nothing in the pilot needs one;
the restore in §8 makes the only other copy the gate ever takes.

**It sits beside `standards/`, not inside it.** A folder inside `standards/` reads as part of the
layer to anything walking it, and the developer who downloaded a zip file has no other way back —
there is no undo underneath them.

Say it once, in the line between that answer and the next question:

> *"Before I change anything: I've put a copy of the rules exactly as they came with the library in
> a folder called `standards-original`, next to the rules themselves. If you ever want them back,
> they're there and I can put them back for you."*

---

## 6. Step G2.n-c — what the check found

Run a compatibility check on any file the developer **replaced** or **pointed at**. Do not run it
after `Change them with me` — an edit made together preserves the shape by construction.

Two halves:

- **Structure** — are the sections all there, in the same order, numbered the same way.
- **What it says** — does the file still answer the questions the build asks of it. This half is
  yours to judge, and it is the half that catches real damage.

**Why the second half is not optional.** An audit-logging build finds the tables to audit by looking
for a name starting with `tbl` or `tlkp` — the naming rules made executable, in two places in the
generated code. A developer who replaces the naming rules with their own, perfectly good, correctly
structured, with no prefix at all, gets a build that runs cleanly and finds nothing. Nothing is
malformed. Only reading the file catches it.

**Ask only when the check found something.** Compose the `Ask` line from what was actually found, in
the developer's words, naming the one thing — this is the only step in the library whose question is
written at run time, and a fixed one ("something was found, what now?") would force them into *Tell
me more* just to learn what they are being asked.

```markdown
**Ask:** Your naming rules don't say how table names start. What should I do?

| Option | Short description |
|---|---|
| `Sort it out with me` | We settle the missing piece now, before anything is built. |
| `Build with them as they are` | I use your rules and tell you afterwards what I couldn't do. |

**Preferred:** `Sort it out with me` — the library's own; the build uses this piece directly.

<details>
<summary>Tell me more about what was found</summary>

[List exactly what was found: sections that are not there, sections in a different order, and
anything the build depends on and cannot find. One line each, in plain words.]

The build finds the tables to audit by looking at how their names start. Your rules don't say
anything about that, so the build has nothing to match on and would find none of your tables.

</details>
```

`Build with them as they are` is a real choice, not a failure. Take it, build, and put every
consequence in the build record under what is left for the developer to do.

---

## 7. Recording the outcome

**Append a dated block to `standards/README.md`.** That file is never walked and never replaced by
this gate, so a replacement of any other file cannot destroy the record.

```markdown
## How the standards were settled

Settled 2026-08-03 while building audit logging. Run any template to change this.

| The file | What you chose |
|---|---|
| `naming-conventions` | Changed together — table names start with `t_` |
| `audit-columns` | Used as it came with the library |
| `error-handling` | Replaced with your own file |
| `query-style` | Used as it came with the library |
| `design-principles` | Used as it came with the library |
| `form-conventions` | Used as it came with the library |
| `startup-conventions` | Read from `D:\HouseRules\startup.md` |

A copy of the rules as they came with the library is in `standards-original`.
```

Replace an existing block; do not add a second one.

**The build record** (`_materialization.md`, "The build record") carries the gate too: its steps and
answers under the decisions taken, everything the compatibility check found under what was checked
before building — including anything the developer chose to build past — and anything a chosen rule
called for that the build could not deliver under what is left for the developer to do.

---

## 8. Step G1-again — the gate on a later run

Asked in place of Step G1 whenever `standards/README.md` carries a settled block.

```markdown
**Ask:** On 3 August you settled how these rules work. Use that again?

| Option | Short description |
|---|---|
| `Use what I settled` | The build follows the rules as you left them. |
| `Change something` | We go through all of the files again and you can change any of them. |
| `Go back to the library's rules` | Your changes are set aside and I use the rules that came with the library. |

**Preferred:** `Use what I settled` — your own answer, from the last time this was asked.

<details>
<summary>Tell me more about what you settled</summary>

[List it, one line per file: what was kept, what was changed, what was replaced, and where anything
that was pointed at lives.]

Going back to the library's rules doesn't throw your version away — it's kept, and you can come back
to it again if you want.

</details>
```

**Use the real date** from the settled block, not the words "last time".

**Offer `Go back to the library's rules` only when `standards-original/` exists and something
actually differs from it.** With nothing to restore the option means nothing, and the slot is better
spent on going back. That is not hiding an option under §10.5 rule 8 — the choice does not exist.

**Restoring keeps the developer's version.** Copy the current `standards/` to
`standards-saved-YYYY-MM-DD/` before overwriting it from `standards-original/`, and say where it
went. A restore that destroys their work is the same mistake the backup exists to prevent, made in
the other direction.

**That saved copy is not a second original**, and its name has to keep saying so — it is what the
developer had, at a date, and any later one is another dated folder beside it. `standards-original/`
is never a destination for any of this.

---

## 9. What to say between the steps

§10.4 rule 1 governs this, and it carries more weight in the walkthrough than anywhere else in the
library: **this is the only place the developer is shown a file before being asked about it.** Each
line does three things — names what was just recorded, says in plain words what the next file
covers, and where it applies, says what that file does to the build in front of them.

Moving from the tracking columns to what happens when something goes wrong:

> *"The five tracking columns stay as they came with the library — that's the set this build writes
> to every time someone changes a record. Next is what happens when the code hits a problem while
> it's running: whether a message appears on screen, and whether the problem gets written down
> somewhere you can look at later. It shows up in every procedure I build for you."*

Moving into the last two files, where the developer is most likely to wonder why they are still
being asked:

> *"That's the five rules we'll need for this build. Two more are part of the same set — they shape
> other templates in the library, just not this build. Settling them now means you won't be asked
> again."*

**After a `Change them with me` edit, say what changed in the build as a result** — not just that an
edit happened. That is where the developer finds out their edit had a consequence, while the file is
still in front of them.
