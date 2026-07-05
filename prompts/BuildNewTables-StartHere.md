# Build a Set of Tables (a *schema*) — Ready-to-Use Prompt

Using this template, you'll create a set of tables for whatever you're building — for example,
adding stocktake scanning to an inventory app, or setting up a catalog for a library. You don't
need to know how to design a database, and you don't need the vocabulary that goes with it. The
hard design decisions are already made. You describe what you want, the AI builds the tables, and
you look them over.

**You will build tables. Your job is to evaluate them — and either approve them or ask for
revisions.** That's the whole loop. You're the one who knows the business; the AI handles the
structure.

You won't write anything technical. You'll fill in four short lines in plain English, paste the
prompt to your AI assistant, and read back what it gives you.

*(Your AI assistant needs to be able to see the library files. If you followed the setup steps in
the README, it can.)*

---

## Step 1 — Fill in the form

Right below is **the form** — the only part you edit. On GitHub or in a Markdown reader it shows
as a shaded box; if you're reading this as plain text, it's everything between the two lines of
three backticks (```). Replace the text in `<angle brackets>`; leave the rest as it is.
*(Open this file in Notepad or any plain-text editor. Don't use a word processor like Word — it
quietly changes some characters as you type, and the form may stop working.)*

```text
- Build: <what you want to create, in plain words — e.g. "add stocktake scanning to an inventory app">
- Standards: <'default', or 'use my own standards'>
- Who this is for: <describe your client — see below>
- Extra options: <'none' to start, or pick from the template's Extra Options — see below>
```

**Build** — Just say what you're making. The AI matches your words to the right template; you don't
pick a file. *"Add stocktake scanning to an inventory application"* selects the Stocktake Scanning
template. *"Set up a catalog for a library"* selects the Library Catalog template. And if nothing in
the library matches what you asked for, the AI won't force a fit — it will say so, and offer to
design your tables from scratch under the same conventions, with the same review. You're covered
either way.

**Standards** — Leave this as `default` to use the library's built-in conventions as they are. Or,
if you've already worked out your own conventions for your AI, say `use my own standards` instead.

**Who this is for** — Describe how your client wants the database to work: what they call things,
and how they want to be able to report on them. **Be as precise as you can — concrete, measurable
terms work far better than vague generalities.** *"Show each product's shrinkage as a percentage
and flag any over 5%"* gives the AI something real to build toward; *"track inventory accuracy"*
does not. The clearer and more specific your description, the closer the first result will be. If
you just want to see the template at work, give the bare minimum here — you can always come back
and expand this part later, when you want more customized output.

**Extra options** — Most templates offer optional add-ons, listed inside the template under a
heading called **Extra Options**. For example, the Library Catalog template offers per-copy
holdings, lending, and keyword tagging as extras. Leave this `none` to start, or name the ones
you want.

---

## Step 2 — Copy and paste

Copy **everything from the form above down to the end of this file** — your filled-in form *and*
the "Instructions to the AI" section — and paste it all to your assistant in one go. The
instructions are what make the result reliable; the assistant needs both. That's it — the AI
takes it from here.

If your assistant can't see the library's files, it will ask for them by name — paste in each
file's contents (or, if it can read your computer's files, give it the library folder's full
location), and carry on.

---

## Step 3 — Read back what you get

The AI hands you the tables in two parts:

- **A diagram** showing the tables and how they connect.
- **A list of every field** in each table, and what it's for.

Look it over. If it fits, approve it. If something's off — a name, a missing piece, a rule that
doesn't match your client — say so, and the AI revises. Nothing is final until you say it is.

The diagram is drawn in a notation called **Mermaid** — it's designed for viewing on GitHub (and
some Markdown readers, such as MarkText), and it may look unfamiliar at first. That's fine: **the
diagram only shows the shape** — which tables exist and how they connect. **The field list is the
authority** on exact data types and what's required or optional.

---

> **Leave the next section as it is for now.** It tells the AI how to do the job correctly, and the
> defaults here are what make the results reliable. Once you've run this prompt a few times and
> you're comfortable with how it works, you can adjust it. Until then, don't change it.

### Instructions to the AI

You are generating a set of tables for the person and project described above. Work strictly from
the matched template and the standards layer — do not invent content beyond what the template
defines and what I've supplied under "Who this is for."

1. **Match and read** the template that fits the "Build" line, and read the **standards** (the
   library default unless I named my own). If you cannot see the library files, list the exact
   files you need — the matched template and the standards — and ask for them before proceeding;
   never guess at their contents.
2. **Honor the template** — its tables, fields, relationships, and business rules are decisions
   already made. Carry them through unless my notes override a specific point.
3. **Apply the standards** — naming, audit columns, and the error-handling pattern come from the
   standards layer, never from the template itself. Apply the field-naming rules (no bare reserved
   or ambiguous words — `Status` becomes `<Entity>StatusID`).
4. **Surface every declared house assumption** the template carries. List each one and ask me to
   confirm or change it before finalizing — these are choices I should see, not inherit silently.
5. **Fold in my specifics and any extra options** I asked for. If I named an option, pull it from
   the template's **Extra Options** section into the live result.
6. **Ask, don't guess.** If the template and my notes leave something genuinely undecided, ask me —
   don't fill the gap with invented content.
7. **No matching template? Don't stop.** Tell me which templates you considered and why each falls
   short. Then offer, as the default next step: with my approval, you'll design it from scratch
   following the same standards layer, producing the same two-part result below for the same
   review — and ask me for the go-ahead before you begin. (Mention the alternatives in passing:
   adapting the nearest template despite the mismatch, or me refining the description.) In a
   from-scratch design, list every assumption you had to invent as **proposed assumptions** for me
   to confirm or change — I should see your choices, not inherit them silently. And once I approve
   a from-scratch design, you may offer to shape it into a template for the library — it would be
   added upon the curator's approval.

Then produce the result in two parts:

1. **A diagram** — a fenced `mermaid` `erDiagram` showing the tables, their keys, how they relate
   (one-to-many, etc.), and how they connect to any existing tables the template builds on. Under
   the diagram, add one line: "This diagram shows the shape only — exact types and
   required/optional are in the field list below."
2. **A field list** — one `| Field | Type | Key / Req | Purpose & rules |` table per table, plus
   each table's indexes, any derived (not stored) values, and the audit columns from the standards
   layer.

Present it for my review. I'll approve it or tell you what to change — this is never finished until
I say so.

Once I approve the design and ask you to build it, **first ask whether I want Access tables (a VBA
sub) or SQL Server (DDL)**, then generate that — carrying the keys, indexes, relationships, and lookup
seed rows. For the Access version, build the tables **in the VBA sub with DAO (`CreateTableDef`), not
with `CREATE TABLE` statements** (the Access engine rejects a `DEFAULT` clause in DDL), and remind me to
**run the sub from a Trusted Location**, or Access blocks the code.
