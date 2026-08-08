# Open Template Scaffolds

**Who reads this:** anyone deciding whether this library is useful to them.

AI-readable templates for the tables behind common Access tasks — a stocktake, a library catalog, and more.

When you start a new part of an Access database, you usually build the same kinds of things from
scratch: the tables, the field names, how they connect. The templates we offer hold those decisions,
already worked out and following solid conventions — so your AI assistant can build them *for* you,
the right way, instead of guessing from a blank page each time.

They're based on real standards, not one-off improvisations. You can use them exactly as they are,
or adjust them to match your own conventions.

And it's more than a set of templates. It's a way to **shape** each one to the job in front of you
— your client, your names, the specifics of this build — without taking it apart.

## New to this? You're who we built it for.

You don't need to be a database designer. If you've outgrown Excel and you're building real Access
applications — especially with an AI assistant helping — this library takes the hardest, most
abstract part, designing the tables, and gives you a running start. You describe what you want in
plain words, the AI builds it, and you look it over and approve or adjust.

## What the templates build — and when to back up first

The library provides templates for most of the kinds of work you do when building an Access database:

- **Design and create relational tables**
- **Write VBA procedures and data macros** — a data macro lives on a table itself and runs whenever a
  record is added, changed, or deleted. Access has other kinds of macro too; those are not covered
  yet, but nothing rules them out.
- **Design and build forms and reports**

Some templates assume you are building new from scratch. Others add features and functionality to a
database you already have. Each template carries information about how it does those things, and your
AI assistant tells you which kind you are getting before it builds anything.

**If you choose a template that retrofits an existing database, you must follow good backup
practices.**

- **Back up before you start.** If your database is split — the tables in one file, the forms and
  code in another — that means the back end *and* every front end.
- **Never run a template on a production database** until you have verified that the template
  produces correct, well-formed results in a backup.
- **Always keep before and after master copies of your databases**, so you can compare what changed
  and go back if you need to.

## Try it in 15 minutes

*For your first try, use this quick path. **Nothing to install.** (The library also ships an
optional server that lets your AI look up templates without opening the files itself — see
[`mcp-server/`](mcp-server/README.md). It changes nothing about what you get, and it can wait.)*

*Never used an AI assistant, or not sure you have one? Read the short **"Before you start"**
section in [`WELCOME.md`](WELCOME.md) first — it takes two minutes and gets you set up.*

1. **Get your own copy of the library — it's just a folder of files.**
    - **Were you handed the library** (say, as a `.zip`)? Just **unzip it** somewhere you'll find
      again (your Documents folder is fine). That unzipped folder is your copy — everything's inside it.
    - **Getting it from GitHub?** Click the green **Code** button → **Download ZIP**, then unzip it —
      or, if you use Git, *fork* it (your own copy you can keep and update) and clone it.
2. **Give your AI assistant the files.** Your assistant needs to be able to read the library's
   files. How depends on which assistant you use — three common cases:
    - **An assistant built into a code editor** (such as Claude Code, or GitHub Copilot in
      VS Code): open the library folder in the editor. The AI assistant can then read everything in it.
    - **A chat assistant that can read files on your computer** (such as Claude Desktop with
      folder access turned on): it can read files, but it doesn't know *where* the library is —
      it won't go looking on its own. Start your chat by telling it the full location, in plain
      words: *"My template library is in `C:\Users\<your name>\Documents\OpenTemplateScaffolds`
      — read what you need from there."* After that, it finds everything itself.
    - **A chat assistant in a browser** (such as Claude or ChatGPT on the web): don't worry about
      giving it files yet. Skip ahead to **step 3 first** — fill in the form there, then copy that
      whole file into the chat and send it. The AI assistant will then tell you, by name, which other
      files it needs. For each one: open it, copy everything (the same select-all, copy, paste as
      in "Before you start"), and paste it into the chat. *(If your assistant can read files on
      your computer — the middle case above — you can skip the copying and just tell it the full
      location of the library folder instead.)*

    Don't worry about getting this perfect. If the AI assistant ever says it can't see a file, just
    paste that file's contents into the chat and carry on.
3. **Open the prompt and fill in four lines.** In the folder you just unzipped, open
   [`prompts/BuildNewTables-StartHere.md`](prompts/BuildNewTables-StartHere.md). Near the top is
   the form — a small box; **these four lines are the only thing you change:**

   ```text
   - Build: <what you want to make, in plain words — e.g. "add stocktake scanning to an inventory app">
   - Standards: default
   - Who this is for: <describe your client — what they call things, and how they want to report>
   - Extra options: none
   ```

   Replace the text inside the `<angle brackets>`; leave everything else as it is. Then copy the
   **whole** prompt (the box *and* the instructions under it) and paste it to your assistant. That's
   all the prompt needs — it tells the AI the rest.
4. **Read back what it builds** — a diagram of the tables, plus a list of every field. Approve it,
   or tell it what to change.

Want to see it work first? Open **[`examples/northwind-stocktake/`](examples/northwind-stocktake/).**
It's the same prompt filled in for a real request, and the tables the AI produced from it. You can compare it
with your own first result. Look for what's the same and what's different.

## Working on a database you already have

The 15-minute path above builds tables from scratch, so there is nothing of yours to protect. Once
you move to a template that retrofits — one that adds VBA, data macros, forms or reports to a
database you already have — set up a working folder first. Everything happens there, on a copy of
your database.

1. **Make a folder for the work.** Anywhere you like; give it a name you will recognise later.
2. **Put your copy of the library inside it.** If you downloaded a `.zip`, unzip it here.
3. **Copy in the databases you want to work on — copies, never the originals.** If your database is
   split, that means the back end *and* a front end. Close every copy in Access before you start; a
   file held open elsewhere will stop a build part-way through.
4. **Open the working folder — the one holding the library and your databases — in your AI
   assistant.** Not your Documents folder, and not the library folder on its own. That is what
   rooting the assistant means: it reads and writes inside that folder, and it can see both the
   library and your databases at once. If you are using a chat assistant in a browser instead, tell
   it the full path of the folder in your first message.
5. **Say what you want in plain words.** For a template that retrofits you do not fill in the
   four-line form from the quick path — you just ask, naming the database: *"Add change auditing to
   `MyDatabase.accdb` in this folder, using the library here."* The assistant finds the matching
   template, asks you the questions it needs answered one at a time, shows you the design, and waits
   for your approval before building anything.

**About the two kinds of server — you need neither.** Everything above works without them. Read
this only if you have already met one.

- **The template library MCP server** ships with the library and lets your AI look up templates and
  standards without reading the files itself. It reads this library's files and **cannot create or
  change anything in your database.** The library ships a configuration file at its root that lets
  some AI clients start it by themselves — but only when the library folder is the one you opened.
  In the arrangement above the library is a subfolder, so nothing appears automatically, and there
  is no error to tell you why. Two ways forward, neither wrong: register it by hand once
  (`mcp-server/setup.ps1` prints exactly what to paste), or skip it.
- **An Access MCP server** is a separate thing, and this library does not ship one: its tools open
  your database and build in it directly. If you have one connected, your AI will say so and ask
  before using it, and you can always say no and import the code yourself instead. If you don't
  have one, nothing changes — your AI hands you the code and tells you how to run it.

## How it works

Two parts, kept separate on purpose:

- **The template** holds the design — the tables, fields, relationships, and rules for one
  particular job. These are decisions already made, and made well.
- **The standards** hold the conventions — how tables and fields are named, how audit columns and
  error handling are done.

When the AI builds your tables, it reads both: the template for *what* to build, the standards for
*how* to name and shape it. **You never have to edit a template to get results; you swap the
standards,** and the same template comes out matching your shop.

Editing a template is for something different: changing the *design* itself — adding a field you
always need, dropping one you never use or changing data types to suit your preferences.
Once you've used a template a few times and know it well, it's yours to customize. Make it your own.

And if nothing in the library matches what you need? The AI will say so plainly — no forced fit —
and offer to design your tables from scratch, following the same conventions, with the same
look-it-over-and-approve flow. Fair warning, though: without a template, you're outside the
library's tested ground. The AI helps, and you still approve everything, but the design is your
own — no template stands behind it.

The library ships with a sensible default set of standards, so you can start the minute you arrive.
However, you will still have the opportunity to make decisions about the shape of the final output.

## What's in here

You only act on a few of these. The rest your AI reads for you, or you can ignore until later.

| In the folder | What it's for | Do you open it? |
|---|---|---|
| **`README.md`** (this file) | Where you start | **Yes — you're reading it** |
| **`WELCOME.md`** | If you have never used an AI assistant — how to get set up | **Yes, if you're new** |
| **`prompts/`** | The prompt you fill in and paste (`BuildNewTables-StartHere.md`) | **Yes — the one you use** |
| **`examples/`** | A finished example, to see it work first | Optional — read to learn |
| `templates/` | The designs your AI builds from | No — the AI reads these |
| `standards/` | The default conventions your AI applies | No — unless you swap your own (later) |
| `CLAUDE.md` | Instructions your AI picks up on its own | No — leave it |
| `mcp-server/` | An optional server that lets your AI look up templates and standards directly. It cannot change your database. | Only if you choose it — has its own README; not needed to start |
| `CONTRIBUTING.md`, `LICENSE` | For people adding templates; the license | No |

### Reading these files

Everything in the library is a Markdown (`.md`) file — plain text with simple formatting marks.
GitHub renders them nicely in your browser, but opened on your desktop (say, in Notepad) they show
the raw marks. If you'd like them to read just as nicely on your desktop, and you don't already
have a Markdown app, two good free options:

- **[MarkText](https://github.com/marktext/marktext/releases)** — a full Markdown editor that
  shows formatted text as you read and write. Free, open source, and actively maintained.
- **A viewer from the Microsoft Store** — search the Store for "markdown viewer" (for example,
  [MarkdownView](https://apps.microsoft.com/detail/9n6pkz6fp1ml)). These are small, free, and open
  a `.md` file with a double-click.

Either way, this is optional — your AI assistant reads these files just fine as they are.

## Making it your own

The default standards are good enough to use on day one — you don't have to change a thing to get
started. When you're ready, replace the files in `standards/` with your own conventions. From then
on, every template you use comes out in your style. The library is yours to keep and adapt; there's
nothing connecting back to us that you have to maintain.

## Contributing and license

Want to add a template or improve one? See [`CONTRIBUTING.md`](CONTRIBUTING.md) — and the developers
who've already shared templates are credited in [`CONTRIBUTORS.md`](CONTRIBUTORS.md). The library is
released under the [MIT license](LICENSE) — free to use, change, and build on.

## More to come

We build to a plan, and we hold ourselves to the same discipline this library is about: we publish
each new piece as we finish and validate it, not before.

The three kinds of work described above — tables, VBA and data macros, forms and reports — are where
this library is headed across the board. They are a direction, not a finished list. Expect more
templates in each of them, and easier ways to put them to work.
