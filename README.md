# Open Template Scaffolds

**Who reads this:** anyone deciding whether this library is useful to them.

Open Template Scaffolds (OTS) are AI-readable templates for the tables behind common Access tasks — a stocktake, a library catalog, and more. As the library grows, we'll keep adding templates for additional objects and functions.

When you start a new part of an Access database, you usually build the same kinds of things from
scratch: the tables, the field names, how they connect, i.e. how they are related. The templates we offer hold those decisions, already worked out and following solid conventions. Your AI assistant can build those objects *for* you, the right way, instead of starting from a blank page each time.

OTS templates are based on proven standards. You can use them exactly as they are,
or adjust them to match your own conventions. That's the "Open" part of Open Template Scaffolds.

It's also more than a set of templates. It's a way to **shape** each one to the job in front of you
— your client, your names, the specifics of this build.

## New to AI assisted development? We built OTS for you.

You don't need to be an experienced database designer, although it can help if you are. Even if you've outgrown Excel and you're starting out building real Access applications — especially with an AI assistant helping — this library takes the hardest, most abstract part, designing the tables, and gives you a running start. You describe what you want in plain words, the AI builds it, and you look it over and approve or adjust. That's the "Template" part of Open Template Scaffolds.

## Which AI assistant you use, and where you run it, changes what you see during a build and how you experience it

The same template, run two different ways, can look very different on screen. Some AI assistants,
especially those built into a code editor, show you a great deal as they work; it looks a lot like a progress report. It includes everything they are considering, and every file as it is written, line by line. Others show you almost none of it. Our library can't control that, and neither can your AI assistant. That's the "Scaffold" part of Open Template Scaffolds. Although the actual build process runs on the same scaffold every time, it can look different while running. But more importantly, it can produce different outputs that function the same.

**Two things do not change. What you get as the output, and how you make decisions.** The build process presents anything you have to decide as a question and it waits for your answer. That means that if text goes past on the screen during a build without pausing for your input, nothing is wrong. You aren't necessarily expected to read it, and you haven't missed any decision you needed to make.

**The library doesn't choose your assistant, and it doesn't choose where you run it.** You might work in a code editor, a desktop app, a browser, or at a command prompt, with whichever assistant you have and however you prefer to work. Each feels different, different amounts of text going past, different ways of answering a question, and this is important to remember, different things that can go wrong.

All of that is yours to pick.

 **The library's responsibility is this: Whichever AI assistant and working environment you choose, you end up with something that meets the same description.** That is not a promise that any two runs produce identical code; they will not. That's the probabilistic nature of AI. Each template says where it leaves the builder free to choose an appropriate path. However, we do promise that both runs satisfy everything the template says it delivers.

**You only need two things to use OTS:** an AI assistant you can run somewhere, and the willingness to work with it. If you don't have one yet, **[`WELCOME.md`](WELCOME.md)** walks you through getting set up. Beyond that the library takes no view — it doesn't recommend an assistant, doesn't recommend where to run it, and doesn't ask you to understand how any of it works underneath.

## What the templates build

The library provides templates for most of the kinds of work you do when building an Access database:

- **Design and create relational tables**
- **Write VBA procedures and data macros** — a data macro lives on a table itself and runs whenever a
  record is added, changed, or deleted. Access has other kinds of macro too; we plan to add more templates, some of which might include those macros.
- **Design and build forms and reports**

Some templates assume you are building new from scratch. Others add features and functionality to a
database you already have. Each template carries information about how it does those things, and your
AI assistant tells you which kind you are getting before it builds anything.

## When to back up

**If you choose a template that retrofits an existing database, you must follow good backup
practices.** These are among the very few absolute statements you'll find in our library, for good reason.

- **Always back up BEFORE you start.** If your database is split — the tables in one file, the forms, reports and
  code in another — that means back up the back end *and* every front end.
- **Never run a template on a production database** until you have verified that the template
  produces correct, well-formed results in a backup.
- **Always keep before and after master copies of your databases**, so you can compare what changed
  and go back if you need to.

## Try it in 15 minutes

*For your first try, we set up a quick path. **There is nothing to install for this.** (The library also ships an
optional server that lets your AI look up templates without opening the files itself — see
[`mcp-server/`](mcp-server/README.md). It doesn't matter for the quick try out; it can wait.)*

*Never used an AI assistant, or not sure you have one? Read the short **"Before you start"**
section in [`WELCOME.md`](WELCOME.md) first — it takes two minutes and gets you set up.*

1. **Your database has to sit in a trusted location because templates work by having Access run VBA code.**
    The library folder itself is nothing but text files, so it needs no trust setting. Put it wherever you like.

    In case you are not familiar, to trust a folder: in Access, File → Options → Trust Center → Trust Center
    Settings → Trusted Locations → Add new location, then choose the folder your database is in.

    Although you can also Enable all Macros, that allows all macros to run in any accdb, which we recommend
    you do not allow.
2. **Get your own copy of the OTS library — it's just a folder of files.**
    - **Did someone hand you the library** (say, as a `.zip`)? Just **unzip it** somewhere you'll find it
      again (your Documents folder is fine). That unzipped folder is your copy of the library — everything you need is inside it.
    - **Do you want to download it from GitHub?** Click the green **Code** button → **Download ZIP**, then unzip it, or if you are familiar with and use Git, *fork* it (your own copy you can keep and update) and clone it.
3. **Give your AI assistant the OTS files.** Your assistant needs to be able to read the library's
   files. How you do that depends partly on which assistant you use. Here are three common cases:
    - **An assistant built into a code editor** (such as Claude Code, or GitHub Copilot in
      VS Code): open the library folder in the editor. The AI assistant can then read everything in that folder.
    - **A chat assistant that can read files on your computer** (such as Claude Desktop with
      folder access turned on): the chat can read files, but it doesn't yet know *where* the library is. Chat is not supposed to go looking on its own and it won't do that. Start your chat by telling it the full location, in plain words: *"My template library is in `C:\Users\<yourname>\Documents\OpenTemplateScaffolds`
      for example. Read what you need from there."* After that, it finds everything itself.
    - **A chat assistant in a browser** (such as Claude or ChatGPT on the web): this case is different,
      so don't worry about giving it files yet. Skip ahead to **step 4 first** — fill in the form there, then copy that whole file into the chat and send it. The AI assistant will then be able to tell you, by name, which other files it needs. For each one: open it, copy everything (the same select-all, copy, paste as
      in "Before you start"), and paste it into the chat. *(Remember, if your assistant can read files on
      your computer — the middle case above — you can skip the copying and just tell it the full
      location of the library folder instead.)*

    Don't worry about getting this perfect right out of the gate.

    If the AI assistant ever says it can't see a file, just paste that file's contents into the chat and carry on.

    When you're ready to try it out, follow the directions in step 4.

4. **Open a prompt and fill in four lines.** In the folder you just unzipped, we have an example to help you
     get started by building new tables for a database. Open [`prompts/BuildNewTables-StartHere.md`](prompts/BuildNewTables-StartHere.md).
     Near the top is the form — a small box; **in a normal build, these four lines are the only thing you change:**

   ```text
   - Build: <what you want to make, in plain words — e.g. "add stocktake scanning to an inventory app">
   - Standards: default
   - Who this is for: <describe your client — what they call things, and how they want to report>
   - Extra options: none
   ```

   To keep it simple for this first try, we suggest you accept *default* standards and *none* for extra
   options. We've already put those two values in the form for you, leaving the other two for you to fill in.

   Replace the text inside the `<angle brackets>`; leave everything else as it is. Then copy the
   **whole** prompt (the box *and* the instructions under it) and paste it to your assistant. That's
   all the prompt needs. It tells the AI the rest.

5. **Read what your assistant builds** It will produce a diagram of the tables you requested, plus a list 
    of every field. Approve it, or tell it what to change. You'll have as many chances to modify as you need. The AI assistant will keep working as long as you need to get it right for you.

Want to see it work first? We provided an example to show you.

Open **[`examples/northwind-stocktake/`](examples/northwind-stocktake/).**
It's the same prompt already filled in for a real request, and the tables the AI produced from it. You can compare it with your own first result. Look for what's the same and what's different.

## Working on a database you already have

The 15-minute path above builds tables from scratch, so there is nothing of yours to protect. Once
you move to a template that retrofits — one that adds VBA, data macros, forms or reports to a
database you already have — set up a working folder first. Everything happens in that working folder
on a throwaway copy of your database.

1. **Make a folder for the work.** Create your working folder anywhere you like; just give it a name
   you will recognise later.
2. **Put your copy of the library inside it.** If you downloaded a `.zip`, unzip it here.
3. **Copy the databases you want to work on into the working folder with the library files. Always use copies, never the originals.** If your database is split, that means the back end *and* a front end. Close every copy in Access before you start; a file held open elsewhere will stop a build part-way through.
4. **Open the working folder in your AI assistant. That's the one holding the library and your databases.** Not your Documents folder, and not the library folder on its own.
   That's what "rooting the assistant" means: it reads and writes inside that folder, and it can see both the
   library and your databases at once. If you are using a chat assistant in a browser instead, tell
   it the full path of the folder in your first message.
5. **Say what you want in plain words.** For a template that retrofits existing databases, you do not fill in the
   four-line form from the quick path — you just ask: *"Add change auditing to
   `MyDatabase.accdb` in this folder, using the library here."* where `MyDatabase.accdb` is the one you want to retrofit. The assistant finds the matching
   template, asks you the questions it needs answered one at a time, shows you the design, and waits
   for your approval before building anything.

**About the two kinds of server** 
You may have heard references to MCP servers, and you may also be wondering if they are involved in the OTS Library. Yes, the templates come with two MCP Servers.

 However, **you need neither.** Everything above works with or without them. You only need to read this section if you already have one and want to know how it's used in the OTS library. Of course, you may want to read anyway to see what they are all about.

- **The template library MCP server** ships with the library. Its job is to let your AI look up templates and
  standards without reading the files itself. It only reads this library's files; it **cannot create or change anything in your database.** The library ships a configuration file at its root that lets
  some AI clients start the library MCP server by themselves, but that only happens when you open the library folder.
  In the examples above the library is a subfolder, so nothing appears automatically, and there
  is nothing that tells you why. You have two ways forward: 
  - register it by hand (`mcp-server/setup.ps1` prints exactly what to paste)
  - skip it. As previously stated, the library runs with or without it.

- **An Access MCP server** is a separate tool. Its tools open your database and build in it directly. Some developers prefer this method; some don't. It's up to you. 
However, this library does not ship with one. We do look for one, and if you have one connected, your AI will say so and ask before using it. You can always say no. It's on you, then, to import the code produced by the library yourself instead. If you don't have one, your AI simply hands you the code and tells you how to run it.

## How it works

There are two parts to an OTS build; we kept them separate because they serve different purposes:

- **The template itself** holds the design: that includes descriptions of the tables, fields, and relationships, plus the rules for one particular job. These decisions are already made, and, we believe, made well.

- **The standards** hold conventions: they specify how tables and fields are named, how audit columns in tables are implemented, how error handling and logging are done in VBA, and many other similar conventions. Standards are not unique to a database. The standards shipped in the libary were defined by the creator of the Open Template Scaffolds project. When you become comfortable with the library, you are free to merge or replace those standards with your own. One size definitely does not fit all.

When the AI builds your tables, it reads both the template and the standards. 

- It looks into the template for *what* to build.
- It consults the standards for *how* to name and shape it.

 **You never have to edit a template to get different results; you edit or swap the standards,** and the same template comes out matching your preferred standards.

Editing a template is for something different: changing the *design* itself — adding a field you
always need in a table, dropping one you never use or changing our data types to suit your preferences.
Once you've used a template a few times and know it well, it's yours to customize. Make it your own.

What if nothing in the library matches what you need? The AI will say so plainly — our intent is never to force a fit. If that happens, the AI is authorized to offer to design your tables from scratch, following the same conventions, with the same look-it-over-and-approve flow. Fair warning, though: without a template, you're outside the library's tested ground. The AI helps, and you still approve everything, but the design is your
own. No template stands behind it.

## What's in here

You only need to act on a few of these items. Your AI reads the rest for you or you can ignore them until later.

| In the folder | What it's for | Do you open it? |
|---|---|---|
| **`README.md`** (this file) | Where you start | **Yes, you're reading it** |
| **`WELCOME.md`** | If you have never used an AI assistant — how to get set up | **Yes, if you're new** |
| **`prompts/`** | The prompt you fill in and paste (`BuildNewTables-StartHere.md`) | **Yes, the one you use  to start** |
| **`examples/`** | A finished example, to see it work first | Optional, read to learn |
| `templates/` | The designs your AI builds from | Not necessary, only the AI reads these |
| `standards/` | The default conventions your AI applies | Not necessary unless you edit or swap in your own (later) |
| `CLAUDE.md` | Instructions your AI picks up on its own | No, leave it to the AI|
| `mcp-server/` | An optional server that lets your AI look up templates and standards directly. It cannot change your database. | Only if you choose it. It has its own README; it's not needed to start |
| `CONTRIBUTING.md`, `LICENSE` | For people adding templates; the license | Not until you are ready to share a template you created |

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

The default standards are good enough to use on day one; you don't have to change a thing to get
started. When you're ready, if you like you can edit them or replace the files in `standards/` with your own conventions. From then on, every template you use comes out in your style. The library is yours to keep and adapt; there's nothing connecting back to us that you have to maintain.

**One warning about replacing the standard files.** The templates expect all of the sections in the standards to be available. So, if you choose to replace them, you'll need to include all of those sections in the standards, if only as placeholders.

## Contributing and license

Want to add a template or improve one? See [`CONTRIBUTING.md`](CONTRIBUTING.md). The developers
who've already shared templates are credited in [`CONTRIBUTORS.md`](CONTRIBUTORS.md). The library is
released under the [MIT license](LICENSE) — free to use, change, and build on.

## More to come

We build to a plan, and we hold ourselves to the same discipline this library is about: we publish
each new piece as we finish and validate it.

The three kinds of work described above — tables, VBA and data macros, forms and reports — are where
this library is headed in our vision for the future. They are a direction, not a finished list. Expect more
templates in each of them, and easier ways to put them to work.
