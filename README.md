# Open Template Scaffolds

**Who reads this:** anyone deciding whether this library is useful to them.

Open Template Scaffolds (OTS) are AI-readable templates for the tables and functions behind common Access tasks — a stocktake, a library catalog, error logging and more. As the library grows, we'll keep adding templates for additional objects and functions.

When you start a new part of an Access database, you usually build the same kinds of things from
scratch: the tables, the field names, how they connect, i.e. how they are related. The templates we offer hold those decisions, already worked out and following solid conventions. Your AI assistant can build those objects *for* you, the right way, instead of starting from a blank page each time.

OTS templates are based on proven standards. You can use them exactly as they are,
or adjust them to match your own conventions. That's the "Open" part of Open Template Scaffolds.

It's also more than a set of templates. It's a way to **shape** each one to the job in front of you
— your client, your names, the specifics of this build.

**Built and tested against [unmateria](https://github.com/unmateria)'s Access MCP server** — the tool
that lets an AI assistant open your database and build in it directly. You do not need it to use this
library; *About the two kinds of server* below says what it does, and how the library works without one.

**Want to download the library and run your first build right now?** See
**[`GETTING-STARTED.md`](GETTING-STARTED.md)** — five short steps, nothing to install.

## New to AI assisted development? We built OTS for you.

You don't need to be an experienced database designer, although it can help if you are. If you've outgrown Excel and you're just starting out building Access applications — especially with an AI assistant helping — this library takes the hardest, most abstract part, designing the tables and logic, and gives you a running start. You describe what you want in plain words, the AI builds it, and you look it over and approve or adjust. That's the "Template" part of Open Template Scaffolds.

## Which AI assistant you use, and where you run it, changes what you see during a build and how you experience it

The same template, run two different ways, can look very different on screen. Some AI assistants,
especially those built into a code editor, show you a great deal as they work; it looks a lot like a progress report. It includes everything they are considering, and every file as it is written, line by line. Others show you almost none of it. Our library can't control that, and neither can your AI assistant. That's the "Scaffold" part of Open Template Scaffolds. Although the actual build process runs on the same scaffold every time, it can look different while running. But more importantly, it can produce different outputs that function the same.

## Two things do not change from build to build

- what you get as the output of a template is the same
- you make decisions about what that output should look like

 **The template's output satisfies the same description each time.**  Whichever AI assistant and working environment you choose, you end up with something that meets the same description.

 That is not a promise that any two runs produce identical code or tables; they will not. That's the probabilistic nature of AI.

 Moreover, variation doesn't mean something went wrong during the build, although is also a possibility. The important point is how the AI responds when something does go wrong *during* a build.

**AI-driven builds can be self-correcting when they have the entire template library available to them.** Different things can go wrong enroute, but an AI-driven build is capable of identifying, diagnosing and remedying problems. When the build finishes, the outcome meets the template's promise.

In addition, each template says where it leaves the builder free to choose an appropriate path. We do promise that each run satisfies everything the template says it delivers.

The templates offer proof of that by running validation checks on every build.

**Your decisions are part of the process.** You are a participant in the build, not an observer. You can't start a build and walk away while it runs because builds depend on you to make those decisions in-flight.

The build process presents anything you have to decide as a question and it waits for your answer. On the other hand, if text goes past on the screen during a build without pausing for your input, nothing is wrong. You aren't necessarily expected to read it, and you haven't missed any decision you needed to make.

**The library doesn't choose your assistant, and it doesn't choose where you run it.** You might work in a code editor, a desktop app, a browser, or at a command prompt, with whichever assistant you have and however you prefer to work. Each feels different: different amounts of text going past, different ways of answering a question. Variation between builds is normal and expected. Variation doesn't necessarily mean something went wrong, although that's also a possibility.

**What we can currently report:** our own trials have run heavily on Claude Code. Our results there have been reliable. However, we haven't put the same volume of runs behind other AI assistants, so we can't yet make the same claim about them. That's a gap in our testing, not a claim that they don't work. In fact, if you use one of the other AIs, we'd love feedback from you about how your AI assistant performed.

**You only need two things to use OTS:** an AI assistant you can run somewhere, and the willingness to work with it. If you don't have one yet, **[`WELCOME.md`](WELCOME.md)** walks you through getting set up. Beyond that the library takes no view — it doesn't recommend an assistant, doesn't recommend where to run it, and doesn't ask you to understand how any of it works underneath.

## What the templates build

The library provides templates for most of the kinds of work you do when building an Access database:

- **Design and create relational tables**
- **Write VBA procedures and data macros**
- **Design and build forms and reports**

Some templates assume you are building new from scratch. Others add features and functionality to a
database you already have. Each template carries information about how it does those things, and your
AI assistant tells you which kind you are getting before it builds anything.

## Backup your databases

**If you choose a template that retrofits an existing database, you must follow good backup practices.**
These are among the very few absolute statements you'll find in our library, for good reason.

- **Always back up BEFORE you start.** If your database is split — the tables in one file, the forms, reports and
  code in another — that means back up the back end *and* front end.
- **Never run a template on a production database** until you have verified that the template
  produces correct, well-formed results in a backup.
- **Always keep before and after master copies of your databases**, so you can compare what changed
  and go back if you need to.

## Verify that every validation check ran

**Each template lists validation checks under *How you validate the template's output*.** Before you
trust a build, look at the build record and make sure your AI assistant ran every single one — not most, and not
"the ones that still applied." All validation checks must be run and the results reported as "Pass", "Fail" or "Not applicable".

An AI session checking its own work can sometimes skip one check and report the skip as a considered decision rather than a gap.
To the AI, that can seem reasonable in the moment. It only shows if you compare what actually ran against the full list the template names. 

If the build record doesn't tell you that all checks *ran* and *passed*, ask directly: *"Did you run every check listed, and did each one pass?"* If the answer is anything but yes, ask why before you treat the build as validated.

## Try it in 15 minutes

*For your first try, we set up a quick start demo. **There is nothing to install for this.** (The library also ships an optional server that lets your AI look up templates without opening the files itself — see
[`mcp-server/`](mcp-server/README.md). It doesn't matter for the quick try out; it can wait.)*

*Never used an AI assistant, or not sure you have one? Read the short **"Before you start"**
section in [`WELCOME.md`](WELCOME.md) first — it takes two minutes and gets you set up.*

Do these five things to set up and run the quick start demo build. If any of those steps aren't clear, detailed instructions are below.

1. Set up a Trusted Location for the Access database the demo will create.
2. Put your copy of the OTS library in a working folder, or start from where it is now.
3. Root your AI session in that working folder.
4. Prompt the AI to build or retrofit an Access database using one of the templates
5. At the end of the build, read the build record it creates

In your prompt, make sure the AI knows where to direct the database it creates.

1. **Your database has to sit in a trusted location because templates work by having Access run VBA code.**
    The library folder itself is nothing but text files, so it needs no trust setting. Put it wherever you like.

    In case you are not familiar, to trust a folder: in Access, File → Options → Trust Center → Trust Center
    Settings → Trusted Locations → Add new location, then choose the folder your database is in.

    Although you can also Enable all Macros, that setting allows all macros to run in any accdb, which we recommend
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

    When you're ready to try out the quick start demo, follow the directions in step 4.

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

## A worked example for the quick start demo

Want to see it work first? We provided an example to show you.

Open **[`examples/northwind-stocktake/`](examples/northwind-stocktake/).**
It's the same prompt already filled in for a real request, and the tables the AI produced from it. You can compare it with your own first result. Look for what's the same and what's different.

## Ready to work on a database you already have?

When you move on to a template that retrofits an existing Access database by adding additional tables, VBA, data macros, forms or reports, set up a working folder in a Trusted Location for it first. Everything happens in that working folder on a throwaway copy of your database.

1. **Make a folder for the work.** Create your working folder anywhere you like; just give it a name
   you will recognise later.
2. **Put your copy of the library inside it.** If you downloaded a `.zip`, unzip it here.
3. **Copy the databases you want to work on into the working folder with the library files. Always use copies, never the originals.** If your database is split, that means the back end *and* a front end. Close every copy in Access before you start; a file held open elsewhere will stop a build part-way through.
4. **Open the working folder inside your AI assistant. That's the one holding the library and your databases.**  Not your Documents folder, and not the library folder on its own.
   That's what "rooting the assistant" means: it reads and writes inside that folder, and it can see both the
   library and your databases at once. If you are using a chat assistant in a browser instead, tell
   it the full path of the folder in your first message.
5. **Say what you want in plain words.** For a template that retrofits existing databases, do not use the
   four-line form from the quick path. 
   Just ask: *"Add change auditing to `MyDatabase.accdb` in this folder, using the library here."* where `MyDatabase.accdb` is the one you want to retrofit. The assistant finds the matching
   template, asks you the questions it needs answered, one at a time, shows you the design, and waits
   for your approval before building anything.

   **Sample prompts for different versions of the templates** are available in prompts/. Use them to get your AI started. You can modify them as you and your AI gain familiarity and experience.

**About the two kinds of server**
You may have heard references to MCP servers, and you may also be wondering if they are involved in the OTS Library. Yes, the templates come with the ability to work with two MCP Servers.

We ship one in the library. The other is optional; you may already have one installed, or you may choose to install one if the benefits it offers seem worthwhile to you.

 **You need neither.** Everything above works with or without them. You only need to read this section if you already have one and want to know how it's used in the OTS library. Of course, you may want to read anyway to see what they are all about.

- **OTS's template library MCP server** ships with the library. Its job is to let your AI look up templates and
  standards without reading the files itself. It only reads this library's files; it **cannot create or change anything in your database.** The library ships a configuration file at its root that lets
  some AI clients start the library MCP server by themselves, but that only happens when you open the library folder.
  In the examples above the library is a subfolder, so nothing appears automatically, and there
  is nothing that tells you why. You have two ways forward: 
  - register it by hand (`mcp-server/setup.ps1` prints exactly what to paste)
  - skip it. As previously stated, the library runs with or without it.

- **An Access MCP server** is a separate tool. Its tools open your database and **build in it directly**. Some developers prefer this method; some don't. It's up to you.
However, this library does not ship with one. We do look for one on your system, and if you have one connected, your AI will say so and ask if you want to use it. You can always say no. It's on you, then, to import the code produced by the library yourself instead. If you don't have one, your AI simply hands you the code and tells you how to run it.

  **With thanks to [unmateria](https://github.com/unmateria).**
  [MCP-Access](https://github.com/unmateria/MCP-Access) is the Access MCP server this library is
  developed and tested against, and it is what makes the build-it-for-you route something we can
  stand behind rather than describe. It is their project, not ours, and they have been generous with
  it. As far as this library is concerned any Access MCP server works the same way; we name this one
  because it is the one we owe.

## How it works

There are two primary parts to an OTS build; we kept them separate because they serve different purposes:

- **The template itself** holds the design: that includes descriptions of the tables, fields, and relationships, plus the rules for one particular job. It can also include VBA and Data Macros and other design patterns. These decisions are already made, and, we believe, made well.

- **The standards** hold conventions: they specify
  - how tables and fields are named
  - how audit columns in tables are implemented
  - how error handling and logging are done in VBA
  - many other similar conventions. 
  
  Standards are not unique to a specific database. The creator of the Open Template Scaffolds project defined the standards shipped in the library. When you become comfortable with the library, you are free to merge or replace those standards with your own. One size definitely does not fit all in the Access world.

When the AI builds your tables, it reads both the template and the standards.

- It looks into the template for *what* to build.
- It consults the standards for *how* to name and shape it.

 **Standards are customizable.** You never have to edit a template to get different results; you edit or swap the standards, and the same template comes out matching your preferred standards.

 **Templates are customizable, but with a different impact.** Editing a template changes the *design* itself
 — adding a field you always need in a table, dropping one you never use or changing our data types to suit your preferences. Once you've used a template a few times and know it well, it's yours to customize. Make it your own.

## OTS is also a method for building

 What if nothing in the library matches what you need? Our intent is never to force a fit to an existing template. If that happens, the AI can offer to design your tables, code and forms, but it doesn't have to do so from scratch. Your AI has access to all of the standards and core templates in the library, along with any prior build-records. That means your AI follows the conventions you've already established. It delivers the same decision-guided, look-it-over-and-approve flow for an ad-hoc design as it would for a core template, just breaking new ground in the process. That, in turn, produces build records that feed back into future builds.

## Build records add up; keep them

Every build produces a build record: what was built, what was checked, and anything that didn't go
as the template described. Don't discard these — your AI assistant puts them in `build-records/<template-name>/` in your copy of the library. **That folder holds one per template, not a new folder for every build** — a template's folder is where all of its history lives, so your AI assistant (yours, or the next person's) can read what past builds of that template ran into instead of repeating the work of finding out.

This holds across templates too, not just within one. A build record for one template can save real
time on another that shares a technique — a Data Macro convention, say. The library ships
`build-records/` empty; it fills up as you use it.

## What's in here

You only need to act on a few of these items. Your AI reads the rest for you or you can ignore them until later.

| In the folder | What it's for | Do you open it? |
|---|---|---|
| **`README.md`** (this file) | Where you start | **Yes, you're reading it** |
| **`WELCOME.md`** | If you have never used an AI assistant — how to get set up | **Yes, if you're new to AI** |
| **`prompts/`** | The prompt you fill in and paste (`BuildNewTables-StartHere.md`) | **Yes, the one you use to start** |
| **`examples/`** | A finished example, to see it work first | Optional, read to learn |
| `templates/` | The designs your AI builds from | Not necessary, only the AI needs to read these |
| `build-records/` | Accumulated records from past builds, one folder per template | Not necessary, your AI reads it before building and adds to it after |
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
started. If you like you can edit them or replace the files in `standards/` with your own conventions. From then on, every template you use comes out in your style. The library is yours to keep and adapt; there's nothing connecting back to us that you have to maintain.

**One warning about replacing the standard files.** The templates expect all of the sections in the standards to be available. So, if you choose to replace them, you'll need to include all of those sections in the standards, if only as placeholders.

## Contributing and license

Want to add a template or improve one? See [`CONTRIBUTING.md`](CONTRIBUTING.md). The developers
who've already shared templates are credited in [`CONTRIBUTORS.md`](CONTRIBUTORS.md). The library is
released under the [MIT license](LICENSE) — free to use, change, and build on.

## More to come

We build to a plan, and we hold ourselves to the same discipline this library is about: we publish
each new piece as we finish and validate it.

The three kinds of work described above — tables, VBA and data macros, forms and reports — are where
this library is headed in our vision for the future. They are a direction, not a finished list. Expect more templates in each of them, and easier ways to put them to work.
