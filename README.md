# Open Template Scaffolds

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

## Try it in 15 minutes

*There are two ways to use this library — for your first try, use this quick path. **Nothing to
install.** (A more automated "server" path is optional and comes later — see
[`mcp-server/`](mcp-server/README.md).)*

1. **Get your own copy of the library — it's just a folder of files.**
    - **Were you handed the library** (say, as a `.zip`)? Just **unzip it** somewhere you'll find
      again (your Documents folder is fine). That unzipped folder is your copy — everything's inside it.
    - **Getting it from GitHub?** Click the green **Code** button → **Download ZIP**, then unzip it —
      or, if you use Git, *fork* it (your own copy you can keep and update) and clone it.
2. **Give your AI assistant the files.** Your assistant needs to be able to read the library's
   files. How depends on which assistant you use:
    - **An assistant built into a code editor** (such as Claude Code, or GitHub Copilot in
      VS Code): open the library folder in the editor. The assistant can then read everything in it.
    - **A chat assistant in a browser or app** (such as Claude or ChatGPT): paste in the contents
      of the files the prompt asks for — the template and the standards — before you run the prompt.

    Don't worry about getting this perfect. If the assistant ever says it can't see a file, just
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

And if nothing in the library matches what you need? You're still covered. The AI will say so
plainly — and, with your go-ahead, design your tables from scratch following the same conventions,
with the same look-it-over-and-approve flow. A missing template never leaves you stranded.

The library ships with a sensible default set of standards, so you can start the minute you arrive.

## What's in here

You only act on a few of these. The rest your AI reads for you, or you can ignore until later.

| In the folder | What it's for | Do you open it? |
|---|---|---|
| **`README.md`** (this file) | Where you start | **Yes — you're reading it** |
| **`prompts/`** | The prompt you fill in and paste (`BuildNewTables-StartHere.md`) | **Yes — the one you use** |
| **`examples/`** | A finished example, to see it work first | Optional — read to learn |
| `templates/` | The designs your AI builds from | No — the AI reads these |
| `standards/` | The default conventions your AI applies | No — unless you swap your own (later) |
| `CLAUDE.md` | Instructions your AI picks up on its own | No — leave it |
| `mcp-server/` | The optional, more-automated server path | Only if you choose it — has its own README; not needed to start |
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

Want to add a template or improve one? See [`CONTRIBUTING.md`](CONTRIBUTING.md). The library is
released under the [MIT license](LICENSE) — free to use, change, and build on.

## More to come

We build to a plan, and we hold ourselves to the same discipline this library is about: we publish
each new piece as we finish and validate it, not before. More templates — and easier ways to put
them to work — are on the way.
