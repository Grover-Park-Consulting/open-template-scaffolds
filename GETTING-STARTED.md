# Getting Started

**Who reads this:** anyone who wants to download Open Template Scaffolds and run their first build.

No installation. No account beyond an AI assistant you already have (or can get free —
see [`WELCOME.md`](WELCOME.md) if you don't have one yet).

## 1. Download the library

Click the green **Code** button above → **Download ZIP**, then unzip it somewhere you'll
find it again (your Documents folder is fine). If you use Git, you can fork and clone instead.

That unzipped folder is your copy of the library — everything you need is inside it.

## 2. Trust a folder for the database

Templates work by having Access run VBA code, so the database the demo creates needs to
sit in a Trusted Location. (The library folder itself is just text files — it needs no
trust setting.)

In Access: **File → Options → Trust Center → Trust Center Settings → Trusted Locations →
Add new location**, then choose the folder your database will go in.

## 3. Open the library folder in your AI assistant

How you do this depends on which assistant you use:

- **Built into a code editor** (Claude Code, GitHub Copilot in VS Code): open the library
  folder in the editor. Done.
- **A chat assistant that can read files on your computer** (Claude Desktop with folder
  access): tell it the full path — *"My template library is in
  `C:\Users\<yourname>\Documents\OpenTemplateScaffolds`. Read what you need from there."*
- **A chat assistant in a browser** (Claude or ChatGPT on the web): skip to step 4, fill in
  the form, and paste that whole file into the chat. The assistant will tell you, by name,
  which other files it needs — open, copy, and paste each one in.

## 4. Fill in four lines and send them

Open [`prompts/BuildNewTables-StartHere.md`](prompts/BuildNewTables-StartHere.md). Near the
top is a small form:

```text
- Build: <what you want to make, in plain words — e.g. "add stocktake scanning to an inventory app">
- Standards: default
- Who this is for: <describe your client — what they call things, and how they want to report>
- Extra options: none
```

For your first try, leave **Standards** and **Extra options** as they are. Replace the two
`<angle bracket>` lines. Copy the **whole** prompt — the form and the instructions below it
— and paste it to your assistant.

## 5. Review what it builds

Your assistant produces a diagram of the tables and a field list. Approve it, or say what to
change — as many rounds as you need. When it's done, read the build record it creates.

## Want to see it work first?

Open [`examples/northwind-stocktake/`](examples/northwind-stocktake/) — the same prompt,
already filled in, with the tables it produced. Compare it against your own first result.

## Working on a database you already have?

This walkthrough builds new tables from scratch. To add features to an existing database
instead, see **"Ready to work on a database you already have?"** in [`README.md`](README.md).
