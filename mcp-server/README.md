# Open Template Scaffolds — The Template Library MCP Server

**Who reads this:** anyone setting up the server so their AI assistant can read the library.

The template library MCP server is a thin, dependency-light **Model Context Protocol (MCP)** server that exposes the Open Template Scaffolds library to MCP-capable AI clients. It reads the library's markdown (`templates/`, `standards/`) as the single source of truth. It ships *inside* the library itself. Therefore, when you make your own copy of the library, the server comes with it. There's nothing separate to download or keep up to date.

**It never touches your database.** Every tool here reads this library's own files. It cannot create a table, import code, or change anything in an `.accdb`. The one tool that opens a database at all, `check_compatibility`, opens it **only** to read what is already there.

Building into a database is the job of **an Access MCP server**. An Access MCP Server is a different thing.
This library does not ship one. However, if you already have one connected, your AI will say so and ask before using it.

## Tools in the template library MCP server

| Tool | Purpose |
|---|---|
| `list_templates` | List the library's templates + key metadata |
| `search_templates` | Find templates by keyword / domain / type, rating the relevance of every result |
| `get_template` | A template composed with the active standards layer (kept separate for easy swapping) |
| `get_standards` | The standards layer which powers from-scratch designs when no template fits |
| `validate` | Validate the format of templates per `templates/_template-schema.md` |
| `check_compatibility(template, db_path)` | Confirm a template's required tables/fields exist in a host Access DB (optional — see below) |

---

## Is this the path for you?

This MCP server is **optional.** It buys you speed of lookup, not a different
result. With it, your AI can find, compose, and check templates without you pointing it at files.

It needs Python and a couple of terminal commands.

**You don't need it to get started** — the [main quick-start](../README.md) uses the templates with just your AI assistant and a copy-paste prompt requiring no setup. However, you can come back here when you want the AI to discover, compose, and validate templates automatically. If you want, you could have someone help you set it up.

## What you'll need to use the template library MCP server

- **Python 3.10 or newer**, installed from [python.org](https://www.python.org/) (on Windows, tick
  *"Add Python to PATH"* during the install).
- A **terminal** — on Windows that's **PowerShell** (search "PowerShell" in the Start menu); on Mac,
  **Terminal**. This is a separate window where you type commands; **it is not the AI chat.**

## Easiest setup — run the helper script

This folder ships a one-time helper, called `setup.ps1`, that checks Python, installs the dependencies, and
**prints the exact block of text you'll need, with the real path to *this* copy already filled in.**
No guessing, no editing a placeholder path.

1. In File Explorer, open this `mcp-server` folder.
2. Click in the address bar, type `powershell`, and press **Enter** — a PowerShell window opens
   already pointed at this folder.
3. Run:
   ```
   powershell -ExecutionPolicy Bypass -File .\setup.ps1
   ```

Then follow along with what it prints. **It never changes a client config for you.** It only shows you what to
paste, so nothing on your machine is altered without your say-so.

*Prefer to do it by hand?* Entering in the terminal `pip install -r requirements.txt` from this folder installs the only two dependencies (`mcp` and `pyyaml`, both pure Python, cross-platform).
After the setup, follow the registration steps below.

## Registering the server with your AI client

The server doesn't do anything on its own — an **AI client** (the app you chat with) starts it and
uses its tools. "Registering" it just means telling your client how to start it. You do this **once**
per client.

The following registration steps for Claude clients will be similar to most other clients.

**Claude Code — automatic, nothing to paste.** This library ships a `.mcp.json` at its root, so when
you open the library folder in Claude Code it offers the `open-template-scaffolds` server on its own.
Approve the one-time trust prompt (or type `/mcp` in a session and approve it there) and the tools are
live. There's no path to enter.

**Claude Desktop — add a few lines to a small settings file.** Desktop doesn't read the library's
`.mcp.json`, so you tell it about the server yourself.

There's no switch to flick or box to tick for this one. Claude Desktop keeps this particular setting
in a small text file, and it gives you a button that opens that file for you. You add a few lines,
save it, and close it. That's the whole job.

- **The easy way:** run `setup.ps1` (above). It prints the exact lines to add, with the real path
  already filled in. Copy them, then follow the steps below to open the file and put them in.
- **Opening the file:** in Claude Desktop, go to **Settings → Developer → Edit Config**. That button
  opens `claude_desktop_config.json` in a text editor (on Windows the file lives under
  `%APPDATA%\Claude\`). Add an `open-template-scaffolds` entry inside `mcpServers`, with
  `"command": "python"` and `"args"` set to the full path of *this copy's* `server.py`. Save the
  file, then close Claude Desktop and open it again.

**Be very careful not to change anything else in `claude_desktop_config.json`. It would be a good idea to make a back up copy before editing it.**

**Any other MCP client** — the pattern is identical: a server named `open-template-scaffolds`, started
with `python <full path>\server.py`. `setup.ps1` prints that exact path; check your client's own
"MCP servers" documentation for where to register it — the model is always the same (the client runs
the command; the tools appear).

---

## Dependencies and your own copy

The **core tools** (discovery, `get_template`, `validate`) need only `mcp` + `pyyaml` — nothing
platform-specific, so **your copy runs anywhere Python does.**

There is one exception: `check_compatibility`. Reading an Access `.accdb` requires an Access driver (Windows + ACE), so `check_compatibility` is an **optional** capability that quietly does nothing when the required Access driver isn't installed. The easy-to-copy core never depends on it. Given that you are installing the OTS library of Access templates, it's highly likely you'll already have the required Access driver available.
