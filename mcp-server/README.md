<!-- DRAFT — pre-release. Finalize at B4-17 before publishing. -->

# GPC Template Library — Reference MCP Server

A thin, dependency-light **Model Context Protocol (MCP)** server that exposes the GPC Template
Library to MCP-capable AI clients. It reads the library's markdown (`templates/`, `standards/`) as the
single source of truth, and it ships *inside* the library itself — so when you make your own copy of
the library, the server comes with it. There's nothing separate to download or keep up to date.

## Tools (built across phase B3)

| Tool | Purpose | Status |
|---|---|---|
| `list_templates` | The library's templates + key metadata | scaffold |
| `search_templates` | Find templates by keyword / domain / type | planned (B3-11) |
| `get_template` | A template composed with the active standards layer | planned (B3-12) |
| `validate` | Format-only validation per `templates/_template-schema.md` | planned (B3-13) |
| `check_compatibility(template, db_path)` | Confirm a template's required tables/fields exist in a host Access DB | planned (B3-14, optional — see below) |

---

## Is this the path for you?

This server is the **more powerful, but more involved** way to use the library. It needs Python and a
couple of terminal commands. **If that's new to you, you don't need it to get started** — the
[main quick-start](../README.md) uses the templates with just your AI assistant and a copy-paste
prompt, no setup at all. Come back here when you want the AI to discover, compose, and validate
templates automatically — or have a more technical colleague set this up once.

## What you'll need

- **Python 3.10 or newer**, installed from [python.org](https://www.python.org/) (on Windows, tick
  *"Add Python to PATH"* during the install).
- A **terminal** — on Windows that's **PowerShell** (search "PowerShell" in the Start menu); on Mac,
  **Terminal**. This is a separate window where you type commands; **it is not the AI chat.**

## Easiest setup — run the helper script

This folder ships a one-time helper, `setup.ps1`, that checks Python, installs the dependencies, and
**prints the exact configuration to paste — with the real path to *this* copy already filled in.** No
guessing, no editing a placeholder path.

1. In File Explorer, open this `mcp-server` folder.
2. Click in the address bar, type `powershell`, and press **Enter** — a PowerShell window opens
   already pointed at this folder.
3. Run:
   ```
   powershell -ExecutionPolicy Bypass -File .\setup.ps1
   ```

Then follow what it prints. **It never changes a client config for you** — it only shows you what to
paste, so nothing on your machine is altered without your say-so.

*Prefer to do it by hand?* `pip install -r requirements.txt` from this folder installs the only two
dependencies (`mcp` and `pyyaml`, both pure Python, cross-platform); then see the registration steps
below.

## Registering the server with your AI client

The server doesn't do anything on its own — an **AI client** (the app you chat with) starts it and
uses its tools. "Registering" it just means telling your client how to start it. You do this **once**
per client.

**Claude Code — automatic, nothing to paste.** This library ships a `.mcp.json` at its root, so when
you open the library folder in Claude Code it offers the `gpc-template-library` server on its own.
Approve the one-time trust prompt (or type `/mcp` in a session and approve it there) and the tools are
live. There's no path to enter.

**Claude Desktop — paste one block.** Desktop doesn't read the library's `.mcp.json`, so it needs a
manual entry. The easy way: run `setup.ps1` (above) and paste the block it prints — the real path is
already filled in. By hand: open **Settings → Developer → Edit Config** (this opens
`claude_desktop_config.json`; on Windows it's under `%APPDATA%\Claude\`), and add a
`gpc-template-library` entry inside `mcpServers`, with `"command": "python"` and `"args"` set to the
full path of *this copy's* `server.py`. Save and restart Claude Desktop.

**Any other MCP client** — the pattern is identical: a server named `gpc-template-library`, started
with `python <full path>\server.py`. `setup.ps1` prints that exact path; check your client's own
"MCP servers" documentation for where to register it — the model is always the same (the client runs
the command; the tools appear).

---

## Dependencies and your own copy

The **core** (discovery, `get_template`, `validate`) needs only `mcp` + `pyyaml` — nothing
platform-specific, so **your copy runs anywhere Python does.** `check_compatibility` is the one
exception: reading an Access `.accdb` needs an Access driver (Windows + ACE), so it's an **optional**
capability that quietly does nothing when the driver isn't present. The easy-to-copy core never
depends on it.
