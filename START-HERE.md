# Start Here — GPC Template Library

*A quick note before you dive in. — George*

Thanks for taking this for a test drive. This is an early (**alpha**) release, so you're among the
first to use it — and your reactions are exactly what I'm after. Nothing here is set in stone.

## What this is

A small library of **AI-readable templates** for the tables behind common Microsoft Access tasks — a
stocktake, a library catalog, and so on. Instead of designing tables from a blank page every time, you
describe what you want in plain words, your AI assistant builds the tables from a template, and you
review and approve. The design decisions and naming conventions are already worked out and baked in.

## Two ways to use it — pick one

**1. Quick path — nothing to install (start here).**
All you need is an AI assistant that can read files (Claude, Claude Code, ChatGPT, Copilot, …).

- Unzip this package somewhere you'll find again (your Documents folder is fine).
- Open **`README.md`** and follow **"Try it in 15 minutes."**

That's the whole setup. Most testers should start here.

**2. Server path — optional, more automated.**
This adds a small "MCP server" that lets the AI discover and validate templates automatically. It needs
Python and a one-time registration — you don't need it for the trial, so try the quick path first.

- **Claude Code:** just open the unzipped folder in it. The library includes a `.mcp.json`, so Claude
  Code offers the server on its own — approve the one-time prompt and you're set. Nothing to type.
- **Claude Desktop (or another client):** open the **`mcp-server/`** folder, run **`setup.ps1`**, and
  paste the configuration it prints (it fills in the right path for you). Full details in
  **`mcp-server/README.md`**.

## What's in the box

You only act on a couple of these; the rest your AI reads for you. `README.md` has the full map — in
short:

- **`README.md`** — the 15-minute quick start; your main entry point.
- **`prompts/`** — the prompt you fill in and paste to your AI.
- **`examples/`** — a finished example, to see it work first.
- **`templates/`, `standards/`** — what the AI reads to build your tables (you don't edit these to start).
- **`mcp-server/`** — the optional server path (`setup.ps1` + its own README).
- **`CLAUDE.md`, `CONTRIBUTING.md`, `LICENSE`** — AI instructions, how to contribute, the license.

## What I'd love to hear

- Where did it work well? Where did it stumble or surprise you?
- Anything confusing in the setup or the prompt?
- Did the tables it built match what you expected?

Rough notes are perfect — even "I got stuck right here" is gold. Email me back whenever.

Thanks again,
George
