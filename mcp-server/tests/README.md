# Checking that the library's tools work — the self-check

The library's optional server comes with a built-in self-check: a set of small
automatic tests that confirm the tools behave the way they should. Think of it
like the compile check in Access — you run it, and it either says *everything's
fine* or points at what's broken.

**Most people never need to run this.** It matters in two situations:

- You **changed something** — in the `mcp-server` folder, or in a template —
  and want to know you didn't break anything.
- You want to **confirm your copy works** before relying on it.

## Before you start

You need the same things the server itself needs — Python and its two
supporting packages. If you already followed the server setup (the
[`mcp-server` README](../README.md), or its `setup.ps1` helper), you have them.
If not, do that setup first and come back.

## Running the self-check

1. **Open a PowerShell window.** On Windows: press the Start button, type
   `powershell`, press Enter. (A terminal inside VS Code works the same way.)

2. **Go to the library's *main* folder — the top one, not this one.** Even
   though the file you're reading lives in `mcp-server\tests`, the tests are
   run from the **top folder of the library** — the one that contains
   `README.md`, `WELCOME.md`, `prompts`, `templates`, and `mcp-server` side by
   side. Type `cd` followed by a space and that folder's location, then press
   Enter. For example — **your path will be different**:

   ```powershell
   cd C:\Users\you\Documents\open-template-scaffolds
   ```

   Not sure you're in the right place? Type `ls` and press Enter — you should
   see `templates` and `mcp-server` in the list. If you see `test_library.py`
   or `fixtures` instead, you're too deep: you're standing in the tests folder
   itself — type `cd ..\..` and press Enter to climb back up two levels.

3. **Run the tests** — copy-paste this line and press Enter:

   ```powershell
   python -m unittest discover -s mcp-server/tests
   ```

   If instead of test results you get an error saying the folder *"does not
   exist"* or *"is not importable"*, you're standing in the wrong folder —
   go back to step 2.

## What you should see

After a second or two, something very close to this:

```text
...........................s
----------------------------------------------------------------------
Ran 28 tests in 0.2s

OK
```

How to read it:

- **Every dot is one test that passed.**
- **`OK` at the bottom is the verdict** — that's the word you're looking for.
- The lone **`s`** means one test **skipped itself on purpose**. That's normal,
  not a problem: it's the test that opens a real Access database, and it stays
  quiet unless you deliberately set one up for it (see the last section).
  The last line may read `OK (skipped=1)` — still a pass.

## What a failure looks like

If something is broken you'll see **`F`** letters mixed in with the dots, a
block of text naming exactly which test failed and why, and **`FAILED`**
instead of `OK` at the bottom. The test names read as plain sentences —
`test_live_corpus_all_pass`, `test_no_match_returns_empty_list` — so the
failure block usually tells you the story on its own.

What to do about it:

- **If you changed something just before this**, the failure is almost
  certainly pointing at that change. Read the named test — its name says what
  it checks — and revisit your edit.
- **If a fresh, unmodified copy fails**, that's worth reporting: open an issue
  on the library's GitHub page and paste in the failure text.

## Optional — the real-database test

One test is different from all the others: it opens an actual Access database
file and confirms the compatibility checker reads it correctly. It's switched
off by default because it needs things not every computer has (Windows, the
Access database engine, and a specially prepared sample database). Unless
you're working on the compatibility checker itself, skipping it is the normal
state — that's the `s` in the output above.

If you *are* working on it, here is the setup, step by step:

1. **Have a Northwind sample database ready.** Any copy works — the standard
   one that Microsoft's *Northwind Developer Edition* template creates is
   fine, exactly as it comes. (The test checks that the compatibility tool
   *tells the truth* about what your database contains — it doesn't require
   the database to contain anything special.)

2. **Tell the test where that file is.** In your PowerShell window, type this
   line — with the location of *your* file between the quotes — and press
   Enter:

   ```powershell
   $env:OTS_TEST_ACCDB = 'C:\the\folder\where\it\lives\Northwind.accdb'
   ```

   This is just a sticky note for the test to read: it lasts until you close
   the window, and it doesn't install or change anything.

3. **Check whether your Office is 32-bit or 64-bit.** Open Access, then:
   **File → Account → About Access**. The very top line of the window that
   opens ends with either *32-bit* or *64-bit*.

4. **Run the tests again from the same window.** If your Office said
   *64-bit*, run them the normal way — you're done. If it said *32-bit*
   (very common for Access), this one test needs a 32-bit Python to match —
   Windows Python installs come with a launcher that can pick it, like this:

   ```powershell
   py -3.12-32 mcp-server/tests/test_compat.py
   ```

   (If that says the version isn't found, you'd need to install a 32-bit
   Python from python.org — only worth doing if you're really digging into
   the compatibility checker.)

The match-your-Office rule applies **only to this one optional test**.
Everything else runs on any Python.

## For contributors — how the tests are organized

- `test_library.py` — the parsing helpers and every formatting rule the
  `validate` tool enforces.
- `test_tools.py` — the six tools the server offers.
- `test_compat.py` — the compatibility checker, including that it fails
  *politely* (with a reason, never a crash) on machines that can't run it.
- `fixtures/broken/` — a set of deliberately broken templates, **one per
  validation rule**, each named for the rule it violates. If you add a rule to
  the validator, add a matching broken file here and list it in
  `EXPECTED_RULE` in `test_library.py` — a test fails if the two lists drift
  apart, so you can't forget.

The tests use only what comes with Python (its built-in `unittest` module) —
there is nothing extra to install, ever.
