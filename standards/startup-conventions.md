# Startup Conventions — GPC Default Standards Layer

> **GPC default; fork-and-replace.** Governs how a generated Access application initializes when it
> opens — the `AutoExec` → `Startup()` convention and the single extensible open-time slot. A forked
> practice that uses a different startup mechanism swaps this file for its own.

---

## 1. Every generated application opens through one entry point

A generated Access application gets an **`AutoExec` macro** whose only action is `RunCode Startup()`
(followed by `StopMacro`). Access runs `AutoExec` automatically on open, so `Startup()` is the one
place open-time initialization happens — no logic is scattered across form Load events or left to run
by chance.

- **`Startup()` is a `Public Function`**, not a Sub — the `RunCode` macro action can only call a
  Function. It returns `Boolean` (True on success) and carries its own `errHandler` block from
  `error-handling.md`.
- It is the **single open-time initialization slot.** Typical things it does, and things to consider
  adding as the application grows:
  - ensure the application's working folders exist (`EnsureAppFolders()`, §2);
  - open the application's startup form — a switchboard, menu, or home form — e.g.
    `DoCmd.OpenForm "AppStartupForm"`;
  - later: relink checks, version checks, environment probes.

  Whatever it does, it happens here, in one predictable place.

## 2. `EnsureAppFolders()` — the idempotent folder-ensure slot

Today `Startup()`'s core job is making sure the application's working folders exist, delegated to
`EnsureAppFolders()`. This routine is **idempotent**: the *same* routine both creates a folder on the
first build and verifies it on every subsequent open. It is the extension point for open-time setup
that must be safe to run repeatedly — add each new working folder here as new folder settings are
introduced.

The idempotence principle generalizes beyond folders: anything `Startup()` does should be safe to run
on every open, because it runs on every open.

## 3. Why this is a standard, not just a build detail

An application that stores a **reference to an external file asset** (a photo folder, an attachments
folder) only works if that folder actually exists on every machine the application runs on.
`Startup()` → `EnsureAppFolders()` is what makes that reliable. The build-time mechanics — how the
`AutoExec` macro is generated — live in `templates/_materialization.md`; this file states the
*convention* a generated application follows.

## 4. Two kinds of folder — ensure them differently

Which folder is which matters most in a **split database**: the normal shape for a multi-user Access
application, where the tables live in one file (the **back end**, on a shared network drive) and each
person runs their own copy of a second file holding the forms and code (the **front end**). The front
end is the file that gets copied to each machine — so anything `EnsureAppFolders()` creates "beside
the application" gets created once per person.

**Shared-content folders** hold files that a row in a shared table points at — official photos,
scanned documents, attachments. There must be **one** such folder for everybody, named by an absolute
path (typically a network location beside the back end) held in a settings table, not a relative path.
`EnsureAppFolders()` **verifies that it exists and is reachable**, and reports plainly if it isn't. It
must not quietly create a local one instead.

> **Why "verify, don't create" is the rule here.** Creating a local folder *succeeds*, which is the
> problem. Ann adds a photo; the picker copies the file into Ann's local folder and stores the file
> name in the shared table. Bob opens the same record, his front end looks in *Bob's* folder, finds
> nothing, and shows a blank. The row is correct, both front ends did exactly what they were told, and
> the file is invisible to everyone but Ann. Nothing raises an error — which is what makes this
> expensive to find.

**Per-user working folders** hold files that belong to one person and one session — exports, temporary
output, a local log. These are created per front end, on each machine, exactly as §2 describes. A
relative path is right for these.

On a **single-file database** — one .accdb holding everything, an acceptable choice for one user —
there is only one file and one person, so the two kinds collapse into one and a relative path serves
both.
