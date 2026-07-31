# Startup Conventions — OTS Default Standards Layer

> **OTS default; fork-and-replace.** Governs how a generated Access application initializes when it
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
  adding as the application grows, **in this order**:
  - confirm the application can reach its data, and reconnect it if it has moved (§5) — this comes
    **first**, because everything after it assumes the data is reachable;
  - ensure the application's working folders exist (`EnsureAppFolders()`, §2);
  - open the application's startup form — a switchboard, menu, or home form — e.g.
    `DoCmd.OpenForm "AppStartupForm"`;
  - later: version checks, environment probes.

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

## 5. When the data has moved — the application reconnects itself

In a split database the front end does not contain the tables; it contains **links** to them. A link
stores the back end's **path, recorded at the moment the link was made**, and nothing keeps that path
up to date afterwards. Move the back end to a new server, rename the share, deploy the application to
a second site, or hand a front end to someone whose drive letters differ, and every link in it points
at a file that is no longer there.

**What that looks like to the person using it:** the application opens normally, and then fails on
the first screen that shows data. It reads as "the application is broken," which is why it generates
a support call rather than a shrug.

A generated application handles this itself, in `Startup()`, **before anything else runs**:

1. **Check first.** Find where the links currently point and confirm that file can actually be
   opened. This happens before folders are ensured and before any form opens, because both of those
   assume the data is reachable — the folder paths are themselves settings stored in the back end.
2. **Reconnect from memory.** If the links are stale, try the path the front end remembers from the
   last time it worked, and reconnect silently if that path is good. Most moves are handled here,
   with nobody seeing anything.
3. **Otherwise ask, and name the path.** Tell the person plainly that the data file cannot be found,
   **including the path that was tried** — the single most useful fact for whoever has to sort it
   out — and offer to let them find it.
4. **Check what they chose before connecting to it.** A file picker will happily return last year's
   backup or a copy sitting on one machine, and the application would then run perfectly against the
   wrong data. Confirm the chosen file really is this application's back end — open it and look for a
   table it must contain — before repointing anything.
5. **Remember only what worked.** Store the new path *after* a reconnection succeeds, never before.
   A remembered bad path is worse than no remembered path, because it gets tried first on every
   future open.

If the person cancels, or the file they choose is the wrong one, **no form opens.** An application
that runs on top of broken links produces errors that look like data problems.

### Where the remembered path is kept

**In the front end, in a table that is not a link** — by default a configuration table named per
`naming-conventions.md` (`USys` prefix), holding one setting row.

> **Why not in the settings table with everything else.** A settings table in the *back end* is right
> for settings everybody shares — change one row and it changes for everyone at once. It is exactly
> wrong for this one: **a setting stored in the back end cannot tell you where the back end is.**
> When the back end is unreachable, that is the one table you cannot read.

Two other places are legitimate, ranked below it rather than ruled out:

| Where | What it buys | What it costs |
|---|---|---|
| **A local table** (the default) | Opens even when every link is broken, because it is local. Changing it is a data edit, not a code change. | Each front-end copy remembers its own answer — which is usually correct, since two people can reach the same back end by different routes. |
| **A text file beside the front end** | Someone can fix it in Notepad **without opening Access at all** — precisely the situation you are in when the application won't start. | File reading in the startup path, and a file that can be deleted or edited into nonsense with nothing to catch it. |
| **A constant in code** | Simplest to write and read; nothing to create, nothing to seed. | Changing where the back end lives means editing code and giving everybody a new front end. |

**On a single-file database this is a no-op.** There are no links, so the check finds nothing to
check and the application carries straight on. Nothing has to be removed or switched off.

**One constraint this places on error handling.** Anything running at this point in startup may be
running while the data is unreachable — that is the whole situation. So the error logger used here
**must not write to the back end**, or it fails inside the handler that was reporting the original
failure and the person sees neither message. A log file beside the front end, or a message box, both
survive it.

The runnable code — the reconnect sequence, the file-picker prompt, and the check that a chosen file
is the right one — is `templates/startup/app-startup-scaffold.md`.
