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

An application that stores a **relative reference to an external file asset** (a photo folder, an
attachments folder) only works if that folder actually exists and is re-created wherever the app is
copied. `Startup()` → `EnsureAppFolders()` is what makes a relative asset reference reliable across
machines. The build-time mechanics — how the `AutoExec` macro is generated — live in
`templates/_materialization.md`; this file states the *convention* a generated application follows.
