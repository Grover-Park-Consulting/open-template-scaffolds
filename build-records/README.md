# Build records

**Who reads this:** your AI assistant, before and after every build; anyone else, to see what past
builds ran into.

This folder holds every build record ever produced against this library, one subfolder per template
slug (e.g. `officiating-assignment-outcome-first/`). It ships empty. Your AI assistant creates a
template's subfolder the first time it builds that template, and adds one dated file to it after
every build from then on — `YYYY-MM-DD-<short-description>.md`.

**One folder per template, never a new folder per build.** The value here is in what accumulates:
a build record from six months ago, for this template or a different one that shares a mechanism,
can save the next build from re-deriving something the hard way. See `CLAUDE.md`, "Build records
accumulate," for what your AI assistant is instructed to do with this folder, and
`templates/_materialization.md`, "The build record," for what each file contains.
