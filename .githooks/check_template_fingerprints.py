#!/usr/bin/env python3
"""Pre-commit guard: a template's text must match the fingerprint taken when its status was set.

A template edited after its status was determined may no longer deserve that
status, and nobody is reliably going to remember to look. This stops the commit
and says what to ask. The same check runs in `validate`, and in CI as the backstop.

Reads the files on disk, not the staged content, so a template changed and not
yet staged is caught too. Where the check cannot run (no pyyaml, a template that
does not parse) it warns and lets the commit through, because CI repeats it.
"""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "mcp-server"))


def main():
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    try:
        import library
        failures = []
        for path, front, _ in library.iter_templates():
            if str(front.get("type", "")).lower() == "spec":
                continue
            steps = library.read_steps(path, front)
            failures += library.fingerprint_errors(str(front.get("template")), path, front, steps)
    except Exception as exc:  # noqa: BLE001 - a guard that crashes must not block every commit
        print(f"pre-commit: template fingerprint check skipped ({exc}); CI will run it.", file=sys.stderr)
        return 0

    if not failures:
        return 0

    out = sys.stderr
    print("", file=out)
    print("COMMIT BLOCKED - a template changed after its status was set:", file=out)
    print("", file=out)
    for message in failures:
        print(f"    {message}", file=out)
    print("", file=out)
    print("To commit anyway, without this check:", file=out)
    print("    git commit --no-verify", file=out)
    print("", file=out)
    return 1


if __name__ == "__main__":
    sys.exit(main())
