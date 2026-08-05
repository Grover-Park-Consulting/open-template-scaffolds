#!/usr/bin/env python3
"""Pre-commit guard: every staged Python file must parse.

Editors that save automatically write a stray paste to disk with no prompt and
no visible mark, so a file can be corrupted without anyone seeing it happen.
This blocks the commit instead of letting it reach the repository.

Checks the *staged* content rather than the file on disk, because the staged
content is what the commit would actually contain.
"""

import ast
import subprocess
import sys


def _git(args):
    return subprocess.run(
        ["git"] + args, capture_output=True, check=True
    ).stdout


def staged_python_files():
    # -z keeps paths intact when they contain spaces or non-ASCII characters.
    # ACM = added / copied / modified; a deleted file has nothing left to parse.
    out = _git(["diff", "--cached", "--name-only", "--diff-filter=ACM", "-z"])
    names = [n.decode("utf-8", "surrogateescape") for n in out.split(b"\0") if n]
    return [n for n in names if n.endswith(".py")]


def check(path):
    """Return (line, message) if the staged file is unparseable, else None."""
    raw = _git(["show", f":{path}"])
    try:
        source = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        return (None, f"not valid UTF-8 ({exc})")
    try:
        ast.parse(source, filename=path)
    except SyntaxError as exc:
        return (exc.lineno, exc.msg)
    return None


def main():
    failures = []
    for path in staged_python_files():
        result = check(path)
        if result:
            failures.append((path, result[0], result[1]))

    if not failures:
        return 0

    out = sys.stderr
    print("", file=out)
    print("COMMIT BLOCKED - staged Python does not parse:", file=out)
    print("", file=out)
    for path, line, msg in failures:
        where = f"{path}:{line}" if line else path
        print(f"    {where}: {msg}", file=out)
    print("", file=out)
    print("Most often this is text pasted into the file by accident.", file=out)
    print("Open it at the line above and look at what is there.", file=out)
    print("", file=out)
    print("To commit anyway, without this check:", file=out)
    print("    git commit --no-verify", file=out)
    print("", file=out)
    return 1


if __name__ == "__main__":
    sys.exit(main())
