#!/usr/bin/env python3
"""Print the size of every part the server can deliver, and how close each is to the limit.

    python mcp-server/part_sizes.py            # parts with less than 5,000 characters to spare
    python mcp-server/part_sizes.py --all      # every part
    python mcp-server/part_sizes.py <template> # one template's parts

A report, not a check: validate (SZ1) is what refuses a part over the limit. This shows
which parts are getting close, so text can be moved into a step or a fact narrowed to the
children a step uses before the limit is reached, not after.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from library import RESPONSE_LIMIT, compose, iter_templates, plan, serialized_size  # noqa: E402

MARGIN = 5000


def sizes(only: str = "") -> list[tuple[int, str, str, str]]:
    rows = []
    for path, front, body in iter_templates():
        tid = str(front.get("template", ""))
        if only and tid != only:
            continue
        for p in plan(path, front, body):
            routes = ["handover"] if p["route"] == "handover" else ["build"] if p["route"] == "build" \
                else ["build", "handover"]
            biggest = max((serialized_size(compose(tid, p["part"], r, True)), r) for r in routes)
            rows.append((biggest[0], tid, p["part"], biggest[1]))
    return sorted(rows, reverse=True)


def main(argv: list[str]) -> int:
    show_all = "--all" in argv
    only = next((a for a in argv if not a.startswith("--")), "")
    rows = sizes(only)
    shown = [r for r in rows if show_all or only or RESPONSE_LIMIT - r[0] < MARGIN]
    for size, tid, part, route in shown:
        print(f"{size:6}  {RESPONSE_LIMIT - size:6} to spare  {tid} / {part} ({route})")
    if not shown:
        print(f"Every part has at least {MARGIN:,} characters to spare (limit {RESPONSE_LIMIT:,}).")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
