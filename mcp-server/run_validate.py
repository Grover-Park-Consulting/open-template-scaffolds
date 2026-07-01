#!/usr/bin/env python3
"""CI gate / local check: is every template in the library well-formed?

Runs exactly the format-only checks behind the MCP `validate` tool (one shared
code path: library.validate_library) and exits non-zero when any template
fails, so a GitHub Actions workflow — or an adopter at a shell — can gate on
it. Needs only pyyaml, never the MCP SDK, and never opens a host database.

Usage:  python mcp-server/run_validate.py
"""

import sys

from library import validate_library


def main() -> int:
    report = validate_library()
    for result in report["results"]:
        print(("PASS" if result["ok"] else "FAIL") + f"  {result['template']}")
        for err in result["errors"]:
            print(f"      {err}")
    verdict = "all well-formed." if report["ok"] else "failures above."
    print(f"\n{report['checked']} template(s) checked; {verdict}")
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
