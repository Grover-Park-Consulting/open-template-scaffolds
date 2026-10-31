#!/usr/bin/env python3
"""Record what happened to a template's text, after the pre-commit check stopped on it.

    python mcp-server/template_status.py <template> typo
    python mcp-server/template_status.py <template> fix|change [--status draft|review|stable]

typo    The change did not alter meaning. Rewrites `status_fingerprint` and nothing else.
fix     A correction. Patch version bump, then a status decision.
change  Anything an adopter would act on. Minor version bump, then a status decision.

For fix and change, the command first prints a decision block and writes nothing. Run it
again with `--status` to apply: version, status, today's date in "Status last
determined", then the fingerprint. `--status stable` is refused unless the newest local
build record carries this text's fingerprint with no NOT PASSED result; add `--by-hand`
when you have compared the evidence yourself.
"""

import datetime
import re
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import library  # noqa: E402

ROOT = library.LIBRARY_ROOT
RECORDS = ROOT / "build-records"
STATUSES = ("draft", "review", "stable")
_DATE = re.compile(r"^(\*{0,2}Status last determined:\*{0,2}[ \t]*)\d{4}-\d{2}-\d{2}", re.M)
_RECORD_FP = re.compile(r"fingerprint\W{0,6}([0-9a-f]{12})", re.I)


def bump(version: str, mode: str) -> str:
    major, minor, patch = (int(x) for x in str(version).split("."))
    return f"{major}.{minor}.{patch + 1}" if mode == "fix" else f"{major}.{minor + 1}.0"


def set_front_line(text: str, key: str, value: str) -> str:
    """Replace `key: ...` inside the front matter; add it after `status:` when it is absent."""
    head, sep, rest = text.partition("\n---")
    line = re.compile(rf"^{key}:[^\r\n]*", re.M)
    if line.search(head):
        return line.sub(f"{key}: {value}", head, count=1) + sep + rest
    after = re.compile(r"^(status:[^\r\n]*)", re.M)
    return after.sub(lambda m: f"{m.group(1)}\n{key}: {value}", head, count=1) + sep + rest


def set_date(text: str, today: str) -> tuple[str, bool]:
    new, n = _DATE.subn(lambda m: m.group(1) + today, text, count=1)
    return new, n == 1


def evidence(slug: str) -> dict | None:
    """Facts from the newest local build record for a template, or None when there is none."""
    folder = RECORDS / slug
    files = sorted(folder.glob("*.md")) if folder.is_dir() else []
    if not files:
        return None
    text = files[-1].read_text(encoding="utf-8", errors="replace")
    found = _RECORD_FP.search(text)
    # A tally such as "Result: PASSED: 62" is not a result; only the bare form counts.
    return {"file": files[-1].name, "passed": len(re.findall(r"Result: PASSED(?!:)", text)),
            "failed": len(re.findall(r"Result: NOT PASSED(?!:)", text)),
            "fingerprint": found.group(1).lower() if found else None}


def matches(ev: dict | None, now: str) -> bool:
    return bool(ev and ev["fingerprint"] == now and ev["failed"] == 0 and ev["passed"] > 0)


def propose(current: str, ev: dict | None, now: str) -> tuple[str, str]:
    if ev is None:
        return current, "no build record for this template on this machine"
    if ev["failed"]:
        return "draft", "the newest record has a NOT PASSED result (_maturity-criteria.md section 3)"
    if ev["fingerprint"] is None:
        return current, "the newest record predates fingerprints; compare its checklist to the template's by hand"
    if ev["fingerprint"] != now:
        return ("draft" if current == "draft" else "review"), "the newest record tested other text: no current evidence"
    return ("stable" if ev["passed"] else current), "the newest record tested this text and every result passed"


def _git(*args: str) -> str:
    try:
        run = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True, timeout=30)
        return run.stdout.strip() if run.returncode == 0 else ""
    except (OSError, subprocess.SubprocessError):
        return ""


def changed_since_status(path: Path) -> list[str]:
    base = _git("log", "-1", "--format=%H", "-G^status_fingerprint:", "--", str(path))
    if not base:
        return ["(git history unavailable)"]
    folder = library.steps_folder(path)
    out = _git("diff", "--numstat", base, "--", str(path), str(folder))
    rows = []
    for line in out.splitlines():
        added, removed, name = line.split("\t", 2)
        rows.append(f"{name}  +{added} -{removed}")
    return rows or ["(none)"]


def decision_block(slug: str, path: Path, front: dict, mode: str, new_version: str, now: str) -> str:
    ev = evidence(slug)
    status, reason = propose(str(front.get("status")), ev, now)
    if ev is None:
        record = "none"
    else:
        fp = ev["fingerprint"]
        tested = "no fingerprint in the record" if fp is None else ("matches" if fp == now else f"differs ({fp})")
        record = f"{ev['file']}: {ev['passed']} passed, {ev['failed']} not passed, fingerprint {tested}"
    lines = [f"DECISION  {slug}  v{front.get('version')} -> v{new_version}  status now: {front.get('status')}",
             "Changed since the status was set:"] + [f"  {r}" for r in changed_since_status(path)] + [
             f"Newest local build record: {record}",
             f"Proposed: {status}  ({reason})",
             "Choices: draft | review | stable" + ("" if matches(ev, now) else "  (stable needs --by-hand)"),
             f"Apply: python mcp-server/template_status.py {slug} {mode} --status {status}"]
    return "\n".join(lines)


def main(argv: list[str]) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    args = [a for a in argv if not a.startswith("--")]
    flags = {a.split("=")[0]: (a.split("=", 1)[1] if "=" in a else "") for a in argv if a.startswith("--")}
    status = flags.get("--status") or (argv[argv.index("--status") + 1] if "--status" in argv else "")
    if status:
        args = [a for a in args if a != status]
    if len(args) != 2 or args[1] not in ("typo", "fix", "change") or (status and status not in STATUSES):
        print(__doc__, file=sys.stderr)
        return 2
    slug, mode = args[0].lower(), args[1]
    for path, front, _ in library.iter_templates():
        if str(front.get("template", "")).lower() == slug:
            break
    else:
        print(f"No template with id '{slug}'.", file=sys.stderr)
        return 2

    def fingerprint() -> str:
        return library.template_fingerprint(path, library.read_steps(path, front))

    def write(transform):
        text = path.read_bytes().decode("utf-8")
        path.write_bytes(transform(text).encode("utf-8"))

    if mode == "typo":
        now = fingerprint()
        write(lambda t: set_front_line(t, "status_fingerprint", now))
        print(f"{slug}: fingerprint {now} written.")
        return 0

    new_version = bump(front.get("version"), mode)
    def fingerprint_after_bump() -> str:
        """The fingerprint the text will have once the version is bumped; the file is restored."""
        original = path.read_bytes()
        try:
            write(lambda t: set_front_line(t, "version", new_version))
            return fingerprint()
        finally:
            path.write_bytes(original)

    if not status:
        print(decision_block(slug, path, front, mode, new_version, fingerprint_after_bump()))
        return 0

    if status == "stable" and "--by-hand" not in flags:
        if not matches(evidence(slug), fingerprint_after_bump()):
            print(f"{slug}: stable refused; no local record carries this text's fingerprint with every "
                  f"result passed. Add --by-hand once you have compared the evidence.", file=sys.stderr)
            return 1
    today = datetime.date.today().isoformat()
    dated = []

    def apply(text: str) -> str:
        text = set_front_line(text, "version", new_version)
        text = set_front_line(text, "status", status)
        text, found = set_date(text, today)
        dated.append(found)
        return text

    write(apply)
    now = fingerprint()
    write(lambda t: set_front_line(t, "status_fingerprint", now))
    print(f"{slug}: v{new_version}, status {status}, fingerprint {now}"
          + (f", dated {today}." if dated[0] else ". No 'Status last determined' line to date."))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
