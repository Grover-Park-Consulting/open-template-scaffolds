"""Keep the local lesson index current.

Who reads this: the AI assistant maintaining the library, or a contributor.

The text files in `lessons/` are the source of truth for the library's lessons.
A lessons database (an accdb) is built from them and holds the developer's own
approvals and lessons. The server reads only the index file this module writes,
never the database. A PowerShell script (`Sync-LessonsDb.ps1`) does the database
work through the database engine's DAO objects, so no Access window starts.

`ensure_index` never raises for an environment problem. It returns the index to
use, or None, with a status and one line saying why.
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import tempfile
from pathlib import Path

from library import split_front_matter

HERE = Path(__file__).resolve().parent
DEFAULT_LESSONS = HERE.parent / "lessons"
SCRIPT = HERE / "Sync-LessonsDb.ps1"
EXIT_LOCKED, EXIT_NO_ENGINE, EXIT_NO_POWERSHELL = 10, 20, 21
TIMEOUT_SECONDS = 90


def _app_folder() -> Path:
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "OpenTemplateScaffolds"


def settings_path() -> Path:
    return _app_folder() / "settings.json"


def default_store() -> Path:
    """The folder offered to the developer as the default; never used until they choose it."""
    return _app_folder() / "lessons"


def chosen_store() -> Path | None:
    """The folder the developer chose for the lessons database, or None if they have not chosen.

    The OTS_LESSONS_STORE environment variable overrides the setting (used by the tests)."""
    env = os.environ.get("OTS_LESSONS_STORE")
    if env:
        return Path(env)
    try:
        value = json.loads(settings_path().read_text(encoding="utf-8")).get("lessons_store")
    except (OSError, ValueError, AttributeError):
        return None
    return Path(value) if value else None


_SYNCED_FOLDERS = ("onedrive", "dropbox", "google drive", "icloud", "box sync")


def set_lessons_store(path: str = "") -> dict:
    """Record the folder for the lessons database, creating it if needed.

    An empty path or "default" chooses the default folder. Raises ValueError when the folder
    cannot be used. Returns {"store", "created", "warnings"}."""
    wanted = default_store() if path.strip() in ("", "default") else Path(path.strip()).expanduser()
    if not wanted.is_absolute():
        raise ValueError(f"'{path}' is not a full folder path (for example C:\\Lessons).")
    created = not wanted.exists()
    try:
        wanted.mkdir(parents=True, exist_ok=True)
        probe = wanted / ".ots-write-test"
        probe.write_text("x", encoding="utf-8")
        probe.unlink()
    except OSError as exc:
        raise ValueError(f"The folder '{wanted}' cannot be created or written to: {exc}") from exc
    warnings = []
    if any(part.lower().startswith(_SYNCED_FOLDERS) for part in wanted.parts):
        warnings.append("This folder looks like a cloud-synced folder. A database in a synced folder can be "
                        "locked or duplicated by the sync client. Use a local folder, or a network share.")
    target = settings_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(".tmp")
    try:
        current = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        current = {}
    current["lessons_store"] = str(wanted)
    tmp.write_text(json.dumps(current, indent=1), encoding="utf-8")
    os.replace(tmp, target)
    return {"store": str(wanted), "created": created, "warnings": warnings}


def _read_mechanisms(lessons_dir: Path) -> list[dict]:
    rows = []
    for line in (lessons_dir / "_mechanisms.md").read_text(encoding="utf-8").splitlines():
        if not line.startswith("|") or line.startswith("|---") or line.startswith("| Canonical"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) == 4:
            rows.append({"name": cells[0], "description": cells[1], "aliases": cells[2], "always": cells[3].lower() == "yes"})
    if not rows:
        raise ValueError("no mechanism rows found in _mechanisms.md")
    return rows


def _read_lookups(lessons_dir: Path) -> dict[str, list[dict]]:
    """Table name -> [{"id", "value"}], from the sections of `_lookups.md`."""
    tables: dict[str, list[dict]] = {}
    current = None
    for line in (lessons_dir / "_lookups.md").read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            current = line[3:].strip()
            tables[current] = []
        elif current and line.startswith("|") and not line.startswith("|---"):
            cells = [c.strip() for c in line.strip().strip("|").split("|")]
            if len(cells) == 2 and cells[0].isdigit():
                tables[current].append({"id": int(cells[0]), "value": cells[1]})
    if not tables or not all(tables.values()):
        raise ValueError("a lookup in _lookups.md has no rows")
    return tables


def read_text_lessons(lessons_dir: Path) -> tuple[list[dict], list[dict], str]:
    """Return (mechanisms, lessons, hash of every source file)."""
    digest = hashlib.sha256()
    mech_path = lessons_dir / "_mechanisms.md"
    digest.update(mech_path.read_bytes())
    digest.update((lessons_dir / "_lookups.md").read_bytes())
    mechanisms = _read_mechanisms(lessons_dir)
    lessons = []
    for path in sorted(lessons_dir.glob("[!_]*.md")):
        digest.update(path.read_bytes())
        front, _ = split_front_matter(path.read_text(encoding="utf-8"))
        lessons.append({
            "slug": front["id"], "mechanism": front["mechanism"], "kind": front["kind"], "confidence": front["confidence"],
            "status": front.get("status", "active"), "lesson": front["lesson"], "applies_when": front.get("applies_when"),
            "triggers_when": front.get("triggers_when"), "evidence": front.get("evidence"),
            "environment": "\n".join(front.get("environment") or []), "sources": "\n".join(front.get("sources") or []),
            "confirmations": front.get("confirmations") or [], "contradictions": front.get("contradictions") or [],
        })
    return mechanisms, lessons, digest.hexdigest()


def _index(mechanisms: list[dict], lessons: list[dict], text_hash: str, source: str, accdb_mtime: float | None) -> dict:
    mech = {m["name"]: {"always": bool(m["always"]), "description": m.get("description") or "", "aliases": [a.strip() for a in (m.get("aliases") or "").split(",") if a.strip()]} for m in mechanisms}
    out = []
    for l in lessons:
        if l.get("status") == "inactive":
            continue
        out.append({
            "id": l["slug"], "mechanism": l["mechanism"], "kind": l["kind"], "confidence": l["confidence"],
            "triggers_when": l.get("triggers_when") or None, "lesson": l["lesson"], "applies_when": l.get("applies_when"),
            "environment": [e for e in (l.get("environment") or "").split("\n") if e], "origin": (l.get("origin") or "Library").lower(),
        })
    return {"text_hash": text_hash, "source": source, "accdb_mtime": accdb_mtime, "mechanisms": mech, "lessons": out}


def _powershells() -> list[str]:
    windir = os.environ.get("SystemRoot", r"C:\Windows")
    tail = r"WindowsPowerShell\v1.0\powershell.exe"
    found = []
    for base in ("Sysnative", "System32", "SysWOW64"):
        p = os.path.join(windir, base, tail)
        if os.path.exists(p) and p not in found:
            found.append(p)
    return found


def _run_sync(db: Path, import_file: Path, export_file: Path) -> tuple[int, str]:
    """Try each PowerShell bitness until one has the database engine. Returns (exit code, message)."""
    shells = _powershells()
    if not shells:
        return EXIT_NO_POWERSHELL, "PowerShell was not found"
    last = (EXIT_NO_ENGINE, "no PowerShell found")
    for ps in shells:
        try:
            proc = subprocess.run(
                [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT),
                 "-DatabasePath", str(db), "-Import", str(import_file), "-Export", str(export_file)],
                capture_output=True, text=True, timeout=TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            return 30, f"sync timed out after {TIMEOUT_SECONDS} seconds"
        except OSError as exc:
            last = (30, str(exc))
            continue
        last = (proc.returncode, (proc.stderr or "").strip())
        if proc.returncode != EXIT_NO_ENGINE:
            return last
    return last


def _write_index(path: Path, index: dict) -> None:
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(index, indent=1), encoding="utf-8")
    os.replace(tmp, path)


def ensure_index(lessons_dir: Path = DEFAULT_LESSONS, store_dir: Path | None = None) -> dict:
    """Return {"status", "message", "index"}. Status is one of
    current, synced, kept-old (database in use), from-text (database unavailable), not-chosen
    (no folder chosen yet; the library's own lessons are used), unavailable."""
    store = store_dir or chosen_store()
    try:
        mechanisms, lessons, text_hash = read_text_lessons(lessons_dir)
    except Exception as exc:  # the library's own text is unreadable
        return {"status": "unavailable", "message": f"Lessons unavailable: {exc}", "index": None}
    if store is None:
        return {"status": "not-chosen", "message": "The developer has not chosen where the lessons database lives.",
                "index": _index(mechanisms, lessons, text_hash, "text", None)}

    db = store / "lessons.accdb"
    index_path = store / "lessons.index.json"
    old = None
    try:
        old = json.loads(index_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        pass
    mtime = db.stat().st_mtime if db.exists() else None
    if old and old.get("source") == "accdb" and old.get("text_hash") == text_hash and mtime is not None and old.get("accdb_mtime") == mtime:
        return {"status": "current", "message": "", "index": old}

    fallback = _index(mechanisms, lessons, text_hash, "text", None)
    try:
        store.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory() as tmp:
            imp, exp = Path(tmp) / "import.json", Path(tmp) / "export.json"
            imp.write_text(json.dumps({"lookups": _read_lookups(lessons_dir), "mechanisms": mechanisms, "lessons": lessons}), encoding="utf-8")
            code, msg = _run_sync(db, imp, exp)
            if code == 0:
                data = json.loads(exp.read_text(encoding="utf-8"))
                index = _index(data["mechanisms"], data["lessons"], text_hash, "accdb", db.stat().st_mtime)
                _write_index(index_path, index)
                return {"status": "synced", "message": "", "index": index}
    except Exception as exc:
        code, msg = 30, str(exc)

    if code == EXIT_LOCKED:
        note = "Close the lessons database to pick up new lessons."
        if old:
            return {"status": "kept-old", "message": note, "index": old}
        return {"status": "from-text", "message": note + " Using the library's own lessons only.", "index": fallback}
    if code == EXIT_NO_POWERSHELL:
        reason = "PowerShell was not found"
    elif code == EXIT_NO_ENGINE:
        reason = "the database engine is not available"
    else:
        reason = msg or "the sync failed"
    if old:
        return {"status": "kept-old", "message": f"Lessons database not updated: {reason}.", "index": old}
    return {"status": "from-text", "message": f"Lessons database not used: {reason}. Using the library's own lessons only.", "index": fallback}


LESSONS_RULE = (
    "Read every lesson below before the first design question. Do not search for lessons yourself. "
    "Each lesson's `action` says what to do. apply: do it and do not alert the developer. state: tell "
    "the developer before the design, as a consideration. ask: when the developer chooses, or has "
    "already chosen, what `triggers_when` names, tell them before they commit what they would have to "
    "reconcile, and let them decide. Let `confidence` set how firmly you word it. A lesson with "
    "`environment` holds only where those facts hold: check them first, and state the lesson as a "
    "consideration if you cannot. A lesson never overrides the template or the standards; if it "
    "contradicts either, tell the developer. If `note` is not empty, give it to the developer in one line. "
    "If `status` is `not-chosen`, ask the developer once, through the selection panel, where the lessons "
    "database should live: `suggested_store`, a folder they name, or not now. Call `set_lessons_store` with "
    "their answer and pass on any `warnings`; on not now, carry on. "
    "Write `record_line` in the build record, adding a clause for each lesson that changed what you "
    "built or asked. The lessons are a consultation, never a condition for building.")


def _action(lesson: dict) -> str:
    if lesson["kind"] == "consequence":
        return "ask"
    return "apply" if lesson["confidence"] == "Observed" else "state"


def select(index: dict, mechanisms: list[str]) -> tuple[list[dict], list[str], int]:
    """The lessons for `mechanisms` plus every always-consulted mechanism.

    Returns (lessons in delivery order, names that match no mechanism, Potential lessons held back).
    Names match a canonical name or an alias, ignoring case.
    """
    names = index.get("mechanisms", {})
    lookup = {n.lower(): n for n in names}
    for n, row in names.items():
        for alias in row.get("aliases", []):
            lookup.setdefault(alias.lower(), n)
    wanted = {n for n, row in names.items() if row.get("always")}
    unknown = []
    for m in mechanisms or []:
        hit = lookup.get(str(m).strip().lower())
        if hit:
            wanted.add(hit)
        else:
            unknown.append(str(m))
    out, held = [], 0
    for l in sorted(index.get("lessons", []), key=lambda x: (x["mechanism"], x["id"])):
        if l["mechanism"] not in wanted:
            continue
        if l["confidence"] == "Potential":
            held += 1
            continue
        out.append({"id": l["id"], "mechanism": l["mechanism"], "kind": l["kind"], "confidence": l["confidence"],
                    "action": _action(l), "lesson": l["lesson"], "applies_when": l.get("applies_when"),
                    "triggers_when": l.get("triggers_when"), "environment": l.get("environment") or [],
                    "origin": l.get("origin", "library")})
    return out, unknown, held


def deliver(mechanisms: list[str], lessons_dir: Path = DEFAULT_LESSONS, store_dir: Path | None = None) -> dict:
    """The answer for a `lessons` part or a `get_lessons` call. Never raises for an environment problem."""
    got = ensure_index(lessons_dir, store_dir)
    if got["index"] is None:
        reason = got["message"] or "the lesson store could not be read"
        return {"status": got["status"], "note": reason, "lessons": [], "mechanisms_unknown": [], "held_back": 0,
                "rule": LESSONS_RULE, "record_line": f"Lessons consulted: none ({reason})"}
    lessons, unknown, held = select(got["index"], mechanisms)
    ids = ", ".join(l["id"] for l in lessons)
    out = {"status": got["status"], "note": got["message"], "lessons": lessons, "mechanisms_unknown": unknown,
           "held_back": held, "rule": LESSONS_RULE, "record_line": f"Lessons consulted: {len(lessons)} ({ids})"}
    if unknown or not mechanisms:
        out["mechanisms_available"] = [{"name": n, "description": row.get("description", "")}
                                       for n, row in got["index"].get("mechanisms", {}).items()]
    if got["status"] == "not-chosen":
        out["suggested_store"] = str(default_store())
    return out
